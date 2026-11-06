"""COM backend — operations that were previously unimplemented.

The upstream COM backend covered drawing I/O, entity creation/query and bulk
selection, but roughly a hundred operations declared on
:class:`~autocad_mcp.backends.base.AutoCADBackend` fell through to the base
class stub and returned ``"Not supported on this backend"``.  This module
implements them against the AutoCAD ActiveX (COM) API.

Split out from ``com_backend.py`` purely for file size; the mixin is mixed
into :class:`COMBackend` and relies on the helpers defined there
(``_validate_connection``, ``_safe_get_property``, ``_apply_properties`` ...).

Conventions
-----------
* Every method returns a :class:`CommandResult`.
* ``self._doc`` / ``self._msp`` are the live document and model space.
* Blocking COM calls stay synchronous (matching the rest of the backend);
  polling loops yield with ``asyncio.sleep`` so the event loop stays alive.
"""

from __future__ import annotations

import asyncio
import math
import os
import tempfile
import time
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

import structlog

from autocad_mcp.backends.base import CommandResult

log = structlog.get_logger()

# INSUNITS enumeration -> friendly name
_UNIT_BY_CODE = {
    0: "unitless", 1: "inches", 2: "feet", 3: "miles", 4: "millimeters",
    5: "centimeters", 6: "meters", 7: "kilometers", 8: "microinches",
    9: "mils", 10: "yards", 11: "angstroms", 12: "nanometers", 13: "microns",
    14: "decimeters", 15: "decameters", 16: "hectometers", 17: "gigameters",
}

_UNIT_CODE = {
    "unitless": 0, "none": 0, "unspecified": 0,
    "in": 1, "inch": 1, "inches": 1,
    "ft": 2, "foot": 2, "feet": 2,
    "mi": 3, "mile": 3, "miles": 3,
    "mm": 4, "millimeter": 4, "millimeters": 4,
    "cm": 5, "centimeter": 5, "centimeters": 5,
    "m": 6, "meter": 6, "meters": 6,
    "km": 7, "kilometer": 7, "kilometers": 7,
}

_LISP_DIR = Path(tempfile.gettempdir()) / "autocad_mcp_lisp"


class COMOperationsMixin:
    """COM implementations for the backend methods that were previously stubs."""

    # ==================================================================
    # Shared helpers
    # ==================================================================

    def _err(self) -> CommandResult | None:
        return self._validate_connection()

    @staticmethod
    def _dbl_array(values: list) -> Any:
        """Pack a flat list of floats into a COM VARIANT double array."""
        import pythoncom
        import win32com.client

        return win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(v) for v in values]
        )

    def _handle_object(self, handle: Any) -> Any:
        """Resolve a handle to a COM entity (direct lookup, not a scan)."""
        return self._doc.HandleToObject(str(handle))

    @staticmethod
    def _dyn(obj: Any) -> Any:
        """Re-wrap a COM object for late binding.

        AutoCAD's type library declares collection ``Item()`` as returning the
        base ``IAcadEntity``, so static (gencache) wrappers hide subtype
        properties such as ``StartPoint``.  Late binding resolves them.
        """
        try:
            import win32com.client.dynamic as _dynamic

            return _dynamic.Dispatch(obj)
        except Exception:
            return obj

    def _iter_msp(self):
        """Yield every entity in model space, late-bound.

        Collection ``Item()`` returns the base ``IAcadEntity`` interface, so
        subtype properties (``TextString``, ``Name``, ``StartPoint``) raise
        AttributeError unless the object is re-wrapped for late binding.
        """
        count = self._msp.Count
        for i in range(count):
            try:
                yield self._dyn(self._msp.Item(i))
            except Exception:
                continue

    def _selection_set(self, handles: list) -> Any:
        """Build a COM SelectionSet from handles. Caller must Delete() it.

        ``AddItems`` (plural) takes a SAFEARRAY of IDispatch, not a list — a
        plain Python list is rejected with "Invalid argument pSelSet".
        """
        import pythoncom
        import win32com.client

        name = f"MCP_SS_{os.getpid()}_{int(time.time() * 1000) % 1000000}"
        try:
            self._doc.SelectionSets.Item(name).Delete()
        except Exception:
            pass
        sset = self._doc.SelectionSets.Add(name)
        entities = []
        for h in handles:
            try:
                entities.append(self._handle_object(h))
            except Exception as exc:
                log.warning("selection_add_failed", handle=str(h), error=str(exc))
        if entities:
            arr = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, entities
            )
            sset.AddItems(arr)
        return sset

    def _delete_selection_set(self, sset: Any) -> None:
        try:
            sset.Delete()
        except Exception:
            pass

    async def _com_retry(self, fn, tries: int = 15, delay: float = 2.0, label: str = ""):
        """Retry a COM call that AutoCAD rejected because it was busy.

        ``RPC_E_CALL_REJECTED`` ("Call was rejected by callee") is transient;
        Autodesk's documented workaround is to retry after a pause.
        """
        last: Exception | None = None
        for attempt in range(tries):
            try:
                return fn()
            except Exception as exc:
                last = exc
                log.warning("com_retry", label=label, attempt=attempt, error=str(exc)[:120])
                if attempt < tries - 1:
                    await asyncio.sleep(delay)
        raise last  # type: ignore[misc]

    async def _await_document_ready(
        self, doc: Any, timeout: float = 120.0, interval: float = 0.5
    ) -> bool:
        """Wait until a freshly opened/created document answers COM.

        ``Documents.Add`` / ``Documents.Open`` return before the document is
        initialised; touching ModelSpace straight away raises
        "Call was rejected by callee".
        """
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            try:
                _ = doc.ModelSpace.Count
                return True
            except Exception:
                await asyncio.sleep(interval)
        return False

    async def _wait_idle(self, timeout: float = 300.0) -> bool:
        """Poll CMDNAMES until the command line is quiescent.

        Note: idleness does *not* imply the command succeeded — an unknown
        command also returns to idle.  Callers that need certainty must
        verify state afterwards.
        """
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            try:
                if self._doc.GetVariable("CMDNAMES") == "":
                    return True
            except Exception:
                pass
            await asyncio.sleep(0.4)
        return False

    async def _send_command(self, command: str, timeout: float = 300.0) -> bool:
        """Send a raw command string and wait for the command line to settle."""
        if not await self._wait_idle(timeout=30.0):
            raise RuntimeError("command line busy before SendCommand")
        self._doc.SendCommand(command)
        return await self._wait_idle(timeout=timeout)

    async def _lisp_eval(self, code: str, timeout: float = 180.0) -> CommandResult:
        """Evaluate AutoLISP in the document and return its printed result.

        The snippet is written to a temp ``.lsp`` and loaded with a single
        short ``(load "...")`` — the corpus-proven pattern.  A footer writes
        the result (or the caught error) to a sibling file so nothing is lost
        if the expression raises.
        """
        _LISP_DIR.mkdir(parents=True, exist_ok=True)
        stamp = f"{os.getpid()}_{int(time.time() * 1000) % 10000000}"
        lsp_path = _LISP_DIR / f"eval_{stamp}.lsp"
        out_path = _LISP_DIR / f"eval_{stamp}.out"
        for p in (lsp_path, out_path):
            if p.exists():
                p.unlink()

        wrapper = (
            "(progn\n"
            f"  (setq *mcp-res* (vl-catch-all-apply '(lambda () {code})))\n"
            f'  (setq *mcp-f* (open "{out_path.as_posix()}" "w"))\n'
            "  (if (vl-catch-all-error-p *mcp-res*)\n"
            '    (write-line (strcat "ERROR: " (vl-catch-all-error-message *mcp-res*)) *mcp-f*)\n'
            "    (write-line (vl-prin1-to-string *mcp-res*) *mcp-f*))\n"
            "  (close *mcp-f*)\n"
            "  (princ)\n"
            ")\n"
        )
        lsp_path.write_text(wrapper, encoding="utf-8", errors="replace")

        await self._send_command(f'(load "{lsp_path.as_posix()}") \n', timeout=60.0)

        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            await asyncio.sleep(0.25)
            try:
                text = out_path.read_text(encoding="utf-8", errors="replace").strip()
            except Exception:
                continue
            if not text:
                continue
            try:
                lsp_path.unlink()
                out_path.unlink()
            except Exception:
                pass
            if text.startswith("ERROR: "):
                return CommandResult(ok=False, error=text[7:].strip())
            return CommandResult(ok=True, payload={"result": text[:100_000]})
        return CommandResult(
            ok=False, error=f"LISP evaluation timed out after {timeout}s"
        )

    # ==================================================================
    # Escape hatches
    # ==================================================================

    async def execute_lisp(self, code: str) -> CommandResult:
        """Execute arbitrary AutoLISP in the active drawing."""
        err = self._err()
        if err:
            return err
        if not code or not code.strip():
            return CommandResult(ok=False, error="Empty LISP code")
        try:
            return await self._lisp_eval(code)
        except Exception as exc:
            return CommandResult(ok=False, error=f"execute_lisp failed: {exc}")

    async def send_command(self, command: str) -> CommandResult:
        """Send a raw AutoCAD command string (newline-separated prompts)."""
        err = self._err()
        if err:
            return err
        if not command or not command.strip():
            return CommandResult(ok=False, error="Empty command")
        try:
            settled = await self._send_command(command)
            return CommandResult(
                ok=True,
                payload={
                    "command": command.replace("\n", "|"),
                    "settled": settled,
                    "note": (
                        "settled=True only means the command line is idle; "
                        "verify drawing state for commands with side effects"
                    ),
                },
            )
        except Exception as exc:
            return CommandResult(ok=False, error=f"send_command failed: {exc}")

    # ==================================================================
    # Drawing-level operations
    # ==================================================================

    async def drawing_units(self, units: Any = None) -> CommandResult:
        """Get or set INSUNITS."""
        err = self._err()
        if err:
            return err
        try:
            if units is None:
                code = int(self._doc.GetVariable("INSUNITS"))
                return CommandResult(ok=True, payload={
                    "insunits": code,
                    "name": _UNIT_BY_CODE.get(code, "unknown"),
                })

            if isinstance(units, str) and units.strip().isdigit():
                units = int(units.strip())
            if isinstance(units, str):
                code = _UNIT_CODE.get(units.strip().lower())
                if code is None:
                    return CommandResult(
                        ok=False,
                        error=f"Unknown unit {units!r}. Use one of {sorted(_UNIT_CODE)}",
                    )
            else:
                code = int(units)

            self._doc.SetVariable("INSUNITS", code)
            actual = int(self._doc.GetVariable("INSUNITS"))
            return CommandResult(ok=True, payload={
                "insunits": actual,
                "name": _UNIT_BY_CODE.get(actual, "unknown"),
                "changed": actual == code,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"drawing_units failed: {exc}")

    async def drawing_limits(
        self,
        x1: float | None = None,
        y1: float | None = None,
        x2: float | None = None,
        y2: float | None = None,
    ) -> CommandResult:
        """Get or set the drawing limits."""
        err = self._err()
        if err:
            return err
        try:
            if any(v is None for v in (x1, y1, x2, y2)):
                lo = self._doc.GetVariable("LIMMIN")
                hi = self._doc.GetVariable("LIMMAX")
                return CommandResult(ok=True, payload={
                    "limmin": [round(float(v), 4) for v in lo],
                    "limmax": [round(float(v), 4) for v in hi],
                })
            self._doc.SetVariable("LIMMIN", self._to_variant_array((float(x1), float(y1))))
            self._doc.SetVariable("LIMMAX", self._to_variant_array((float(x2), float(y2))))
            return CommandResult(ok=True, payload={
                "limmin": [float(x1), float(y1)],
                "limmax": [float(x2), float(y2)],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"drawing_limits failed: {exc}")

    async def drawing_audit(self, fix: bool = False) -> CommandResult:
        """Run AUDIT. With ``fix=True`` AutoCAD repairs errors it finds."""
        err = self._err()
        if err:
            return err
        try:
            before = int(self._msp.Count)
            try:
                has_errors = bool(self._doc.AuditInfo(bool(fix)))
                method = "AuditInfo"
            except Exception:
                await self._send_command(
                    f"_.AUDIT\n{'_y' if fix else '_n'}\n", timeout=600.0
                )
                has_errors = None
                method = "AUDIT command"
            after = int(self._msp.Count)
            return CommandResult(ok=True, payload={
                "fix": fix,
                "method": method,
                "errors_found": has_errors,
                "entities_before": before,
                "entities_after": after,
                "entities_removed": before - after,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"drawing_audit failed: {exc}")

    async def drawing_wblock(self, handles: list, path: str) -> CommandResult:
        """Write the given entities out to a separate DWG."""
        err = self._err()
        if err:
            return err
        if not handles:
            return CommandResult(ok=False, error="No handles supplied")
        sset = None
        try:
            out = Path(path)
            out.parent.mkdir(parents=True, exist_ok=True)
            sset = self._selection_set(handles)
            self._doc.Wblock(str(out), sset)
            return CommandResult(ok=True, payload={
                "path": str(out),
                "written": out.exists(),
                "size": out.stat().st_size if out.exists() else 0,
                "count": len(handles),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"drawing_wblock failed: {exc}")
        finally:
            if sset is not None:
                self._delete_selection_set(sset)

    # ==================================================================
    # XREF management
    # ==================================================================

    def _xref_records(self) -> list[dict]:
        out = []
        try:
            blocks = self._doc.Blocks
        except Exception:
            return out
        for i in range(blocks.Count):
            try:
                blk = blocks.Item(i)
                if not blk.IsXRef:
                    continue
                path = ""
                try:
                    path = blk.XRefPath or ""
                except Exception:
                    pass
                out.append({
                    "name": blk.Name,
                    "path": path,
                    "resolved": bool(path) and os.path.exists(path),
                })
            except Exception:
                continue
        return out

    async def xref_list(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            recs = self._xref_records()
            return CommandResult(ok=True, payload={
                "count": len(recs),
                "xrefs": recs,
                "unresolved": [r["name"] for r in recs if not r["resolved"]],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"xref_list failed: {exc}")

    async def xref_attach(
        self, path: str, x: float = 0, y: float = 0,
        scale: float = 1.0, overlay: bool = False,
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        if not os.path.exists(path):
            return CommandResult(ok=False, error=f"Xref file not found: {path}")
        try:
            before = {r["name"] for r in self._xref_records()}
            mode = "_o" if overlay else "_a"  # overlay / attach
            await self._send_command(
                f"_.-XREF\n{mode}\n{Path(path).as_posix()}\n"
                f"{float(x)},{float(y)}\n{float(scale)}\n0\n",
                timeout=180.0,
            )
            after = {r["name"] for r in self._xref_records()}
            new = sorted(after - before)
            return CommandResult(ok=True, payload={
                "attached": new,
                "path": path,
                "overlay": overlay,
                "count": len(after),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"xref_attach failed: {exc}")

    async def xref_detach(self, name: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            before = len(self._xref_records())
            await self._send_command(f"_.-XREF\n_d\n{name}\n", timeout=180.0)
            after = self._xref_records()
            return CommandResult(ok=True, payload={
                "detached": name,
                "before": before,
                "after": len(after),
                "removed": not any(r["name"] == name for r in after),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"xref_detach failed: {exc}")

    async def xref_bind(self, name: str, insert: bool = False) -> CommandResult:
        """Bind an xref into the drawing.

        An xref with no path or a missing source file cannot be bound —
        AutoCAD rejects it — so that is reported rather than attempted.
        """
        err = self._err()
        if err:
            return err
        rec = next((r for r in self._xref_records() if r["name"] == name), None)
        if rec is None:
            return CommandResult(ok=False, error=f"No xref named {name!r}")
        if not rec["resolved"]:
            return CommandResult(
                ok=False,
                error=(
                    f"Cannot bind {name!r}: no resolvable source file "
                    f"(path={rec['path']!r})"
                ),
            )
        try:
            before = {r["name"] for r in self._xref_records()}
            await self._send_command(f"_.-XREF\n_b\n{name}\n", timeout=600.0)
            after = {r["name"] for r in self._xref_records()}
            return CommandResult(ok=True, payload={
                "bound": name,
                "still_xref": name in after,
                "before": sorted(before),
                "after": sorted(after),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"xref_bind failed: {exc}")

    async def xref_reload(self, name: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            await self._send_command(f"_.-XREF\n_r\n{name}\n", timeout=300.0)
            return CommandResult(ok=True, payload={"reloaded": name})
        except Exception as exc:
            return CommandResult(ok=False, error=f"xref_reload failed: {exc}")

    async def xref_path_update(self, name: str, new_path: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            await self._send_command(
                f"_.-XREF\n_p\n{name}\n{Path(new_path).as_posix()}\n", timeout=300.0
            )
            rec = next((r for r in self._xref_records() if r["name"] == name), None)
            return CommandResult(ok=True, payload={
                "name": name,
                "new_path": new_path,
                "record": rec,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"xref_path_update failed: {exc}")

    async def xref_query_entities(
        self, name: str, entity_type: str | None = None, layer: str | None = None
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            blk = self._doc.Blocks.Item(name)
            found = []
            for i in range(blk.Count):
                try:
                    e = self._dyn(blk.Item(i))
                except Exception:
                    continue
                if entity_type and e.EntityName.upper() != entity_type.upper():
                    continue
                if layer and str(e.Layer).lower() != layer.lower():
                    continue
                found.append({"handle": e.Handle, "type": e.EntityName, "layer": e.Layer})
                if len(found) >= 500:
                    break
            return CommandResult(ok=True, payload={
                "xref": name,
                "count": len(found),
                "entities": found,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"xref_query_entities failed: {exc}")

    # ==================================================================
    # Queries
    # ==================================================================

    async def query_entity_geometry(self, entity_id: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            e = self._dyn(self._handle_object(entity_id))
            t = str(e.EntityName).upper()
            geo: dict[str, Any] = {"handle": e.Handle, "type": e.EntityName}
            if t == "ACDBLINE":
                geo["start"] = [round(v, 6) for v in e.StartPoint]
                geo["end"] = [round(v, 6) for v in e.EndPoint]
                geo["length"] = round(float(e.Length), 6)
                geo["angle_deg"] = round(math.degrees(float(e.Angle)), 6)
            elif t in ("ACDBCIRCLE", "ACDBARC"):
                geo["center"] = [round(v, 6) for v in e.Center]
                geo["radius"] = round(float(e.Radius), 6)
                if t == "ACDBCIRCLE":
                    geo["area"] = round(float(e.Area), 6)
                    geo["circumference"] = round(float(e.Circumference), 6)
                else:
                    geo["start_angle_deg"] = round(math.degrees(float(e.StartAngle)), 6)
                    geo["end_angle_deg"] = round(math.degrees(float(e.EndAngle)), 6)
                    geo["arc_length"] = round(float(e.ArcLength), 6)
            elif t in ("ACDBLWPOLYLINE", "ACDBPOLYLINE", "ACDB2DPOLYLINE"):
                coords = [round(float(v), 6) for v in e.Coordinates]
                geo["closed"] = bool(self._safe_get_property(e, "Closed", False))
                geo["vertex_count"] = len(coords) // 2
                geo["coordinates"] = coords
                geo["area"] = self._safe_get_property(e, "Area")
                geo["length"] = self._safe_get_property(e, "Length")
            elif t in ("ACDBTEXT", "ACDBMTEXT"):
                geo["position"] = [round(v, 6) for v in e.InsertionPoint]
                geo["text"] = e.TextString
                geo["height"] = self._safe_get_property(e, "Height")
            else:
                try:
                    lo, hi = e.GetBoundingBox()
                    geo["bounding_box"] = {
                        "min": [round(float(v), 6) for v in lo],
                        "max": [round(float(v), 6) for v in hi],
                    }
                except Exception:
                    geo["bounding_box"] = None
            return CommandResult(ok=True, payload=geo)
        except Exception as exc:
            return CommandResult(ok=False, error=f"query_entity_geometry failed: {exc}")

    async def query_layer_summary(self, layer: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            by_type: Counter = Counter()
            xs: list[float] = []
            ys: list[float] = []
            count = 0
            for e in self._iter_msp():
                try:
                    if str(e.Layer).lower() != layer.lower():
                        continue
                    by_type[e.EntityName] += 1
                    count += 1
                    lo, hi = e.GetBoundingBox()
                    xs.extend([float(lo[0]), float(hi[0])])
                    ys.extend([float(lo[1]), float(hi[1])])
                except Exception:
                    continue
            bbox = None
            if xs:
                bbox = {
                    "min": [round(min(xs), 4), round(min(ys), 4)],
                    "max": [round(max(xs), 4), round(max(ys), 4)],
                }
            return CommandResult(ok=True, payload={
                "layer": layer,
                "count": count,
                "by_type": dict(by_type),
                "bounding_box": bbox,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"query_layer_summary failed: {exc}")

    async def query_block_tree(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            blocks = self._doc.Blocks
            tree = []
            for i in range(blocks.Count):
                try:
                    blk = blocks.Item(i)
                    nested = set()
                    for j in range(blk.Count):
                        try:
                            child = blk.Item(j)
                            if str(child.EntityName).upper() == "ACDBBLOCKREFERENCE":
                                nested.add(child.Name)
                        except Exception:
                            continue
                    tree.append({
                        "name": blk.Name,
                        "is_xref": bool(self._safe_get_property(blk, "IsXRef", False)),
                        "is_layout": bool(self._safe_get_property(blk, "IsLayout", False)),
                        "entity_count": int(blk.Count),
                        "nested_blocks": sorted(nested),
                    })
                except Exception:
                    continue
            used = Counter()
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() == "ACDBBLOCKREFERENCE":
                        used[e.Name] += 1
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "block_count": len(tree),
                "used_in_modelspace": dict(used),
                "blocks": tree,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"query_block_tree failed: {exc}")

    async def query_drawing_metadata(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            si = self._doc.SummaryInfo
            fields = {}
            for attr in ("Title", "Subject", "Author", "Keywords", "Comments",
                         "LastSavedBy", "RevisionNumber", "HyperlinkBase"):
                try:
                    fields[attr.lower()] = getattr(si, attr)
                except Exception:
                    fields[attr.lower()] = None
            return CommandResult(ok=True, payload={
                "name": self._doc.Name,
                "path": self._doc.FullName,
                "saved": bool(self._doc.Saved),
                "summary": fields,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"query_drawing_metadata failed: {exc}")

    def _collection_summary(self, coll: Any, attrs: tuple[str, ...]) -> list[dict]:
        out = []
        for i in range(coll.Count):
            try:
                item = coll.Item(i)
            except Exception:
                continue
            rec = {"name": self._safe_get_property(item, "Name")}
            for a in attrs:
                rec[a] = self._safe_get_property(item, a)
            out.append(rec)
        return out

    async def query_text_styles(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            styles = self._collection_summary(
                self._doc.TextStyles,
                ("FontFile", "BigFontFile", "Height", "Width", "ObliqueAngle"),
            )
            return CommandResult(ok=True, payload={"count": len(styles), "styles": styles})
        except Exception as exc:
            return CommandResult(ok=False, error=f"query_text_styles failed: {exc}")

    async def query_dimension_styles(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            styles = self._collection_summary(
                self._doc.DimStyles,
                ("ScaleFactor", "TextHeight", "ArrowheadSize", "DecimalSeparator"),
            )
            return CommandResult(ok=True, payload={"count": len(styles), "styles": styles})
        except Exception as exc:
            return CommandResult(ok=False, error=f"query_dimension_styles failed: {exc}")

    async def query_linetypes(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            lts = self._collection_summary(
                self._doc.Linetypes, ("Description", "LinetypeGeneration", "PaperSpaceUnits")
            )
            return CommandResult(ok=True, payload={"count": len(lts), "linetypes": lts})
        except Exception as exc:
            return CommandResult(ok=False, error=f"query_linetypes failed: {exc}")

    # ==================================================================
    # Geometry
    # ==================================================================

    async def geometry_distance(
        self, x1: float, y1: float, x2: float, y2: float
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        dx, dy = float(x2) - float(x1), float(y2) - float(y1)
        return CommandResult(ok=True, payload={
            "distance": round(math.hypot(dx, dy), 6),
            "dx": round(dx, 6),
            "dy": round(dy, 6),
            "angle_deg": round(math.degrees(math.atan2(dy, dx)), 6),
        })

    async def geometry_area(self, entity_id: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            e = self._dyn(self._handle_object(entity_id))
            area = self._safe_get_property(e, "Area")
            if area is None:
                return CommandResult(
                    ok=False, error=f"{e.EntityName} has no area property"
                )
            return CommandResult(ok=True, payload={
                "handle": e.Handle, "type": e.EntityName, "area": float(area),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"geometry_area failed: {exc}")

    async def geometry_length(self, entity_id: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            e = self._dyn(self._handle_object(entity_id))
            for attr in ("Length", "ArcLength", "Circumference", "Perimeter"):
                val = self._safe_get_property(e, attr)
                if val is not None:
                    return CommandResult(ok=True, payload={
                        "handle": e.Handle, "type": e.EntityName,
                        "length": float(val), "property": attr,
                    })
            return CommandResult(
                ok=False, error=f"{e.EntityName} has no length property"
            )
        except Exception as exc:
            return CommandResult(ok=False, error=f"geometry_length failed: {exc}")

    async def geometry_bounding_box(
        self, entity_id: str | None = None, layer: str | None = None
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            if entity_id:
                e = self._dyn(self._handle_object(entity_id))
                lo, hi = e.GetBoundingBox()
                return CommandResult(ok=True, payload={
                    "scope": "entity",
                    "handle": e.Handle,
                    "min": [round(float(v), 6) for v in lo],
                    "max": [round(float(v), 6) for v in hi],
                })
            xs: list[float] = []
            ys: list[float] = []
            zs: list[float] = []
            n = 0
            for e in self._iter_msp():
                try:
                    if layer and str(e.Layer).lower() != layer.lower():
                        continue
                    lo, hi = e.GetBoundingBox()
                    xs.extend([float(lo[0]), float(hi[0])])
                    ys.extend([float(lo[1]), float(hi[1])])
                    zs.extend([float(lo[2]), float(hi[2])])
                    n += 1
                except Exception:
                    continue
            if not xs:
                return CommandResult(ok=True, payload={"scope": "empty", "count": 0})
            return CommandResult(ok=True, payload={
                "scope": "layer" if layer else "drawing",
                "layer": layer,
                "count": n,
                "min": [round(min(xs), 6), round(min(ys), 6), round(min(zs), 6)],
                "max": [round(max(xs), 6), round(max(ys), 6), round(max(zs), 6)],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"geometry_bounding_box failed: {exc}")

    async def geometry_polyline_info(self, entity_id: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            e = self._dyn(self._handle_object(entity_id))
            if str(e.EntityName).upper() not in (
                "ACDBLWPOLYLINE", "ACDBPOLYLINE", "ACDB2DPOLYLINE", "ACDB3DPOLYLINE"
            ):
                return CommandResult(
                    ok=False, error=f"Not a polyline: {e.EntityName}"
                )
            coords = [float(v) for v in e.Coordinates]
            pts = [coords[i:i + 2] for i in range(0, len(coords), 2)]
            return CommandResult(ok=True, payload={
                "handle": e.Handle,
                "type": e.EntityName,
                "vertex_count": len(pts),
                "closed": bool(self._safe_get_property(e, "Closed", False)),
                "coordinates": [[round(a, 6), round(b, 6)] for a, b in pts],
                "area": self._safe_get_property(e, "Area"),
                "length": self._safe_get_property(e, "Length"),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"geometry_polyline_info failed: {exc}")

    # ==================================================================
    # Search
    # ==================================================================

    async def find_text(
        self, pattern: str, case_sensitive: bool = False, max_results: int = 50,
        zoom_to_first: bool = True, zoom_height: float = 600.0,
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            needle = pattern if case_sensitive else pattern.lower()
            hits = []
            for e in self._iter_msp():
                try:
                    t = str(e.EntityName).upper()
                    if t not in ("ACDBTEXT", "ACDBMTEXT", "ACDBATTRIBUTE"):
                        continue
                    txt = str(e.TextString)
                    hay = txt if case_sensitive else txt.lower()
                    if needle not in hay:
                        continue
                    hits.append({
                        "handle": e.Handle,
                        "type": e.EntityName,
                        "layer": e.Layer,
                        "text": txt,
                        "position": [round(float(v), 4) for v in e.InsertionPoint],
                    })
                    if len(hits) >= max_results:
                        break
                except Exception:
                    continue
            if hits and zoom_to_first:
                try:
                    p = hits[0]["position"]
                    self._doc.Application.ZoomCenter(
                        self._to_variant_array((p[0], p[1])), float(zoom_height)
                    )
                except Exception:
                    pass
            return CommandResult(ok=True, payload={
                "pattern": pattern, "count": len(hits), "matches": hits,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"find_text failed: {exc}")

    def _bbox_of(self, e: Any) -> tuple[tuple, tuple] | None:
        try:
            lo, hi = e.GetBoundingBox()
            return (tuple(float(v) for v in lo), tuple(float(v) for v in hi))
        except Exception:
            return None

    async def search_by_window(
        self, x1: float, y1: float, x2: float, y2: float
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            lo_x, hi_x = sorted((float(x1), float(x2)))
            lo_y, hi_y = sorted((float(y1), float(y2)))
            hits = []
            for e in self._iter_msp():
                bb = self._bbox_of(e)
                if bb is None:
                    continue
                (bx0, by0, _), (bx1, by1, _) = bb
                if bx1 < lo_x or bx0 > hi_x or by1 < lo_y or by0 > hi_y:
                    continue
                hits.append({"handle": e.Handle, "type": e.EntityName, "layer": e.Layer})
            return CommandResult(ok=True, payload={"count": len(hits), "entities": hits})
        except Exception as exc:
            return CommandResult(ok=False, error=f"search_by_window failed: {exc}")

    async def search_by_proximity(
        self, x: float, y: float, radius: float
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            cx, cy, r = float(x), float(y), float(radius)
            hits = []
            for e in self._iter_msp():
                bb = self._bbox_of(e)
                if bb is None:
                    continue
                (bx0, by0, _), (bx1, by1, _) = bb
                # distance from point to bbox (0 if inside)
                dx = max(bx0 - cx, 0.0, cx - bx1)
                dy = max(by0 - cy, 0.0, cy - by1)
                d = math.hypot(dx, dy)
                if d <= r:
                    hits.append({
                        "handle": e.Handle, "type": e.EntityName, "layer": e.Layer,
                        "distance": round(d, 6),
                    })
            hits.sort(key=lambda h: h["distance"])
            return CommandResult(ok=True, payload={
                "count": len(hits), "radius": r, "entities": hits,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"search_by_proximity failed: {exc}")

    async def search_by_block_name(self, block_name: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            hits = []
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() != "ACDBBLOCKREFERENCE":
                        continue
                    if str(e.Name).lower() != block_name.lower():
                        continue
                    hits.append({
                        "handle": e.Handle, "layer": e.Layer,
                        "position": [round(float(v), 4) for v in e.InsertionPoint],
                        "rotation": self._safe_get_property(e, "Rotation"),
                    })
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "block": block_name, "count": len(hits), "inserts": hits,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"search_by_block_name failed: {exc}")

    async def search_by_attribute(
        self, tag: str | None = None, value: str | None = None
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            hits = []
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() != "ACDBBLOCKREFERENCE":
                        continue
                    atts = e.GetAttributes()
                except Exception:
                    continue
                for a in atts:
                    try:
                        if tag and str(a.TagString).lower() != tag.lower():
                            continue
                        if value and str(a.TextString).lower() != value.lower():
                            continue
                        hits.append({
                            "handle": e.Handle,
                            "block": e.Name,
                            "layer": e.Layer,
                            "tag": a.TagString,
                            "value": a.TextString,
                        })
                    except Exception:
                        continue
            return CommandResult(ok=True, payload={
                "tag": tag, "value": value, "count": len(hits), "matches": hits,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"search_by_attribute failed: {exc}")

    async def search_by_handle_list(self, handles: list) -> CommandResult:
        err = self._err()
        if err:
            return err
        found, missing = [], []
        for h in handles:
            try:
                e = self._dyn(self._handle_object(h))
                found.append({
                    "handle": e.Handle, "type": e.EntityName, "layer": e.Layer,
                })
            except Exception:
                missing.append(str(h))
        return CommandResult(ok=True, payload={
            "found": found, "missing": missing, "count": len(found),
        })

    # ==================================================================
    # Selection / bulk modification
    # ==================================================================

    async def select_filter(
        self, entity_type: str | None = None, layer: str | None = None,
        color: Any = None, x1: float | None = None, y1: float | None = None,
        x2: float | None = None, y2: float | None = None,
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            want_color = self._get_color_index(color)
            window = None
            if None not in (x1, y1, x2, y2):
                window = (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
            hits = []
            for e in self._iter_msp():
                try:
                    if entity_type and str(e.EntityName).upper() != entity_type.upper():
                        continue
                    if layer and str(e.Layer).lower() != layer.lower():
                        continue
                    if want_color is not None:
                        if self._safe_get_property(e, "color") != want_color:
                            continue
                    if window:
                        bb = self._bbox_of(e)
                        if bb is None:
                            continue
                        (bx0, by0, _), (bx1, by1, _) = bb
                        if bx1 < window[0] or bx0 > window[2] or by1 < window[1] or by0 > window[3]:
                            continue
                    hits.append({"handle": e.Handle, "type": e.EntityName, "layer": e.Layer})
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "count": len(hits), "handles": [h["handle"] for h in hits],
                "entities": hits[:1000],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"select_filter failed: {exc}")

    async def bulk_move(self, handles: list, dx: float, dy: float) -> CommandResult:
        err = self._err()
        if err:
            return err
        moved, errors = [], []
        for h in handles:
            try:
                e = self._dyn(self._handle_object(h))
                e.Move(self._to_variant_array((0.0, 0.0, 0.0)),
                       self._to_variant_array((float(dx), float(dy), 0.0)))
                moved.append(str(h))
            except Exception as exc:
                errors.append(f"{h}: {exc}")
        return CommandResult(ok=True, payload={
            "moved": moved, "errors": errors, "moved_count": len(moved),
        })

    async def bulk_copy(self, handles: list, dx: float, dy: float) -> CommandResult:
        err = self._err()
        if err:
            return err
        created, errors = [], []
        for h in handles:
            try:
                e = self._dyn(self._handle_object(h))
                new = e.Copy()
                new.Move(self._to_variant_array((0.0, 0.0, 0.0)),
                         self._to_variant_array((float(dx), float(dy), 0.0)))
                created.append(new.Handle)
            except Exception as exc:
                errors.append(f"{h}: {exc}")
        return CommandResult(ok=True, payload={
            "created": created, "errors": errors, "created_count": len(created),
        })

    async def entity_set_property(
        self, entity_id: str, property_name: str, value: Any
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            e = self._dyn(self._handle_object(entity_id))
            prop_map = {
                "layer": "Layer", "color": "color", "linetype": "Linetype",
                "lineweight": "Lineweight", "visible": "Visible",
                "linetypescale": "LinetypeScale",
            }
            prop = prop_map.get(str(property_name).lower(), property_name)
            before = self._safe_get_property(e, prop)
            setattr(e, prop, value)
            return CommandResult(ok=True, payload={
                "handle": e.Handle, "property": prop,
                "before": before, "after": self._safe_get_property(e, prop),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"entity_set_property failed: {exc}")

    async def entity_set_text(self, entity_id: str, text: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            e = self._dyn(self._handle_object(entity_id))
            if str(e.EntityName).upper() not in (
                "ACDBTEXT", "ACDBMTEXT", "ACDBATTRIBUTE", "ACDBATTRIBUTEDEFINITION"
            ):
                return CommandResult(
                    ok=False, error=f"{e.EntityName} has no text content"
                )
            before = e.TextString
            e.TextString = text
            return CommandResult(ok=True, payload={
                "handle": e.Handle, "before": before, "after": e.TextString,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"entity_set_text failed: {exc}")

    async def entity_explode(self, entity_id: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            e = self._dyn(self._handle_object(entity_id))
            new = e.Explode()
            handles = []
            for item in new or []:
                try:
                    handles.append(item.Handle)
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "exploded": e.Handle, "created": handles, "count": len(handles),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"entity_explode failed: {exc}")

    async def entity_join(self, entity_ids: list) -> CommandResult:
        err = self._err()
        if err:
            return err
        if len(entity_ids) < 2:
            return CommandResult(ok=False, error="Join needs at least two entities")
        try:
            sset = self._selection_set(entity_ids)
            try:
                names = ",".join(f'"{h}"' for h in entity_ids)
                lisp = (
                    f"(progn (setq ss (ssadd)) "
                    f"(foreach h (list {names}) (ssadd (handent h) ss)) "
                    f'(command "_.JOIN" ss "") "ok")'
                )
                result = await self._lisp_eval(lisp, timeout=120.0)
            finally:
                self._delete_selection_set(sset)
            return CommandResult(ok=result.ok, payload=result.payload, error=result.error)
        except Exception as exc:
            return CommandResult(ok=False, error=f"entity_join failed: {exc}")

    async def entity_fillet(
        self, entity_id1: str, entity_id2: str, radius: float
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        lisp = (
            f'(progn (setvar "FILLETRAD" {float(radius)}) '
            f'(command "_.FILLET" (handent "{entity_id1}") (handent "{entity_id2}")) '
            '"ok")'
        )
        try:
            return await self._lisp_eval(lisp, timeout=120.0)
        except Exception as exc:
            return CommandResult(ok=False, error=f"entity_fillet failed: {exc}")

    async def entity_chamfer(
        self, entity_id1: str, entity_id2: str, dist1: float, dist2: float
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        lisp = (
            f'(progn (setvar "CHAMFERA" {float(dist1)}) '
            f'(setvar "CHAMFERB" {float(dist2)}) '
            f'(command "_.CHAMFER" (handent "{entity_id1}") (handent "{entity_id2}")) '
            '"ok")'
        )
        try:
            return await self._lisp_eval(lisp, timeout=120.0)
        except Exception as exc:
            return CommandResult(ok=False, error=f"entity_chamfer failed: {exc}")

    async def entity_break_at(
        self, entity_id: str, x: float, y: float
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        lisp = (
            f'(progn (command "_.BREAK" (handent "{entity_id}") '
            f'(list {float(x)} {float(y)} 0.0) (list {float(x)} {float(y)} 0.0)) "ok")'
        )
        try:
            return await self._lisp_eval(lisp, timeout=120.0)
        except Exception as exc:
            return CommandResult(ok=False, error=f"entity_break_at failed: {exc}")

    async def _trim_extend(
        self, command: str, entity_id: str, boundary_id: str
    ) -> CommandResult:
        """Classic (non-quick) TRIM/EXTEND: boundary first, then target."""
        lisp = (
            '(progn (setq old (getvar "TRIMEXTENDMODE")) '
            '(setvar "TRIMEXTENDMODE" 0) '
            f'(command "_.{command}" (handent "{boundary_id}") "" '
            f'(handent "{entity_id}") "") '
            '(setvar "TRIMEXTENDMODE" old) "ok")'
        )
        return await self._lisp_eval(lisp, timeout=120.0)

    async def entity_trim(self, entity_id: str, boundary_id: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            return await self._trim_extend("TRIM", entity_id, boundary_id)
        except Exception as exc:
            return CommandResult(ok=False, error=f"entity_trim failed: {exc}")

    async def entity_extend(self, entity_id: str, boundary_id: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            return await self._trim_extend("EXTEND", entity_id, boundary_id)
        except Exception as exc:
            return CommandResult(ok=False, error=f"entity_extend failed: {exc}")

    # ==================================================================
    # Layers (rename / merge / visibility)
    # ==================================================================

    async def layer_rename(self, old_name: str, new_name: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            self._doc.Layers.Item(old_name).Name = new_name
            return CommandResult(ok=True, payload={
                "old": old_name, "new": new_name,
                "exists": self._safe_get_property(
                    self._doc.Layers.Item(new_name), "Name"
                ) == new_name,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layer_rename failed: {exc}")

    async def layer_merge(self, source_layer: str, target_layer: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            try:
                self._doc.Layers.Item(target_layer)
            except Exception:
                self._doc.Layers.Add(target_layer)
            moved = 0
            for e in self._iter_msp():
                try:
                    if str(e.Layer).lower() == source_layer.lower():
                        e.Layer = target_layer
                        moved += 1
                except Exception:
                    continue
            deleted = False
            try:
                self._doc.Layers.Item(source_layer).Delete()
                deleted = True
            except Exception:
                pass
            return CommandResult(ok=True, payload={
                "source": source_layer, "target": target_layer,
                "entities_moved": moved, "source_deleted": deleted,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layer_merge failed: {exc}")

    async def layer_visibility(self, name: str, visible: bool) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            layer = self._doc.Layers.Item(name)
            layer.LayerOn = bool(visible)
            return CommandResult(ok=True, payload={
                "layer": name, "visible": bool(layer.LayerOn),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layer_visibility failed: {exc}")

    # ==================================================================
    # Validation / QC
    # ==================================================================

    async def validate_zero_length(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            hits = []
            for e in self._iter_msp():
                t = str(e.EntityName).upper()
                try:
                    if t == "ACDBLINE":
                        sp, ep = e.StartPoint, e.EndPoint
                        if math.dist(sp, ep) < 1e-9:
                            hits.append({"handle": e.Handle, "type": e.EntityName,
                                         "layer": e.Layer, "issue": "zero-length line"})
                    elif t in ("ACDBCIRCLE", "ACDBARC"):
                        if abs(float(e.Radius)) < 1e-9:
                            hits.append({"handle": e.Handle, "type": e.EntityName,
                                         "layer": e.Layer, "issue": "zero radius"})
                    elif t == "ACDBPOLYLINE":
                        coords = [float(v) for v in e.Coordinates]
                        pts = [coords[i:i + 2] for i in range(0, len(coords), 2)]
                        uniq = {tuple(round(v, 9) for v in p) for p in pts}
                        if len(uniq) < 2:
                            hits.append({"handle": e.Handle, "type": e.EntityName,
                                         "layer": e.Layer, "issue": "degenerate polyline"})
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "count": len(hits), "entities": hits[:500],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"validate_zero_length failed: {exc}")

    def _geometry_signature(self, e: Any, tol: float) -> str | None:
        """Round geometry to *tol* so near-identical entities collide."""
        t = str(e.EntityName).upper()
        try:
            if t == "ACDBLINE":
                pts = [e.StartPoint, e.EndPoint]
            elif t == "ACDBCIRCLE":
                pts = [e.Center, (float(e.Radius), 0.0, 0.0)]
            elif t == "ACDBPOLYLINE":
                c = [float(v) for v in e.Coordinates]
                pts = [tuple(c[i:i + 2]) for i in range(0, len(c), 2)]
            elif t in ("ACDBTEXT", "ACDBMTEXT"):
                pts = [e.InsertionPoint]
                return f"{t}|{str(e.TextString).strip()}|" + "|".join(
                    f"{round(float(v) / tol)}" for v in pts[0]
                )
            else:
                return None
            flat = []
            for p in pts:
                flat.extend(round(float(v) / tol) for v in p)
            return f"{t}|" + "|".join(str(v) for v in flat)
        except Exception:
            return None

    async def validate_duplicates(self, tolerance: float = 0.001) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            tol = float(tolerance) if tolerance else 0.001
            groups: defaultdict[str, list] = defaultdict(list)
            for e in self._iter_msp():
                sig = self._geometry_signature(e, tol)
                if sig:
                    groups[sig].append({"handle": e.Handle, "type": e.EntityName,
                                        "layer": e.Layer})
            dupes = {sig: items for sig, items in groups.items() if len(items) > 1}
            total = sum(len(v) - 1 for v in dupes.values())
            return CommandResult(ok=True, payload={
                "tolerance": tol,
                "duplicate_groups": len(dupes),
                "redundant_entities": total,
                "groups": list(dupes.values())[:100],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"validate_duplicates failed: {exc}")

    async def validate_layer_standards(self, allowed_layers: list) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            allowed = {str(a).lower() for a in allowed_layers}
            used: Counter = Counter()
            for e in self._iter_msp():
                try:
                    used[str(e.Layer)] += 1
                except Exception:
                    continue
            offenders = [
                {"layer": name, "entities": n}
                for name, n in used.items() if name.lower() not in allowed
            ]
            offenders.sort(key=lambda r: -r["entities"])
            return CommandResult(ok=True, payload={
                "allowed_count": len(allowed),
                "layers_used": len(used),
                "offending_layers": offenders,
                "compliant": not offenders,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"validate_layer_standards failed: {exc}")

    async def validate_text_standards(
        self, allowed_styles: list | None = None,
        allowed_heights: list | None = None,
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            styles = {str(s).lower() for s in allowed_styles} if allowed_styles else None
            heights = {round(float(h), 4) for h in allowed_heights} if allowed_heights else None
            offenders = []
            checked = 0
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() not in (
                        "ACDBTEXT", "ACDBMTEXT", "ACDBATTRIBUTE"
                    ):
                        continue
                    checked += 1
                    style = str(self._safe_get_property(e, "StyleName", "") or "")
                    height = self._safe_get_property(e, "Height")
                    issues = []
                    if styles is not None and style.lower() not in styles:
                        issues.append(f"style {style!r}")
                    if heights is not None and height is not None:
                        if round(float(height), 4) not in heights:
                            issues.append(f"height {round(float(height), 4)}")
                    if issues:
                        offenders.append({
                            "handle": e.Handle, "layer": e.Layer,
                            "style": style, "height": height, "issues": issues,
                        })
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "checked": checked,
                "offenders": offenders[:500],
                "offender_count": len(offenders),
                "compliant": not offenders,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"validate_text_standards failed: {exc}")

    async def validate_connectivity(
        self, layer: str | None = None, tolerance: float = 0.01
    ) -> CommandResult:
        """Report line endpoints that are not shared with any other endpoint."""
        err = self._err()
        if err:
            return err
        try:
            tol = float(tolerance) if tolerance else 0.01
            key = lambda p: (round(float(p[0]) / tol), round(float(p[1]) / tol))  # noqa: E731
            endpoint_count: Counter = Counter()
            segments = []
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() != "ACDBLINE":
                        continue
                    if layer and str(e.Layer).lower() != layer.lower():
                        continue
                    a, b = key(e.StartPoint), key(e.EndPoint)
                    endpoint_count[a] += 1
                    endpoint_count[b] += 1
                    segments.append((e.Handle, e.Layer, a, b))
                except Exception:
                    continue
            dangling = []
            for handle, lay, a, b in segments:
                loose = [p for p in (a, b) if endpoint_count[p] == 1]
                if loose:
                    dangling.append({"handle": handle, "layer": lay,
                                     "dangling_endpoints": len(loose)})
            return CommandResult(ok=True, payload={
                "layer": layer,
                "segments_checked": len(segments),
                "dangling_segments": len(dangling),
                "tolerance": tol,
                "entities": dangling[:500],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"validate_connectivity failed: {exc}")

    async def validate_orphaned_entities(self) -> CommandResult:
        """Entities on layers that are off/frozen, plus empty block references."""
        err = self._err()
        if err:
            return err
        try:
            layer_state = {}
            for i in range(self._doc.Layers.Count):
                ly = self._doc.Layers.Item(i)
                layer_state[str(ly.Name)] = {
                    "on": bool(self._safe_get_property(ly, "LayerOn", True)),
                    "frozen": bool(self._safe_get_property(ly, "Freeze", False)),
                }
            hidden_layers = {n for n, s in layer_state.items()
                             if not s["on"] or s["frozen"]}
            hidden_count: Counter = Counter()
            empty_blocks = []
            for e in self._iter_msp():
                try:
                    lay = str(e.Layer)
                    if lay in hidden_layers:
                        hidden_count[lay] += 1
                    if str(e.EntityName).upper() == "ACDBBLOCKREFERENCE":
                        try:
                            if int(self._doc.Blocks.Item(e.Name).Count) == 0:
                                empty_blocks.append({"handle": e.Handle, "block": e.Name})
                        except Exception:
                            pass
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "hidden_layers": sorted(hidden_layers),
                "entities_on_hidden_layers": dict(hidden_count),
                "empty_block_references": empty_blocks[:200],
                "empty_block_count": len(empty_blocks),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"validate_orphaned_entities failed: {exc}")

    async def validate_attribute_completeness(
        self, required_tags: list | None = None
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            required = [str(t).lower() for t in required_tags] if required_tags else None
            incomplete, complete = [], 0
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() != "ACDBBLOCKREFERENCE":
                        continue
                    atts = e.GetAttributes()
                except Exception:
                    continue
                values = {}
                for a in atts:
                    try:
                        values[str(a.TagString).lower()] = str(a.TextString)
                    except Exception:
                        continue
                tags = required if required is not None else list(values)
                missing = [t for t in tags if not str(values.get(t, "")).strip()]
                if missing:
                    incomplete.append({"handle": e.Handle, "block": e.Name,
                                       "missing": missing})
                else:
                    complete += 1
            return CommandResult(ok=True, payload={
                "complete": complete,
                "incomplete": incomplete[:300],
                "incomplete_count": len(incomplete),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"validate_attribute_completeness failed: {exc}")

    async def validate_qc_report(self, allowed_layers: list | None = None) -> CommandResult:
        err = self._err()
        if err:
            return err

        async def section(coro) -> dict:
            """Run one check; never let a single failure sink the report."""
            try:
                res = await coro
                if res.ok and isinstance(res.payload, dict):
                    return res.payload
                return {"error": res.error or "no payload"}
            except Exception as exc:  # noqa: BLE001
                return {"error": f"{type(exc).__name__}: {exc}"}

        try:
            report: dict[str, Any] = {}
            report["zero_length"] = await section(self.validate_zero_length())
            report["duplicates"] = await section(self.validate_duplicates())
            report["orphaned"] = await section(self.validate_orphaned_entities())
            report["connectivity"] = await section(self.validate_connectivity())
            report["attributes"] = await section(self.validate_attribute_completeness())
            if allowed_layers:
                report["layer_standards"] = await section(
                    self.validate_layer_standards(allowed_layers)
                )

            def num(sec: str, key: str) -> int:
                val = report.get(sec, {})
                if not isinstance(val, dict):
                    return 0
                return int(val.get(key, 0) or 0)

            issues = (
                num("zero_length", "count")
                + num("duplicates", "redundant_entities")
                + num("orphaned", "empty_block_count")
                + num("connectivity", "dangling_segments")
                + num("attributes", "incomplete_count")
            )
            failed = [k for k, v in report.items()
                      if isinstance(v, dict) and v.get("error")]
            return CommandResult(ok=True, payload={
                "issue_total": issues,
                "clean": issues == 0 and not failed,
                "failed_sections": failed,
                "sections": report,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"validate_qc_report failed: {exc}")

    # ==================================================================
    # Export / reporting
    # ==================================================================

    async def export_layer_report(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            per_layer: defaultdict[str, Counter] = defaultdict(Counter)
            for e in self._iter_msp():
                try:
                    per_layer[str(e.Layer)][e.EntityName] += 1
                except Exception:
                    continue
            rows = [
                {"layer": name, "total": sum(c.values()), "by_type": dict(c)}
                for name, c in sorted(per_layer.items())
            ]
            return CommandResult(ok=True, payload={
                "layer_count": len(rows), "layers": rows,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"export_layer_report failed: {exc}")

    async def export_block_count(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            counts: Counter = Counter()
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() == "ACDBBLOCKREFERENCE":
                        counts[e.Name] += 1
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "distinct_blocks": len(counts),
                "total_inserts": sum(counts.values()),
                "blocks": dict(sorted(counts.items(), key=lambda kv: -kv[1])),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"export_block_count failed: {exc}")

    async def export_bom(self, block_names: list | None = None) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            wanted = {str(b).lower() for b in block_names} if block_names else None
            rows: defaultdict[str, dict] = defaultdict(
                lambda: {"block": "", "count": 0, "layers": Counter(), "attributes": Counter()}
            )
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() != "ACDBBLOCKREFERENCE":
                        continue
                    name = str(e.Name)
                    if wanted and name.lower() not in wanted:
                        continue
                    row = rows[name]
                    row["block"] = name
                    row["count"] += 1
                    row["layers"][str(e.Layer)] += 1
                    try:
                        for a in e.GetAttributes():
                            row["attributes"][str(a.TagString)] += 1
                    except Exception:
                        pass
                except Exception:
                    continue
            out = []
            for row in rows.values():
                out.append({
                    "block": row["block"],
                    "count": row["count"],
                    "layers": dict(row["layers"]),
                    "attribute_tags": dict(row["attributes"]),
                })
            out.sort(key=lambda r: -r["count"])
            return CommandResult(ok=True, payload={
                "distinct_blocks": len(out),
                "total_inserts": sum(r["count"] for r in out),
                "rows": out,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"export_bom failed: {exc}")

    async def export_data_extract(
        self, entity_type: str | None = None, layer: str | None = None,
        properties: list | None = None,
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            wanted = [str(p) for p in properties] if properties else None
            rows = []
            for e in self._iter_msp():
                try:
                    if entity_type and str(e.EntityName).upper() != entity_type.upper():
                        continue
                    if layer and str(e.Layer).lower() != layer.lower():
                        continue
                    row = {"handle": e.Handle, "type": e.EntityName, "layer": e.Layer}
                    if wanted:
                        for p in wanted:
                            row[p] = self._safe_get_property(e, p)
                    else:
                        for p in ("Layer", "color", "Linetype", "Lineweight"):
                            row[p.lower()] = self._safe_get_property(e, p)
                    rows.append(row)
                    if len(rows) >= 5000:
                        break
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "count": len(rows), "rows": rows,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"export_data_extract failed: {exc}")

    # ==================================================================
    # Blocks
    # ==================================================================

    async def block_define(self, name: str, entities: list) -> CommandResult:
        """Create a block definition from simple entity specs.

        Each spec: ``{"type": "line"|"circle"|"polyline"|"text", ...}``
        """
        err = self._err()
        if err:
            return err
        if not name:
            return CommandResult(ok=False, error="Block name required")
        try:
            try:
                existing = self._doc.Blocks.Item(name)
                if existing:
                    return CommandResult(
                        ok=False, error=f"Block {name!r} already exists"
                    )
            except Exception:
                pass
            blk = self._doc.Blocks.Add(self._to_variant_array((0.0, 0.0, 0.0)), name)
            created = 0
            for spec in entities or []:
                kind = str(spec.get("type", "")).lower()
                if kind == "line":
                    blk.AddLine(
                        self._to_variant_array((spec["x1"], spec["y1"], spec.get("z1", 0.0))),
                        self._to_variant_array((spec["x2"], spec["y2"], spec.get("z2", 0.0))),
                    )
                elif kind == "circle":
                    blk.AddCircle(
                        self._to_variant_array((spec["cx"], spec["cy"], spec.get("cz", 0.0))),
                        float(spec["radius"]),
                    )
                elif kind == "polyline":
                    pts = spec["points"]
                    flat = [float(v) for p in pts for v in (p[0], p[1])]
                    pl = blk.AddLightWeightPolyline(self._dbl_array(flat))
                    pl.Closed = bool(spec.get("closed", False))
                elif kind == "text":
                    blk.AddText(
                        str(spec.get("text", "")),
                        self._to_variant_array((spec["x"], spec["y"], spec.get("z", 0.0))),
                        float(spec.get("height", 1.0)),
                    )
                else:
                    continue
                created += 1
            return CommandResult(ok=True, payload={
                "block": name, "entities_created": created,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"block_define failed: {exc}")

    async def block_insert_with_attributes(
        self, name: str, x: float, y: float, scale: float = 1.0,
        rotation: float = 0.0, attributes: dict | None = None,
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            ref = self._msp.InsertBlock(
                self._to_variant_array((float(x), float(y), 0.0)),
                name, float(scale), float(scale), float(scale), float(rotation),
            )
            applied, missing = {}, []
            if attributes:
                try:
                    atts = ref.GetAttributes()
                except Exception:
                    atts = []
                by_tag = {}
                for a in atts:
                    try:
                        by_tag[str(a.TagString).lower()] = a
                    except Exception:
                        continue
                for tag, value in attributes.items():
                    a = by_tag.get(str(tag).lower())
                    if a is None:
                        missing.append(tag)
                    else:
                        a.TextString = str(value)
                        applied[tag] = str(value)
            return CommandResult(ok=True, payload={
                "handle": ref.Handle, "block": name,
                "attributes_applied": applied, "attributes_missing": missing,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"block_insert_with_attributes failed: {exc}")

    # ==================================================================
    # Dimensions & leaders
    # ==================================================================

    async def create_dimension_aligned(
        self, x1: float, y1: float, x2: float, y2: float, offset: float
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            dx, dy = float(x2) - float(x1), float(y2) - float(y1)
            length = math.hypot(dx, dy) or 1.0
            nx, ny = -dy / length, dx / length
            mid = ((float(x1) + float(x2)) / 2 + nx * offset,
                   (float(y1) + float(y2)) / 2 + ny * offset)
            dim = self._msp.AddDimAligned(
                self._to_variant_array((float(x1), float(y1), 0.0)),
                self._to_variant_array((float(x2), float(y2), 0.0)),
                self._to_variant_array((mid[0], mid[1], 0.0)),
            )
            return CommandResult(ok=True, payload={
                "handle": dim.Handle, "measurement": round(length, 6),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"create_dimension_aligned failed: {exc}")

    async def create_dimension_angular(
        self, cx: float, cy: float, x1: float, y1: float, x2: float, y2: float
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            vertex = self._to_variant_array((float(cx), float(cy), 0.0))
            p1 = self._to_variant_array((float(x1), float(y1), 0.0))
            p2 = self._to_variant_array((float(x2), float(y2), 0.0))
            tx = (float(x1) + float(x2)) / 2 + (float(x1) - float(cx)) * 0.2
            ty = (float(y1) + float(y2)) / 2 + (float(y1) - float(cy)) * 0.2
            dim = self._msp.AddDimAngular(
                vertex, p1, p2, self._to_variant_array((tx, ty, 0.0))
            )
            a1 = math.atan2(float(y1) - float(cy), float(x1) - float(cx))
            a2 = math.atan2(float(y2) - float(cy), float(x2) - float(cx))
            return CommandResult(ok=True, payload={
                "handle": dim.Handle,
                "angle_deg": round(abs(math.degrees(a2 - a1)), 6),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"create_dimension_angular failed: {exc}")

    async def create_dimension_radius(
        self, cx: float, cy: float, radius: float, angle: float
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            center = self._to_variant_array((float(cx), float(cy), 0.0))
            rad = math.radians(float(angle))
            chord = (float(cx) + float(radius) * math.cos(rad),
                     float(cy) + float(radius) * math.sin(rad))
            dim = self._msp.AddDimRadial(
                center, self._to_variant_array((chord[0], chord[1], 0.0)),
                float(radius) * 0.5,
            )
            return CommandResult(ok=True, payload={
                "handle": dim.Handle, "radius": float(radius),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"create_dimension_radius failed: {exc}")

    async def create_leader(self, points: list, text: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            if len(points) < 2:
                return CommandResult(ok=False, error="Leader needs at least 2 points")
            flat = []
            for p in points:
                flat.extend([float(p[0]), float(p[1]), 0.0])
            import pythoncom as _pc
            import win32com.client as _w32

            arr = _w32.VARIANT(_pc.VT_ARRAY | _pc.VT_R8, flat)
            last = points[-1]
            note = self._msp.AddMText(
                self._to_variant_array((float(last[0]) + 2.0, float(last[1]) + 2.0, 0.0)),
                20.0, str(text),
            )
            note.Height = 1.0
            leader = self._msp.AddLeader(arr, note, 0)  # 0 = acLineWithArrow
            return CommandResult(ok=True, payload={
                "leader": leader.Handle, "text": note.Handle,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"create_leader failed: {exc}")

    # ==================================================================
    # View control
    # ==================================================================

    async def zoom_center(self, x: float, y: float, height: float) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            self._doc.Application.ZoomCenter(
                self._to_variant_array((float(x), float(y), 0.0)), float(height)
            )
            return CommandResult(ok=True, payload={
                "center": [float(x), float(y)], "height": float(height),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"zoom_center failed: {exc}")

    async def zoom_scale(self, factor: float) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            # acZoomScaledRelative = 1
            self._doc.Application.ZoomScaled(float(factor), 1)
            return CommandResult(ok=True, payload={"factor": float(factor)})
        except Exception as exc:
            return CommandResult(ok=False, error=f"zoom_scale failed: {exc}")

    async def pan(self, dx: float, dy: float) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            view = self._doc.ActiveViewport
            center = list(view.Center)
            height = float(view.Height)
            new_center = (center[0] + float(dx), center[1] + float(dy), 0.0)
            self._doc.Application.ZoomCenter(
                self._to_variant_array(new_center), height
            )
            return CommandResult(ok=True, payload={
                "from": [round(float(center[0]), 4), round(float(center[1]), 4)],
                "to": [round(new_center[0], 4), round(new_center[1], 4)],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"pan failed: {exc}")

    # ==================================================================
    # Find & replace
    # ==================================================================

    async def find_replace_text(
        self, find: str, replace: str, layer: str | None = None
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            changed = []
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() not in (
                        "ACDBTEXT", "ACDBMTEXT", "ACDBATTRIBUTE"
                    ):
                        continue
                    if layer and str(e.Layer).lower() != layer.lower():
                        continue
                    old = str(e.TextString)
                    if find not in old:
                        continue
                    new = old.replace(find, replace)
                    e.TextString = new
                    changed.append({"handle": e.Handle, "before": old, "after": new})
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "find": find, "replace": replace,
                "replaced": len(changed), "entities": changed[:500],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"find_replace_text failed: {exc}")

    async def find_replace_attribute(
        self, tag: str, find: str, replace: str
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            changed = []
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() != "ACDBBLOCKREFERENCE":
                        continue
                    atts = e.GetAttributes()
                except Exception:
                    continue
                for a in atts:
                    try:
                        if str(a.TagString).lower() != tag.lower():
                            continue
                        old = str(a.TextString)
                        if find not in old:
                            continue
                        a.TextString = old.replace(find, replace)
                        changed.append({
                            "handle": e.Handle, "block": e.Name, "tag": a.TagString,
                            "before": old, "after": a.TextString,
                        })
                    except Exception:
                        continue
            return CommandResult(ok=True, payload={
                "tag": tag, "find": find, "replace": replace,
                "replaced": len(changed), "matches": changed[:500],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"find_replace_attribute failed: {exc}")

    # ==================================================================
    # Layouts / paper space
    # ==================================================================

    def _layout_by_name(self, name: str) -> Any:
        layouts = self._doc.Layouts
        for i in range(layouts.Count):
            lay = layouts.Item(i)
            if str(lay.Name).lower() == str(name).lower():
                return lay
        raise KeyError(name)

    async def layout_list(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            out = []
            active = str(self._safe_get_property(self._doc, "ActiveLayout", ""))
            try:
                active = str(self._doc.ActiveLayout.Name)
            except Exception:
                pass
            for i in range(self._doc.Layouts.Count):
                lay = self._doc.Layouts.Item(i)
                out.append({
                    "name": lay.Name,
                    "tab_order": self._safe_get_property(lay, "TabOrder"),
                    "active": str(lay.Name) == active,
                })
            return CommandResult(ok=True, payload={
                "count": len(out), "active": active, "layouts": out,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layout_list failed: {exc}")

    async def layout_create(self, name: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            lay = self._doc.Layouts.Add(str(name))
            return CommandResult(ok=True, payload={
                "created": lay.Name, "count": self._doc.Layouts.Count,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layout_create failed: {exc}")

    async def layout_switch(self, name: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            lay = self._layout_by_name(name)
            self._doc.ActiveLayout = lay
            return CommandResult(ok=True, payload={
                "active": str(self._doc.ActiveLayout.Name),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layout_switch failed: {exc}")

    async def layout_delete(self, name: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            before = self._doc.Layouts.Count
            self._layout_by_name(name).Delete()
            return CommandResult(ok=True, payload={
                "deleted": name, "before": before,
                "after": self._doc.Layouts.Count,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layout_delete failed: {exc}")

    async def layout_viewport_create(
        self, x: float, y: float, width: float, height: float, scale: float = 1.0
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            ps = self._doc.PaperSpace
            vp = ps.AddPViewport(
                self._to_variant_array((float(x), float(y), 0.0)),
                float(width), float(height),
            )
            try:
                vp.ViewHeight = float(height) / float(scale) if scale else float(height)
            except Exception:
                pass
            vp.DisplayLocked = False
            return CommandResult(ok=True, payload={
                "handle": vp.Handle, "center": [float(x), float(y)],
                "width": float(width), "height": float(height), "scale": float(scale),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layout_viewport_create failed: {exc}")

    async def layout_viewport_set_scale(
        self, viewport_id: str, scale: float
    ) -> CommandResult:
        """Set a viewport's scale.

        ``ViewHeight`` is read-only over COM in AutoCAD 2025, so the scale is
        applied through ``CustomScale`` and verified by reading the resulting
        view height back.
        """
        err = self._err()
        if err:
            return err
        try:
            vp = self._dyn(self._handle_object(viewport_id))
            paper_h = float(self._safe_get_property(vp, "Height") or 0.0)
            before = self._safe_get_property(vp, "ViewHeight")
            applied = None
            # CustomScale is paper:model, so 1:scale => 1/scale
            for name, value in (("CustomScale", 1.0 / float(scale)),):
                try:
                    setattr(vp, name, value)
                    applied = name
                    break
                except Exception:
                    continue
            if applied is None:
                try:
                    vp.StandardScale = int(scale)
                    applied = "StandardScale"
                except Exception as exc:
                    return CommandResult(
                        ok=False, error=f"layout_viewport_set_scale failed: {exc}"
                    )
            # Verify by reading CustomScale back (ViewHeight is not exposed to
            # COM in AutoCAD 2025, and is unreliable through LISP here).
            custom = self._safe_get_property(vp, "CustomScale")
            effective = round(1.0 / custom, 4) if custom else None
            verified = (
                effective is not None and abs(effective - float(scale)) < 1e-3
            )
            return CommandResult(ok=True, payload={
                "handle": viewport_id, "requested_scale": float(scale),
                "applied_via": applied,
                "custom_scale": custom,
                "effective_scale": effective,
                "verified": verified,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layout_viewport_set_scale failed: {exc}")

    async def layout_viewport_lock(
        self, viewport_id: str, lock: bool = True
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            vp = self._dyn(self._handle_object(viewport_id))
            vp.DisplayLocked = bool(lock)
            return CommandResult(ok=True, payload={
                "handle": viewport_id,
                "locked": bool(self._safe_get_property(vp, "DisplayLocked")),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layout_viewport_lock failed: {exc}")

    async def layout_page_setup(
        self, name: str, paper_size: str | None = None, orientation: str | None = None
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            lay = self._layout_by_name(name)
            changes = {}
            if paper_size:
                canonical = str(paper_size).lower().replace("x", "_x_").replace(" ", "_")
                for i in range(lay.GetCanonicalMediaNames().Count):
                    cand = lay.GetCanonicalMediaNames().Item(i)
                    if canonical in str(cand).lower():
                        lay.CanonicalMediaName = cand
                        changes["paper_size"] = cand
                        break
            if orientation:
                # acPlotOrientation: 0 = portrait, 1 = landscape
                val = 1 if str(orientation).lower().startswith("land") else 0
                lay.PlotOrientation = val
                changes["orientation"] = "landscape" if val else "portrait"
            return CommandResult(ok=True, payload={
                "layout": lay.Name, "changes": changes,
                "canonical_media": self._safe_get_property(lay, "CanonicalMediaName"),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layout_page_setup failed: {exc}")

    async def layout_titleblock_fill(
        self, layout_name: str, attributes: dict | None = None
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            lay = self._layout_by_name(layout_name)
            wanted = {str(k).lower(): v for k, v in (attributes or {}).items()}
            applied, missing = {}, set(wanted)
            for i in range(lay.Block.Count):
                try:
                    ref = lay.Block.Item(i)
                    if str(ref.EntityName).upper() != "ACDBBLOCKREFERENCE":
                        continue
                    for a in ref.GetAttributes():
                        tag = str(a.TagString).lower()
                        if tag in wanted:
                            a.TextString = str(wanted[tag])
                            applied[tag] = str(wanted[tag])
                            missing.discard(tag)
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "layout": layout_name, "applied": applied,
                "missing_tags": sorted(missing),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layout_titleblock_fill failed: {exc}")

    async def layout_batch_plot(
        self, layouts: list | None = None, output_path: str | None = None
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            names = [str(n) for n in layouts] if layouts else None
            out_dir = Path(output_path) if output_path else Path(tempfile.gettempdir())
            out_dir.mkdir(parents=True, exist_ok=True)
            plotted, errors = [], []
            for i in range(self._doc.Layouts.Count):
                lay = self._doc.Layouts.Item(i)
                name = str(lay.Name)
                if names and name not in names:
                    continue
                try:
                    target = out_dir / f"{name}.pdf"
                    self._doc.ActiveLayout = lay
                    plot = self._doc.Plot
                    plot.PlotToFile(str(target))
                    plotted.append({"layout": name, "path": str(target),
                                    "exists": target.exists()})
                except Exception as exc:
                    errors.append(f"{name}: {exc}")
            return CommandResult(ok=True, payload={
                "plotted": plotted, "errors": errors, "count": len(plotted),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"layout_batch_plot failed: {exc}")

    # ==================================================================
    # Equipment / tag placement
    # ==================================================================

    async def equipment_find(
        self, pattern: str, case_sensitive: bool = False,
        search_scope: str = "all", zoom_to_first: bool = True,
        zoom_height: float = 600.0, max_results: int = 50,
    ) -> CommandResult:
        """Find equipment tags by text or block attribute, with wildcards."""
        err = self._err()
        if err:
            return err
        try:
            import fnmatch
            import re

            rx = fnmatch.translate(pattern) if any(c in pattern for c in "*?") else re.escape(pattern)
            regex = re.compile(rx, 0 if case_sensitive else re.IGNORECASE)
            results = []
            for e in self._iter_msp():
                if len(results) >= max_results:
                    break
                try:
                    t = str(e.EntityName).upper()
                    if t in ("ACDBTEXT", "ACDBMTEXT"):
                        txt = str(e.TextString)
                        if regex.search(txt):
                            results.append({
                                "handle": e.Handle, "source": "text", "text": txt,
                                "layer": e.Layer,
                                "position": [round(float(v), 4) for v in e.InsertionPoint],
                            })
                    elif t == "ACDBBLOCKREFERENCE" and search_scope in ("all", "attributes"):
                        for a in e.GetAttributes():
                            val = str(a.TextString)
                            if regex.search(val):
                                results.append({
                                    "handle": e.Handle, "source": "attribute",
                                    "block": e.Name, "tag": a.TagString, "text": val,
                                    "layer": e.Layer,
                                    "position": [round(float(v), 4) for v in e.InsertionPoint],
                                })
                                break
                except Exception:
                    continue
            if results and zoom_to_first:
                try:
                    p = results[0]["position"]
                    self._doc.Application.ZoomCenter(
                        self._to_variant_array((p[0], p[1])), float(zoom_height)
                    )
                except Exception:
                    pass
            return CommandResult(ok=True, payload={
                "pattern": pattern, "count": len(results), "results": results,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"equipment_find failed: {exc}")

    async def equipment_inspect(
        self, x: float, y: float, view_width: float = 600.0,
        view_height: float = 600.0, infer_center: bool = True,
        handle: str | None = None,
    ) -> CommandResult:
        """Inspect a block reference near a point (or by handle)."""
        err = self._err()
        if err:
            return err
        try:
            target = None
            if handle:
                target = self._dyn(self._handle_object(handle))
            else:
                best = None
                for e in self._iter_msp():
                    try:
                        if str(e.EntityName).upper() != "ACDBBLOCKREFERENCE":
                            continue
                        bb = self._bbox_of(e)
                        if bb is None:
                            continue
                        (x0, y0, _), (x1, y1, _) = bb
                        if x0 <= float(x) <= x1 and y0 <= float(y) <= y1:
                            d = math.hypot((x0 + x1) / 2 - float(x), (y0 + y1) / 2 - float(y))
                            if best is None or d < best[0]:
                                best = (d, e)
                    except Exception:
                        continue
                target = best[1] if best else None
            if target is None:
                return CommandResult(ok=True, payload={
                    "found": False, "point": [float(x), float(y)],
                })
            attrs = {}
            try:
                for a in target.GetAttributes():
                    attrs[str(a.TagString)] = str(a.TextString)
            except Exception:
                pass
            return CommandResult(ok=True, payload={
                "found": True,
                "handle": target.Handle,
                "block": self._safe_get_property(target, "Name"),
                "layer": self._safe_get_property(target, "Layer"),
                "position": [round(float(v), 4) for v in target.InsertionPoint],
                "attributes": attrs,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"equipment_inspect failed: {exc}")

    async def place_equipment_tag(
        self, cx: float, cy: float, cz: float = 0.0, tag: str = "",
        cube_size: float = 24.0, direction: str = "right", text_height: float = 8.0,
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            half = float(cube_size) / 2.0
            box = self._msp.AddBox(
                self._to_variant_array((float(cx) - half, float(cy) - half, float(cz))),
                float(cube_size), float(cube_size), float(cube_size),
            )
            right = str(direction).lower() != "left"
            sign = 1.0 if right else -1.0
            tx = float(cx) + sign * (half + 2.0)
            note = self._msp.AddMText(
                self._to_variant_array((tx, float(cy), float(cz))),
                float(cube_size) * 2, str(tag),
            )
            note.Height = float(text_height)
            leader = self._msp.AddLine(
                self._to_variant_array((float(cx) + sign * half, float(cy), float(cz))),
                self._to_variant_array((tx, float(cy), float(cz))),
            )
            return CommandResult(ok=True, payload={
                "cube": box.Handle, "text": note.Handle, "leader": leader.Handle,
                "tag": tag, "center": [float(cx), float(cy), float(cz)],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"place_equipment_tag failed: {exc}")

    async def batch_find_and_tag(
        self, tags: list, cube_size: float = 24.0,
        direction: str = "right", text_height: float = 8.0,
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            placed, not_found = [], []
            for tag in tags or []:
                if not tag:
                    continue
                found = await self.find_text(
                    str(tag), case_sensitive=False, max_results=1, zoom_to_first=False
                )
                if not (found.ok and found.payload.get("count")):
                    not_found.append(str(tag))
                    continue
                pos = found.payload["matches"][0]["position"]
                res = await self.place_equipment_tag(
                    pos[0], pos[1], pos[2] if len(pos) > 2 else 0.0,
                    str(tag), cube_size, direction, text_height,
                )
                if res.ok:
                    entry = dict(res.payload)
                    entry["tag"] = str(tag)
                    placed.append(entry)
                else:
                    not_found.append(f"{tag} (place failed: {res.error})")
            return CommandResult(ok=True, payload={
                "placed": placed, "placed_count": len(placed),
                "not_found": not_found,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"batch_find_and_tag failed: {exc}")

    # ==================================================================
    # Natural-language dispatch
    # ==================================================================

    async def execute_natural_command(self, command: str) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            from autocad_mcp.nlp.processor import NLPProcessor

            parsed = NLPProcessor(strict_mode=False).parse_command(command)
            if parsed.confidence < 0.2:
                return CommandResult(
                    ok=False,
                    error=f"Could not parse command: {command}",
                    payload={"operation": parsed.operation,
                             "confidence": parsed.confidence},
                )
            p = parsed.parameters
            op = parsed.operation
            if op == "create_line":
                result = await self.create_line(
                    p.get("x1", 0), p.get("y1", 0), p.get("x2", 100),
                    p.get("y2", 100), layer=p.get("layer"),
                )
            elif op == "create_circle":
                result = await self.create_circle(
                    p.get("cx", 0), p.get("cy", 0), p.get("radius", 50),
                    layer=p.get("layer"),
                )
            elif op == "create_polyline":
                result = await self.create_polyline(
                    p.get("points", []), p.get("closed", False), layer=p.get("layer")
                )
            elif op == "create_rectangle":
                result = await self.create_rectangle(
                    p.get("x1", 0), p.get("y1", 0), p.get("x2", 100),
                    p.get("y2", 100), layer=p.get("layer"),
                )
            elif op == "layer_create":
                result = await self.layer_create(
                    p.get("name", "NewLayer"), color=p.get("color")
                )
            else:
                result = CommandResult(
                    ok=False,
                    error=f"Parsed as {op!r} but no COM handler is wired for it",
                    payload={"parsed": {"operation": op, "parameters": p}},
                )
            return result
        except Exception as exc:
            return CommandResult(ok=False, error=f"execute_natural_command failed: {exc}")

    # ==================================================================
    # Electrical (design aids — reference values, not a code substitute)
    # ==================================================================

    # Copper conductor resistance, ohms per 1000 ft @ 75 C
    _R_PER_KFT = {
        "14": 3.07, "12": 1.93, "10": 1.21, "8": 0.764, "6": 0.491,
        "4": 0.308, "3": 0.245, "2": 0.194, "1": 0.154, "1/0": 0.122,
        "2/0": 0.0967, "3/0": 0.0766, "4/0": 0.0608,
    }
    # THHN copper ampacity, 90 C column
    _AMPACITY_THHN = {
        "14": 25, "12": 30, "10": 35, "8": 50, "6": 65, "4": 85, "3": 100,
        "2": 115, "1": 130, "1/0": 150, "2/0": 175, "3/0": 200, "4/0": 230,
    }
    # THHN conductor area, square inches
    _WIRE_AREA_IN2 = {
        "14": 0.0097, "12": 0.0133, "10": 0.0211, "8": 0.0366, "6": 0.0507,
        "4": 0.0824, "3": 0.0973, "2": 0.1158, "1": 0.1562, "1/0": 0.1855,
        "2/0": 0.2223, "3/0": 0.2679, "4/0": 0.3237,
    }
    # EMT internal area at 40% fill, square inches
    _CONDUIT_40PCT = {
        "1/2": 0.122, "3/4": 0.213, "1": 0.346, "1-1/4": 0.598,
        "1-1/2": 0.814, "2": 1.342, "2-1/2": 1.948, "3": 2.907,
        "3-1/2": 3.895, "4": 5.153,
    }

    async def electrical_nec_lookup(
        self, table: str, parameters: dict | None = None
    ) -> CommandResult:
        params = parameters or {}
        try:
            key = str(table).lower()
            if "310" in key or "ampacity" in key:
                gauge = str(params.get("gauge", params.get("size", ""))).lower()
                if gauge:
                    if gauge not in self._AMPACITY_THHN:
                        return CommandResult(
                            ok=False,
                            error=f"Unknown gauge {gauge!r}; known: {sorted(self._AMPACITY_THHN)}",
                        )
                    return CommandResult(ok=True, payload={
                        "table": "310.16", "gauge": gauge,
                        "ampacity_thhn_90c": self._AMPACITY_THHN[gauge],
                    })
                return CommandResult(ok=True, payload={
                    "table": "310.16", "ampacity_thhn_90c": self._AMPACITY_THHN,
                })
            if "conduit" in key or "chapter 9" in key or "fill" in key:
                return CommandResult(ok=True, payload={
                    "table": "Chapter 9 Table 4 (EMT, 40% fill)",
                    "conduit_40pct_in2": self._CONDUIT_40PCT,
                    "conductor_area_in2": self._WIRE_AREA_IN2,
                })
            return CommandResult(
                ok=False,
                error=f"Unknown NEC table {table!r}. Try '310.16' or 'conduit fill'.",
            )
        except Exception as exc:
            return CommandResult(ok=False, error=f"electrical_nec_lookup failed: {exc}")

    async def electrical_voltage_drop(
        self, voltage: float, current: float, wire_gauge: str, length: float,
        phase: int = 1, power_factor: float = 1.0,
    ) -> CommandResult:
        try:
            gauge = str(wire_gauge).lower()
            if gauge not in self._R_PER_KFT:
                return CommandResult(
                    ok=False,
                    error=f"Unknown gauge {wire_gauge!r}; known: {sorted(self._R_PER_KFT)}",
                )
            r = self._R_PER_KFT[gauge]
            L = float(length)
            I = float(current)
            if int(phase) == 3:
                vd = 1.732 * I * r * L / 1000.0 * float(power_factor)
            else:
                vd = 2.0 * I * r * L / 1000.0 * float(power_factor)
            v = float(voltage) or 1.0
            pct = vd / v * 100.0
            return CommandResult(ok=True, payload={
                "voltage": float(voltage), "current": I, "gauge": gauge,
                "length_ft": L, "phase": int(phase),
                "voltage_drop": round(vd, 3),
                "voltage_drop_pct": round(pct, 3),
                "within_3pct": pct <= 3.0,
                "end_voltage": round(v - vd, 3),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"electrical_voltage_drop failed: {exc}")

    async def electrical_conduit_fill(
        self, conduit_size: str, conduit_type: str = "EMT",
        wire_gauges: list | None = None,
    ) -> CommandResult:
        try:
            size = str(conduit_size).replace('"', "").strip()
            capacity = self._CONDUIT_40PCT.get(size)
            if capacity is None:
                return CommandResult(
                    ok=False,
                    error=f"Unknown conduit size {conduit_size!r}; known: {sorted(self._CONDUIT_40PCT)}",
                )
            gauges = [str(g).lower() for g in (wire_gauges or [])]
            used = 0.0
            detail = []
            for g in gauges:
                area = self._WIRE_AREA_IN2.get(g)
                if area is None:
                    return CommandResult(ok=False, error=f"Unknown wire gauge {g!r}")
                used += area
                detail.append({"gauge": g, "area_in2": area})
            return CommandResult(ok=True, payload={
                "conduit_size": size, "conduit_type": conduit_type,
                "capacity_40pct_in2": capacity,
                "conductor_area_in2": round(used, 4),
                "fill_pct": round(used / capacity * 100.0, 2) if capacity else None,
                "compliant": used <= capacity,
                "conductors": detail,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"electrical_conduit_fill failed: {exc}")

    async def electrical_load_calc(self, devices: list | None = None) -> CommandResult:
        try:
            total_va = 0.0
            rows = []
            for d in devices or []:
                qty = float(d.get("quantity", 1) or 1)
                va = float(d.get("va", d.get("watts", 0)) or 0)
                total_va += qty * va
                rows.append({
                    "device": d.get("name", d.get("type", "device")),
                    "quantity": qty, "va_each": va, "va_total": qty * va,
                })
            return CommandResult(ok=True, payload={
                "total_va": round(total_va, 2),
                "total_kva": round(total_va / 1000.0, 3),
                "at_120v_amps": round(total_va / 120.0, 2) if total_va else 0,
                "at_208v_3ph_amps": round(total_va / (208 * 1.732), 2) if total_va else 0,
                "devices": rows,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"electrical_load_calc failed: {exc}")

    async def electrical_panel_schedule_gen(
        self, panel_block: str, x: float = 0, y: float = 0
    ) -> CommandResult:
        """Generate a simple panel schedule as text next to a panel block."""
        err = self._err()
        if err:
            return err
        try:
            circuits = []
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() not in ("ACDBTEXT", "ACDBMTEXT"):
                        continue
                    txt = str(e.TextString).strip()
                    if txt.upper().startswith(("CKT", "CIRCUIT", "P-", "PANEL")):
                        circuits.append({"text": txt, "handle": e.Handle,
                                         "layer": e.Layer})
                except Exception:
                    continue
            body = "\\P".join(c["text"] for c in circuits) or "No circuits found"
            note = self._msp.AddMText(
                self._to_variant_array((float(x), float(y), 0.0)), 120.0,
                f"{panel_block}\\P{body}",
            )
            note.Height = 2.5
            return CommandResult(ok=True, payload={
                "panel": panel_block, "schedule_handle": note.Handle,
                "circuits": len(circuits),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"electrical_panel_schedule_gen failed: {exc}")

    async def electrical_symbol_insert(
        self, symbol_type: str, x: float, y: float, scale: float = 1.0,
        rotation: float = 0.0, layer: str | None = None,
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            ref = self._msp.InsertBlock(
                self._to_variant_array((float(x), float(y), 0.0)),
                symbol_type, float(scale), float(scale), float(scale), float(rotation),
            )
            self._apply_properties(ref, layer=layer)
            return CommandResult(ok=True, payload={
                "handle": ref.Handle, "symbol": symbol_type,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"electrical_symbol_insert failed: {exc}")

    async def electrical_wire_number_assign(
        self, layer: str, prefix: str = "W", start_num: int = 1
    ) -> CommandResult:
        """Label each line on *layer* with a sequential wire number."""
        err = self._err()
        if err:
            return err
        try:
            n = int(start_num)
            labelled = []
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() != "ACDBLINE":
                        continue
                    if str(e.Layer).lower() != str(layer).lower():
                        continue
                    mid = ((e.StartPoint[0] + e.EndPoint[0]) / 2,
                           (e.StartPoint[1] + e.EndPoint[1]) / 2)
                    label = self._msp.AddText(
                        f"{prefix}{n}", self._to_variant_array((mid[0], mid[1], 0.0)), 0.1
                    )
                    labelled.append({"line": e.Handle, "text": label.Handle, "number": f"{prefix}{n}"})
                    n += 1
                except Exception:
                    continue
            return CommandResult(ok=True, payload={
                "layer": layer, "labelled": len(labelled),
                "last_number": f"{prefix}{n - 1}" if labelled else None,
                "items": labelled[:500],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"electrical_wire_number_assign failed: {exc}")

    async def electrical_circuit_trace(
        self, start_entity: str, layer: str | None = None
    ) -> CommandResult:
        """Walk connected line endpoints from a starting entity."""
        err = self._err()
        if err:
            return err
        try:
            tol = 0.01
            key = lambda p: (round(float(p[0]) / tol), round(float(p[1]) / tol))  # noqa: E731
            adjacency: defaultdict[tuple, list] = defaultdict(list)
            for e in self._iter_msp():
                try:
                    if str(e.EntityName).upper() != "ACDBLINE":
                        continue
                    if layer and str(e.Layer).lower() != layer.lower():
                        continue
                    a, b = key(e.StartPoint), key(e.EndPoint)
                    adjacency[a].append((e.Handle, b))
                    adjacency[b].append((e.Handle, a))
                except Exception:
                    continue
            start = self._dyn(self._handle_object(start_entity))
            seen = {start.Handle}
            frontier = [key(start.StartPoint), key(start.EndPoint)]
            order = [start.Handle]
            while frontier:
                node = frontier.pop()
                for handle, other in adjacency.get(node, []):
                    if handle in seen:
                        continue
                    seen.add(handle)
                    order.append(handle)
                    frontier.append(other)
            return CommandResult(ok=True, payload={
                "start": start_entity, "circuit_size": len(order),
                "handles": order, "layer": layer,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"electrical_circuit_trace failed: {exc}")

    # ==================================================================
    # P&ID
    # ==================================================================

    _PID_LAYERS = {
        "process": "P-PROC", "instrument": "P-INST", "equipment": "P-EQUIP",
        "text": "P-TEXT", "flow": "P-FLOW",
    }

    async def pid_setup_layers(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            created, existing = [], []
            for name in self._PID_LAYERS.values():
                try:
                    self._doc.Layers.Item(name)
                    existing.append(name)
                except Exception:
                    self._doc.Layers.Add(name)
                    created.append(name)
            return CommandResult(ok=True, payload={
                "created": created, "existing": existing,
                "layers": list(self._PID_LAYERS.values()),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"pid_setup_layers failed: {exc}")

    async def pid_list_symbols(self, category: str) -> CommandResult:
        try:
            from autocad_mcp.pid.cto_library import list_categories, list_symbols

            symbols = list_symbols(category)
            return CommandResult(ok=True, payload={
                "category": category, "symbols": symbols, "count": len(symbols),
                "categories": list_categories(),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"pid_list_symbols failed: {exc}")

    async def pid_insert_symbol(
        self, category: str, symbol: str, x: float, y: float,
        scale: float = 1.0, rotation: float = 0.0,
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            from autocad_mcp.pid.cto_library import symbol_path

            src = symbol_path(category, symbol)
            if not Path(src).exists():
                return CommandResult(
                    ok=False, error=f"Symbol file not found: {src}"
                )
            before = {str(self._msp.Item(i).Handle) for i in range(self._msp.Count)}
            await self._send_command(
                f"_.-INSERT\n{Path(src).as_posix()}\n"
                f"{float(x)},{float(y)}\n{float(scale)}\n{float(rotation)}\n",
                timeout=180.0,
            )
            after = [str(self._msp.Item(i).Handle) for i in range(self._msp.Count)]
            new = [h for h in after if h not in before]
            return CommandResult(ok=True, payload={
                "category": category, "symbol": symbol,
                "inserted": new, "source": str(src),
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"pid_insert_symbol failed: {exc}")

    async def _pid_place(
        self, category: str, symbol: str, x: float, y: float,
        scale: float = 1.0, rotation: float = 0.0, layer: str | None = None,
        attributes: dict | None = None,
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            from autocad_mcp.pid.cto_library import symbol_path

            src = Path(symbol_path(category, symbol))
            if not src.exists():
                return CommandResult(ok=False, error=f"Symbol not found: {src}")
            try:
                self._doc.Layers.Item(layer or self._PID_LAYERS["equipment"])
            except Exception:
                self._doc.Layers.Add(layer or self._PID_LAYERS["equipment"])
            ref = self._msp.InsertBlock(
                self._to_variant_array((float(x), float(y), 0.0)),
                src.stem, float(scale), float(scale), float(scale), float(rotation),
            )
            self._apply_properties(ref, layer=layer or self._PID_LAYERS["equipment"])
            applied = {}
            if attributes:
                try:
                    by_tag = {str(a.TagString).lower(): a for a in ref.GetAttributes()}
                    for k, v in attributes.items():
                        a = by_tag.get(str(k).lower())
                        if a is not None:
                            a.TextString = str(v)
                            applied[k] = str(v)
                except Exception:
                    pass
            return CommandResult(ok=True, payload={
                "handle": ref.Handle, "category": category, "symbol": symbol,
                "attributes_applied": applied,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"pid insert failed: {exc}")

    async def pid_insert_valve(
        self, x: float, y: float, valve_type: str, rotation: float = 0.0,
        attributes: dict | None = None,
    ) -> CommandResult:
        return await self._pid_place("valves", valve_type, x, y, 1.0, rotation,
                                     self._PID_LAYERS["process"], attributes)

    async def pid_insert_pump(
        self, x: float, y: float, pump_type: str, rotation: float = 0.0,
        attributes: dict | None = None,
    ) -> CommandResult:
        return await self._pid_place("pumps", pump_type, x, y, 1.0, rotation,
                                     self._PID_LAYERS["equipment"], attributes)

    async def pid_insert_tank(
        self, x: float, y: float, tank_type: str, scale: float = 1.0,
        attributes: dict | None = None,
    ) -> CommandResult:
        return await self._pid_place("tanks", tank_type, x, y, scale, 0.0,
                                     self._PID_LAYERS["equipment"], attributes)

    async def pid_insert_instrument(
        self, x: float, y: float, instrument_type: str, rotation: float = 0.0,
        tag_id: str = "", range_value: str = "",
    ) -> CommandResult:
        attrs = {"TAG": tag_id, "RANGE": range_value} if (tag_id or range_value) else None
        return await self._pid_place("instruments", instrument_type, x, y, 1.0,
                                     rotation, self._PID_LAYERS["instrument"], attrs)

    async def pid_draw_process_line(
        self, x1: float, y1: float, x2: float, y2: float
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            try:
                self._doc.Layers.Item(self._PID_LAYERS["process"])
            except Exception:
                self._doc.Layers.Add(self._PID_LAYERS["process"])
            line = self._msp.AddLine(
                self._to_variant_array((float(x1), float(y1), 0.0)),
                self._to_variant_array((float(x2), float(y2), 0.0)),
            )
            line.Layer = self._PID_LAYERS["process"]
            return CommandResult(ok=True, payload={
                "handle": line.Handle, "layer": self._PID_LAYERS["process"],
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"pid_draw_process_line failed: {exc}")

    async def pid_connect_equipment(
        self, x1: float, y1: float, x2: float, y2: float
    ) -> CommandResult:
        res = await self.pid_draw_process_line(x1, y1, x2, y2)
        if res.ok:
            res.payload["connected"] = True
        return res

    async def pid_add_flow_arrow(
        self, x: float, y: float, rotation: float = 0.0
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            size = 2.0
            pts = [(-size, -size / 2), (0.0, 0.0), (-size, size / 2)]
            flat = []
            for px, py in pts:
                flat.extend([float(x) + px, float(y) + py])
            pl = self._msp.AddLightWeightPolyline(self._dbl_array(flat))
            try:
                pl.Layer = self._PID_LAYERS["flow"]
            except Exception:
                pass
            if rotation:
                pl.Rotate(self._to_variant_array((float(x), float(y), 0.0)),
                          math.radians(float(rotation)))
            return CommandResult(ok=True, payload={"handle": pl.Handle})
        except Exception as exc:
            return CommandResult(ok=False, error=f"pid_add_flow_arrow failed: {exc}")

    async def pid_add_equipment_tag(
        self, x: float, y: float, tag: str, description: str = ""
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            text = f"{tag}\\P{description}" if description else tag
            note = self._msp.AddMText(
                self._to_variant_array((float(x), float(y), 0.0)), 40.0, text
            )
            note.Height = 2.5
            try:
                note.Layer = self._PID_LAYERS["text"]
            except Exception:
                pass
            return CommandResult(ok=True, payload={"handle": note.Handle, "tag": tag})
        except Exception as exc:
            return CommandResult(ok=False, error=f"pid_add_equipment_tag failed: {exc}")

    async def pid_add_line_number(
        self, x: float, y: float, line_num: str, spec: str
    ) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            note = self._msp.AddText(
                f"{line_num} {spec}".strip(),
                self._to_variant_array((float(x), float(y), 0.0)), 2.0,
            )
            try:
                note.Layer = self._PID_LAYERS["text"]
            except Exception:
                pass
            return CommandResult(ok=True, payload={"handle": note.Handle})
        except Exception as exc:
            return CommandResult(ok=False, error=f"pid_add_line_number failed: {exc}")

    # ==================================================================
    # MagiCAD (plugin commands driven through the command line)
    # ==================================================================

    async def magicad_status(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            probe = await self._lisp_eval(
                '(if (member "MagiCAD" (mapcar (quote strcase) (arx))) "loaded" "not-loaded")',
                timeout=60.0,
            )
            arx = await self._lisp_eval("(arx)", timeout=60.0)
            loaded = None
            if arx.ok:
                loaded = [p for p in str(arx.payload.get("result", "")).split()
                          if "magicad" in p.lower()]
            return CommandResult(ok=True, payload={
                "backend": "com",
                "probe": probe.payload if probe.ok else probe.error,
                "magicad_arx": loaded or [],
                "note": "MagiCAD commands are dispatched via the AutoCAD command line",
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"magicad_status failed: {exc}")

    async def magicad_run(self, command: str, args: str | None = None) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            cmd = f"_{command}" + (f" {args}" if args else "")
            await self._send_command(f"{cmd}\n", timeout=600.0)
            return CommandResult(ok=True, payload={
                "command": command, "args": args or "",
                "sent": cmd,
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"magicad_run failed: {exc}")

    async def magicad_list_commands(self) -> CommandResult:
        return CommandResult(ok=True, payload={
            "known": [
                "MAGIADDDEVICE", "MAGIADDPIPE", "MAGIADDUCT", "MAGICABLE",
                "MAGISECTION", "MAGIUPDATEDRAWING", "MAGIIFCEXPORT",
                "MAGISPRINKLERCALC", "MAGIVIEWMODE", "MAGISTOREY",
            ],
            "note": "Pass any MagiCAD command name to magicad_run",
        })

    async def magicad_update_drawing(self, flags: str | None = None) -> CommandResult:
        return await self.magicad_run(
            "MAGIUPDATEDRAWING", flags or "1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1"
        )

    async def magicad_cleanup(self, options: str | None = None) -> CommandResult:
        return await self.magicad_run("MAGICLEANUP", options or "Y Y Y Y 8")

    async def magicad_ifc_export(self, mode: str = "current") -> CommandResult:
        return await self.magicad_run("MAGIIFCEXPORT", str(mode))

    async def magicad_view_mode(self, mode: str, type: str = "D") -> CommandResult:
        return await self.magicad_run("MAGIVIEWMODE", f"{mode} {type}")

    async def magicad_change_storey(self, storey: str) -> CommandResult:
        return await self.magicad_run("MAGISTOREY", str(storey))

    async def magicad_section_update(self) -> CommandResult:
        return await self.magicad_run("MAGISECTION")

    async def magicad_fix_errors(self) -> CommandResult:
        return await self.magicad_run("MAGIFIXERRORS")

    async def magicad_show_all(self) -> CommandResult:
        return await self.magicad_run("MAGISHOWALL")

    async def magicad_clear_garbage(self) -> CommandResult:
        return await self.magicad_run("MAGICLEARGARBAGE")

    async def magicad_disconnect_project(self) -> CommandResult:
        return await self.magicad_run("MAGIDISCONNECT")

    async def magicad_project_info(self) -> CommandResult:
        err = self._err()
        if err:
            return err
        try:
            layers = []
            for i in range(self._doc.Layers.Count):
                name = str(self._doc.Layers.Item(i).Name)
                if name.upper().startswith(("MAGI", "MEP", "VVS")):
                    layers.append(name)
            return CommandResult(ok=True, payload={
                "drawing": self._doc.Name,
                "magi_layers": layers,
                "layer_count": len(layers),
                "note": "Full project metadata lives in the MagiCAD project database",
            })
        except Exception as exc:
            return CommandResult(ok=False, error=f"magicad_project_info failed: {exc}")
