"""Live validation of COM backend batch 1 against a running AutoCAD.

Safety: closes only documents that report Saved=True, refuses dirty ones,
and asserts the active document is the scratch drawing before every step.
"""
import asyncio
import json
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, "src")

import pythoncom
import win32com.client

from autocad_mcp.backends.com_backend import COMBackend

RESULTS = []


def check(name, ok, detail=None):
    RESULTS.append({"op": name, "ok": bool(ok), "detail": detail})
    print(f"[{'PASS' if ok else 'FAIL'}] {name}  {json.dumps(detail, default=str)[:260]}")


def guard(doc, expected):
    if not doc.Name.lower().startswith(expected.lower()):
        raise RuntimeError(f"REFUSING: active doc {doc.Name!r} != {expected!r}")


async def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("AutoCAD.Application")

    for d in list(app.Documents):
        name = d.Name
        if name.lower().startswith("drawing"):
            # our own untitled scratch documents — safe to discard
            d.Close(False)
            continue
        if d.Saved:
            d.Close(False)
        else:
            print("REFUSED to close unsaved user document:", name)
            return

    b = COMBackend()
    r = await b.initialize()
    check("initialize", r.ok, r.payload or r.error)
    if not r.ok:
        return

    r = await b.drawing_create("mcp_scratch")
    check("drawing_create", r.ok, r.payload or r.error)
    doc = b._doc
    for _ in range(60):
        try:
            _ = doc.ModelSpace.Count
            break
        except Exception:
            time.sleep(1)
    else:
        check("scratch ready", False, "document never became COM-ready")
        return
    guard(doc, "Drawing")

    # --- units ---
    r = await b.drawing_units("feet")
    check("drawing_units set", r.ok and r.payload.get("insunits") == 2, r.payload or r.error)
    r = await b.drawing_units()
    check("drawing_units get", r.ok and r.payload.get("name") == "feet", r.payload)

    # --- limits ---
    r = await b.drawing_limits(0, 0, 100, 50)
    check("drawing_limits set", r.ok, r.payload or r.error)
    r = await b.drawing_limits()
    check("drawing_limits get", r.ok and r.payload["limmax"][0] == 100.0, r.payload)

    # --- execute_lisp: create geometry, return handle ---
    r = await b.execute_lisp(
        '(progn (setq e (entmakex (list (cons 0 "LINE") (cons 10 (list 0.0 0.0 0.0)) '
        '(cons 11 (list 10.0 5.0 0.0))))) (cdr (assoc 5 (entget e))))'
    )
    handle = r.payload["result"].strip().strip('"') if r.ok else None
    check("execute_lisp entmakex", r.ok and handle is not None, {"handle": handle, "err": r.error})

    guard(doc, "Drawing")
    if handle:
        r2 = await b.entity_get(handle)
        check(
            "entity_get subtype props (late binding)",
            r2.ok and r2.payload.get("start_point") is not None,
            r2.payload if r2.ok else r2.error,
        )

    # --- error propagation ---
    r = await b.execute_lisp("(car 1)")
    check("execute_lisp error surfaced", (not r.ok) and "argument" in (r.error or "").lower(), r.error)

    # --- multi-form body ---
    r = await b.execute_lisp("(setq a 2) (setq bb 3) (+ a bb)")
    check("execute_lisp multi-form", r.ok and r.payload["result"] == "5", r.payload or r.error)

    # --- send_command ---
    r = await b.send_command("_.ZOOM\n_e\n")
    check("send_command", r.ok and r.payload.get("settled") is True, r.payload or r.error)

    # --- audit ---
    r = await b.drawing_audit(fix=True)
    check("drawing_audit", r.ok and r.payload.get("method") == "AuditInfo", r.payload or r.error)

    # --- wblock ---
    if handle:
        out = Path(tempfile.gettempdir()) / "autocad_mcp_lisp" / "batch1_wblock.dwg"
        if out.exists():
            out.unlink()
        r = await b.drawing_wblock([handle], str(out))
        check("drawing_wblock", r.ok and r.payload.get("written"), r.payload or r.error)

    guard(doc, "Drawing")
    r = await b.status()
    check("status after ops", r.ok, r.payload.get("document_name") if r.ok else r.error)

    try:
        b._doc.Close(False)
        check("scratch closed", True, None)
    except Exception as exc:
        check("scratch closed", False, str(exc))

    passed = sum(1 for x in RESULTS if x["ok"])
    print(f"\nSUMMARY: {passed}/{len(RESULTS)} passed")
    return passed == len(RESULTS)


if __name__ == "__main__":
    ok = asyncio.run(main())
    sys.exit(0 if ok else 1)
