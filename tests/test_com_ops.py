"""Unit tests for COM backend operations that were previously unimplemented.

These run without AutoCAD: a fake COM object graph stands in for the
Application / Document / ModelSpace / Layers / Blocks collections.  The fakes
record calls so we can assert on what the backend actually asked AutoCAD to do.
"""

from __future__ import annotations

import asyncio
import math

import pytest

from autocad_mcp.backends.com_ops import COMOperationsMixin


# ---------------------------------------------------------------------------
# Fake COM object graph
# ---------------------------------------------------------------------------


class FakeCollection(list):
    """Minimal COM collection: supports Item() and Count."""

    @property
    def Count(self) -> int:
        return len(self)

    def Item(self, key):
        if isinstance(key, int):
            return self[key]
        for item in self:
            if getattr(item, "Name", None) == key:
                return item
        raise KeyError(key)

    def Add(self, *args, **kwargs):
        obj = FakeNamed(args[0] if args else "new")
        self.append(obj)
        return obj

    def Delete(self):  # pragma: no cover - only used via SelectionSet
        pass


class FakeNamed:
    def __init__(self, name: str):
        self.Name = name


class FakeEntity:
    def __init__(self, handle: str, entity_name: str = "AcDbLine", layer: str = "0"):
        self.Handle = handle
        self.EntityName = entity_name
        self.Layer = layer
        self.color = 256
        self.Linetype = "ByLayer"
        self.Lineweight = -1

    def GetBoundingBox(self):
        return ((0.0, 0.0, 0.0), (1.0, 1.0, 0.0))

    def Delete(self):
        self._deleted = True


class FakeSelectionSets(list):
    def Add(self, name):
        ss = FakeSelectionSet(name)
        self.append(ss)
        return ss


class FakeSelectionSet:
    def __init__(self, name):
        self.Name = name
        self.items = []
        self.deleted = False

    def AddItems(self, arr):
        # AutoCAD takes a SAFEARRAY of IDispatch here, not a list.
        items = getattr(arr, "value", arr)
        self.items.extend(list(items))

    def Delete(self):
        self.deleted = True


class FakeModelSpace(FakeCollection):
    pass


class FakeDocument:
    def __init__(self, entities=None):
        self.Name = "test.dwg"
        self.FullName = "C:/tmp/test.dwg"
        self.ReadOnly = False
        self.Saved = True
        self.ModelSpace = FakeModelSpace(entities or [])
        self.Layers = FakeCollection([FakeNamed("0")])
        self.Blocks = FakeCollection()
        self.SelectionSets = FakeSelectionSets()
        self.vars = {
            "INSUNITS": 0,
            "MEASUREMENT": 0,
            "CMDNAMES": "",
            "ERRNO": 0,
            "LIMMIN": (0.0, 0.0, 0.0),
            "LIMMAX": (12.0, 9.0, 0.0),
        }
        self.sent_commands: list[str] = []
        self.wblock_calls: list[tuple] = []
        self.saved_path = None
        self.audit_errors = False

    def AuditInfo(self, fix):
        self.audit_fixed = fix
        return self.audit_errors

    def GetVariable(self, name):
        return self.vars.get(name, 0)

    def SetVariable(self, name, value):
        self.vars[name] = value

    def HandleToObject(self, handle):
        for e in self.ModelSpace:
            if e.Handle == str(handle):
                return e
        raise KeyError(handle)

    def SendCommand(self, text):
        self.sent_commands.append(text)

    def Wblock(self, path, sset):
        self.wblock_calls.append((path, [i.Handle for i in sset.items]))

    def Regen(self, mode=1):
        pass


class FakeApp:
    def __init__(self, doc):
        self.Name = "AutoCAD"
        self.Version = "25.0"
        self.ActiveDocument = doc
        self.Documents = FakeCollection([doc])


class Harness(COMOperationsMixin):
    """COMOperationsMixin with the bits of COMBackend it depends on."""

    def __init__(self, doc):
        self._app = FakeApp(doc)
        self._doc = doc
        self._msp = doc.ModelSpace

    def _validate_connection(self):
        return None

    @staticmethod
    def _safe_get_property(entity, prop_name, default=None):
        try:
            return getattr(entity, prop_name)
        except Exception:
            return default

    @staticmethod
    def _to_variant_array(point):
        return tuple(point)


def run(coro):
    return asyncio.run(coro)


# ---------------------------------------------------------------------------
# Helpers / escape hatches
# ---------------------------------------------------------------------------


def test_send_command_records_and_reports():
    doc = FakeDocument()
    h = Harness(doc)
    result = run(h.send_command("_.ZOOM\n_e\n"))
    assert result.ok
    assert doc.sent_commands == ["_.ZOOM\n_e\n"]
    assert result.payload["settled"] is True


def test_send_command_rejects_empty():
    h = Harness(FakeDocument())
    assert not run(h.send_command("   ")).ok


def test_lisp_eval_writes_wrapper_and_returns_result(tmp_path, monkeypatch):
    """execute_lisp must load a temp .lsp and surface the printed result."""
    doc = FakeDocument()
    h = Harness(doc)

    monkeypatch.setattr("autocad_mcp.backends.com_ops._LISP_DIR", tmp_path)

    # Simulate AutoCAD executing the wrapper: capture the load command, then
    # write the result file the footer would have produced.
    def fake_send(text):
        doc.sent_commands.append(text)
        lsp = text.split('"')[1]
        out = lsp.replace(".lsp", ".out")
        tmp_path.joinpath(out).write_text('"hello"', encoding="utf-8")

    doc.SendCommand = fake_send

    result = run(h.execute_lisp('(strcat "he" "llo")'))
    assert result.ok, result.error
    assert result.payload["result"] == '"hello"'
    assert "(load " in doc.sent_commands[0]


def test_lisp_eval_surfaces_error(tmp_path, monkeypatch):
    doc = FakeDocument()
    h = Harness(doc)
    monkeypatch.setattr("autocad_mcp.backends.com_ops._LISP_DIR", tmp_path)

    def fake_send(text):
        lsp = text.split('"')[1]
        tmp_path.joinpath(lsp.replace(".lsp", ".out")).write_text(
            "ERROR: bad argument type", encoding="utf-8"
        )

    doc.SendCommand = fake_send
    result = run(h.execute_lisp("(car 1)"))
    assert not result.ok
    assert "bad argument type" in result.error


# ---------------------------------------------------------------------------
# Drawing-level operations
# ---------------------------------------------------------------------------


def test_drawing_units_get():
    doc = FakeDocument()
    doc.vars["INSUNITS"] = 2
    result = run(Harness(doc).drawing_units())
    assert result.ok
    assert result.payload["name"] == "feet"


@pytest.mark.parametrize("value,expected", [
    ("feet", 2), ("ft", 2), ("inches", 1), ("mm", 4), ("meters", 6), (4, 4),
])
def test_drawing_units_set(value, expected):
    doc = FakeDocument()
    result = run(Harness(doc).drawing_units(value))
    assert result.ok, result.error
    assert doc.vars["INSUNITS"] == expected


def test_drawing_units_rejects_unknown():
    result = run(Harness(FakeDocument()).drawing_units("furlongs"))
    assert not result.ok
    assert "Unknown unit" in result.error


def test_drawing_limits_roundtrip():
    doc = FakeDocument()
    h = Harness(doc)
    assert run(h.drawing_limits()).payload["limmax"][0] == 12.0
    result = run(h.drawing_limits(0, 0, 100, 50))
    assert result.ok
    assert doc.vars["LIMMAX"] == (100.0, 50.0)


def test_drawing_audit_prefers_auditinfo():
    doc = FakeDocument()
    result = run(Harness(doc).drawing_audit(fix=True))
    assert result.ok
    assert result.payload["method"] == "AuditInfo"
    assert doc.audit_fixed is True
    assert doc.sent_commands == []  # no command-line round trip


def test_drawing_audit_falls_back_to_command():
    class NoAudit(FakeDocument):
        def AuditInfo(self, fix):
            raise AttributeError("AuditInfo unavailable")

    d = NoAudit()
    result = run(Harness(d).drawing_audit(fix=False))
    assert result.ok
    assert result.payload["method"] == "AUDIT command"
    assert d.sent_commands == ["_.AUDIT\n_n\n"]


def test_drawing_wblock_builds_selection_set_and_cleans_up(tmp_path):
    doc = FakeDocument([FakeEntity("A1"), FakeEntity("A2")])
    h = Harness(doc)
    out = tmp_path / "out.dwg"
    result = run(h.drawing_wblock(["A1", "A2"], str(out)))
    assert result.ok, result.error
    assert doc.wblock_calls[0][1] == ["A1", "A2"]
    # selection set must be deleted even on success
    assert all(ss.deleted for ss in doc.SelectionSets)


def test_drawing_wblock_requires_handles():
    assert not run(Harness(FakeDocument()).drawing_wblock([], "x.dwg")).ok
