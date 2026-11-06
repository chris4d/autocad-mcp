"""Unit tests for the validation / electrical / layout COM operations.

Pure-logic paths only — no AutoCAD required.
"""

from __future__ import annotations

import asyncio
import math

import pytest

from tests.test_com_ops import FakeDocument, FakeEntity, Harness, run


class GeomEntity(FakeEntity):
    """Fake entity carrying the properties the validators read."""

    def __init__(self, handle, entity_name, layer="0", **props):
        super().__init__(handle, entity_name, layer)
        for k, v in props.items():
            setattr(self, k, v)


def line(handle, start, end, layer="0"):
    return GeomEntity(handle, "AcDbLine", layer, StartPoint=start, EndPoint=end)


def circle(handle, center, radius, layer="0"):
    return GeomEntity(
        handle, "AcDbCircle", layer, Center=center, Radius=radius,
        Area=math.pi * radius * radius, Circumference=2 * math.pi * radius,
    )


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------


def test_validate_zero_length_finds_zero_line_and_zero_radius():
    doc = FakeDocument([
        line("A", (0, 0, 0), (10, 0, 0)),
        line("B", (5, 5, 0), (5, 5, 0)),
        circle("C", (0, 0, 0), 0.0),
    ])
    result = run(Harness(doc).validate_zero_length())
    assert result.ok
    assert result.payload["count"] == 2
    issues = {e["handle"] for e in result.payload["entities"]}
    assert issues == {"B", "C"}


def test_validate_duplicates_groups_identical_lines():
    doc = FakeDocument([
        line("A", (0, 0, 0), (10, 0, 0)),
        line("B", (0, 0, 0), (10, 0, 0)),
        line("C", (0, 0, 0), (20, 0, 0)),
    ])
    result = run(Harness(doc).validate_duplicates(tolerance=0.001))
    assert result.ok
    assert result.payload["duplicate_groups"] == 1
    assert result.payload["redundant_entities"] == 1
    handles = {e["handle"] for e in result.payload["groups"][0]}
    assert handles == {"A", "B"}


def test_validate_duplicates_respects_tolerance():
    doc = FakeDocument([
        line("A", (0, 0, 0), (10, 0, 0)),
        line("B", (0.0004, 0, 0), (10, 0, 0)),
    ])
    loose = run(Harness(doc).validate_duplicates(tolerance=0.01))
    tight = run(Harness(doc).validate_duplicates(tolerance=0.0001))
    assert loose.payload["duplicate_groups"] == 1
    assert tight.payload["duplicate_groups"] == 0


def test_validate_connectivity_flags_dangling_ends():
    doc = FakeDocument([
        line("A", (0, 0, 0), (10, 0, 0)),
        line("B", (10, 0, 0), (10, 10, 0)),
        line("C", (50, 50, 0), (60, 60, 0)),
    ])
    result = run(Harness(doc).validate_connectivity())
    assert result.ok
    assert result.payload["segments_checked"] == 3
    # A and B each have one loose end; C has two -> all three are dangling
    assert result.payload["dangling_segments"] == 3


def test_validate_layer_standards_lists_offenders():
    doc = FakeDocument([line("A", (0, 0, 0), (1, 0, 0), layer="BAD"),
                        line("B", (0, 0, 0), (1, 0, 0), layer="GOOD")])
    result = run(Harness(doc).validate_layer_standards(["GOOD"]))
    assert result.ok
    assert [o["layer"] for o in result.payload["offending_layers"]] == ["BAD"]
    assert result.payload["compliant"] is False


def test_validate_qc_report_aggregates():
    doc = FakeDocument([
        line("A", (0, 0, 0), (10, 0, 0)),
        line("B", (0, 0, 0), (10, 0, 0)),
        line("Z", (1, 1, 0), (1, 1, 0)),
    ])
    result = run(Harness(doc).validate_qc_report())
    assert result.ok
    assert result.payload["issue_total"] >= 2
    assert result.payload["clean"] is False


# ---------------------------------------------------------------------------
# Geometry
# ---------------------------------------------------------------------------


def test_geometry_distance_pythagoras():
    result = run(Harness(FakeDocument()).geometry_distance(0, 0, 3, 4))
    assert result.ok
    assert result.payload["distance"] == 5.0
    assert result.payload["angle_deg"] == pytest.approx(53.130102, abs=1e-4)


# ---------------------------------------------------------------------------
# Electrical (pure computation)
# ---------------------------------------------------------------------------


def test_electrical_voltage_drop_single_phase():
    h = Harness(FakeDocument())
    result = run(h.electrical_voltage_drop(120, 16, "12", 100, 1))
    assert result.ok
    # 2 * 16A * 1.93 ohm/kft * 0.1 kft = 6.176 V
    assert result.payload["voltage_drop"] == pytest.approx(6.176, abs=1e-3)
    assert result.payload["within_3pct"] is False


def test_electrical_voltage_drop_three_phase_scales_by_sqrt3():
    h = Harness(FakeDocument())
    one = run(h.electrical_voltage_drop(480, 10, "10", 100, 1)).payload
    three = run(h.electrical_voltage_drop(480, 10, "10", 100, 3)).payload
    assert three["voltage_drop"] == pytest.approx(one["voltage_drop"] * 1.732 / 2, rel=1e-3)


def test_electrical_voltage_drop_rejects_unknown_gauge():
    result = run(Harness(FakeDocument()).electrical_voltage_drop(120, 10, "99", 100))
    assert not result.ok
    assert "Unknown gauge" in result.error


def test_electrical_conduit_fill_compliance():
    h = Harness(FakeDocument())
    ok = run(h.electrical_conduit_fill('3/4"', "EMT", ["12", "12", "12"]))
    assert ok.ok and ok.payload["compliant"] is True
    over = run(h.electrical_conduit_fill('1/2"', "EMT", ["8"] * 6))
    assert over.ok and over.payload["compliant"] is False


def test_electrical_load_calc_totals():
    result = run(Harness(FakeDocument()).electrical_load_calc(
        [{"name": "lamp", "quantity": 4, "va": 60},
         {"name": "fan", "quantity": 1, "va": 100}]
    ))
    assert result.ok
    assert result.payload["total_va"] == 340.0
    assert result.payload["at_120v_amps"] == pytest.approx(2.83, abs=0.01)


def test_electrical_nec_lookup_known_and_unknown():
    h = Harness(FakeDocument())
    good = run(h.electrical_nec_lookup("310.16", {"gauge": "12"}))
    assert good.ok and good.payload["ampacity_thhn_90c"] == 30
    bad = run(h.electrical_nec_lookup("310.16", {"gauge": "99"}))
    assert not bad.ok


# ---------------------------------------------------------------------------
# Layout viewport scale verification
# ---------------------------------------------------------------------------


class FakeViewport(FakeEntity):
    def __init__(self, handle, height=6.0):
        super().__init__(handle, "AcDbViewport")
        self.Height = height
        self._custom = 1.0

    @property
    def CustomScale(self):
        return self._custom

    @CustomScale.setter
    def CustomScale(self, value):
        self._custom = value


def test_layout_viewport_set_scale_verifies_readback():
    vp = FakeViewport("VP1", height=6.0)
    doc = FakeDocument([vp])
    doc.HandleToObject = lambda h: vp if h == "VP1" else (_ for _ in ()).throw(KeyError(h))
    result = run(Harness(doc).layout_viewport_set_scale("VP1", 2.0))
    assert result.ok, result.error
    assert result.payload["custom_scale"] == pytest.approx(0.5)
    assert result.payload["effective_scale"] == pytest.approx(2.0)
    assert result.payload["verified"] is True
