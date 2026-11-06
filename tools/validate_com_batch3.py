"""Live validation — batches 4 & 5: validate, export, blocks, dims, layout, equipment, electrical, pid."""
import asyncio
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, "src")

import pythoncom
import win32com.client
from win32com.client import VARIANT

from autocad_mcp.backends.com_backend import COMBackend

RESULTS = []


def check(name, ok, detail=None):
    RESULTS.append({"op": name, "ok": bool(ok)})
    print(f"[{'PASS' if ok else 'FAIL'}] {name}  {str(detail)[:200]}")


def pt(*v):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list(v))


async def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("AutoCAD.Application")
    for d in list(app.Documents):
        if d.Name.lower().startswith("drawing") or d.Saved:
            d.Close(False)
        else:
            print("REFUSED: unsaved user doc", d.Name)
            return False

    b = COMBackend()
    await b.initialize()
    r = await b.drawing_create("scratch")
    if not r.ok:
        check("drawing_create", False, r.error)
        return False
    doc = b._doc
    for _ in range(60):
        try:
            _ = doc.ModelSpace.Count
            break
        except Exception:
            time.sleep(1)
    if not doc.Name.lower().startswith("drawing"):
        print("REFUSING:", doc.Name)
        return False
    msp = doc.ModelSpace

    # fixtures: two touching lines, one duplicate, one zero-length, text, block
    l1 = msp.AddLine(pt(0, 0, 0), pt(10, 0, 0))
    l2 = msp.AddLine(pt(10, 0, 0), pt(10, 10, 0))
    dup = msp.AddLine(pt(0, 0, 0), pt(10, 0, 0))
    zero = msp.AddLine(pt(5, 5, 0), pt(5, 5, 0))
    txt = msp.AddText("VALVE-01", pt(2, 2, 0), 1.0)
    doc.Blocks.Add(pt(0, 0, 0), "EQBLK").AddLine(pt(0, 0, 0), pt(2, 0, 0))
    ins = msp.InsertBlock(pt(30, 30, 0), "EQBLK", 1.0, 1.0, 1.0, 0.0)

    # --- validation ---
    r = await b.validate_zero_length()
    check("validate_zero_length", r.ok and r.payload["count"] == 1, r.payload["count"])

    r = await b.validate_duplicates()
    check("validate_duplicates", r.ok and r.payload["duplicate_groups"] >= 1, r.payload)

    r = await b.validate_connectivity()
    check("validate_connectivity", r.ok and r.payload["segments_checked"] >= 3, r.payload)

    r = await b.validate_orphaned_entities()
    check("validate_orphaned_entities", r.ok, r.payload)

    r = await b.validate_layer_standards(["0"])
    check("validate_layer_standards", r.ok and r.payload["compliant"], r.payload)

    r = await b.validate_text_standards(allowed_heights=[1.0])
    check("validate_text_standards", r.ok and r.payload["offender_count"] == 0, r.payload)

    r = await b.validate_attribute_completeness()
    check("validate_attribute_completeness", r.ok, r.payload)

    r = await b.validate_qc_report()
    check("validate_qc_report", r.ok and r.payload["issue_total"] >= 1, r.payload["issue_total"])

    # --- export ---
    r = await b.export_layer_report()
    check("export_layer_report", r.ok and r.payload["layer_count"] >= 1, r.payload["layer_count"])

    r = await b.export_block_count()
    check("export_block_count", r.ok and r.payload["blocks"].get("EQBLK") == 1, r.payload)

    r = await b.export_bom()
    check("export_bom", r.ok and r.payload["total_inserts"] == 1, r.payload)

    r = await b.export_data_extract(entity_type="AcDbLine")
    check("export_data_extract", r.ok and r.payload["count"] == 4, r.payload["count"])

    # --- blocks ---
    r = await b.block_define("NEWBLK", [
        {"type": "line", "x1": 0, "y1": 0, "x2": 5, "y2": 0},
        {"type": "circle", "cx": 0, "cy": 0, "radius": 1},
    ])
    check("block_define", r.ok and r.payload["entities_created"] == 2, r.payload)

    r = await b.block_insert_with_attributes("NEWBLK", 60, 60)
    check("block_insert_with_attributes", r.ok, r.payload)

    # --- dimensions ---
    r = await b.create_dimension_aligned(0, 0, 10, 0, 2)
    check("create_dimension_aligned", r.ok and r.payload["measurement"] == 10.0, r.payload)

    r = await b.create_dimension_angular(0, 0, 10, 0, 0, 10)
    check("create_dimension_angular", r.ok, r.payload)

    r = await b.create_dimension_radius(20, 20, 5, 45)
    check("create_dimension_radius", r.ok, r.payload)

    r = await b.create_leader([[0, 0, 0], [5, 5, 0]], "NOTE")
    check("create_leader", r.ok, r.payload)

    # --- view ---
    r = await b.zoom_center(0, 0, 100)
    check("zoom_center", r.ok, r.payload)
    r = await b.zoom_scale(0.8)
    check("zoom_scale", r.ok, r.payload)
    r = await b.pan(1, 1)
    check("pan", r.ok, r.payload)

    # --- find/replace ---
    r = await b.find_replace_text("VALVE", "GATE")
    check("find_replace_text", r.ok and r.payload["replaced"] == 1, r.payload)

    r = await b.find_replace_attribute("TAG", "a", "b")
    check("find_replace_attribute", r.ok, r.payload)

    # --- layouts ---
    r = await b.layout_list()
    check("layout_list", r.ok and r.payload["count"] >= 1, r.payload["count"])

    r = await b.layout_create("MCP_TEST_LAYOUT")
    check("layout_create", r.ok, r.payload)

    r = await b.layout_switch("MCP_TEST_LAYOUT")
    check("layout_switch", r.ok and r.payload["active"] == "MCP_TEST_LAYOUT", r.payload)

    r = await b.layout_viewport_create(5, 5, 8, 6, 1.0)
    vp = r.payload.get("handle") if r.ok else None
    check("layout_viewport_create", r.ok and vp, r.payload)

    if vp:
        r = await b.layout_viewport_set_scale(vp, 2.0)
        check("layout_viewport_set_scale", r.ok, r.payload)
        r = await b.layout_viewport_lock(vp, True)
        check("layout_viewport_lock", r.ok and r.payload["locked"], r.payload)

    r = await b.layout_titleblock_fill("MCP_TEST_LAYOUT", {"TITLE": "X"})
    check("layout_titleblock_fill", r.ok, r.payload)

    r = await b.layout_delete("MCP_TEST_LAYOUT")
    check("layout_delete", r.ok, r.payload)
    await b.layout_switch("Model")

    # --- equipment ---
    r = await b.equipment_find("VALVE*")
    check("equipment_find", r.ok and r.payload["count"] == 0, r.payload)

    r = await b.equipment_find("GATE*")
    check("equipment_find text hit", r.ok and r.payload["count"] == 1, r.payload["count"])

    r = await b.equipment_inspect(0, 0, handle=ins.Handle)
    check("equipment_inspect", r.ok and r.payload["found"], r.payload)

    r = await b.place_equipment_tag(100, 100, 0, "EQ-1", 24, "right", 8)
    check("place_equipment_tag", r.ok and r.payload["cube"], r.payload)

    r = await b.batch_find_and_tag(["GATE-01"])
    check("batch_find_and_tag", r.ok and r.payload["placed_count"] == 1, r.payload)

    # --- electrical (pure) ---
    r = await b.electrical_nec_lookup("310.16", {"gauge": "12"})
    check("electrical_nec_lookup", r.ok and r.payload["ampacity_thhn_90c"] == 30, r.payload)

    r = await b.electrical_voltage_drop(120, 16, "12", 100, 1)
    check("electrical_voltage_drop", r.ok and r.payload["voltage_drop"] > 0, r.payload)

    r = await b.electrical_conduit_fill('3/4"', "EMT", ["12", "12", "12"])
    check("electrical_conduit_fill", r.ok and r.payload["compliant"], r.payload)

    r = await b.electrical_load_calc([{"name": "lamp", "quantity": 4, "va": 60}])
    check("electrical_load_calc", r.ok and r.payload["total_va"] == 240.0, r.payload)

    r = await b.electrical_wire_number_assign("0", "W", 1)
    check("electrical_wire_number_assign", r.ok and r.payload["labelled"] >= 3, r.payload)

    r = await b.electrical_circuit_trace(l1.Handle)
    check("electrical_circuit_trace", r.ok and r.payload["circuit_size"] >= 2, r.payload)

    r = await b.electrical_symbol_insert("EQBLK", 80, 80)
    check("electrical_symbol_insert", r.ok, r.payload)

    # --- pid ---
    r = await b.pid_setup_layers()
    check("pid_setup_layers", r.ok, r.payload)

    r = await b.pid_draw_process_line(0, 50, 50, 50)
    check("pid_draw_process_line", r.ok, r.payload)

    r = await b.pid_connect_equipment(50, 50, 80, 50)
    check("pid_connect_equipment", r.ok and r.payload["connected"], r.payload)

    r = await b.pid_add_flow_arrow(25, 50)
    check("pid_add_flow_arrow", r.ok, r.payload)

    r = await b.pid_add_equipment_tag(0, 55, "P-101", "PUMP")
    check("pid_add_equipment_tag", r.ok, r.payload)

    r = await b.pid_add_line_number(0, 60, "1-P-001", '2"')
    check("pid_add_line_number", r.ok, r.payload)

    r = await b.pid_list_symbols("valves")
    check("pid_list_symbols", r.ok, {"count": r.payload.get("count") if r.ok else None})

    # --- magicad ---
    r = await b.magicad_status()
    check("magicad_status", r.ok, r.payload)
    r = await b.magicad_list_commands()
    check("magicad_list_commands", r.ok, len(r.payload.get("known", [])))
    r = await b.magicad_project_info()
    check("magicad_project_info", r.ok, r.payload)

    doc.Close(False)
    passed = sum(1 for x in RESULTS if x["ok"])
    print(f"\nSUMMARY: {passed}/{len(RESULTS)} passed")
    return passed == len(RESULTS)


if __name__ == "__main__":
    ok = asyncio.run(main())
    sys.exit(0 if ok else 1)
