"""Live validation — batches 2 & 3: xref, query, geometry, search, select, layers."""
import asyncio
import sys
import time

sys.path.insert(0, "src")

import pythoncom
import win32com.client
from win32com.client import VARIANT

from autocad_mcp.backends.com_backend import COMBackend

RESULTS = []


def check(name, ok, detail=None):
    RESULTS.append({"op": name, "ok": bool(ok)})
    print(f"[{'PASS' if ok else 'FAIL'}] {name}  {str(detail)[:220]}")


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
        print("REFUSING: active doc", doc.Name)
        return False
    msp = doc.ModelSpace

    # --- scratch geometry ---
    ln = msp.AddLine(pt(0, 0, 0), pt(10, 0, 0))
    ln2 = msp.AddLine(pt(0, 5, 0), pt(10, 5, 0))
    ci = msp.AddCircle(pt(20, 20, 0), 5)
    pl = msp.AddLightWeightPolyline(
        VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0, 10, 10, 10, 10, 0])
    )
    pl.Closed = True
    txt = msp.AddText("Hello World", pt(2, 2, 0), 1.0)
    txt.Layer = "TEXTLAYER" if False else txt.Layer

    blk = doc.Blocks.Add(pt(100, 100, 0), "TESTBLK")
    blk.AddLine(pt(0, 0, 0), pt(5, 0, 0))
    ins = msp.InsertBlock(pt(50, 50, 0), "TESTBLK", 1.0, 1.0, 1.0, 0.0)

    # --- geometry ---
    r = await b.geometry_distance(0, 0, 3, 4)
    check("geometry_distance", r.ok and r.payload["distance"] == 5.0, r.payload)

    r = await b.geometry_area(ci.Handle)
    check("geometry_area", r.ok and abs(r.payload["area"] - 78.5398) < 0.01, r.payload)

    r = await b.geometry_length(ln.Handle)
    check("geometry_length", r.ok and r.payload["length"] == 10.0, r.payload)

    r = await b.geometry_bounding_box(ln.Handle)
    check("geometry_bounding_box entity", r.ok and r.payload["max"][0] == 10.0, r.payload)

    r = await b.geometry_bounding_box()
    check("geometry_bounding_box drawing", r.ok and r.payload["count"] >= 5, r.payload)

    r = await b.geometry_polyline_info(pl.Handle)
    check("geometry_polyline_info", r.ok and r.payload["vertex_count"] == 3, r.payload)

    # --- search ---
    r = await b.find_text("hello")
    check("find_text", r.ok and r.payload["count"] == 1, r.payload)

    r = await b.search_by_window(-1, -1, 11, 6)
    check("search_by_window", r.ok and r.payload["count"] >= 2, r.payload["count"])

    r = await b.search_by_proximity(20, 20, 1)
    check("search_by_proximity", r.ok and r.payload["count"] >= 1, r.payload["count"])

    r = await b.search_by_block_name("TESTBLK")
    check("search_by_block_name", r.ok and r.payload["count"] == 1, r.payload)

    r = await b.search_by_handle_list([ln.Handle, "ZZZZ"])
    check("search_by_handle_list", r.ok and r.payload["count"] == 1
          and r.payload["missing"] == ["ZZZZ"], r.payload)

    # --- select / modify ---
    r = await b.select_filter(entity_type="AcDbLine")
    check("select_filter", r.ok and r.payload["count"] == 2, r.payload["count"])

    r = await b.bulk_move([ln2.Handle], 0, 100)
    moved_ok = False
    if r.ok:
        after = await b.geometry_bounding_box(ln2.Handle)
        moved_ok = after.ok and abs(after.payload["min"][1] - 105.0) < 0.01
    check("bulk_move", r.ok and moved_ok, r.payload)

    r = await b.bulk_copy([ln.Handle], 0, 20)
    check("bulk_copy", r.ok and r.payload["created_count"] == 1, r.payload)

    r = await b.entity_set_property(ln.Handle, "color", 1)
    check("entity_set_property", r.ok and r.payload["after"] == 1, r.payload)

    r = await b.entity_set_text(txt.Handle, "Changed")
    check("entity_set_text", r.ok and r.payload["after"] == "Changed", r.payload)

    r = await b.entity_explode(ins.Handle)
    check("entity_explode", r.ok and r.payload["count"] >= 1, r.payload)

    # --- queries ---
    r = await b.query_layer_summary("0")
    check("query_layer_summary", r.ok and r.payload["count"] >= 1, r.payload["count"])

    r = await b.query_block_tree()
    check("query_block_tree", r.ok and r.payload["block_count"] >= 2, r.payload["block_count"])

    r = await b.query_drawing_metadata()
    check("query_drawing_metadata", r.ok and r.payload["name"] == doc.Name, r.payload)

    r = await b.query_text_styles()
    check("query_text_styles", r.ok and r.payload["count"] >= 1, r.payload["count"])

    r = await b.query_dimension_styles()
    check("query_dimension_styles", r.ok and r.payload["count"] >= 1, r.payload["count"])

    r = await b.query_linetypes()
    check("query_linetypes", r.ok and r.payload["count"] >= 1, r.payload["count"])

    r = await b.query_entity_geometry(ln.Handle)
    check("query_entity_geometry", r.ok and r.payload["length"] == 10.0, r.payload)

    # --- layers ---
    doc.Layers.Add("SRC")
    ln.Layer = "SRC"
    r = await b.layer_visibility("SRC", False)
    check("layer_visibility", r.ok and r.payload["visible"] is False, r.payload)
    await b.layer_visibility("SRC", True)

    r = await b.layer_rename("SRC", "SRC2")
    check("layer_rename", r.ok and r.payload["exists"], r.payload)

    r = await b.layer_merge("SRC2", "0")
    check("layer_merge", r.ok and r.payload["entities_moved"] >= 1, r.payload)

    # --- xref ---
    r = await b.xref_list()
    check("xref_list", r.ok and r.payload["count"] == 0, r.payload)

    doc.Close(False)
    passed = sum(1 for x in RESULTS if x["ok"])
    print(f"\nSUMMARY: {passed}/{len(RESULTS)} passed")
    return passed == len(RESULTS)


if __name__ == "__main__":
    ok = asyncio.run(main())
    sys.exit(0 if ok else 1)
