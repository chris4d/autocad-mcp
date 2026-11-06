"""Summarise a dwg_scan3 corpus into a Revit-link-warning triage report.

Reads every <base>.scan.json in the scan output directory, prints the counts
that matter for DWG sanitising (far-from-origin, z-spread, ambiguous units,
hatches, xrefs, proxy/3D/raster/OLE entities), and writes _corpus_summary.json
alongside them.

Usage:
  uv run python tools/aggregate_corpus.py [scan-dir]

scan-dir defaults to ../scan-results, matching dwg_scan3.py.
"""
import glob
import json
import os
import sys

DEFAULT_OUT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "scan-results"))
PROXY_KEYS = ("ACAD_PROXY_ENTITY", "3DSOLID", "REGION", "OLE2FRAME", "RASTERIMAGE")


def load_reports(out_dir):
    reps = []
    for sp in glob.glob(os.path.join(out_dir, "*.scan.json")):
        with open(sp, encoding="utf-8") as fh:
            reps.append(json.load(fh))
    return reps


def main():
    out_dir = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_OUT
    reps = load_reports(out_dir)
    errs = [r for r in reps if r.get("scan_error")]
    ok = [r for r in reps if not r.get("scan_error")]
    print(f"scanned: {len(reps)} | ok: {len(ok)} | errors: {len(errs)}")
    for r in errs:
        print(f"  ERROR {r['name']}: {r['scan_error']}")

    over20 = sorted([r for r in ok if r.get("max_dist_from_origin_mi", 0) > 20],
                    key=lambda r: -r.get("max_dist_from_origin_mi", 0))
    far = [r for r in ok if 2 < r.get("max_dist_from_origin_mi", 0) <= 20]
    zspread = [r for r in ok if r.get("z_spread", 0) and abs(r["z_spread"]) > 0.1]
    units_bad = [r for r in ok if r.get("insunits", r.get("INSUNITS")) in (0, 4, None)]
    hatch = [r for r in ok if r.get("hatch_count", 0) > 0]
    xref = [r for r in ok if r.get("xrefs")]
    proxy = [r for r in ok if any(r.get(k) for k in PROXY_KEYS)]

    print(f"\n>20mi from origin: {len(over20)}")
    for r in over20[:15]:
        print(f"  {r.get('max_dist_from_origin_mi'):>12.2f}mi  {r['name'][:70]}")
    print(f"\n2-20mi (far): {len(far)}")
    print(f"z-spread > 0.1: {len(zspread)}")
    for r in zspread[:10]:
        print(f"  z={r['z_spread']:>10.3f}  {r['name'][:70]}")
    print(f"\nINSUNITS 0/4/None (ambiguous): {len(units_bad)}")
    print(f"with hatches: {len(hatch)} (total hatch count: {sum(r['hatch_count'] for r in hatch)})")
    print(f"hatch zero-loop: {sum(r.get('hatch_zero_loop', 0) for r in ok)}")
    print(f"with xrefs: {len(xref)}")
    noref = [r for r in xref if any(not x.get("path") or x.get("path", "").strip() in ('.', '', '.\\')
                                    for x in r['xrefs'])]
    print(f"  (xref-less-path or bare: {len(noref)})")
    print(f"unsupported/proxy/3d/raster/ole: {len(proxy)}")
    for r in proxy[:10]:
        parts = [k for k in PROXY_KEYS if r.get(k)]
        print(f"  {dict((k, r[k]) for k in parts)} {r['name'][:50]}")

    summary = os.path.join(out_dir, "_corpus_summary.json")
    with open(summary, "w", encoding="utf-8") as fh:
        json.dump({"over20": [r["name"] for r in over20], "far": [r["name"] for r in far],
                   "zspread": [r["name"] for r in zspread], "units_bad": [r["name"] for r in units_bad],
                   "errors": [r["name"] for r in errs]}, fh, indent=2)
    print(f"\nsaved {summary}")


if __name__ == "__main__":
    main()
