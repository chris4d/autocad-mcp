"""DWG Sanitizer Stage-1 scanner v3 — per-file COM + per-file census load, wedge-proof.

Per file:
  1. ensure acad alive & COM-responsive (else kill + relaunch + wait ready)
  2. close stale read-only docs
  3. Documents.Open(read-only) via EnsureDispatch
  4. readiness wait (ModelSpace.Count queryable)
  5. one tiny SendCommand: (load census_one.lsp) -> census_<base>.json in this
     doc's own LISP namespace (ssget works per-document; no cross-doc ssget)
  6. poll the result file (no COM in the hot loop) up to 300s
  7. doc.Close(False) via COM (proven safe; vla-close CRASHES AutoCAD)
  8. COM/queue stall detection -> kill acad, relaunch, retry file once

Requirements:
  Windows, a full (non-LT) AutoCAD install, and pywin32. AutoCAD is reused if
  already running, otherwise launched. It may be killed and relaunched when it
  stops answering COM, so do not run this against a session with unsaved work.

Usage:
  uv run python tools/dwg_scan3.py <root-dir> [--limit N] [--out DIR]

Output: one <base>.scan.json per drawing in the output dir (default
  ../scan-results), plus the transient census_one.lsp / census_<base>.json
  handshake files.
"""
import glob
import json
import os
import subprocess
import sys
import time

import pythoncom
import win32com.client

MILE = 1609.344
DEFAULT_OUTDIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "scan-results"))


def find_acad_exe():
    """Locate acad.exe: $ACAD_EXE, else newest AutoCAD <year> install.

    Matches "AutoCAD 20*" only -- AutoCAD LT ("AutoCAD LT 20xx") has no COM
    server, so it is not a usable target for this scanner.
    """
    env = os.environ.get("ACAD_EXE")
    if env:
        return env
    pattern = os.path.join(
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        "Autodesk", "AutoCAD 20*", "acad.exe")
    found = sorted(glob.glob(pattern))
    return found[-1] if found else r"C:\Program Files\Autodesk\AutoCAD 2025\acad.exe"


ACAD_EXE = find_acad_exe()

CENSUS_ONE = r"""(progn
  (setq *r* (open "RESULTFILE" "w"))
  (setq *err* (lambda (m)
    (write-line (strcat "\"lisp_error\": \"" (vl-string-translate "\"" "'" (if m m "?")) "\",") *r*)
    (write-line "\"done\": true" *r*)
    (close *r*)
    (princ)
  ))
  (setq *ss* (vl-catch-all-apply 'ssget (list "X")))
  (setq *tot* (if (and *ss* (not (vl-catch-all-error-p *ss*))) (sslength *ss*) 0))
  (write-line "{" *r*)
  (write-line (strcat "\"total\": " (rtos *tot* 2 0) ",") *r*)
  (foreach *typ* (list "HATCH" "3DSOLID" "REGION" "ACAD_PROXY_ENTITY" "RASTERIMAGE" "OLE2FRAME")
    (setq *s2* (vl-catch-all-apply 'ssget (list "X" (list (cons 0 *typ*)))))
    (setq *n2* (if (and *s2* (not (vl-catch-all-error-p *s2*))) (sslength *s2*) 0))
    (if (> *n2* 0)
      (write-line (strcat "\"" (strcase *typ*) "\": " (rtos *n2* 2 0) ",") *r*)
    )
  )
  (setq *e1* (vl-catch-all-apply 'getvar (list "EXTMIN")))
  (setq *e2* (vl-catch-all-apply 'getvar (list "EXTMAX")))
  (if (and (listp *e1*) (listp *e2*))
    (progn
      (write-line (strcat "\"extmin\": [" (rtos (car *e1*) 2 4) "," (rtos (cadr *e1*) 2 4) "," (rtos (caddr *e1*) 2 4) "],") *r*)
      (write-line (strcat "\"extmax\": [" (rtos (car *e2*) 2 4) "," (rtos (cadr *e2*) 2 4) "," (rtos (caddr *e2*) 2 4) "],") *r*)
    )
  )
  (write-line (strcat "\"insunits\": " (rtos (getvar "INSUNITS") 2 0) ",") *r*)
  (write-line (strcat "\"measurement\": " (rtos (getvar "MEASUREMENT") 2 0) ",") *r*)
  (setq *hc* 0 *hz* 0)
  (setq *hss* (vl-catch-all-apply 'ssget (list "X" (list (cons 0 "HATCH")))))
  (if (and *hss* (not (vl-catch-all-error-p *hss*)))
    (progn
      (setq *i* 0)
      (repeat (sslength *hss*)
        (setq *en* (ssname *hss* *i*))
        (setq *ed* (entget *en*))
        (setq *hc* (1+ *hc*))
        (setq *lp* (cdr (assoc 91 *ed*)))
        (if (and *lp* (= *lp* 0)) (setq *hz* (1+ *hz*)))
        (setq *i* (1+ *i*))
      )
    )
  )
  (write-line (strcat "\"hatch_count\": " (rtos *hc* 2 0) ",") *r*)
  (write-line (strcat "\"hatch_zero_loop\": " (rtos *hz* 2 0) ",") *r*)
  (write-line "\"xrefs\": [" *r*)
  (setq *xf* T)
  (setq *blk* (tblnext "BLOCK" T))
  (while *blk*
    (setq *xflag* (cdr (assoc 70 *blk*)))
    (if (= 4 (logand 4 *xflag*))
      (progn
        (if (not *xf*) (write-line "," *r*))
        (setq *xf* nil)
        (write-line (strcat "{\"name\": \"" (vl-string-translate "\"" "'" (cdr (assoc 2 *blk*))) "\", \"path\": \"" (vl-string-translate "\"" "'" (vl-string-translate "\\" "/" (cdr (assoc 1 *blk*)))) "\"}") *r*)
      )
    )
    (setq *blk* (tblnext "BLOCK"))
  )
  (write-line "]," *r*)
  (write-line "\"done\": true" *r*)
  (write-line "}" *r*)
  (close *r*)
  (princ)
)
"""


def busy(call, tries=45, wait=2.0):
    """Patient caller-side retry — AutoCAD rejects COM while busy (Autodesk-documented)."""
    last = None
    for _ in range(tries):
        try:
            return call()
        except (pythoncom.com_error, AttributeError) as exc:  # noqa: PERF203
            last = exc
            time.sleep(wait)
    raise last


def kill_acad():
    subprocess.run(["taskkill", "/F", "/IM", "acad.exe"],
                   capture_output=True, timeout=60)
    time.sleep(5)


def launch_acad(timeout=120):
    subprocess.Popen([ACAD_EXE, "/nologo"])
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            pythoncom.CoInitialize()
            app = win32com.client.gencache.EnsureDispatch(
                win32com.client.GetActiveObject("AutoCAD.Application"))
            n = app.Documents.Count
            if n >= 1:
                _ = app.ActiveDocument.Name
                return app
        except Exception:  # noqa: BLE001
            time.sleep(5)
    raise TimeoutError("acad never came up")


def ensure_acad(app):
    """Return a live app object; kill+relaunch if dead/unresponsive."""
    try:
        if app is not None:
            _ = busy(lambda: app.Documents.Count, tries=2, wait=1.0)
            return app
    except Exception:  # noqa: BLE001
        pass
    kill_acad()
    return launch_acad()


def close_stale(app):
    try:
        for d in list(busy(lambda: app.Documents, tries=3)):
            nm = str(d.Name)
            if not nm.startswith("Drawing"):
                try:
                    d.Close(False)
                except Exception:  # noqa: BLE001
                    pass
    except Exception:  # noqa: BLE001
        pass


def scan_v3(dwg_path, outdir=DEFAULT_OUTDIR):
    report = {"file": dwg_path, "name": os.path.splitext(os.path.basename(dwg_path))[0]}
    for attempt in (1, 2):
        try:
            return _scan_once(dwg_path, report, attempt, outdir)
        except Exception as exc:  # noqa: BLE001
            report["scan_error"] = f"attempt{attempt}: {type(exc).__name__}: {exc}"
            kill_acad()
            if attempt == 2:
                return report
    return report


def _scan_once(dwg_path, report, attempt, outdir):
    census_lsp = os.path.join(outdir, "census_one.lsp")
    pythoncom.CoInitialize()
    app = ensure_acad(globals().get("_APP"))
    globals()["_APP"] = app
    close_stale(app)
    doc = busy(lambda: app.Documents.Open(dwg_path, True))
    base = busy(lambda: doc.Name)
    # re-fetch stable proxy
    def refetch():
        for d in app.Documents:
            if d.Name == base:
                return d
        return doc
    doc = busy(refetch)
    # readiness
    deadline = time.time() + 180
    ready = False
    while time.time() < deadline:
        try:
            _ = busy(lambda: doc.ModelSpace.Count, tries=1, wait=0.5)
            ready = True
            break
        except Exception:  # noqa: BLE001
            time.sleep(3)
    if not ready:
        raise TimeoutError(f"doc not ready: {base}")
    # census via per-file load
    result_file = os.path.join(outdir, f"census_{report['name'][:80]}.json")
    if os.path.exists(result_file):
        os.remove(result_file)
    try:
        for var, val in (("EXPERT", 5), ("PROXYNOTICE", 0), ("FILEDIA", 0)):
            try:
                busy(lambda: doc.SetVariable(var, val), tries=3, wait=1.0)
            except Exception:  # noqa: BLE001
                pass
        lsp = CENSUS_ONE.replace("RESULTFILE", result_file.replace("\\", "/"))
        with open(census_lsp, "w", encoding="ascii", errors="replace") as fh:
            fh.write(lsp)
        busy(lambda: doc.SendCommand(
            f'(load "{census_lsp.replace(os.sep, "/")}") \n'), tries=5, wait=4.0)
    except Exception as exc:  # noqa: BLE001
        report["send_error"] = repr(exc)
    # poll result file — NO COM in hot loop
    data = None
    deadline = time.time() + 300
    while time.time() < deadline:
        time.sleep(2)
        try:
            txt = open(result_file, encoding="utf-8").read()
            if '"done": true' in txt:
                data = json.loads(txt)
                break
        except Exception:  # noqa: BLE001
            continue
    if data:
        report.update(data)
    else:
        raise TimeoutError(f"census not complete: {base}")
    # derived metrics
    try:
        if "extmin" not in report:
            ex1 = busy(lambda: [round(v, 2) for v in doc.GetVariable("EXTMIN")], tries=3)
            ex2 = busy(lambda: [round(v, 2) for v in doc.GetVariable("EXTMAX")], tries=3)
            report["extmin"], report["extmax"] = ex1, ex2
        ex1, ex2 = report["extmin"], report["extmax"]
        worst = max(ex1[0] ** 2 + ex1[1] ** 2 + ex1[2] ** 2,
                    ex2[0] ** 2 + ex2[1] ** 2 + ex2[2] ** 2) ** 0.5
        report["max_dist_from_origin_mi"] = round(worst / MILE, 2)
        report["z_spread"] = round(ex2[2] - ex1[2], 3)
    except Exception as exc:  # noqa: BLE001
        report["extents_error"] = str(exc)
    try:
        busy(lambda: doc.Close(False), tries=10, wait=3.0)
    except Exception:  # noqa: BLE001
        pass
    return report


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print("usage: dwg_scan3.py <root-dir> [--limit N] [--out DIR]")
        sys.exit(0 if len(sys.argv) > 1 else 1)
    root = sys.argv[1]
    limit = 999999
    if "--limit" in sys.argv:
        limit = int(sys.argv[sys.argv.index("--limit") + 1])
    outdir = DEFAULT_OUTDIR
    if "--out" in sys.argv:
        outdir = os.path.abspath(sys.argv[sys.argv.index("--out") + 1])
    os.makedirs(outdir, exist_ok=True)
    files = [f for f in glob.glob(os.path.join(root, "**", "*.dwg"), recursive=True)
             if "\\Archive\\" not in f and not f.lower().endswith(".bak")]
    files.sort(key=lambda f: os.path.getsize(f))
    done = 0
    for f in files:
        base = os.path.splitext(os.path.basename(f))[0][:80]
        out = os.path.join(outdir, base + ".scan.json")
        if os.path.exists(out) and os.path.getsize(out) > 100:
            continue
        report = scan_v3(f, outdir)
        with open(out, "w", encoding="utf-8") as fh:
            json.dump(report, fh, indent=2, ensure_ascii=False)
        tag = "ERROR" if report.get("scan_error") else "OK"
        flags = []
        if report.get("max_dist_from_origin_mi", 0) > 20:
            flags.append(">20mi")
        elif report.get("max_dist_from_origin_mi", 0) > 2:
            flags.append("far")
        if report.get("hatch_count", 0):
            flags.append(f"hatch:{report.get('hatch_count')}")
        if report.get("xrefs"):
            flags.append(f"xref:{len(report['xrefs'])}")
        units = report.get("INSUNITS", report.get("insunits"))
        if units in (0, 4, None):
            flags.append(f"units:{units}")
        print(f"[{tag}] {os.path.getsize(f)//1024:>6}KB {base[:60]} "
              f"{','.join(flags)}{report.get('scan_error', '')}", flush=True)
        done += 1
        if done >= limit:
            break
    print(f"scanned {done} this run")


if __name__ == "__main__":
    main()
