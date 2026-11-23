# Self-contained bundle

The SMS Toolkit suite does not install Python, `uv`, or wheels on the target
machine. Instead, `build-bundle.ps1` produces a single relocatable zip:

```
autocad-mcp-v<version>-win64.zip
  run-autocad-mcp.cmd
  python\              portable CPython (python-build-standalone)
  app\                 autocad_mcp + every dependency (installed with --target)
  lisp-code\ config\
  LICENSE NOTICE README.md
  VERSION.txt BUILD-INFO.txt
```

The suite extracts this to `{app}\autocad-mcp` and points MCP clients at
`{app}\autocad-mcp\run-autocad-mcp.cmd`.

## Build locally

```powershell
.\packaging\build-bundle.ps1                 # current pyproject version
.\packaging\build-bundle.ps1 -Version 5.1.0
.\packaging\build-bundle.ps1 -KeepWork       # keep build\bundle for inspection
```

Output: `dist\autocad-mcp-v<version>-win64.zip` plus a `.sha256` sidecar.

## Notes

- The runtime is pinned in `build-bundle.ps1` (`$PbsTag` / `$PythonVersion`).
  Bump both deliberately; the suite bakes the runtime into every installed copy.
- Dependencies are installed with `uv pip install --target` (or bundled pip) so
  there is no venv and no absolute `pyvenv.cfg` to break on relocation.
- pywin32's `.pth` is not processed under `--target`; `run-autocad-mcp.cmd`
  adds `app\win32` and `app\win32\lib` to `PYTHONPATH` and
  `app\pywin32_system32` to `PATH` instead of running the admin-only
  post-install.
- The shim forces `AUTOCAD_MCP_BACKEND=com`; this bundle targets full AutoCAD,
  not AutoCAD LT's File IPC path or the headless ezdxf path.
- Tag `v<version>` to publish the zip via `.github/workflows/release-bundle.yml`.
  The tag must match `pyproject.toml`; the build fails otherwise.
