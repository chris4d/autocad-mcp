@echo off
rem Launch the self-contained AutoCAD MCP server bundle.
rem Layout: this file sits next to python\ and app\; app\ holds autocad_mcp
rem and every dependency installed with --target (fully relocatable).
setlocal
set "ROOT=%~dp0"
rem pywin32's .pth is ignored under --target, so add its win32 dirs explicitly.
set "PYTHONPATH=%ROOT%app;%ROOT%app\win32;%ROOT%app\win32\lib"
set "PYTHONNOUSERSITE=1"
rem pywin32 ships its DLLs in app\pywin32_system32; put them on PATH instead of
rem running the admin-only pywin32 post-install.
set "PATH=%ROOT%app\pywin32_system32;%ROOT%python;%ROOT%python\Scripts;%PATH%"
rem COM backend only (the suite ships the full-AutoCAD path, not File IPC/ezdxf).
set "AUTOCAD_MCP_BACKEND=com"
if not defined AUTOCAD_MCP_CAD_TYPE set "AUTOCAD_MCP_CAD_TYPE=autocad"
rem matplotlib needs a writable cache; keep it out of the install dir.
set "MPLCONFIGDIR=%LOCALAPPDATA%\autocad-mcp\mpl"
if not exist "%MPLCONFIGDIR%" mkdir "%MPLCONFIGDIR%" >nul 2>&1
"%ROOT%python\python.exe" -m autocad_mcp %*
