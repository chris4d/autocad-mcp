# Build a self-contained, relocatable Windows x64 bundle of the AutoCAD MCP
# server: portable CPython + autocad_mcp + all runtime dependencies + shim.
#
# The resulting zip is what the SMS Toolkit suite pins and stages:
#   autocad-mcp-v<version>-win64.zip
# Layout inside the zip (extracted to {app}\autocad-mcp by the suite):
#   run-autocad-mcp.cmd
#   python\            portable CPython (python-build-standalone)
#   app\               autocad_mcp + deps installed with --target
#   lisp-code\ config\ LICENSE NOTICE README.md
#   VERSION.txt BUILD-INFO.txt
#
# Usage:
#   .\packaging\build-bundle.ps1                 # build current pyproject version
#   .\packaging\build-bundle.ps1 -Version 5.1.0
#   .\packaging\build-bundle.ps1 -PythonArchive C:\path\cpython-...tar.gz
param(
    [string]$Version = '',
    [string]$PythonArchive = '',
    [string]$OutputDir = '',
    # Pinned portable CPython. Bump deliberately; the suite consumes the zip,
    # so the runtime is baked into every installed copy.
    [string]$PbsTag = '20261123',
    [string]$PythonVersion = '3.12.16',
    [switch]$KeepWork
)
$ErrorActionPreference = 'Stop'
# Native tools (uv, tar, python) write progress to stderr; do not let that
# become a terminating error. $LASTEXITCODE is the source of truth below.
if (Get-Variable PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue) {
    $PSNativeCommandUseErrorActionPreference = $false
}
$root = Split-Path $PSScriptRoot -Parent

if ($Version -match '^v') { $Version = $Version.Substring(1) }
if (-not $Version) {
    $m = Select-String -Path (Join-Path $root 'pyproject.toml') -Pattern '^version\s*=\s*"([^"]+)"'
    if (-not $m) { throw 'Could not read version from pyproject.toml' }
    $Version = $m.Matches[0].Groups[1].Value
}
$declared = (Select-String -Path (Join-Path $root 'pyproject.toml') -Pattern '^version\s*=\s*"([^"]+)"').Matches[0].Groups[1].Value
if ($declared -ne $Version) {
    throw "Version mismatch: requested $Version but pyproject.toml declares $declared"
}

$work = Join-Path $root 'build\bundle'
$cache = Join-Path $root 'build\cache'
$dist = if ($OutputDir) { $OutputDir } else { Join-Path $root 'dist' }
$null = New-Item -ItemType Directory -Force $work, $cache, $dist

# Run a native tool without letting its stderr become a terminating error.
# Returns the combined output; throws on a nonzero exit code.
function Invoke-NativeChecked {
    param([scriptblock]$Block, [string]$What)
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & $Block 2>&1
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $old
    }
    foreach ($line in @($output)) { Write-Host $line }
    if ($code -ne 0) { throw "$What failed (exit $code)" }
    return $output
}

# ---- 1) portable CPython -----------------------------------------------------
$pyDir = Join-Path $work 'python'
if (-not (Test-Path (Join-Path $pyDir 'python.exe'))) {
    if ($PythonArchive) {
        $archive = $PythonArchive
        if (-not (Test-Path -LiteralPath $archive)) { throw "PythonArchive not found: $archive" }
    } else {
        $name = "cpython-$PythonVersion+$PbsTag-x86_64-pc-windows-msvc-install_only.tar.gz"
        $archive = Join-Path $cache $name
        if (-not (Test-Path -LiteralPath $archive)) {
            $url = "https://github.com/astral-sh/python-build-standalone/releases/download/$PbsTag/" + ($name -replace '\+', '%2B')
            Write-Host "Downloading portable CPython $PythonVersion ($PbsTag)"
            Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $archive
        }
    }
    Write-Host "Extracting $archive"
    Invoke-NativeChecked { & tar.exe -xzf $archive -C $work } "tar extract $archive" | Out-Null
    if (-not (Test-Path (Join-Path $pyDir 'python.exe'))) {
        throw "Expected python\python.exe after extracting $archive"
    }
}
$pyExe = Join-Path $pyDir 'python.exe'
$pyActual = (& $pyExe --version 2>&1) -replace '^Python\s+', ''
Write-Host "Bundled Python: $pyActual"

# ---- 2) install the project + dependencies into app\ ------------------------
# --target keeps everything relocatable (no venv, no absolute pyvenv.cfg).
$appDir = Join-Path $work 'app'
if (Test-Path -LiteralPath $appDir) { Remove-Item -LiteralPath $appDir -Recurse -Force }
$uv = Get-Command uv -ErrorAction SilentlyContinue
if ($uv) {
    Write-Host 'Installing project into app\ with uv'
    Invoke-NativeChecked { & $uv.Source pip install --target $appDir --python $pyExe $root } 'uv install' | Out-Null
} else {
    Write-Host 'uv not found; installing project into app\ with bundled pip'
    $pipOk = $false
    try { Invoke-NativeChecked { & $pyExe -m pip --version } 'pip check' | Out-Null; $pipOk = $true } catch { }
    if (-not $pipOk) { Invoke-NativeChecked { & $pyExe -m ensurepip --upgrade } 'ensurepip' | Out-Null }
    Invoke-NativeChecked { & $pyExe -m pip install --target $appDir $root } 'pip install' | Out-Null
}
if (-not (Test-Path (Join-Path $appDir 'autocad_mcp\__main__.py'))) {
    throw 'app\autocad_mcp\__main__.py missing after install'
}
# Smoke test: import the package and the COM dependency from the staged app.
$smoke = "import autocad_mcp, win32com.client, ezdxf, matplotlib, PIL, openpyxl, structlog; print('ok')"
$env:PYTHONPATH = @(
    $appDir,
    (Join-Path $appDir 'win32'),
    (Join-Path $appDir 'win32\lib')
) -join ';'
$env:PATH = (Join-Path $appDir 'pywin32_system32') + ';' + $env:PATH
$out = Invoke-NativeChecked { & $pyExe -c $smoke } 'bundle smoke test'
if (($out | Out-String) -notmatch 'ok') {
    throw "Bundle smoke test did not report ok:`n$($out | Out-String)"
}
Remove-Item Env:PYTHONPATH -ErrorAction SilentlyContinue

# ---- 3) bundle assets --------------------------------------------------------
foreach ($item in 'lisp-code', 'config') {
    $src = Join-Path $root $item
    if (Test-Path -LiteralPath $src) {
        Copy-Item -LiteralPath $src -Destination $work -Recurse -Force
    }
}
foreach ($file in 'LICENSE', 'NOTICE', 'README.md') {
    $src = Join-Path $root $file
    if (Test-Path -LiteralPath $src) { Copy-Item -LiteralPath $src -Destination $work -Force }
}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'run-autocad-mcp.cmd') -Destination $work -Force

Set-Content -LiteralPath (Join-Path $work 'VERSION.txt') -Value $Version -Encoding ASCII -NoNewline
@(
    "version=$Version"
    "python=$pyActual"
    "pbs_tag=$PbsTag"
    "built_utc=$((Get-Date).ToUniversalTime().ToString('o'))"
) | Set-Content -LiteralPath (Join-Path $work 'BUILD-INFO.txt') -Encoding ASCII

# ---- 4) prune caches (keep the zip lean) ------------------------------------
Get-ChildItem -LiteralPath $work -Recurse -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -eq '__pycache__' } |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Get-ChildItem -LiteralPath $work -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -in '.pyc', '.pyo' } |
    Remove-Item -Force -ErrorAction SilentlyContinue

# ---- 5) zip + checksum -------------------------------------------------------
# Never ship the downloaded runtime archive (older builds staged it in $work).
Get-ChildItem -LiteralPath $work -Filter '*.tar.gz' -File -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue
$zip = Join-Path $dist "autocad-mcp-v$Version-win64.zip"
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
Write-Host "Compressing -> $zip"
Compress-Archive -Path (Join-Path $work '*') -DestinationPath $zip -CompressionLevel Optimal
$hash = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLower()
$shaFile = "$zip.sha256"
Set-Content -LiteralPath $shaFile -Value ("{0}  {1}" -f $hash, (Split-Path $zip -Leaf)) -Encoding ASCII

$sizeMb = [math]::Round((Get-Item -LiteralPath $zip).Length / 1MB, 1)
Write-Host "Built $zip ($sizeMb MB)"
Write-Host "SHA-256: $hash"
Write-Host "Checksum file: $shaFile"

if (-not $KeepWork) {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
