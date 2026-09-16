param(
  [string]$Template = "",
  [string]$Source = "",
  [string]$Out = ""
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
if (-not $Template) { $Template = Join-Path $Root "templates\everglades.pptx" }
if (-not $Source) { $Source = Join-Path $Root "fixtures\sample.pptx" }
if (-not $Out) { $Out = Join-Path $Root ".tmp\rebrand" }
New-Item -ItemType Directory -Force -Path $Out | Out-Null

$VenvUnix = Join-Path $Root "converter\.venv\bin\python"
$VenvWin = Join-Path $Root "converter\.venv\Scripts\python.exe"
if ($env:PYTHON) { $Py = $env:PYTHON }
elseif (Test-Path $VenvWin) { $Py = $VenvWin }
elseif (Test-Path $VenvUnix) { $Py = $VenvUnix }
else { $Py = "python" }

& $Py (Join-Path $Root "rebrand\inspect_template.py") $Template --out-dir $Out
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& $Py (Join-Path $Root "rebrand\inspect_source.py") $Source --out-dir $Out
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& $Py (Join-Path $Root "rebrand\plan.py") --out-dir $Out
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& $Py (Join-Path $Root "rebrand\build.py") --out-dir $Out --source $Source --template $Template
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
& $Py (Join-Path $Root "rebrand\qa.py") --out-dir $Out --source $Source
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Host "Wrote $Out\OUTPUT.pptx and $Out\QA\qa-report.md"
