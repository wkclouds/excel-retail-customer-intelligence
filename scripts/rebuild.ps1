$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$runtimeRoot = Join-Path ([Environment]::GetFolderPath('UserProfile')) '.cache\codex-runtimes\codex-primary-runtime\dependencies'
$pythonExe = Join-Path $runtimeRoot 'python\python.exe'
$nodeExe = Join-Path $runtimeRoot 'node\bin\node.exe'
$runtimeModules = Join-Path $runtimeRoot 'node\node_modules'
if (!(Test-Path -LiteralPath $nodeExe) -or !(Test-Path -LiteralPath $pythonExe)) {
    throw 'The Codex dependency runtime is required to rebuild this workbook. The delivered XLSX opens independently in Excel.'
}
$junction = Join-Path $projectRoot 'node_modules'
if (!(Test-Path -LiteralPath $junction)) {
    New-Item -ItemType Junction -Path $junction -Target $runtimeModules | Out-Null
}
Push-Location $projectRoot
try {
    & $pythonExe scripts/prepare_data.py
    if ($LASTEXITCODE -ne 0) { throw 'Data preparation failed.' }
    & $nodeExe scripts/build_workbook.mjs
    if ($LASTEXITCODE -ne 0) { throw 'Workbook build failed.' }
    & $pythonExe scripts/validate_export.py
    if ($LASTEXITCODE -ne 0) { throw 'Export validation failed.' }
} finally {
    Pop-Location
}
