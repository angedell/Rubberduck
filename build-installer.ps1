# Compiles Rubberduck.Deployment\InnoSetup\Rubberduck.Installer.Build.iss with Inno Setup 6 or 7.
# Inno Setup 7 rejects two things in the checked-in script, so a temporary patched copy is compiled
# and then deleted. The tracked script is never modified.
#   1. "ignoreversion" combined with "replacesameversion" (ignoreversion already implies it)
#   2. "type HINSTANCE = THandle;" (predefined in Inno Setup 7)
$ErrorActionPreference = 'Stop'

$innoDir = Join-Path $PSScriptRoot 'Rubberduck.Deployment\InnoSetup'
$script  = Join-Path $innoDir 'Rubberduck.Installer.Build.iss'
$tmp     = Join-Path $innoDir 'Rubberduck.Installer.Build.tmp.iss'
$exe     = Join-Path $innoDir 'Installers\Rubberduck.Setup.exe'

$iscc = @(
    'C:\Program Files\Inno Setup 7\ISCC.exe',
    'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $iscc) { Write-Host 'Inno Setup (6 or 7) not found.'; exit 1 }
if (-not (Test-Path (Join-Path $PSScriptRoot 'Rubberduck.Deployment\bin\Rubberduck.dll'))) {
    Write-Host 'Rubberduck.Deployment\bin\Rubberduck.dll is missing - build the solution first.'; exit 1
}

if (Test-Path $exe) { Remove-Item $exe }

$text = Get-Content $script -Raw -Encoding UTF8
if ($iscc -like '*Inno Setup 7*') {
    $text = $text -replace '(ignoreversion[^;]*?)\s+replacesameversion', '$1'
    $text = $text -replace '(?m)^type\r?\n\s+HINSTANCE = THandle;\r?\n', ''
}
Set-Content $tmp $text -Encoding UTF8 -NoNewline

try {
    & $iscc /Q $tmp
    $rc = $LASTEXITCODE
} finally {
    Remove-Item $tmp -ErrorAction SilentlyContinue
}

if ($rc -ne 0 -or -not (Test-Path $exe)) { Write-Host "Installer build failed (ISCC exit code $rc)."; exit 1 }
exit 0
