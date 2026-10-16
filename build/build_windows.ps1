# Build восьминіг for Windows: a portable folder + zip, and a setup.exe if Inno Setup 6 is installed.
#   powershell -ExecutionPolicy Bypass -File build\build_windows.ps1
# Needs Python 3.10+ (python.org installer, "Add to PATH"), internet for pip the first time.
# Output: dist\vosmynih-<version>-windows-x64.zip  and  dist\vosmynih-<version>-windows-x64-setup.exe
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
$Version = (Select-String -Path app\vosmynih_app.py -Pattern '^VERSION = "(.*)"').Matches[0].Groups[1].Value
$Cache = if ($env:VOSMYNIH_CACHE) { $env:VOSMYNIH_CACHE } else { Join-Path $env:LOCALAPPDATA "vosmynih" }
$Venv = Join-Path $Cache "build-venv"

Write-Host "==> build environment in $Cache"
if (-not (Test-Path $Venv)) { python -m venv $Venv }
& "$Venv\Scripts\python.exe" -m pip install -q --upgrade pip
& "$Venv\Scripts\python.exe" -m pip install -q -r requirements-app.txt pyinstaller fonttools

Write-Host "==> assets"
& "$Venv\Scripts\python.exe" tools\make_assets.py

Write-Host "==> bundle"
$Work = Join-Path $Cache "pyinstaller"
& "$Venv\Scripts\pyinstaller.exe" --noconfirm --clean --windowed --name vosmynih `
  --icon "$Root\app\assets\icon.ico" --add-data "$Root\app\ui;ui" --paths "$Root" --collect-data pptx `
  --exclude-module cairosvg --exclude-module cairocffi --exclude-module tkinter `
  --workpath $Work --specpath $Work --distpath dist "$Root\app\vosmynih_app.py"

if ($env:SELFTEST_DECK) {
  Write-Host "==> self-test with $env:SELFTEST_DECK"
  & "dist\vosmynih\vosmynih.exe" --selftest $env:SELFTEST_DECK "$Root\dist\selftest.pptx" | Select-String SELFTEST
}

Write-Host "==> zip"
$Zip = "dist\vosmynih-$Version-windows-x64.zip"
if (Test-Path $Zip) { Remove-Item $Zip }
Compress-Archive -Path dist\vosmynih\* -DestinationPath $Zip

$Iscc = @("${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe", "$env:ProgramFiles\Inno Setup 6\ISCC.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($Iscc) {
  Write-Host "==> installer"
  & $Iscc "/DAppVersion=$Version" "/DSourceDir=$Root\dist\vosmynih" "/O$Root\dist" build\installer.iss
} else {
  Write-Host "Inno Setup 6 not found: skipped setup.exe (portable zip is ready). Get it at https://jrsoftware.org/isdl.php"
}
Write-Host "==> done"
