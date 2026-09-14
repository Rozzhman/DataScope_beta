param(
    [string]$IsccPath = ""
)

$ErrorActionPreference = "Stop"

$issFile = "installer\DataScope.iss"
if (-not (Test-Path $issFile)) {
    throw "Inno script not found: $issFile"
}

if (-not (Test-Path "dist\DataScope.exe")) {
    throw "Onefile EXE not found: dist\DataScope.exe. Run build_onefile.ps1 first."
}

if (-not (Test-Path ".env")) {
    throw ".env not found in project root. Installer expects it to package app configuration."
}

if (-not $IsccPath) {
    $candidates = @(
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe"
    )
    foreach ($p in $candidates) {
        if (Test-Path $p) {
            $IsccPath = $p
            break
        }
    }
}

if (-not $IsccPath) {
    throw "ISCC.exe not found. Install Inno Setup 6 and rerun."
}

& $IsccPath $issFile

Write-Host "Done. Installer path: dist\installer\DataScopeSetup.exe"
