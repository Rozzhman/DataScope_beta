param(
    [string]$PythonExe = "python",
    [string]$AppName = "DataScope",
    [string]$EntryPoint = "dorabotka.py"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $EntryPoint)) {
    throw "Entry point not found: $EntryPoint"
}

$pyBase = & $PythonExe -c "import sys; print(sys.base_prefix)"
$tclDir = Join-Path $pyBase "tcl\tcl8.6"
$tkDir = Join-Path $pyBase "tcl\tk8.6"

$iconPath = ""
if (Test-Path "icon_result.ico") {
    $iconPath = "icon_result.ico"
} elseif (Test-Path "absolute_v5/ui/assets/icon_result.ico") {
    $iconPath = "absolute_v5/ui/assets/icon_result.ico"
} elseif (Test-Path "icon.ico") {
    $iconPath = "icon.ico"
} elseif (Test-Path "absolute_v5/ui/assets/icon.ico") {
    $iconPath = "absolute_v5/ui/assets/icon.ico"
}

$cmd = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--windowed",
    "--clean",
    "--name", $AppName,
    "--hidden-import", "tkinter",
    "--hidden-import", "_tkinter",
    "--collect-data", "tkinter",
    "--add-data", "absolute_v5/ui/assets;absolute_v5/ui/assets",
    $EntryPoint
)

if (Test-Path $tclDir) {
    $cmd += @("--add-data", "$tclDir;tcl/tcl8.6")
}
if (Test-Path $tkDir) {
    $cmd += @("--add-data", "$tkDir;tcl/tk8.6")
}

if ($iconPath -ne "") {
    $cmd = @(
        "-m", "PyInstaller",
        "--noconfirm",
        "--windowed",
        "--clean",
        "--name", $AppName,
        "--hidden-import", "tkinter",
        "--hidden-import", "_tkinter",
        "--collect-data", "tkinter",
        "--add-data", "absolute_v5/ui/assets;absolute_v5/ui/assets",
        "--icon", $iconPath,
        $EntryPoint
    )

    if (Test-Path $tclDir) {
        $cmd += @("--add-data", "$tclDir;tcl/tcl8.6")
    }
    if (Test-Path $tkDir) {
        $cmd += @("--add-data", "$tkDir;tcl/tk8.6")
    }
}

& $PythonExe @cmd
