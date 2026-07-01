# Tao shortcut tren Desktop — goi tu "Cai dat lan dau.bat"
$root = Split-Path $PSScriptRoot -Parent
$bat = Join-Path $root "Mo Excel Analyst.bat"

if (-not (Test-Path $bat)) {
    Write-Host "  Khong tim thay Mo Excel Analyst.bat"
    exit 0
}

$desktop = [Environment]::GetFolderPath("Desktop")
$lnk = Join-Path $desktop "Excel Analyst.lnk"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($lnk)
$shortcut.TargetPath = $bat
$shortcut.WorkingDirectory = $root
$shortcut.WindowStyle = 1
$shortcut.Description = "Excel Analyst Tool — Data Tools Suite"
$shortcut.Save()

Write-Host "  Da tao shortcut: $lnk"
