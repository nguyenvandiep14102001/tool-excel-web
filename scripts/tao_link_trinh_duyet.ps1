# Tạo file shortcut .url trên Desktop — user chỉ double-click mở trình duyệt
param(
    [string]$ServerUrl = ""
)

if (-not $ServerUrl) {
    $ServerUrl = Read-Host "Nhap link server (vd: http://192.168.1.50:8501)"
}

if ($ServerUrl -notmatch "^https?://") {
    Write-Host "Link phai bat dau bang http:// hoac https://"
    exit 1
}

$desktop = [Environment]::GetFolderPath("Desktop")
$urlFile = Join-Path $desktop "Excel Analyst.url"

@"
[InternetShortcut]
URL=$ServerUrl
IconIndex=0
"@ | Set-Content -Path $urlFile -Encoding ASCII

Write-Host ""
Write-Host "Da tao: $urlFile"
Write-Host "Gui file nay cho user — ho chi can double-click."
