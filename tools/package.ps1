#requires -Version 5.1
# Đóng gói bản phát hành: dist\SoanThaoChuan-<phiên bản>.zip gồm .dotm, config\, install.ps1 và hướng dẫn.
# Chạy sau tools\build.ps1 và tests\run-tests.ps1 xanh. Không ký số (xem docs\11, mục Macro bị chặn).
param(
    [string]$Version = '0.6.1',
    [string]$Dotm = ''    # tệp .dotm ĐÃ KÝ (xem docs\12-KY-SO.md); bỏ trống = dist\SoanThaoChuan.dotm chưa ký
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$dist = Join-Path $root 'dist'
foreach ($f in 'SoanThaoChuan.dotm','config') { if (-not (Test-Path (Join-Path $dist $f))) { throw "Thiếu dist\$f - hãy chạy tools\build.ps1 trước." } }
$stage = Join-Path $dist "pkg-$Version"
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory $stage | Out-Null
if (-not $Dotm) { $Dotm = Join-Path $dist 'SoanThaoChuan.dotm' }
if (-not (Test-Path $Dotm)) { throw "Không thấy $Dotm" }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$z = [IO.Compression.ZipFile]::OpenRead([string]$Dotm)
try { $isSigned = ($null -ne $z.GetEntry('word/vbaProjectSignature.bin')) } finally { $z.Dispose() }
"Chữ ký số VBA trong gói: $(if ($isSigned) { 'CÓ' } else { 'CHƯA CÓ' })"
Copy-Item $Dotm (Join-Path $stage 'SoanThaoChuan.dotm')
Copy-Item (Join-Path $dist 'config') $stage -Recurse
Copy-Item (Join-Path $PSScriptRoot 'install.ps1') $stage
Copy-Item (Join-Path $root 'docs\11-HUONG-DAN-SU-DUNG.md') (Join-Path $stage 'HUONG-DAN.md')
$zip = Join-Path $dist "SoanThaoChuan-$Version.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Remove-Item $stage -Recurse -Force
"{0} ({1:N1} MB)" -f $zip, ((Get-Item $zip).Length / 1MB)
