#requires -Version 5.1
<#
 Cài / gỡ add-in "Soạn thảo chuẩn" vào thư mục STARTUP của Word (chỉ người dùng hiện tại, không cần quyền admin).
 Chạy: powershell -ExecutionPolicy Bypass -File install.ps1            (cài hoặc cập nhật, có hỏi xác nhận)
       powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall  (gỡ)
 Tham số: -Yes (không hỏi), -DryRun (chỉ in việc sẽ làm), -Target <thư mục> (thử nghiệm, mặc định là STARTUP của Word).
 Nguồn: SoanThaoChuan.dotm + thư mục config cạnh script, hoặc ..\dist khi chạy từ kho mã.
 Cấu hình cũ (kể cả spell.vi.ini bạn đã sửa) được sao lưu vào %APPDATA%\SoanThaoChuan\backup\<ngày-giờ> trước khi ghi đè.
 Không chạm tới cài đặt bảo mật/registry của Word. Chỉ ghi: <STARTUP>\SoanThaoChuan.dotm và <STARTUP>\config\.
#>
param(
    [switch]$Uninstall,
    [switch]$Yes,
    [switch]$DryRun,
    [string]$Target = (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP')
)
$ErrorActionPreference = 'Stop'

function Confirm-Step([string]$msg) {
    if ($Yes -or $DryRun) { return $true }
    $a = Read-Host "$msg [y/N]"
    return ($a -match '^(y|yes|c|co|có)$')
}

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Word đang mở. Hãy lưu và đóng Word rồi chạy lại.'
}

$dotmName = 'SoanThaoChuan.dotm'
$dstDotm = Join-Path $Target $dotmName
$dstCfg  = Join-Path $Target 'config'
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$bak = Join-Path $env:APPDATA "SoanThaoChuan\backup\$stamp"

if ($Uninstall) {
    "Sẽ gỡ: $dstDotm và $dstCfg (sao lưu config vào $bak)"
    if (-not (Test-Path $dstDotm) -and -not (Test-Path $dstCfg)) { 'Không thấy add-in trong thư mục đích. Không có gì để gỡ.'; return }
    if (-not (Confirm-Step 'Gỡ add-in?')) { 'Đã hủy.'; return }
    if ($DryRun) { '[DryRun] không thay đổi gì.'; return }
    if (Test-Path $dstCfg) { New-Item -ItemType Directory -Force $bak | Out-Null; Copy-Item $dstCfg $bak -Recurse -Force; Remove-Item $dstCfg -Recurse -Force }
    if (Test-Path $dstDotm) { Remove-Item $dstDotm -Force }
    "Đã gỡ. Bản sao lưu cấu hình: $bak"
    return
}

# --- tìm nguồn
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcDir = $null
foreach ($c in @($here, (Join-Path (Split-Path $here -Parent) 'dist'))) {
    if ((Test-Path (Join-Path $c $dotmName)) -and (Test-Path (Join-Path $c 'config'))) { $srcDir = $c; break }
}
if (-not $srcDir) { throw "Không tìm thấy $dotmName và thư mục config cạnh script hoặc trong ..\dist. Hãy build (tools\build.ps1) hoặc dùng gói đã đóng." }
$srcDotm = Join-Path $srcDir $dotmName
$srcCfg = Join-Path $srcDir 'config'

# --- kiểm tra tệp nguồn đủ
foreach ($f in 'messages.vi.ini','refs.vi.ini','spell.vi.ini','bigram.vi.txt') {
    if (-not (Test-Path (Join-Path $srcCfg $f))) { throw "Thiếu tệp cấu hình: $f" }
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zipSig = [IO.Compression.ZipFile]::OpenRead($srcDotm)
try { $signed = ($null -ne $zipSig.GetEntry('word/vbaProjectSignature.bin')) } finally { $zipSig.Dispose() }
"Nguồn : $srcDir"
if ($signed) { 'Chữ ký số VBA: CÓ (kiểm tra người ký trong Word: Developer > Macro Security / VBE > Tools > Digital Signature).' }
else { 'Chữ ký số VBA: CHƯA CÓ - Word có thể chặn macro; xem docs\12-KY-SO.md.' }
"Đích   : $Target"
if (Test-Path $dstDotm) { 'Đang có bản cũ -> sẽ cập nhật.' } else { 'Chưa có -> cài mới.' }
if (Test-Path $dstCfg) { "Cấu hình hiện có sẽ được sao lưu vào: $bak" }
'Lưu ý: nếu Word chặn macro, cần đặt thư mục STARTUP (hoặc tệp) vào vùng tin cậy / ký số - xem docs\11-HUONG-DAN-SU-DUNG.md.'
if (-not (Confirm-Step 'Cài add-in?')) { 'Đã hủy.'; return }
if ($DryRun) { '[DryRun] không thay đổi gì.'; return }

New-Item -ItemType Directory -Force $Target | Out-Null
if (Test-Path $dstCfg) { New-Item -ItemType Directory -Force $bak | Out-Null; Copy-Item $dstCfg $bak -Recurse -Force }
Copy-Item $srcDotm $dstDotm -Force
Copy-Item $srcCfg $Target -Recurse -Force
"Đã cài: $dstDotm"
'Mở lại Word -> thẻ "Soạn thảo chuẩn".'
