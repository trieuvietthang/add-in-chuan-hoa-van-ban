# Thư viện dùng chung cho tools/ và tests/ (PowerShell 5.1). Dot-source: . "$PSScriptRoot\lib.ps1"
# Luật: không chạm vào phiên Word của người dùng; mọi thao tác trên bản sao ở thư mục tạm.

$script:RepoRoot = Split-Path -Parent $PSScriptRoot

# LƯU Ý COM: đường dẫn truyền vào Word (SaveAs2/Open) phải là [string] thuần. Kết quả Join-Path là PSObject -> SaveAs2 treo vô hạn (spike S7).
function Get-TempWorkDir {
    $d = Join-Path ([IO.Path]::GetTempPath()) ('wordaddin-' + [guid]::NewGuid().ToString('N').Substring(0,8))
    New-Item -ItemType Directory -Force $d | Out-Null
    return $d
}

function Assert-NoUserWord {
    # Từ chối chạy nếu Word của người dùng đang mở (tránh COM dính vào phiên đó).
    for ($i = 0; $i -lt 10 -and (Get-Process WINWORD -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 500 }  # chờ tiến trình vừa Quit thoát hẳn
    $p = Get-Process WINWORD -ErrorAction SilentlyContinue
    if ($p) {
        throw "Word đang chạy (PID $($p.Id -join ',')). Hãy lưu và đóng Word rồi chạy lại; script không tự đóng Word của bạn."
    }
}

function Start-TestWord([switch]$AllowMacros) {
    Assert-NoUserWord
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $w.DisplayAlerts = 0   # wdAlertsNone
    # ForceDisable (3) làm treo SaveAs2 khi tài liệu có VBA (spike S7) -> chỉ bật khi không build/lưu macro
    if (-not $AllowMacros) { $w.AutomationSecurity = 3 }
    return $w
}

function Stop-TestWord($w) {
    if ($null -eq $w) { return }
    try { $w.Quit([ref]0) } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w) } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # Tiến trình mồ côi do chính script tạo: chỉ dọn khi không còn cửa sổ nào (an toàn vì Assert-NoUserWord đã bảo đảm không có Word của người dùng lúc bắt đầu)
    Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { -not $_.MainWindowHandle } | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Write-Result($name, $ok, $detail = '') {
    $s = if ($ok) { 'PASS' } else { 'FAIL' }
    '{0,-28} {1}  {2}' -f $name, $s, $detail
}
