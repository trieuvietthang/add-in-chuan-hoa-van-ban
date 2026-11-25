# Test M1: KB-A1 (Hợp đồng), KB-A2 (Quy định), KB-A7 (Công văn), KB-A3 (tăng/giảm cấp), KB-A4 (chuyển số gõ tay),
# idempotent, Undo, chữ không đổi. Văn bản tổng hợp. Dot-source từ run-tests.ps1; trả về mảng kết quả.

function Invoke-Macro($w, [string]$name, [object[]]$margs) {
    # Application.Run tìm module theo tài liệu đang active -> kích hoạt tài liệu add-in trước mỗi lần gọi
    if ($global:AddinDoc) { $global:AddinDoc.Activate() }
    return $w.GetType().InvokeMember('Run', 'InvokeMethod', $null, $w, (@($name) + $margs))
}
function New-TestDoc($w, [string[]]$texts) { $d = $w.Documents.Add(); $d.Content.Text = ($texts -join "`r"); return $d }
function Get-ListStrings($d) { @(1..$d.Paragraphs.Count | ForEach-Object { [string]$d.Paragraphs.Item($_).Range.ListFormat.ListString }) }
function Get-Styles($d) { @(1..$d.Paragraphs.Count | ForEach-Object { [string]$d.Paragraphs.Item($_).Style.NameLocal }) }
function Get-Texts($d) { @(1..$d.Paragraphs.Count | ForEach-Object { ([string]$d.Paragraphs.Item($_).Range.Text) -replace '[\r\a]', '' }) }
function Set-Level($w, $d, [int]$i, [int]$level, [string]$profile) {
    [void](Invoke-Macro $w 'modApi.ApplyLevelParas' @([string]$d.Name, $i, $i, $level, $profile))
}
function Join-Arr($a) { ($a -join ' | ') }

function Invoke-M1Tests($w) {
    $res = New-Object System.Collections.Generic.List[object]
    function Add-Res($n, $ok, $d = '') { $res.Add([pscustomobject]@{ Name = $n; Ok = [bool]$ok; Detail = $d }); [Console]::Error.WriteLine("  .. $n $ok $d") }
    function Check($n, $exp, $act) { $e = Join-Arr $exp; $a = Join-Arr $act; Add-Res $n ($e -eq $a) $(if ($e -ne $a) { "mong đợi [$e] thực tế [$a]" } else { '' }) }

    $body = 'Pham vi', 'Khoan mot', 'Khoan hai', 'Diem a', 'Tiet i', 'Dieu hai', 'Khoan ba'
    $hd = 1, 2, 2, 3, 4, 1, 2

    # --- KB-A1 Hợp đồng
    try {
        $d = New-TestDoc $w $body
        for ($i = 1; $i -le 7; $i++) { Set-Level $w $d $i $hd[$i-1] 'HopDong' }
        $ls = Get-ListStrings $d
        Check 'KB-A1 happy' @('Điều 1.', '1.1', '1.2', 'a)', '(i)', 'Điều 2.', '2.1') $ls
        Check 'KB-A1 style' @('01Dieu', '02Khoan', '02Khoan', '03Diem', '04Tiet', '01Dieu', '02Khoan') (Get-Styles $d)
        Check 'KB-A1 text-unchanged' $body (Get-Texts $d)
        $lt1 = $d.ListTemplates.Count
        # idempotent: áp lần 2
        for ($i = 1; $i -le 7; $i++) { Set-Level $w $d $i $hd[$i-1] 'HopDong' }
        Check 'KB-A1 idempotent' $ls (Get-ListStrings $d)
        Add-Res 'KB-A1 idempotent-lt' ($d.ListTemplates.Count -eq $lt1) "ListTemplates $lt1 -> $($d.ListTemplates.Count)"
        # chèn thêm Điều giữa -> số sau tự cập nhật
        $d.Paragraphs.Item(4).Range.InsertParagraphBefore()
        Set-Level $w $d 4 1 'HopDong'
        Check 'KB-A1 insert-renumber' @('Điều 1.', '1.1', '1.2', 'Điều 2.', 'a)', '(i)', 'Điều 3.', '3.1') (Get-ListStrings $d)
        $d.Close(0)
        # undo
        $d = New-TestDoc $w $body
        $d.UndoClear()
        Set-Level $w $d 1 1 'HopDong'
        [void]$d.Undo(1)
        Check 'KB-A1 undo' @('Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal') (Get-Styles $d)
        $d.Close(0)
    } catch { Add-Res 'KB-A1' $false $_.Exception.Message }

    # --- KB-A2 Quy định: Chương/Điều/Khoản (Điều liên tục xuyên Chương)
    try {
        $t = 'Quy dinh chung', 'Pham vi', 'Khoan', 'Diem', 'Chuong hai', 'Dieu moi', 'Khoan moi'
        $lv = 1, 3, 4, 5, 1, 3, 4
        $d = New-TestDoc $w $t
        for ($i = 1; $i -le 7; $i++) { Set-Level $w $d $i $lv[$i-1] 'QuyDinh' }
        Check 'KB-A2 happy' @('CHƯƠNG I', 'Điều 1.', '1.', 'a)', 'CHƯƠNG II', 'Điều 2.', '1.') (Get-ListStrings $d)
        $d.Close(0)
    } catch { Add-Res 'KB-A2' $false $_.Exception.Message }

    # --- KB-A7 Công văn
    try {
        $t = 'Muc lon', 'Muc nho', 'Y chi tiet', 'Diem', 'Tiet', 'Gach', 'Muc nho hai'
        $lv = 1, 2, 3, 4, 5, 6, 2
        $d = New-TestDoc $w $t
        for ($i = 1; $i -le 7; $i++) { Set-Level $w $d $i $lv[$i-1] 'CongVan' }
        Check 'KB-A7 happy' @('I.', '1.', '1.1.', 'a)', '(i)', '-', '2.') (Get-ListStrings $d)
        $d.Close(0)
    } catch { Add-Res 'KB-A7' $false $_.Exception.Message }

    # --- KB-A3 tăng/giảm cấp
    try {
        $d = New-TestDoc $w @('Dieu', 'Khoan', 'Khoan 2')
        Set-Level $w $d 1 1 'HopDong'; Set-Level $w $d 2 2 'HopDong'; Set-Level $w $d 3 2 'HopDong'
        [void](Invoke-Macro $w 'modApi.ChangeLevelParas' @([string]$d.Name, 3, 3, 1, 'HopDong'))
        Check 'KB-A3 tang-cap' @('Điều 1.', '1.1', 'a)') (Get-ListStrings $d)
        [void](Invoke-Macro $w 'modApi.ChangeLevelParas' @([string]$d.Name, 3, 3, -1, 'HopDong'))
        Check 'KB-A3 giam-cap' @('Điều 1.', '1.1', '1.2') (Get-ListStrings $d)
        $d.Close(0)
    } catch { Add-Res 'KB-A3' $false $_.Exception.Message }

    # --- KB-A4 chuyển số gõ tay -> list thật (nhảy số 1,2,4 được đánh lại)
    try {
        $typed = 'Điều 1. Pham vi', '1.1 Khoan mot', '1.2. Khoan hai', 'a) diem', '(i) tiet', 'Điều 4. Doi tuong', '4.1 Khoan', 'Doan van thuong khong danh so'
        $d = New-TestDoc $w $typed
        $pv = [string](Invoke-Macro $w 'modApi.PreviewTypedParas' @([string]$d.Name, 0, 0, 'HopDong'))
        Add-Res 'KB-A4 preview' ($pv -ne '' -and ($pv -split "`n").Count -ge 7) "dòng xem trước: $(($pv -split "`n").Count)"
        $n = (Invoke-Macro $w 'modApi.ConvertTypedParas' @([string]$d.Name, 0, 0, 'HopDong'))
        Add-Res 'KB-A4 count' ($n -eq 7) "đã chuyển $n"
        Check 'KB-A4 list' @('Điều 1.', '1.1', '1.2', 'a)', '(i)', 'Điều 2.', '2.1', '') (Get-ListStrings $d)
        Check 'KB-A4 text-sau-bo-so' @('Pham vi', 'Khoan mot', 'Khoan hai', 'diem', 'tiet', 'Doi tuong', 'Khoan', 'Doan van thuong khong danh so') (Get-Texts $d)
        $d.Close(0)
    } catch { Add-Res 'KB-A4' $false $_.Exception.Message }

    # --- KB-A4 undo: một Ctrl+Z khôi phục chữ gõ tay và style
    try {
        $typed = 'Điều 1. Pham vi', '1.1 Khoan mot', 'Doan van'
        $d = New-TestDoc $w $typed
        $d.UndoClear()
        [void](Invoke-Macro $w 'modApi.ConvertTypedParas' @([string]$d.Name, 0, 0, 'HopDong'))
        [void]$d.Undo(1)
        Check 'KB-A4 undo' $typed (Get-Texts $d)
        Check 'KB-A4 undo-style' @('Normal', 'Normal', 'Normal') (Get-Styles $d)
        $d.Close(0)
    } catch { Add-Res 'KB-A4 undo' $false $_.Exception.Message }

    # --- S2 trong thực tế: đổi hồ sơ trên cùng tài liệu (Điều: cấp 1 -> cấp 3)
    try {
        $d = New-TestDoc $w @('Dieu', 'Khoan')
        Set-Level $w $d 1 1 'HopDong'; Set-Level $w $d 2 2 'HopDong'
        Set-Level $w $d 1 3 'QuyDinh'; Set-Level $w $d 2 4 'QuyDinh'
        Check 'PROFILE switch HD->QD' @('Điều 1.', '1.') (Get-ListStrings $d)
        Set-Level $w $d 1 1 'HopDong'; Set-Level $w $d 2 2 'HopDong'
        Check 'PROFILE switch QD->HD' @('Điều 1.', '1.1') (Get-ListStrings $d)
        $d.Close(0)
    } catch { Add-Res 'PROFILE switch' $false $_.Exception.Message }

    # --- Lưu .docx, mở lại, áp lần nữa: không sinh thêm ListTemplate, số vẫn đúng
    try {
        $d = New-TestDoc $w $body
        for ($i = 1; $i -le 7; $i++) { Set-Level $w $d $i $hd[$i-1] 'HopDong' }
        $p = [string](Join-Path ([IO.Path]::GetTempPath()) ('m1-' + [guid]::NewGuid().ToString('N').Substring(0,6) + '.docx'))
        $d.SaveAs2($p, 16); $c1 = $d.ListTemplates.Count; $d.Close(0)
        $d = $w.Documents.Open($p)
        Set-Level $w $d 1 1 'HopDong'
        Add-Res 'PERSIST reopen-lt' ($d.ListTemplates.Count -eq $c1) "ListTemplates $c1 -> $($d.ListTemplates.Count)"
        Check 'PERSIST reopen-list' @('Điều 1.', '1.1', '1.2', 'a)', '(i)', 'Điều 2.', '2.1') (Get-ListStrings $d)
        $d.Close(0); Remove-Item $p -ErrorAction SilentlyContinue
    } catch { Add-Res 'PERSIST' $false $_.Exception.Message }

    return $res
}
