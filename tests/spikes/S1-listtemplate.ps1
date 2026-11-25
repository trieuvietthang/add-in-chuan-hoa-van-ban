# Spike S1/S2: ListTemplate đa cấp gắn style; một style chuyển cấp giữa hai hồ sơ. Chỉ dùng văn bản tổng hợp.
. "$PSScriptRoot\..\..\tools\lib.ps1"
$w = $null
try {
    $w = Start-TestWord
    $d = $w.Documents.Add()
    # Style theo ADR-007
    $names = '00Chuong','01Dieu','02Khoan','03Diem','04Tiet'
    foreach ($n in $names) { try { [void]$d.Styles.Item($n) } catch { [void]$d.Styles.Add($n, 1) } }

    # --- Hồ sơ HopDong: Điều=cấp1, Khoản=cấp2 (%1.%2), Điểm=cấp3 (a), Tiết=cấp4 (i)
    $lt = $d.ListTemplates.Add($true, 'HopDong')
    $fmt = @('Điều %1.', '%1.%2', '%3)', '(%4)')
    $sty = @('01Dieu','02Khoan','03Diem','04Tiet')
    $nst = @(0,0,4,2)   # wdListNumberStyleArabic=0, LowercaseLetter=4, LowercaseRoman=2
    for ($i=1; $i -le 4; $i++) {
        $lv = $lt.ListLevels.Item($i)
        $lv.NumberFormat = $fmt[$i-1]
        $lv.NumberStyle = $nst[$i-1]
        $lv.TrailingCharacter = 0   # tab
        $lv.NumberPosition = 0
        $lv.TextPosition = 36
        $lv.TabPosition = 36
        $lv.LinkedStyle = $sty[$i-1]
    }
    # Nội dung giả
    $content = @('Tieu de dieu A','Khoan 1','Khoan 2','Diem a','Tiet i','Tieu de dieu B','Khoan 1')
    $stl = @('01Dieu','02Khoan','02Khoan','03Diem','04Tiet','01Dieu','02Khoan')
    $d.Content.Text = ($content -join "`r")
    for ($i=1; $i -le $content.Count; $i++) { $d.Paragraphs.Item($i).Style = $stl[$i-1] }
    $out = for ($i=1; $i -le $content.Count; $i++) { $p=$d.Paragraphs.Item($i); '{0} | {1} | lvl {2}' -f $p.Range.ListFormat.ListString, $p.Style.NameLocal, $p.Range.ListFormat.ListLevelNumber }
    'S1 kết quả HopDong:'; $out

    # --- Idempotent: dùng lại template, kiểm tra ListTemplates.Count
    'ListTemplates trong tài liệu: ' + $d.ListTemplates.Count

    # --- S2: hồ sơ QuyDinh — Điều = cấp 3 (Chương=1, Mục=2, Điều=3). Gắn lại 01Dieu sang template khác
    $lt2 = $d.ListTemplates.Add($true, 'QuyDinh')
    $fmt2 = @('CHƯƠNG %1','Mục %2.','Điều %3.','%4.')
    $sty2 = @('00Chuong','','01Dieu','02Khoan')
    $nst2 = @(1,0,0,0)  # Chương: Roman uppercase=1
    for ($i=1; $i -le 4; $i++) {
        $lv = $lt2.ListLevels.Item($i)
        $lv.NumberFormat = $fmt2[$i-1]; $lv.NumberStyle = $nst2[$i-1]
        $lv.NumberPosition = 0; $lv.TextPosition = 36; $lv.TabPosition = 36
        if ($sty2[$i-1]) { $lv.LinkedStyle = $sty2[$i-1] }
    }
    # Điều đếm liên tục xuyên Chương: ResetOnHigher của cấp 3 = 0
    $lt2.ListLevels.Item(3).ResetOnHigher = 0
    $d.Content.Text = (@('Chuong dau','Dieu 1','Khoan','Chuong hai','Dieu 2') -join "`r")
    $s3 = '00Chuong','01Dieu','02Khoan','00Chuong','01Dieu'
    for ($i=1; $i -le 5; $i++) { $d.Paragraphs.Item($i).Style = $s3[$i-1] }
    'S2 kết quả QuyDinh (sau khi gắn lại style):'
    for ($i=1; $i -le 5; $i++) { $p=$d.Paragraphs.Item($i); '{0} | {1} | lvl {2}' -f $p.Range.ListFormat.ListString, $p.Style.NameLocal, $p.Range.ListFormat.ListLevelNumber }
    $d.Close(0)
} catch { 'LỖI: ' + $_.Exception.Message; $_.ScriptStackTrace }
finally { Stop-TestWord $w }
