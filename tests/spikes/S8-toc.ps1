# Spike S8 (lần 2): TOC theo style tùy biến. Chèn đoạn Normal ở đầu rồi mới đặt field.
. "$PSScriptRoot\..\..\tools\lib.ps1"
$w = $null
try {
    $w = Start-TestWord
    $d = $w.Documents.Add()
    foreach ($n in '00Chuong','01Dieu','02Khoan') { try { [void]$d.Styles.Item($n) } catch { [void]$d.Styles.Add($n, 1) } }
    $d.Content.Text = (@('MUC LUC','Chuong mot','Dieu mot','Khoan','Dieu hai') -join "`r")
    $s = 'Normal','00Chuong','01Dieu','02Khoan','01Dieu'
    for ($i=1; $i -le 5; $i++) { $d.Paragraphs.Item($i).Style = $s[$i-1] }
    $d.Styles.Item('00Chuong').ParagraphFormat.OutlineLevel = 1
    $d.Styles.Item('01Dieu').ParagraphFormat.OutlineLevel = 2
    $sep = $w.International(17)
    $r = $d.Paragraphs.Item(1).Range; $r.MoveEnd(1,-1) | Out-Null; $r.Collapse(0)
    $code = 'TOC \h \z \t "00Chuong{0}1{0}01Dieu{0}2"' -f $sep
    $f = $d.Fields.Add($r, -1, $code, $false); [void]$f.Update()
    'A) \t only -> ' + ($f.Result.Text -replace "`r",' / ')
    $f.Code.Text = ' TOC \h \z \t "00Chuong{0}1{0}01Dieu{0}2" \u ' -f $sep
    [void]$f.Update()
    'B) \t + \u  -> ' + ($f.Result.Text -replace "`r",' / ')
    $f.Code.Text = ' TOC \h \z \u '
    [void]$f.Update()
    'C) \u only (outline level) -> ' + ($f.Result.Text -replace "`r",' / ')
    $d.Close(0)
} finally { Stop-TestWord $w }
