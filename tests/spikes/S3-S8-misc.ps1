# Spike S3 (REF), S4 (xóa định dạng ký tự), S5 (proofing tiếng Việt), S6 (Undo), S8 (TOC). Văn bản tổng hợp.
. "$PSScriptRoot\..\..\tools\lib.ps1"

function New-Template($d) {
    foreach ($n in '00Chuong','01Dieu','02Khoan') { try { [void]$d.Styles.Item($n) } catch { [void]$d.Styles.Add($n, 1) } }
    $lt = $d.ListTemplates.Add($true, 'HD')
    $fmt = 'Điều %1.','%1.%2'; $sty = '01Dieu','02Khoan'
    for ($i=1; $i -le 2; $i++) { $lv=$lt.ListLevels.Item($i); $lv.NumberFormat=$fmt[$i-1]; $lv.NumberStyle=0; $lv.NumberPosition=0; $lv.TextPosition=36; $lv.TabPosition=36; $lv.LinkedStyle=$sty[$i-1] }
}
function Fill($d) {
    $d.Content.Text = (@('Pham vi','Khoan mot','Khoan hai','Doi tuong','Khoan ba') -join "`r")
    $s = '01Dieu','02Khoan','02Khoan','01Dieu','02Khoan'
    for ($i=1; $i -le 5; $i++) { $d.Paragraphs.Item($i).Style = $s[$i-1] }
}
function Try-Spike($name, [scriptblock]$sb) { "=== $name"; try { & $sb } catch { 'LỖI: ' + $_.Exception.Message } }

$w = $null
try {
    $w = Start-TestWord
    'Word ' + $w.Version + ' | wdListSeparator = [' + $w.International(17) + ']'

    Try-Spike 'S3 REF' {
        $d = $w.Documents.Add(); New-Template $d; Fill $d
        $bm = $d.Bookmarks.Add('_Ref_dieu2', $d.Paragraphs.Item(4).Range)
        $bm2 = $d.Bookmarks.Add('_Ref_k3', $d.Paragraphs.Item(5).Range)
        $d.Content.InsertParagraphAfter()
        foreach ($sw in '\r','\n','\w','\h') {
            $r = $d.Paragraphs.Item($d.Paragraphs.Count).Range; $r.Collapse(0)
            $r.InsertAfter("[$sw]=")
            $r = $d.Paragraphs.Item($d.Paragraphs.Count).Range; $r.MoveEnd(1,-1) | Out-Null; $r.Collapse(0)
            $f = $d.Fields.Add($r, -1, "REF _Ref_dieu2 $sw", $false); [void]$f.Update()
            'Điều 2 ' + $sw + ' -> [' + $f.Result.Text + ']'
            $d.Content.InsertParagraphAfter()
        }
        foreach ($sw in '\r','\w') {
            $r = $d.Paragraphs.Item($d.Paragraphs.Count).Range; $r.MoveEnd(1,-1)|Out-Null; $r.Collapse(0)
            $f = $d.Fields.Add($r, -1, "REF _Ref_k3 $sw", $false); [void]$f.Update()
            'Khoản 3 ' + $sw + ' -> [' + $f.Result.Text + ']'
            $d.Content.InsertParagraphAfter()
        }
        $d.Close(0)
    }

    Try-Spike 'S4 xóa định dạng ký tự giữ B/I/U' {
        $d = $w.Documents.Add(); $n = 400
        $d.Content.Text = ((1..$n | % { "Doan so $_ chu dam chu thuong" }) -join "`r")
        $d.Content.Font.Name = 'Arial'; $d.Content.Font.Size = 17; $d.Content.Font.Color = 255
        for ($i=1; $i -le $n; $i+=10) { $d.Paragraphs.Item($i).Range.Words.Item(3).Font.Bold = 1 }  # 40 từ đậm
        $sw = [Diagnostics.Stopwatch]::StartNew()
        # Cách A: đặt lại thuộc tính ký tự trên toàn vùng (không chạm Bold)
        $d.Content.Font.Name = 'Times New Roman'; $d.Content.Font.Size = 12; $d.Content.Font.Color = -16777216
        $ta = $sw.Elapsed.TotalSeconds
        $bold = 0; for ($i=1; $i -le $n; $i+=10) { if ($d.Paragraphs.Item($i).Range.Words.Item(3).Font.Bold -ne 0) { $bold++ } }
        'Cách A (set trên Content): {0:N2}s; font={1} size={2} color={3}; từ đậm còn giữ: {4}/40' -f $ta, $d.Paragraphs.Item(7).Range.Font.Name, $d.Paragraphs.Item(7).Range.Font.Size, $d.Paragraphs.Item(7).Range.Font.Color, $bold
        # Cách B: Font.Reset() — để so sánh
        $d.Content.Font.Reset()
        $bold2 = 0; for ($i=1; $i -le $n; $i+=10) { if ($d.Paragraphs.Item($i).Range.Words.Item(3).Font.Bold -ne 0) { $bold2++ } }
        'Cách B (Font.Reset): từ đậm còn giữ: ' + $bold2 + '/40 (kỳ vọng 0)'
        $d.Close(0)
    }

    Try-Spike 'S5 proofing + tập âm tiết' {
        $has = @()
        foreach ($lid in 1066, 1033) {
            try { $l = $w.Languages.Item($lid); $has += ('{0}: {1}; spellDict={2}' -f $lid, $l.NameLocal, ($l.ActiveSpellingDictionary -ne $null)) } catch { $has += "$lid : lỗi $($_.Exception.Message)" }
        }
        $has
        # Hiệu năng tra tập âm tiết (mô phỏng 10.000 từ, tập 8.000 mục)
        $set = New-Object 'System.Collections.Generic.HashSet[string]'
        1..8000 | % { [void]$set.Add("am$_") }
        $sw = [Diagnostics.Stopwatch]::StartNew(); $miss = 0
        1..10000 | % { if (-not $set.Contains("am$($_ % 9000)")) { $miss++ } }
        'Tra 10.000 từ: {0:N3}s (miss={1})' -f $sw.Elapsed.TotalSeconds, $miss
        $f = Join-Path (Get-TempWorkDir) 'u.txt'
        [IO.File]::WriteAllText($f, "đường`nngười`nViệt", (New-Object Text.UTF8Encoding $false))
        'Đọc UTF-8: ' + ([IO.File]::ReadAllLines($f, [Text.Encoding]::UTF8) -join ',')
    }

    Try-Spike 'S6 UndoRecord' {
        $d = $w.Documents.Add(); New-Template $d; Fill $d
        $before = @(1..5 | % { $p=$d.Paragraphs.Item($_); $p.Style.NameLocal + '|' + $p.Range.ListFormat.ListString + '|' + $p.Range.Text.Trim() }) -join ';'
        $d.Content.Font.Name = 'Verdana'
        $d.UndoClear()
        $before = @(1..5 | % { $p=$d.Paragraphs.Item($_); $p.Style.NameLocal + '|' + $p.Range.ListFormat.ListString + '|' + $p.Range.Font.Name + '|' + $p.Range.Text.Trim() }) -join ';'
        $ur = $w.UndoRecord; $ur.StartCustomRecord('Spike S6')
        $d.Content.Font.Name = 'Times New Roman'
        $d.Paragraphs.Item(2).Style = '01Dieu'
        $d.Paragraphs.Item(3).Style = '01Dieu'
        $d.PageSetup.LeftMargin = 99
        $d.Content.InsertParagraphAfter()
        $r = $d.Paragraphs.Item($d.Paragraphs.Count).Range; [void]$d.Fields.Add($r, -1, 'PAGE', $false)
        $ur.EndCustomRecord()
        'Sau lệnh: ' + $d.Paragraphs.Item(2).Range.ListFormat.ListString + ' lề=' + $d.PageSetup.LeftMargin + ' fields=' + $d.Fields.Count
        $okU = $d.Undo(1)
        $after = @(1..5 | % { $p=$d.Paragraphs.Item($_); $p.Style.NameLocal + '|' + $p.Range.ListFormat.ListString + '|' + $p.Range.Font.Name + '|' + $p.Range.Text.Trim() }) -join ';'
        'Undo(1) trả: ' + $okU + '; paragraphs=' + $d.Paragraphs.Count + ' fields=' + $d.Fields.Count + ' lề=' + $d.PageSetup.LeftMargin
        'Trạng thái giống trước: ' + ($before -eq $after)
        if ($before -ne $after) { 'TRƯỚC: ' + $before; 'SAU  : ' + $after }
        $d.Close(0)
    }

    Try-Spike 'S8 TOC' {
        $d = $w.Documents.Add(); New-Template $d; Fill $d
        $sep = $w.International(17)
        $d.Content.InsertParagraphBefore()
        $r = $d.Paragraphs.Item(1).Range; $r.Collapse(1)
        $code = 'TOC \h \z \t "01Dieu{0}1{0}00Chuong{0}1" \u' -f $sep
        $f = $d.Fields.Add($r, -1, $code, $false); [void]$f.Update()
        'Mã: ' + $f.Code.Text
        'Kết quả TOC: ' + ($f.Result.Text -replace "`r",' / ')
        $d.Close(0)
    }
} finally { Stop-TestWord $w }
