# Test M2: KB-B1..B3, KB-F1 (chuẩn hóa trang/phông/giãn dòng/số trang), idempotent, Undo, chữ không đổi.
# Văn bản tổng hợp. Dot-source từ run-tests.ps1 (dùng Invoke-Macro, Get-* của m1-numbering.ps1).

function Get-Snap($d) {
    $ps = @(1..$d.Paragraphs.Count | ForEach-Object {
            $p = $d.Paragraphs.Item($_)
            '{0}/{1}/{2:N1}/{3:N1}/{4:N1}' -f $p.Range.Font.Name, $p.Range.Font.Size, $p.LineSpacing, $p.SpaceAfter, $p.SpaceBefore })
    $s = $d.Sections.Item(1).PageSetup
    $ft = [int]$d.Sections.Item(1).Footers.Item(1).Range.Fields.Count
    return ($ps -join ';') + ('#{0:N1}/{1:N1}/{2:N1}/{3:N1}/{4:N1}#F{5}' -f $s.PageWidth, $s.TopMargin, $s.BottomMargin, $s.LeftMargin, $s.RightMargin, $ft)
}

function New-M2Doc($w) {
    $t = 'Tieu de', 'Pham vi ap dung', 'Khoan mot', 'Diem a', 'Doan van thuong co chu dam trong do', 'Doan cuoi'
    $d = New-TestDoc $w $t
    $lv = 1, 1, 2, 3
    for ($i = 1; $i -le 4; $i++) { Set-Level $w $d $i $lv[$i-1] 'HopDong' }
    # định dạng rác: font/cỡ lạ, giãn dòng đơn, lề Letter
    $d.Paragraphs.Item(2).Range.Font.Name = 'Calibri'; $d.Paragraphs.Item(2).Range.Font.Size = 14
    $r5 = $d.Paragraphs.Item(5).Range; $r5.Font.Name = 'Arial'; $r5.Font.Size = 11
    $r5.ParagraphFormat.SpaceAfter = 0; $r5.ParagraphFormat.LineSpacingRule = 0
    $wd = $d.Range($r5.Start + 20, $r5.Start + 25); $wd.Font.Bold = $true
    $d.Paragraphs.Item(6).Range.Font.Size = 16
    $ps = $d.Sections.Item(1).PageSetup
    $ps.PaperSize = 2   # Letter
    $ps.TopMargin = 72; $ps.BottomMargin = 72; $ps.LeftMargin = 72; $ps.RightMargin = 72
    [void]$d.Comments.Add($d.Paragraphs.Item(3).Range, 'nhan xet')
    [void]$d.Footnotes.Add($d.Paragraphs.Item(4).Range.Characters.Item(1), [Type]::Missing, 'chu thich')
    $d.Content.InsertParagraphAfter()
    $tb = $d.Tables.Add($d.Paragraphs.Last.Range, 1, 2)
    $tb.Cell(1, 1).Range.Text = 'O mot'; $tb.Cell(1, 2).Range.Text = 'O hai'
    $tb.Range.Font.Name = 'Verdana'; $tb.Range.Font.Size = 9
    return $d
}

function Invoke-M2Tests($w) {
    $res = New-Object System.Collections.Generic.List[object]
    function Add-Res($n, $ok, $d = '') { $res.Add([pscustomobject]@{ Name = $n; Ok = [bool]$ok; Detail = $d }); [Console]::Error.WriteLine("  .. $n $ok $d") }
    function Near($a, $b, $tol = 0.6) { [math]::Abs([double]$a - [double]$b) -le $tol }
    $cm = 28.3465

    try {
        $d = New-M2Doc $w
        $texts0 = Get-Texts $d; $ls0 = Get-ListStrings $d
        $cm0 = $d.Comments.Count; $fn0 = $d.Footnotes.Count; $tb0 = $d.Tables.Count
        $snap0 = Get-Snap $d

        # --- Xem trước: có nội dung, không sửa văn bản
        $pv = [string](Invoke-Macro $w 'modApi.PreviewNormalize' @([string]$d.Name, 'HopDong', 'ND30'))
        Add-Res 'KB-F1 preview' ($pv -match 'norm\.page' -and $pv -match 'norm\.font') ($pv -replace "`n", ' ')
        Add-Res 'KB-F1 preview-khong-sua' ((Get-Snap $d) -eq $snap0)

        # --- Áp ND30
        $d.UndoClear()
        $n = [int](Invoke-Macro $w 'modApi.NormalizeDoc' @([string]$d.Name, 'HopDong', 'ND30'))
        Add-Res 'KB-B1 count' ($n -gt 0) "thay đổi $n"
        $s = $d.Sections.Item(1).PageSetup
        Add-Res 'KB-B1 a4' ((Near $s.PageWidth (21 * $cm)) -and (Near $s.PageHeight (29.7 * $cm))) "$($s.PageWidth)x$($s.PageHeight)"
        Add-Res 'KB-B1 le' ((Near $s.TopMargin (2 * $cm)) -and (Near $s.BottomMargin (2 * $cm)) -and (Near $s.LeftMargin (3 * $cm)) -and (Near $s.RightMargin (2 * $cm))) "$($s.TopMargin)/$($s.BottomMargin)/$($s.LeftMargin)/$($s.RightMargin)"
        $bad = @(1..$d.Paragraphs.Count | Where-Object { $r = $d.Paragraphs.Item($_).Range; $r.Font.Name -ne 'Times New Roman' -or $r.Font.Size -ne 13 })
        Add-Res 'KB-B2 font-size-0-sai' ($bad.Count -eq 0) "đoạn sai: $($bad -join ',')"
        Add-Res 'KB-B2 giu-in-dam' ($d.Range($d.Paragraphs.Item(5).Range.Start + 20, $d.Paragraphs.Item(5).Range.Start + 25).Font.Bold -eq -1)
        $p5 = $d.Paragraphs.Item(5)
        Add-Res 'KB-B3 giãn dòng 1.15' ((Near $p5.LineSpacing 13.8 0.2) -and $p5.LineSpacingRule -eq 5) "$($p5.LineSpacing)/$($p5.LineSpacingRule)"
        Add-Res 'KB-B3 thut-dau-dong' (Near $p5.FirstLineIndent (1.27 * $cm) 0.6) "FirstLineIndent $($p5.FirstLineIndent)"
        Add-Res 'KB-B3 so-giu-nguyen' ((Join-Arr (Get-ListStrings $d)) -eq (Join-Arr $ls0)) (Join-Arr (Get-ListStrings $d))
        Add-Res 'KB-F1 dieu-thut' ((Near $d.Paragraphs.Item(1).FirstLineIndent 0 0.6) -and (Near $d.Paragraphs.Item(1).LeftIndent 0 0.6)) "$($d.Paragraphs.Item(1).FirstLineIndent)/$($d.Paragraphs.Item(1).LeftIndent)"
        Add-Res 'KB-F1 chu-khong-doi' ((Join-Arr (Get-Texts $d)) -eq (Join-Arr $texts0))
        Add-Res 'KB-F1 khong-mat-cmt-fn-bang' ($d.Comments.Count -eq $cm0 -and $d.Footnotes.Count -eq $fn0 -and $d.Tables.Count -eq $tb0) "$($d.Comments.Count)/$($d.Footnotes.Count)/$($d.Tables.Count)"
        $sec = $d.Sections.Item(1)
        Add-Res 'KB-B1 so-trang' ($sec.Footers.Item(1).Range.Fields.Count -eq 1 -and $sec.PageSetup.DifferentFirstPageHeaderFooter -ne 0) "fields=$($sec.Footers.Item(1).Range.Fields.Count)"
        $snap1 = Get-Snap $d

        # --- Idempotent
        $n2 = [int](Invoke-Macro $w 'modApi.NormalizeDoc' @([string]$d.Name, 'HopDong', 'ND30'))
        Add-Res 'KB-F1 idempotent-count' ($n2 -eq 0) "lần 2 thay đổi $n2"
        Add-Res 'KB-F1 idempotent-snap' ((Get-Snap $d) -eq $snap1)
        $pv2 = [string](Invoke-Macro $w 'modApi.PreviewNormalize' @([string]$d.Name, 'HopDong', 'ND30'))
        Add-Res 'KB-F1 preview-sach' ($pv2.Trim() -eq '') $pv2

        # --- ISO 9001 (cỡ 12, giãn 1.2)
        $n3 = [int](Invoke-Macro $w 'modApi.NormalizeDoc' @([string]$d.Name, 'HopDong', 'ISO9001'))
        $bad = @(1..$d.Paragraphs.Count | Where-Object { $d.Paragraphs.Item($_).Range.Font.Size -ne 12 })
        Add-Res 'KB-B2 ISO co-12' ($bad.Count -eq 0 -and $n3 -gt 0) "đoạn sai: $($bad -join ',')"
        Add-Res 'KB-B3 ISO gian-1.2' (Near $d.Paragraphs.Item(5).LineSpacing 14.4 0.2) "$($d.Paragraphs.Item(5).LineSpacing)"
        $d.Close(0)

        # --- Undo: 1 lần Ctrl+Z về đúng trạng thái ban đầu
        $d = New-M2Doc $w
        $snapU = Get-Snap $d
        $d.UndoClear()
        [void](Invoke-Macro $w 'modApi.NormalizeDoc' @([string]$d.Name, 'HopDong', 'ND30'))
        [void]$d.Undo(1)
        Add-Res 'KB-F1 undo' ((Get-Snap $d) -eq $snapU) ((Get-Snap $d))
        $d.Close(0)
    } catch { Add-Res 'M2' $false $_.Exception.Message }

    return $res
}
