# Test M3: KB-C1 (chèn/cập nhật mục lục theo hồ sơ), idempotent, Undo, chữ không đổi. Văn bản tổng hợp.

function Get-TocLines($d) {
    if ($d.TablesOfContents.Count -eq 0) { return @() }
    return @(($d.TablesOfContents.Item(1).Range.Text -split "[\r\a]") | Where-Object { $_ -ne '' })
}

function Invoke-M3Tests($w) {
    $res = New-Object System.Collections.Generic.List[object]
    function Add-Res($n, $ok, $d = '') { $res.Add([pscustomobject]@{ Name = $n; Ok = [bool]$ok; Detail = $d }); [Console]::Error.WriteLine("  .. $n $ok $d") }

    try {
        # --- Hợp đồng: TOC chỉ có Điều
        $t = 'Pham vi ap dung', 'Khoan mot', 'Diem a', 'Doi tuong', 'Khoan hai'
        $lv = 1, 2, 3, 1, 2
        $d = New-TestDoc $w $t
        for ($i = 1; $i -le 5; $i++) { Set-Level $w $d $i $lv[$i-1] 'HopDong' }
        $n0 = $d.Paragraphs.Count
        $d.UndoClear()
        $c = [int](Invoke-Macro $w 'modApi.InsertTocAt' @([string]$d.Name, 1, 'HopDong'))
        Add-Res 'KB-C1 chen' ($c -eq 1 -and $d.TablesOfContents.Count -eq 1) "fields=$c"
        $L = Get-TocLines $d
        Add-Res 'KB-C1 chi-co-dieu' ($L.Count -eq 2 -and $L[0] -like '*Pham vi ap dung*' -and $L[1] -like '*Doi tuong*') ($L -join ' | ')
        Add-Res 'KB-C1 so-dieu' ($L[0] -like ('*' + [char]0x0110 + 'i*u 1*') -and $L[1] -like ('*' + [char]0x0110 + 'i*u 2*')) ($L -join ' | ')
        Add-Res 'KB-C1 so-trang' ($L[0] -match '\d\s*$') ($L[0])
        Add-Res 'KB-C1 tieu-de' ($d.Paragraphs.Item(1).Style.NameLocal -eq 'STC_TocTitle' -and $d.Paragraphs.Item(1).Range.Text -like ('M*C L*C*')) $d.Paragraphs.Item(1).Range.Text
        $texts = Get-Texts $d
        Add-Res 'KB-C1 chu-khong-doi' ($texts[-5..-1] -join '|' -eq ($t -join '|')) ($texts[-5..-1] -join '|')
        $code = $d.Fields.Item(1).Code.Text
        Add-Res 'KB-C1 ma-field' ($code -match '\\t "01Dieu,1"') $code
        $snapT = ($d.Content.Text)

        # idempotent: chèn lần 2 vẫn 1 TOC, nội dung không đổi
        $c2 = [int](Invoke-Macro $w 'modApi.InsertTocAt' @([string]$d.Name, 1, 'HopDong'))
        Add-Res 'KB-C1 idempotent' ($d.TablesOfContents.Count -eq 1 -and $c2 -eq 1 -and $d.Content.Text -eq $snapT) "toc=$($d.TablesOfContents.Count)"

        # thêm Điều mới -> cập nhật -> 3 mục
        $d.Paragraphs.Last.Range.InsertParagraphAfter()
        $d.Paragraphs.Last.Range.InsertBefore('Dieu moi')
        Set-Level $w $d $d.Paragraphs.Count 1 'HopDong'
        [void](Invoke-Macro $w 'modApi.UpdateTocDoc' @([string]$d.Name))
        $L = Get-TocLines $d
        Add-Res 'KB-C1 cap-nhat' ($L.Count -eq 3 -and $L[2] -like '*Dieu moi*') ($L -join ' | ')
        $d.Close(0)

        # undo
        $d = New-TestDoc $w $t
        for ($i = 1; $i -le 5; $i++) { Set-Level $w $d $i $lv[$i-1] 'HopDong' }
        $d.UndoClear()
        [void](Invoke-Macro $w 'modApi.InsertTocAt' @([string]$d.Name, 1, 'HopDong'))
        [void]$d.Undo(1)
        Add-Res 'KB-C1 undo' ($d.TablesOfContents.Count -eq 0 -and $d.Paragraphs.Count -eq $n0) "toc=$($d.TablesOfContents.Count) paras=$($d.Paragraphs.Count)/$n0"
        $d.Close(0)
    } catch { Add-Res 'KB-C1' $false $_.Exception.Message }

    # --- Con trỏ trong bảng đầu văn bản (khối quốc hiệu): chèn SAU bảng, không lỗi; đếm nguồn
    try {
        $t = 'Quoc hieu', 'Pham vi ap dung', 'Khoan mot', 'Doi tuong'
        $d = New-TestDoc $w $t
        $tb = $d.Tables.Add($d.Range(0, 0), 1, 1)
        $tb.Cell(1, 1).Range.Text = 'Quoc hieu'
        $texts = Get-Texts $d
        $iP = 1 + [array]::IndexOf($texts, 'Pham vi ap dung'); $iK = 1 + [array]::IndexOf($texts, 'Khoan mot'); $iD = 1 + [array]::IndexOf($texts, 'Doi tuong')
        Set-Level $w $d $iP 1 'HopDong'; Set-Level $w $d $iK 2 'HopDong'; Set-Level $w $d $iD 1 'HopDong'
        $ns = [int](Invoke-Macro $w 'modApi.CountTocSources' @([string]$d.Name, 'HopDong'))
        Add-Res 'KB-C1 dem-nguon' ($ns -eq 2) "nguon=$ns"
        $c = [int](Invoke-Macro $w 'modApi.InsertTocAt' @([string]$d.Name, 1, 'HopDong'))
        $tocStart = $d.TablesOfContents.Item(1).Range.Start
        Add-Res 'KB-C1 trong-bang-chen-sau-bang' ($c -eq 1 -and $tocStart -ge $d.Tables.Item(1).Range.End) "toc=$tocStart bangEnd=$($d.Tables.Item(1).Range.End)"
        Add-Res 'KB-C1 trong-bang-giu-bang' ($d.Tables.Count -eq 1 -and $d.Tables.Item(1).Range.Text -like 'Quoc hieu*') $d.Tables.Item(1).Range.Text
        $L = Get-TocLines $d
        Add-Res 'KB-C1 trong-bang-noi-dung' ($L.Count -eq 2) ($L -join ' | ')
        $d.Close(0)
        # không có nguồn -> đếm 0
        $d = New-TestDoc $w @('Dieu 1. Go tay', 'Doan thuong')
        $ns0 = [int](Invoke-Macro $w 'modApi.CountTocSources' @([string]$d.Name, 'HopDong'))
        Add-Res 'KB-C1 dem-nguon-0' ($ns0 -eq 0) "nguon=$ns0"
        $d.Close(0)
    } catch { Add-Res 'KB-C1 trong-bang' $false $_.Exception.Message }

    # --- Style không bị nhiễm ngắt trang: (a) tạo mới khi vùng chọn có ngắt trang, (b) chữa style cũ đã nhiễm
    try {
        $d = New-TestDoc $w @('Doan co ngat trang', 'Pham vi', 'Khoan')
        $d.Paragraphs.Item(1).PageBreakBefore = -1; $d.Paragraphs.Item(1).Range.Select()
        Set-Level $w $d 2 1 'HopDong'; Set-Level $w $d 3 2 'HopDong'
        [void](Invoke-Macro $w 'modApi.InsertTocAt' @([string]$d.Name, 1, 'HopDong'))
        $bad = @('01Dieu', '02Khoan', 'STC_TocTitle') | Where-Object { $d.Styles.Item($_).ParagraphFormat.PageBreakBefore -ne 0 }
        Add-Res 'KB-C1 style-khong-ngat-trang' ($bad.Count -eq 0) ($bad -join ',')
        $d.Close(0)
        $d = New-TestDoc $w @('Pham vi', 'Khoan')
        $s = $d.Styles.Add('01Dieu', 1); $s.ParagraphFormat.PageBreakBefore = -1
        Set-Level $w $d 1 1 'HopDong'
        Add-Res 'KB-C1 chua-style-nhiem' ($d.Styles.Item('01Dieu').ParagraphFormat.PageBreakBefore -eq 0) "pbb=$($d.Styles.Item('01Dieu').ParagraphFormat.PageBreakBefore)"
        $d.Close(0)
    } catch { Add-Res 'KB-C1 style-ngat-trang' $false $_.Exception.Message }

    # --- Quy định: Chương + Điều
    try {
        $t = 'Quy dinh chung', 'Pham vi', 'Khoan', 'Chuong hai', 'Dieu moi'
        $lv = 1, 3, 4, 1, 3
        $d = New-TestDoc $w $t
        for ($i = 1; $i -le 5; $i++) { Set-Level $w $d $i $lv[$i-1] 'QuyDinh' }
        [void](Invoke-Macro $w 'modApi.InsertTocAt' @([string]$d.Name, 1, 'QuyDinh'))
        $L = Get-TocLines $d
        Add-Res 'KB-C1 quy-dinh' ($L.Count -eq 4 -and $L[0] -like '*Quy dinh chung*' -and $L[1] -like '*Pham vi*' -and $L[2] -like '*Chuong hai*') ($L -join ' | ')
        $d.Close(0)
    } catch { Add-Res 'KB-C1 quy-dinh' $false $_.Exception.Message }

    return $res
}
