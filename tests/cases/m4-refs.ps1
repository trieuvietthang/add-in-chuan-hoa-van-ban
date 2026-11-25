# Test M4: KB-D2 (kiểm tra dẫn chiếu), KB-A6 (chuyển REF, cập nhật, chèn), idempotent, Undo, chữ không đổi.
# Văn bản tổng hợp, có lỗi cố ý. Dot-source từ run-tests.ps1.

function New-RefDoc($w) {
    $VD = [string][char]0x0110
    $t = @(
        'Pham vi',                                                          # 1  Dieu 1
        ('Thuc hien theo ' + $VD + 'i' + [char]0x1EC1 + 'u 2 va kho' + [char]0x1EA3 + 'n 1 ' + $VD + 'i' + [char]0x1EC1 + 'u 2 cua Hop dong nay'),   # 2 khoan 1.1: hai dan chieu dung
        ('Theo ' + $VD + 'i' + [char]0x1EC1 + 'u 9 la sai'),                  # 3 R1
        'Doi tuong',                                                         # 4  Dieu 2
        ('Can cu kho' + [char]0x1EA3 + 'n 5 ' + $VD + 'i' + [char]0x1EC1 + 'u 2 la sai'),   # 5 R5 (Dieu 2 chi co 3 khoan)
        ('Theo ' + $VD + 'i' + [char]0x1EC1 + 'u 5 c' + [char]0x1EE7 + 'a Lu' + [char]0x1EAD + 't Dat dai khong phai loi'),   # 6 ngoai
        ('Xem ' + $VD + 'i' + [char]0x1EC1 + 'u 1 (Doi tuong) la sai ten'),    # 7 R2
        ('Theo ' + $VD + 'i' + [char]0x1EC1 + 'u 62 ' + [char]0x0111 + [char]0x1EBF + 'n ' + $VD + 'i' + [char]0x1EC1 + 'u 65 Quy' + [char]0x1EBF + 't ' + [char]0x0111 + [char]0x1ECB + 'nh so 746')   # 8 chuoi ngoai: khong bao loi
    )
    $d = New-TestDoc $w $t
    $lv = 1, 2, 2, 1, 2, 2, 2
    for ($i = 1; $i -le 7; $i++) { Set-Level $w $d $i $lv[$i-1] 'HopDong' }
    return $d
}

function Invoke-M4Tests($w) {
    $res = New-Object System.Collections.Generic.List[object]
    function Add-Res($n, $ok, $d = '') { $res.Add([pscustomobject]@{ Name = $n; Ok = [bool]$ok; Detail = $d }); [Console]::Error.WriteLine("  .. $n $ok $d") }
    function Rules($s) { @($s -split "`n" | Where-Object { $_ -ne '' } | ForEach-Object { ($_ -split '\|')[0] }) }

    try {
        $d = New-RefDoc $w
        $texts0 = Get-Texts $d
        $out = [string](Invoke-Macro $w 'modApi.CheckRefsDoc' @([string]$d.Name, 'HopDong'))
        $r = Rules $out
        $cnt = @{ R1 = 0; R2 = 0; R4 = 0; R5 = 0 }; foreach ($x in $r) { $cnt[$x]++ }
        Add-Res 'KB-D2 R1' ($cnt.R1 -eq 1) "R1=$($cnt.R1)"
        Add-Res 'KB-D2 R5' ($cnt.R5 -eq 1) "R5=$($cnt.R5)"
        Add-Res 'KB-D2 R2' ($cnt.R2 -eq 1) "R2=$($cnt.R2)"
        Add-Res 'KB-D2 R4 go-tay-hop-le' ($cnt.R4 -eq 2) "R4=$($cnt.R4)"
        Add-Res 'KB-D2 chuoi-dieu-den-dieu-ngoai' ((($out -split "`n") | Where-Object { $_ -like 'R1|*|62' -or $_ -like 'R1|*|65' }).Count -eq 0) $out.Replace("`n", ' ')
        Add-Res 'KB-D2 bo-qua-van-ban-ngoai' ((($out -split "`n") | Where-Object { $_ -like 'R1|*|5' }).Count -eq 0) $out.Replace("`n", ' ')

        # --- Chuyển REF
        $d.UndoClear()
        $n = [int](Invoke-Macro $w 'modApi.ConvertRefsDoc' @([string]$d.Name, 'HopDong'))
        $refs = @($d.Fields | Where-Object { $_.Type -eq 3 })
        Add-Res 'KB-A6 chuyen-REF' ($n -eq 2 -and $refs.Count -eq 2) "chuyen=$n fields=$($refs.Count)"
        Add-Res 'KB-A6 chu-khong-doi' ((Join-Arr (Get-Texts $d)) -eq (Join-Arr $texts0)) ((Get-Texts $d)[1])
        $n2 = [int](Invoke-Macro $w 'modApi.ConvertRefsDoc' @([string]$d.Name, 'HopDong'))
        Add-Res 'KB-A6 idempotent' ($n2 -eq 0 -and @($d.Fields | Where-Object { $_.Type -eq 3 }).Count -eq 2) "lần 2=$n2"
        $r2 = Rules ([string](Invoke-Macro $w 'modApi.CheckRefsDoc' @([string]$d.Name, 'HopDong')))
        Add-Res 'KB-A6 con-lai-sau-chuyen' (@($r2 | Where-Object { $_ -eq 'R4' }).Count -eq 0 -and @($r2 | Where-Object { $_ -eq 'R1' }).Count -eq 1) ($r2 -join ',')

        # --- Đổi số: chèn Điều mới trước Điều 2 -> REF tự theo
        $d.Paragraphs.Item(4).Range.InsertParagraphBefore()
        $d.Paragraphs.Item(4).Range.InsertBefore('Dieu moi chen')
        Set-Level $w $d 4 1 'HopDong'
        [void](Invoke-Macro $w 'modApi.UpdateRefsDoc' @([string]$d.Name))
        $p2 = ($d.Paragraphs.Item(2).Range.Text -replace '[\r\a]', '')
        $VD = [string][char]0x0110
        Add-Res 'KB-A6 cap-nhat-theo-so-moi' ($p2 -like ("*Thuc hien theo $VD*u 3 va kho*n 1 $VD*u 3*")) $p2

        # --- Chèn dẫn chiếu mới
        $d.Paragraphs.Item(3).Range.InsertParagraphAfter()
        $pos = $d.Paragraphs.Item(4).Range.Start
        $txt = [string](Invoke-Macro $w 'modApi.InsertRefDoc' @([string]$d.Name, $pos, 1, 'HopDong'))
        Add-Res 'KB-A6 chen' ($txt -like "$VD*u 1") "ket qua=[$txt]"
        $none = [string](Invoke-Macro $w 'modApi.InsertRefDoc' @([string]$d.Name, $pos, 99, 'HopDong'))
        Add-Res 'KB-A6 chen-khong-co-dich' ($none -eq '') "[$none]"
        $d.Close(0)

        # --- Undo
        $d = New-RefDoc $w
        $d.UndoClear()
        [void](Invoke-Macro $w 'modApi.ConvertRefsDoc' @([string]$d.Name, 'HopDong'))
        [void]$d.Undo(1)
        Add-Res 'KB-A6 undo' (@($d.Fields | Where-Object { $_.Type -eq 3 }).Count -eq 0 -and (Join-Arr (Get-Texts $d)) -eq (Join-Arr (Get-Texts $d))) "fields=$(@($d.Fields).Count)"
        $d.Close(0)
    } catch { Add-Res 'M4' $false $_.Exception.Message }

    # --- Quy định: Điều ở cấp 3, REF cho ra "Điều N"
    try {
        $VD = [string][char]0x0110
        $t = 'Quy dinh chung', 'Pham vi', 'Chuong hai', 'Dieu moi', ('Theo ' + $VD + 'i' + [char]0x1EC1 + 'u 2 la dung')
        $lv = 1, 3, 1, 3, 4
        $d = New-TestDoc $w $t
        for ($i = 1; $i -le 5; $i++) { Set-Level $w $d $i $lv[$i-1] 'QuyDinh' }
        $n = [int](Invoke-Macro $w 'modApi.ConvertRefsDoc' @([string]$d.Name, 'QuyDinh'))
        $res5 = ($d.Paragraphs.Item(5).Range.Text -replace '[\r\a]', '')
        Add-Res 'KB-A6 quy-dinh REF' ($n -eq 1 -and $res5 -eq $t[4]) "n=$n [$res5]"
        $d.Close(0)
    } catch { Add-Res 'KB-A6 quy-dinh' $false $_.Exception.Message }

    return $res
}
