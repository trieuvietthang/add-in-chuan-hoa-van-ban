# Test M5: KB-D3 (kiểm tra chính tả offline: S1–S6 kiểu chữ, W1 cặp sai-đúng, W2 nhất quán, W3 âm tiết, W4 thuật ngữ).
# Văn bản tổng hợp, có lỗi cố ý. Tệp UTF-8 có BOM. Dot-source từ run-tests.ps1.

function Invoke-M5Tests($w) {
    $res = New-Object System.Collections.Generic.List[object]
    function Add-Res($n, $ok, $d = '') { $res.Add([pscustomobject]@{ Name = $n; Ok = [bool]$ok; Detail = $d }); [Console]::Error.WriteLine("  .. $n $ok $d") }
    function Cnt($out) {
        $h = @{}
        foreach ($ln in ($out -split "`n")) { if ($ln -ne '') { $k = ($ln -split '\|')[0]; $h[$k] = 1 + [int]$h[$k] } }
        return $h
    }

    try {
        # --- W3 thuần: bảng âm tiết đúng / sai
        $valid = 'hợp','đồng','văn','bản','điều','khoản','quy','định','thanh','toán','nghiêng','nghĩa','ngành','giữ','gìn','giám','đốc','trách','nhiệm','xem','xét','chuẩn','bị','khuyến','khích','quyền','lợi','người','nước','hoà','hòa','hóa','hoá','loại','thuế','xử','lý','cộng','tiếng','nghiệp','việc','Việt','được','bên','của','mua','kia','thuộc','ngoài','uy','tín','huyền','quyết','khuya','yêu','uống','đường','lượng','giêng','giường','gia','không','ghi','kiểm','tra','quê','quân','khoẻ','khỏe','chính','sách','thực','hiện','nhà','ở','ủy','tòa','án','sửa','chữa','đảm','bảo','nhưng','phải','lực','tên','chức','danh','phòng','ban','trưởng','phó','nội','bộ','quản','trị','kỷ','luật','khen','thưởng','bổ','sung','xác','định','hiệu','tạm','thời','quốc','quở','quơ','quá','quả','quý'
        $badList = @('đồg','nghành','tóan','khôg','gỉam','qủa','cêm','nghay','hợpp','thựch','quuy','hưóng','ngi','đưong')
        $okAll = $true; $failV = @()
        foreach ($s in $valid) { if ([string](Invoke-Macro $w 'modApi.SyllableOk' @([string]$s)) -ne '1') { $okAll = $false; $failV += $s } }
        Add-Res 'KB-D3 W3 am-tiet-dung' $okAll ($failV -join ',')
        $okBad = $true; $failB = @()
        foreach ($s in $badList) { if ([string](Invoke-Macro $w 'modApi.SyllableOk' @([string]$s)) -ne '0') { $okBad = $false; $failB += $s } }
        Add-Res 'KB-D3 W3 am-tiet-sai' $okBad ($failB -join ',')

        # --- Văn bản có lỗi cố ý
        $t = @(
            'Hợp đồng  này có hai dấu cách.',                                   # 1 S1
            'Hai bên ký kết .',                                                  # 2 S2
            'Hai bên ký,kết hợp đồng.Xong việc',                                  # 3 S3 x2
            'Căn cứ của của hợp đồng.',                                         # 4 S4
            'Công việc được thực hiện dần dần theo kế hoạch 1,5 triệu lúc 12:30 tại TP.HCM, xem http://a.com/x.Y và a@b.vn.',  # 5 sạch
            'Theo (khoản 2 của quy định này.',                                  # 6 S5
            'a) Hợp đồng có hiệu lực.',                                          # 7 sạch (a) gõ tay)
            'Hợp đồng có hiệu lực. điều 5 quy định về thanh toán.',             # 8 S6
            'Việc xem sét hồ sơ.',                                               # 9 W1
            'Theo qui định và quy định này.',                                   # 10 W2
            'bên A và Bên B ký.',                                                # 11 W4
            'Hợp đồg và nghành nghề, tóan.',                                    # 12 W3 x3
            'Điều 5. Người lao động được hưởng quyền lợi theo quy định của pháp luật. Hội nghị quốc tế của BÊN B và BÊN A.',  # 13 sạch (quốc, tiêu đề in hoa)
            'Xem ',                                                              # 14 có field
            'Theo quy đinh này, hợp đòng có hiệu lực và trách nhiêm của hai bên.',  # 15 W5 x3
            'Người lao động có quyền lợi và nghĩa vụ theo quy định của công ty.'   # 16 sạch
        )
        $d = New-TestDoc $w $t
        $r14 = $d.Paragraphs.Item(14).Range; $r14.MoveEnd(1, -1) | Out-Null; $r14.Collapse(0)
        [void]$d.Fields.Add($r14, -1, 'IF 1 = 1 "Dieu" "zzz  qqq,www xxxg"', $false)
        $d.Fields.Update() | Out-Null
        $texts0 = Get-Texts $d
        $out = [string](Invoke-Macro $w 'modApi.CheckSpellDoc' @([string]$d.Name, 'HopDong'))
        $c = Cnt $out
        Add-Res 'KB-D3 S1' ($c['S1'] -eq 1) "S1=$($c['S1'])"
        Add-Res 'KB-D3 S2' ($c['S2'] -eq 1) "S2=$($c['S2'])"
        Add-Res 'KB-D3 S3' ($c['S3'] -eq 2) "S3=$($c['S3'])"
        Add-Res 'KB-D3 S4' ($c['S4'] -eq 1) "S4=$($c['S4'])"
        Add-Res 'KB-D3 S5' ($c['S5'] -eq 1) "S5=$($c['S5'])"
        Add-Res 'KB-D3 S6' ($c['S6'] -eq 1) "S6=$($c['S6'])"
        Add-Res 'KB-D3 W1' ($c['W1'] -eq 1) "W1=$($c['W1'])"
        Add-Res 'KB-D3 W2' ($c['W2'] -eq 1) "W2=$($c['W2'])"
        Add-Res 'KB-D3 W3' ($c['W3'] -eq 3) "W3=$($c['W3'])"
        Add-Res 'KB-D3 W4' ($c['W4'] -eq 1) "W4=$($c['W4'])"
        Add-Res 'KB-D3 W5' ($c['W5'] -eq 3) "W5=$($c['W5'])"
        $total = 0; foreach ($k in $c.Keys) { $total += $c[$k] }
        Add-Res 'KB-D3 khong-bao-nham' ($total -eq 16) "tong=$total $($out.Replace("`n", ' '))"
        # vị trí báo nằm đúng đoạn có lỗi (đoạn 9: W1)
        $w1 = ($out -split "`n" | Where-Object { $_ -like 'W1|*' } | Select-Object -First 1)
        $p9s = $d.Paragraphs.Item(9).Range.Start; $p9e = $d.Paragraphs.Item(9).Range.End
        $pp = $w1 -split '\|'
        Add-Res 'KB-D3 vi-tri' (([int]$pp[1] -ge $p9s) -and ([int]$pp[2] -le $p9e)) $w1
        Add-Res 'KB-D3 chu-khong-doi' ((Join-Arr (Get-Texts $d)) -ceq (Join-Arr $texts0))
        $out2 = [string](Invoke-Macro $w 'modApi.CheckSpellDoc' @([string]$d.Name, 'HopDong'))
        Add-Res 'KB-D3 idempotent' ($out2 -ceq $out)
        $d.Close(0)

        # --- Hiệu năng: 3.000 đoạn
        $big = New-Object string[] 3000
        for ($i = 0; $i -lt 3000; $i++) { $big[$i] = 'Điều 5. Người lao động được hưởng quyền lợi theo quy định của pháp luật hiện hành.' }
        $d2 = New-TestDoc $w $big
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $o3 = [string](Invoke-Macro $w 'modApi.CheckSpellDoc' @([string]$d2.Name, 'HopDong'))
        $sw.Stop()
        Add-Res 'KB-E1 3000-doan' (($sw.Elapsed.TotalSeconds -le 5) -and ($o3 -eq '')) ("{0:N1}s [{1}]" -f $sw.Elapsed.TotalSeconds, $o3)
        # Lần đầu trong phiên Word: phải nạp lại cấu hình + bảng bigram (người dùng gặp ở lần bấm đầu tiên)
        [void](Invoke-Macro $w 'modSpell.ResetConfig' @())
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $o4 = [string](Invoke-Macro $w 'modApi.CheckSpellDoc' @([string]$d2.Name, 'HopDong'))
        $sw.Stop()
        Add-Res 'KB-E1 3000-doan-lan-dau' (($sw.Elapsed.TotalSeconds -le 5) -and ($o4 -eq '')) ("{0:N1}s [{1}]" -f $sw.Elapsed.TotalSeconds, $o4)
        $d2.Close(0)
    } catch { Add-Res 'M5 harness' $false $_.Exception.Message }
    return $res
}
