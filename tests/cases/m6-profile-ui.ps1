# Test M6: form "Sua ho so" (modIniEdit + frmProfile). Chay tren ban sao config trong thu muc tam. Dot-source tu run-tests.ps1.

function Invoke-M6Tests($w, $workDir) {
    $res = New-Object System.Collections.Generic.List[object]
    function Add-Res($n, $ok, $d = '') { $res.Add([pscustomobject]@{ Name = $n; Ok = [bool]$ok; Detail = $d }); [Console]::Error.WriteLine("  .. $n $ok $d") }
    function Val($sec, $key, $v) { [string](Invoke-Macro $w 'modApi.IniValidate' @([string]$sec, [string]$key, [string]$v)) }

    $cfg2 = [string](Join-Path $workDir 'config-ui')
    $cfgOrig = [string](Join-Path $workDir 'config')
    try {
        # --- Kiem tra gia tri
        Add-Res 'M6 validate-hop-le' ((Val 'page' 'top' '2.5') -eq '' -and (Val 'base' 'size' '13') -eq '' -and (Val 'level2' 'numstyle' 'arabic') -eq '' -and (Val 'level3' 'bold' '1') -eq '' -and (Val 'base' 'linespacing' '1,15') -eq '')
        Add-Res 'M6 validate-sai' ((Val 'page' 'top' 'abc') -ne '' -and (Val 'page' 'top' '99') -ne '' -and (Val 'level2' 'numstyle' 'foo') -ne '' -and (Val 'level3' 'bold' '2') -ne '' -and (Val 'level3' 'outline' '1.5') -ne '' -and (Val 'page' 'top' "1`n2") -ne '')
        Add-Res 'M6 validate-khong-rang-buoc' ((Val 'level3' 'format' 'Dieu %3.') -eq '' -and (Val 'khac' 'khoa' 'bat ky') -eq '')

        # --- Ghi INI giu chu thich/thu tu/UTF-8
        $f = [string](Join-Path $workDir 'ND30-copy.ini')
        Copy-Item (Join-Path $cfgOrig 'standards\ND30.ini') $f
        $orig = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8)
        $n1 = [string](Invoke-Macro $w 'modApi.IniSave' @($f, 'base', 'size', '12'))
        $new = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8)
        $bytes = [IO.File]::ReadAllBytes($f)
        $oldL = $orig -split "`r?`n"; $newL = $new -split "`r?`n"
        $diff = @(); for ($i = 0; $i -lt $oldL.Count; $i++) { if ($oldL[$i] -ne $newL[$i]) { $diff += $i } }
        Add-Res 'M6 ini-ghi-1-dong' ($n1 -eq '1' -and $diff.Count -eq 1 -and $newL[$diff[0]] -eq 'size=12' -and $oldL.Count -eq $newL.Count) "n=$n1 diff=$($diff -join ',')"
        Add-Res 'M6 ini-giu-utf8-khong-bom' (($bytes[0] -ne 0xEF) -and ($newL[0] -eq $oldL[0]) -and ($newL[0] -match '[^\x00-\x7F]'))
        Add-Res 'M6 ini-bak' ((Test-Path "$f.bak") -and ([IO.File]::ReadAllText("$f.bak", [Text.Encoding]::UTF8) -eq $orig))
        $n2 = [string](Invoke-Macro $w 'modApi.IniSave' @($f, 'base', 'size', '12'))
        Add-Res 'M6 ini-idempotent' ($n2 -eq '0' -and ([IO.File]::ReadAllText($f, [Text.Encoding]::UTF8) -eq $new))
        $n3 = [string](Invoke-Macro $w 'modApi.IniSave' @($f, 'page', 'size', 'A4'))
        Add-Res 'M6 ini-khong-doi-muc-khac' ($n3 -eq '0' -and ($new -match '(?m)^size=A4'))

        # --- Form tren ban sao config
        if (Test-Path $cfg2) { Remove-Item $cfg2 -Recurse -Force }
        Copy-Item $cfgOrig $cfg2 -Recurse
        [void](Invoke-Macro $w 'modConfig.SetConfigDir' @($cfg2))
        $nd = [string](Join-Path $cfg2 'standards\ND30.ini')
        $r0 = [string](Invoke-Macro $w 'modApi.FormTest' @([int]0, 'base', '', ''))
        Add-Res 'M6 form-nap' ($r0 -eq '6|3|6|') $r0
        $r1 = [string](Invoke-Macro $w 'modApi.FormTest' @([int]0, 'base', 'size', '14'))
        $t1 = [IO.File]::ReadAllText($nd, [Text.Encoding]::UTF8)
        Add-Res 'M6 form-luu' (($t1 -match '(?m)^size=14\r?$') -and ($t1 -match '(?m)^size=A4\r?$') -and ($r1 -match '\b1\b')) $r1
        $r2 = [string](Invoke-Macro $w 'modApi.FormTest' @([int]0, 'base', 'size', '99'))
        $t2 = [IO.File]::ReadAllText($nd, [Text.Encoding]::UTF8)
        Add-Res 'M6 form-tu-choi-gia-tri-sai' ($t2 -eq $t1) $r2
        $qd = [string](Join-Path $cfg2 'profiles\QuyDinh.ini')
        $q0 = [IO.File]::ReadAllText($qd, [Text.Encoding]::UTF8)
        $r3 = [string](Invoke-Macro $w 'modApi.FormTest' @([int]3, 'level3', 'bold', '0'))
        $q1 = [IO.File]::ReadAllText($qd, [Text.Encoding]::UTF8)
        $a0 = $q0 -split "`r?`n"; $a1 = $q1 -split "`r?`n"
        $dq = @(); for ($i = 0; $i -lt $a0.Count; $i++) { if ($a0[$i] -ne $a1[$i]) { $dq += $i } }
        Add-Res 'M6 form-ho-so-dung-cap' ($dq.Count -eq 1 -and $a0[$dq[0]] -eq 'bold=1' -and $a1[$dq[0]] -eq 'bold=0') "dong=$($dq -join ',')"
        # chi dong cua level3 doi: tim muc cua dong do
        $sec = ''; for ($i = 0; $i -le $dq[0]; $i++) { if ($a1[$i] -match '^\[(.+)\]') { $sec = $Matches[1] } }
        Add-Res 'M6 form-dung-muc-level3' ($sec -eq 'level3') $sec
        $r4 = [string](Invoke-Macro $w 'modApi.FormTest' @([int]3, 'level3', 'style', 'XYZ'))
        Add-Res 'M6 form-chi-doc-khong-ghi' (([IO.File]::ReadAllText($qd, [Text.Encoding]::UTF8) -eq $q1)) $r4
        $r5 = [string](Invoke-Macro $w 'modApi.FormTest' @([int]5, 'w5', 'min_ratio', '30'))
        $sp = [IO.File]::ReadAllText((Join-Path $cfg2 'spell.vi.ini'), [Text.Encoding]::UTF8)
        Add-Res 'M6 form-chinh-ta-w5' (($sp -match '(?m)^min_ratio=30\r?$') -and ($r5 -match '^6\|4\|3\|')) $r5
        # chinh ta van chay sau khi doi cau hinh
        $d = $w.Documents.Add(); $d.Content.Text = "Hop dong  co hai dau cach."
        $out = [string](Invoke-Macro $w 'modApi.CheckSpellDoc' @([string]$d.Name, 'HopDong'))
        Add-Res 'M6 chinh-ta-sau-doi-config' ($out -match '^S1\|') $out
        $d.Close(0)
    } catch { Add-Res 'M6 harness' $false $_.Exception.Message }
    finally {
        try { [void](Invoke-Macro $w 'modConfig.SetConfigDir' @($cfgOrig)) } catch {}
    }
    return $res
}
