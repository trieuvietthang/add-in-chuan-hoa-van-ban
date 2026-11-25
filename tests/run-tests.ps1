# Chạy toàn bộ test. Chạy trong tiến trình con có timeout nếu gọi từ công cụ khác (xem docs/04).
param([switch]$SkipWord)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\..\tools\lib.ps1"
. "$PSScriptRoot\cases\m1-numbering.ps1"
. "$PSScriptRoot\cases\m2-normalize.ps1"
. "$PSScriptRoot\cases\m3-toc.ps1"
. "$PSScriptRoot\cases\m4-refs.ps1"
. "$PSScriptRoot\cases\m5-spell.ps1"
. "$PSScriptRoot\cases\m6-profile-ui.ps1"
$pass = 0; $fail = 0; $skip = 0
$sw = [Diagnostics.Stopwatch]::StartNew()
function Report($name, $ok, $detail = '') { Write-Result $name $ok $detail; if ($ok) { $script:pass++ } else { $script:fail++ } }

if ($SkipWord) { Write-Result 'word-tests' $true 'SKIP (-SkipWord)'; $skip++ }
else {
    $w = $null; $dir = Get-TempWorkDir
    try {
        # 1. Word khởi động
        $w = Start-TestWord -AllowMacros
        $d = $w.Documents.Add(); $d.Content.Text = 'xin chao'
        Report 'hello-word' (($d.Paragraphs.Count -ge 1) -and ($d.Content.Text -like 'xin chao*')) ("Word " + $w.Version)
        $d.Close(0)

        # 2. Build vào thư mục tạm (kèm config/)
        $dotm = [string](Join-Path $dir 'test.dotm')
        Stop-TestWord $w; $w = $null
        $null = & "$PSScriptRoot\..\tools\build.ps1" -Out $dotm
        Report 'build-dotm' ((Test-Path $dotm) -and (Test-Path (Join-Path $dir 'config\profiles\HopDong.ini')))

        # 3. Mở add-in, chạy macro
        $w = Start-TestWord -AllowMacros
        $dm = $w.Documents.Open($dotm, $false, $true)
        $global:AddinDoc = $dm
        $res = $w.Run('modStrings.Hello')
        Report 'macro-hello' ($res -eq ('Xin ch' + [char]0x00E0 + 'o'))

        # 4. inspect
        $d = $w.Documents.Add(); $d.Content.Text = "Tieu de`rNoi dung"; $docx = [string](Join-Path $dir 'x.docx'); $d.SaveAs2($docx, 16); $d.Close(0)
        $json = & "$PSScriptRoot\..\tools\inspect.ps1" -Path $docx -Word $w | Out-String | ConvertFrom-Json
        Report 'inspect-json' (($json.paragraphCount -eq 2) -and ($json.paragraphs[0].style -eq 'Normal') -and ($json.sections[0].pageWidth -gt 0))

        # 4b. Ribbon: module biên dịch được + mọi callback trong XML (trong gói .dotm) có hàm tương ứng
        $global:AddinDoc.Activate()
        $ping = $w.GetType().InvokeMember('Run', 'InvokeMethod', $null, $w, @('modRibbon.Ping'))
        Report 'ribbon-module-ping' ($ping -eq 'ok')
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($dotm)
        try { $e = $zip.GetEntry('customUI/customUI14.xml'); $sr = New-Object IO.StreamReader($e.Open()); $xmlText = $sr.ReadToEnd(); $sr.Dispose() } finally { $zip.Dispose() }
        $xml = New-Object Xml.XmlDocument; $xml.LoadXml($xmlText)
        $src = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\src\ui\modRibbon.bas'))
        $missing = @()
        foreach ($attr in 'onAction', 'getLabel', 'getVisible', 'getSelectedItemID', 'onLoad') {
            foreach ($n in $xml.SelectNodes("//*[@$attr]")) {
                $proc = ($n.GetAttribute($attr) -split '\.')[-1]
                if ($src -notmatch "(?m)^\s*Public (Sub|Function) $proc\(") { $missing += $proc }
            }
        }
        Report 'ribbon-callbacks' ($missing.Count -eq 0) ($missing -join ',')

        # 5. M1
        foreach ($r in (Invoke-M1Tests $w)) { Report $r.Name $r.Ok $r.Detail }
        # 6. M2
        foreach ($r in (Invoke-M2Tests $w)) { Report $r.Name $r.Ok $r.Detail }
        # 7. M3
        foreach ($r in (Invoke-M3Tests $w)) { Report $r.Name $r.Ok $r.Detail }
        # 8. M4
        foreach ($r in (Invoke-M4Tests $w)) { Report $r.Name $r.Ok $r.Detail }
        # 9. M5
        foreach ($r in (Invoke-M5Tests $w)) { Report $r.Name $r.Ok $r.Detail }
        # 10. M6 (form sua ho so)
        foreach ($r in (Invoke-M6Tests $w $dir)) { Report $r.Name $r.Ok $r.Detail }
        $dm.Close(0)
    } catch { Report 'harness' $false $_.Exception.Message }
    finally { Stop-TestWord $w; Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue }
}
"Tổng: $pass PASS, $fail FAIL, $skip SKIP  ({0:N1}s)" -f $sw.Elapsed.TotalSeconds
if ($fail -gt 0) { exit 1 }
