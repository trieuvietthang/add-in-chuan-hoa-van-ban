# Spike S7: tạo .dotm qua COM (kèm module VBA => thử AccessVBOM), chèn customUI14.xml vào gói zip, mở lại và kiểm tra.
. "$PSScriptRoot\..\..\tools\lib.ps1"
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$dir = Get-TempWorkDir
$dotm = [string](Join-Path $dir 'spike.dotm')
$w = $null
try {
    $w = Start-TestWord -AllowMacros
    $d = $w.Documents.Add()
    # 1) AccessVBOM
    try {
        $m = $d.VBProject.VBComponents.Add(1)
        $m.Name = 'modSpike'
        $m.CodeModule.AddFromString("Public Function Hello() As String`r`n    Hello = ""xin chao""`r`nEnd Function")
        'VBProject: OK (AccessVBOM có hiệu lực), modules=' + $d.VBProject.VBComponents.Count
    } catch { 'VBProject: LỖI -> ' + $_.Exception.Message }
    "saving $dotm"; [Console]::Out.Flush()
    $d.SaveAs2($dotm, 15); "saved"   # wdFormatXMLTemplateMacroEnabled = 15 (13 là docm)
    $d.Close(0)
    Stop-TestWord $w; $w = $null
    'Đã lưu: ' + (Test-Path $dotm) + ' ' + (Get-Item $dotm).Length + ' byte'

    # 2) Chèn customUI14.xml
    $xml = @'
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon><tabs><tab id="tabSpike" label="Soạn thảo chuẩn"><group id="grpA" label="Đánh số"><button id="btn1" label="Thử" size="large" imageMso="HappyFace"/></group></tab></tabs></ribbon>
</customUI>
'@
    $zip = [IO.Compression.ZipFile]::Open($dotm, 'Update')
    try {
        $e = $zip.CreateEntry('customUI/customUI14.xml')
        $sw = New-Object IO.StreamWriter($e.Open(), (New-Object Text.UTF8Encoding $false)); $sw.Write($xml); $sw.Dispose()
        $rels = $zip.GetEntry('_rels/.rels'); $rd = New-Object IO.StreamReader($rels.Open()); $txt = $rd.ReadToEnd(); $rd.Dispose()
        $rels.Delete()
        $add = '<Relationship Id="rIdUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
        $txt = $txt.Replace('</Relationships>', $add + '</Relationships>')
        $e2 = $zip.CreateEntry('_rels/.rels'); $sw = New-Object IO.StreamWriter($e2.Open(), (New-Object Text.UTF8Encoding $false)); $sw.Write($txt); $sw.Dispose()
    } finally { $zip.Dispose() }
    'Đã chèn customUI14.xml'

    # 3) Mở lại: Word đọc được gói? macro còn không?
    $w = Start-TestWord
    $d2 = $w.Documents.Open([string]$dotm, $false, $true)   # ReadOnly
    'Mở lại OK: ' + $d2.Name + ' | HasVBProject=' + $d2.HasVBProject
    try { 'Module còn: ' + (($d2.VBProject.VBComponents | % Name) -join ',') } catch { 'Đọc VBProject: ' + $_.Exception.Message }
    $d2.Close(0)
    # Ribbon chỉ hiện khi nạp làm add-in (STARTUP/AddIns); kiểm tra bằng AddIns.Install là bước sau, cần quyền ghi cấu hình -> hỏi người dùng
} catch { 'LỖI: ' + $_.Exception.Message; $_.ScriptStackTrace }
finally { Stop-TestWord $w; Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue }
