# Build: src/ -> dist/SoanThaoChuan.dotm (qua COM Word). Không sửa .dotm bằng tay; đổi nguồn rồi build lại.
# Cần AccessVBOM = 1 và Word của người dùng đã đóng (xem lib.ps1).
param([string]$Out)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\lib.ps1"
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$root = $script:RepoRoot
if (-not $Out) { $Out = Join-Path $root 'dist\SoanThaoChuan.dotm' }
$Out = [string]$Out
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
$tmp = [string](Join-Path (Get-TempWorkDir) 'build.dotm')

$files = Get-ChildItem (Join-Path $root 'src') -Recurse -Include *.bas, *.cls, *.frm | Sort-Object FullName
if (-not $files) { throw 'Không có tệp nguồn VBA trong src/.' }
$uiXml = Join-Path $root 'src\ui\customUI14.xml'

# Lint: VBA cấm khai báo cấp module (Private/Public/Dim/Const biến) sau thủ tục đầu tiên -> "Compile error in hidden module".
foreach ($f in $files) {
    $seenProc = $false
    $n = 0
    foreach ($ln in [IO.File]::ReadAllLines($f.FullName)) {
        $n++
        if ($ln -match '^(Public |Private |Friend )?(Static )?(Sub|Function|Property (Get|Let|Set))\s') { $seenProc = $true; continue }
        if ($seenProc -and $ln -match '^(Public |Private )?(Dim|Const)\s|^(Public|Private)\s+(?!Sub|Function|Property|Declare|Type|Enum)\w') {
            throw "Lint: $($f.Name):$n khai báo cấp module nằm sau thủ tục: $ln"
        }
    }
}

$w = $null
try {
    $w = Start-TestWord -AllowMacros
    $d = $w.Documents.Add()
    try { $null = $d.VBProject } catch { throw 'Không truy cập được VBProject: bật AccessVBOM (docs/04) rồi đóng/mở lại Word.' }
    foreach ($f in $files) {
        # Import yêu cầu tệp ANSI; kiểm tra không có ký tự ngoài ASCII (luật: tiếng Việt qua ChrW)
        $bytes = [IO.File]::ReadAllBytes($f.FullName)
        if ($bytes | Where-Object { $_ -gt 127 } | Select-Object -First 1) { throw "Tệp $($f.Name) có ký tự ngoài ASCII; dùng ChrW trong modStrings." }
        # Chuẩn hóa CRLF ASCII vào bản tạm (VBA Import cần CRLF)
        $tmpSrc = [string](Join-Path (Split-Path $tmp) $f.Name)
        $txt = [Text.Encoding]::ASCII.GetString($bytes) -replace "`r`n", "`n" -replace "`n", "`r`n"
        [IO.File]::WriteAllText($tmpSrc, $txt, [Text.Encoding]::ASCII)
        if ($f.Extension -eq '.frm') {
            # Import .frm không chạy được khi không có .frx/trong Word ẩn -> tạo UserForm trống bằng Add(3) rồi nạp mã (điều khiển dựng lúc chạy).
            $h = $txt
            $comp = $d.VBProject.VBComponents.Add(3)
            $comp.Name = [IO.Path]::GetFileNameWithoutExtension($f.Name)
            if ($h -match 'ClientWidth\s*=\s*(\d+)') { $comp.Properties.Item('Width').Value = [double]([int]$Matches[1] / 20 + 6) }
            if ($h -match 'ClientHeight\s*=\s*(\d+)') { $comp.Properties.Item('Height').Value = [double]([int]$Matches[1] / 20 + 28) }
            if ($h -match 'StartUpPosition\s*=\s*(\d)') { $comp.Properties.Item('StartUpPosition').Value = [int]$Matches[1] }
            $code = ($h -split "(?m)^Attribute VB_Exposed = (True|False)\r?\n", 3)[-1]
            $cm = $comp.CodeModule
            if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
            $cm.AddFromString($code)
        } else {
            [void]$d.VBProject.VBComponents.Import($tmpSrc)
        }
        "import: $($f.Name)"
    }
    $d.SaveAs2($tmp, 15)    # 15 = wdFormatXMLTemplateMacroEnabled
    $d.Close(0)
} finally { Stop-TestWord $w }

# Chèn ribbon vào gói
if (Test-Path $uiXml) {
    $zip = [IO.Compression.ZipFile]::Open($tmp, 'Update')
    try {
        $old = $zip.GetEntry('customUI/customUI14.xml'); if ($old) { $old.Delete() }
        $e = $zip.CreateEntry('customUI/customUI14.xml')
        $s = $e.Open(); $b = [IO.File]::ReadAllBytes($uiXml); $s.Write($b, 0, $b.Length); $s.Dispose()
        $re = $zip.GetEntry('_rels/.rels'); $r = New-Object IO.StreamReader($re.Open()); $txt = $r.ReadToEnd(); $r.Dispose(); $re.Delete()
        if ($txt -notmatch 'customUI14') {
            $txt = $txt.Replace('</Relationships>', '<Relationship Id="rIdUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>')
        }
        $e2 = $zip.CreateEntry('_rels/.rels'); $sw = New-Object IO.StreamWriter($e2.Open(), (New-Object Text.UTF8Encoding $false)); $sw.Write($txt); $sw.Dispose()
    } finally { $zip.Dispose() }
    'ribbon: đã chèn customUI14.xml'
}
Copy-Item $tmp $Out -Force
# config/ (hồ sơ, thông báo) nằm cạnh tệp .dotm
$cfgDst = Join-Path (Split-Path $Out) 'config'
if (Test-Path $cfgDst) { Remove-Item $cfgDst -Recurse -Force }
Copy-Item (Join-Path $root 'config') $cfgDst -Recurse -Force
Remove-Item (Split-Path $tmp) -Recurse -Force -ErrorAction SilentlyContinue
"Build xong: $Out ($((Get-Item $Out).Length) byte)"
