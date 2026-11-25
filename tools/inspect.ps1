# Dump cấu trúc một .docx ra JSON để so golden. Mở BẢN SAO ở thư mục tạm, chỉ đọc; không sửa gốc.
# -NoText: không xuất chữ nội dung (dùng cho văn bản thật — luật 9). Mặc định xuất 30 ký tự đầu để định vị (chỉ dùng với fixture tổng hợp).
param([Parameter(Mandatory)][string]$Path, [string]$OutFile, [switch]$NoText, $Word)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\lib.ps1"

function Get-DocStructure($w, [string]$p, [bool]$noText) {
    $dir = Get-TempWorkDir
    $copy = [string](Join-Path $dir ([IO.Path]::GetFileName($p)))
    Copy-Item -LiteralPath $p -Destination $copy
    $d = $w.Documents.Open($copy, $false, $true)   # ReadOnly
    try {
        $paras = New-Object System.Collections.Generic.List[object]
        $i = 0
        foreach ($pa in $d.Paragraphs) {
            $i++
            $r = $pa.Range; $pf = $pa.Format
            $txt = ''
            if (-not $noText) { $t = $r.Text -replace '[\r\n\a]', ''; $txt = $t.Substring(0, [Math]::Min(30, $t.Length)) }
            $paras.Add([ordered]@{
                index = $i
                style = $pa.Style.NameLocal
                listString = $r.ListFormat.ListString
                listLevel = $r.ListFormat.ListLevelNumber
                leftIndent = [Math]::Round($pf.LeftIndent, 2)
                firstLineIndent = [Math]::Round($pf.FirstLineIndent, 2)
                alignment = $pf.Alignment
                spaceBefore = [Math]::Round($pf.SpaceBefore, 2)
                spaceAfter = [Math]::Round($pf.SpaceAfter, 2)
                lineSpacing = [Math]::Round($pf.LineSpacing, 2)
                lineRule = $pf.LineSpacingRule
                fontName = $r.Font.Name
                fontSize = $r.Font.Size
                text = $txt
            })
        }
        $secs = foreach ($s in $d.Sections) { $ps = $s.PageSetup; [ordered]@{
            pageWidth = [Math]::Round($ps.PageWidth,1); pageHeight = [Math]::Round($ps.PageHeight,1)
            top = [Math]::Round($ps.TopMargin,1); bottom = [Math]::Round($ps.BottomMargin,1)
            left = [Math]::Round($ps.LeftMargin,1); right = [Math]::Round($ps.RightMargin,1)
            differentFirstPage = [bool]$ps.DifferentFirstPageHeaderFooter } }
        $fields = foreach ($f in $d.Fields) { [ordered]@{ type = $f.Type; code = ($f.Code.Text.Trim()) } }
        [ordered]@{
            paragraphCount = $d.Paragraphs.Count
            tables = $d.Tables.Count; footnotes = $d.Footnotes.Count; comments = $d.Comments.Count
            revisions = $d.Revisions.Count; inlineShapes = $d.InlineShapes.Count
            listTemplates = $d.ListTemplates.Count
            sections = @($secs); fields = @($fields); paragraphs = $paras.ToArray()
        }
    } finally { $d.Close(0); Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue }
}

$own = $false; $w = $Word
try {
    if (-not $w) { $w = Start-TestWord; $own = $true }
    $obj = Get-DocStructure $w ([string](Resolve-Path -LiteralPath $Path)) ([bool]$NoText)
    $json = $obj | ConvertTo-Json -Depth 6
    if ($OutFile) { [IO.File]::WriteAllText([string]$OutFile, $json, (New-Object Text.UTF8Encoding $false)); "Đã ghi $OutFile" } else { $json }
} finally { if ($own) { Stop-TestWord $w } }
