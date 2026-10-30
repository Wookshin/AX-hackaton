#Requires -Version 7.2
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\AX-hackathon-proposal.docx')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$proposalPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'AX-hackathon-proposal.md'
$proposal = [System.IO.File]::ReadAllText($proposalPath)
# 제출 문서는 제목과 공식 질문 1~7번으로 구성한다. 웹용 핵심 요약은 제외한다.
$titleMatch = [regex]::Match($proposal, '(?m)^# .+\r?$')
$firstQuestion = [regex]::Match($proposal, '(?m)^\*\*1\. .+\*\*\r?$')
if (-not $titleMatch.Success -or -not $firstQuestion.Success) { throw '제출 문서의 제목 또는 첫 질문을 찾을 수 없습니다.' }
$submission = $titleMatch.Value.Trim() + "`n`n" + $proposal.Substring($firstQuestion.Index)
$fragment = [xml]('<root>' + (ConvertFrom-Markdown -InputObject $submission).Html + '</root>')
$script:wordLinks = [System.Collections.Generic.List[object]]::new()
$wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$relationshipsNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$fontProperties = '<w:rFonts w:ascii="바탕체" w:hAnsi="바탕체" w:eastAsia="바탕체" w:cs="바탕체"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/>'

function Convert-XmlText([string]$Text) {
    return [System.Security.SecurityElement]::Escape($Text)
}

function New-WordRun([string]$Text, [bool]$Bold = $false) {
    $emphasis = if ($Bold) { '<w:b/><w:bCs/>' } else { '' }
    return '<w:r><w:rPr>' + $fontProperties + $emphasis + '</w:rPr><w:t xml:space="preserve">' + (Convert-XmlText $Text) + '</w:t></w:r>'
}

function Convert-WordInline([System.Xml.XmlNode]$Node, [bool]$Bold = $false) {
    if ($Node.NodeType -in @([System.Xml.XmlNodeType]::Text, [System.Xml.XmlNodeType]::Whitespace, [System.Xml.XmlNodeType]::SignificantWhitespace)) {
        return New-WordRun $Node.Value $Bold
    }
    if ($Node.Name -eq 'br') { return '<w:r><w:rPr>' + $fontProperties + '</w:rPr><w:br/></w:r>' }
    $content = [System.Text.StringBuilder]::new()
    foreach ($child in $Node.ChildNodes) {
        [void]$content.Append((Convert-WordInline $child ($Bold -or $Node.Name -in @('strong', 'b'))))
    }
    if ($Node.Name -eq 'a' -and $Node.HasAttribute('href')) {
        $relationshipId = 'rIdLink' + ($script:wordLinks.Count + 1)
        $script:wordLinks.Add([pscustomobject]@{ Id = $relationshipId; Target = $Node.GetAttribute('href') })
        return '<w:hyperlink r:id="' + $relationshipId + '">' + $content.ToString() + '</w:hyperlink>'
    }
    return $content.ToString()
}

function New-WordParagraph([System.Xml.XmlNode]$Node, [string]$Style = 'Normal', [string]$Prefix = '', [bool]$InTable = $false, [bool]$Bold = $false) {
    $properties = '<w:pStyle w:val="' + $Style + '"/>'
    if ($InTable) { $properties += '<w:spacing w:after="40" w:line="280" w:lineRule="auto"/>' }
    if ($InTable -and $Bold) { $properties += '<w:keepNext/>' }
    if ($Prefix) { $properties += '<w:ind w:left="420" w:hanging="420"/>' }
    $runs = if ($Prefix) { New-WordRun $Prefix } else { '' }
    $runs += Convert-WordInline $Node $Bold
    return '<w:p><w:pPr>' + $properties + '</w:pPr>' + $runs + '</w:p>'
}

function Convert-WordTable([System.Xml.XmlNode]$Node) {
    $rows = @($Node.SelectNodes('thead/tr|tbody/tr|tr'))
    $columnCount = $rows[0].SelectNodes('th|td').Count
    $ratios = switch ($columnCount) {
        2 { @(0.35, 0.65) }
        3 { @(0.21, 0.34, 0.45) }
        4 { @(0.18, 0.24, 0.40, 0.18) }
        default { @(1..$columnCount | ForEach-Object { 1.0 / $columnCount }) }
    }
    $widths = @($ratios | ForEach-Object { [int][Math]::Round(9866 * $_) })
    $widths[-1] += 9866 - ($widths | Measure-Object -Sum).Sum
    $table = [System.Text.StringBuilder]::new()
    [void]$table.Append('<w:tbl><w:tblPr><w:tblW w:w="9866" w:type="dxa"/><w:tblBorders>')
    foreach ($edge in @('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) {
        [void]$table.Append('<w:' + $edge + ' w:val="single" w:sz="4" w:color="808080"/>')
    }
    [void]$table.Append('</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>')
    foreach ($width in $widths) { [void]$table.Append('<w:gridCol w:w="' + $width + '"/>') }
    [void]$table.Append('</w:tblGrid>')
    foreach ($row in $rows) {
        $cells = @($row.SelectNodes('th|td'))
        $isHeader = $cells[0].Name -eq 'th'
        [void]$table.Append('<w:tr><w:trPr><w:cantSplit/>')
        if ($isHeader) { [void]$table.Append('<w:tblHeader/>') }
        [void]$table.Append('</w:trPr>')
        for ($index = 0; $index -lt $cells.Count; $index++) {
            [void]$table.Append('<w:tc><w:tcPr><w:tcW w:w="' + $widths[$index] + '" w:type="dxa"/>')
            if ($isHeader) { [void]$table.Append('<w:shd w:val="clear" w:fill="EEEEEE"/>') }
            [void]$table.Append('<w:vAlign w:val="top"/></w:tcPr>')
            [void]$table.Append((New-WordParagraph $cells[$index] -InTable $true -Bold $isHeader))
            [void]$table.Append('</w:tc>')
        }
        [void]$table.Append('</w:tr>')
    }
    [void]$table.Append('</w:tbl>')
    return $table.ToString()
}

$body = [System.Text.StringBuilder]::new()
foreach ($node in $fragment.DocumentElement.ChildNodes) {
    switch ($node.Name) {
        'h1' { [void]$body.Append((New-WordParagraph $node 'DocumentTitle')) }
        { $_ -in @('h2', 'h3', 'h4') } { [void]$body.Append((New-WordParagraph $node 'SectionTitle')) }
        'p' {
            $style = if ($node.InnerText -match '^[1-7]\. ') { 'QuestionTitle' } elseif ($node.InnerText -match '^[①②③④] ') { 'SectionTitle' } else { 'Normal' }
            [void]$body.Append((New-WordParagraph $node $style))
        }
        'table' { [void]$body.Append((Convert-WordTable $node)) }
        { $_ -in @('ol', 'ul') } {
            $number = 1
            foreach ($item in $node.SelectNodes('li')) {
                $prefix = if ($node.Name -eq 'ol') { [string]$number + '. ' } else { '• ' }
                [void]$body.Append((New-WordParagraph $item -Prefix $prefix))
                $number++
            }
        }
        'blockquote' { foreach ($paragraph in $node.SelectNodes('p')) { [void]$body.Append((New-WordParagraph $paragraph 'Guidance')) } }
        '#whitespace' { }
        default { throw "지원하지 않는 Word 변환 요소입니다: $($node.Name)" }
    }
}

$document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="' + $wordNamespace + '" xmlns:r="' + $relationshipsNamespace + '"><w:body>' + $body.ToString() + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1020" w:right="1020" w:bottom="1020" w:left="1020" w:header="480" w:footer="480" w:gutter="0"/></w:sectPr></w:body></w:document>'
$styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="' + $wordNamespace + '"><w:docDefaults><w:rPrDefault><w:rPr>' + $fontProperties + '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="320" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
$styles += '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr>' + $fontProperties + '</w:rPr></w:style>'
$styles += '<w:style w:type="paragraph" w:styleId="Guidance"><w:name w:val="양식 안내"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:keepLines/><w:ind w:left="120" w:right="120"/><w:shd w:val="clear" w:fill="F2F4F1"/></w:pPr><w:rPr>' + $fontProperties + '<w:color w:val="59635B"/></w:rPr></w:style>'
# Word의 제목 접기 기능이 생기지 않도록 본문 수준의 사용자 지정 제목 서식을 사용한다.
foreach ($heading in @(@('DocumentTitle', '기획서 제목'), @('QuestionTitle', '문항 제목'), @('SectionTitle', '본문 소제목'))) {
    $alignment = if ($heading[0] -eq 'DocumentTitle') { '<w:jc w:val="center"/>' } else { '' }
    $styles += '<w:style w:type="paragraph" w:styleId="' + $heading[0] + '"><w:name w:val="' + $heading[1] + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="160"/>' + $alignment + '<w:outlineLvl w:val="9"/></w:pPr><w:rPr>' + $fontProperties + '<w:b/><w:bCs/></w:rPr></w:style>'
}
$styles += '</w:styles>'
$documentRelationships = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="' + $relationshipsNamespace + '/styles" Target="styles.xml"/>'
foreach ($link in $script:wordLinks) {
    $documentRelationships += '<Relationship Id="' + $link.Id + '" Type="' + $relationshipsNamespace + '/hyperlink" Target="' + (Convert-XmlText $link.Target) + '" TargetMode="External"/>'
}
$documentRelationships += '</Relationships>'
$timestamp = [DateTime]::UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
$parts = [ordered]@{
    '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'
    '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDocument" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rIdApp" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'
    'word/document.xml' = $document
    'word/styles.xml' = $styles
    'word/_rels/document.xml.rels' = $documentRelationships
    'docProps/core.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>AX 해커톤 아이디어 기획서</dc:title><dc:creator>WorkMemory AI</dc:creator><dc:description>바탕체 14포인트, 편집 가능한 표를 포함한 제출용 기획서</dc:description><dcterms:created xsi:type="dcterms:W3CDTF">' + $timestamp + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + $timestamp + '</dcterms:modified></cp:coreProperties>'
    'docProps/app.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>WorkMemory AI proposal export</Application></Properties>'
}

# OOXML 패키지를 직접 작성하므로 Word 설치나 추가 패키지가 필요하지 않다.
$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$file = [System.IO.File]::Open($resolvedOutputPath, [System.IO.FileMode]::Create)
try {
    $archive = [System.IO.Compression.ZipArchive]::new($file, [System.IO.Compression.ZipArchiveMode]::Create, $true)
    try {
        foreach ($part in $parts.GetEnumerator()) {
            [void][xml]$part.Value
            $entry = $archive.CreateEntry($part.Key, [System.IO.Compression.CompressionLevel]::Optimal)
            $writer = [System.IO.StreamWriter]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
            try { $writer.Write($part.Value) } finally { $writer.Dispose() }
        }
    } finally { $archive.Dispose() }
} finally { $file.Dispose() }
Write-Output "Word 생성 완료: $resolvedOutputPath (바탕체 14pt, 표 $($fragment.SelectNodes('//table').Count)개)"
