#Requires -Version 7.2
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$outputPath = [System.IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $outputPath) {
    if (-not (Test-Path -LiteralPath $outputPath -PathType Container)) {
        throw '출력 경로는 새 폴더 또는 빈 폴더여야 합니다.'
    }
    if (@(Get-ChildItem -LiteralPath $outputPath -Force).Count -ne 0) {
        throw '공개할 파일이 섞이지 않도록 빈 출력 폴더를 지정해 주세요.'
    }
} else {
    New-Item -ItemType Directory -Path $outputPath | Out-Null
}

$indexPath = Join-Path $outputPath 'index.html'
& (Join-Path $PSScriptRoot 'export-proposal-html.ps1') -OutputPath $indexPath
$html = [System.IO.File]::ReadAllText($indexPath)
$sourceRepositoryUrl = 'https://github.com/Wookshin/AX-hackaton'
$html = [regex]::Replace(
    $html,
    'href="(AX-hackathon-(?:proposal|implementation-notes|review)\.md)(#[^\"]*)?"',
    {
        param($link)
        'href="{0}/blob/main/{1}{2}"' -f $sourceRepositoryUrl, $link.Groups[1].Value, $link.Groups[2].Value
    }
)
if ($html -match 'href="(?!https?://|#)[^\"]+"') {
    throw '배포 대상에 포함되지 않은 상대 링크가 남아 있습니다.'
}
[System.IO.File]::WriteAllText($indexPath, $html, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllText((Join-Path $outputPath '.nojekyll'), '')
Write-Output "Pages 배포 파일 생성 완료: $outputPath"
