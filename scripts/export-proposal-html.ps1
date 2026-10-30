#Requires -Version 7.2
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\AX-hackathon-proposal.html')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$proposalPath = Join-Path $repositoryRoot 'AX-hackathon-proposal.md'
$reviewPath = Join-Path $repositoryRoot 'AX-hackathon-review.md'
$proposal = [System.IO.File]::ReadAllText($proposalPath)
$review = [System.IO.File]::ReadAllText($reviewPath)

function Convert-DocumentFragment([string]$Markdown) {
    $html = (ConvertFrom-Markdown -InputObject $Markdown).Html
    $html = $html.Replace('<table>', '<div class="table-wrap" role="region" aria-label="가로로 넘겨 볼 수 있는 표" tabindex="0"><table>')
    $html = $html.Replace('</table>', '</table></div>')
    return [regex]::Replace($html, '<a href="(https?://[^\"]+)"', '<a target="_blank" rel="noopener noreferrer" href="$1"')
}

function Convert-InlinePlainText([string]$Text) {
    $text = [regex]::Replace($Text, '\[([^\]]+)\]\((https?://[^)\s]+)\)', '$1 ($2)')
    return $text.Replace('**', '').Replace('`', '')
}

function Convert-ProposalPlainText([string]$Markdown) {
    $lines = [System.Collections.Generic.List[string]]::new()
    $tableHeaders = $null
    foreach ($line in ($Markdown -split '\r?\n')) {
        if ($line.Trim() -match '^\|.*\|$') {
            $cells = @($line.Trim().Trim('|').Split('|') | ForEach-Object { Convert-InlinePlainText $_.Trim() })
            if (@($cells | Where-Object { $_ -notmatch '^:?-{3,}:?$' }).Count -eq 0) { continue }
            if ($null -eq $tableHeaders) {
                $tableHeaders = $cells
                continue
            }
            if ($cells.Count -ne $tableHeaders.Count) { throw '텍스트로 변환할 표의 열 수를 확인해 주세요.' }
            for ($column = 0; $column -lt $cells.Count; $column++) {
                $lines.Add(('{0}: {1}' -f $tableHeaders[$column], $cells[$column]))
            }
            $lines.Add('')
            continue
        }
        $tableHeaders = $null
        $text = [regex]::Replace($line, '^\s{0,3}#{1,6}\s+', '')
        $text = [regex]::Replace($text, '^\s*[-+*]\s+', '• ')
        $lines.Add((Convert-InlinePlainText $text))
    }
    $text = [regex]::Replace(($lines -join "`n").Trim(), '\n{3,}', "`n`n")
    return $text.Replace("`n", "`r`n") + "`r`n"
}

$proposalText = Convert-ProposalPlainText $proposal
# UTF-8 BOM과 Windows 줄바꿈으로 메모장과 Word에서 한글을 바로 읽을 수 있게 한다.
$textEncoding = [System.Text.UTF8Encoding]::new($true)
$proposalTextBytes = [byte[]]($textEncoding.GetPreamble() + $textEncoding.GetBytes($proposalText))
$proposalTextBase64 = [Convert]::ToBase64String($proposalTextBytes)
$textOutputPath = Join-Path (Split-Path -Parent ([System.IO.Path]::GetFullPath($OutputPath))) 'AX-hackathon-proposal.txt'

$sectionTitles = @('문제와 주요 사용자', 'AI 활용 계획', '2일 개발 계획', '기대효과와 검증', '한 줄 제안', '데이터와 보호', '확장과 차별점')
$questions = [regex]::Matches($proposal, '(?m)^\*\*([1-7])\. (.+)\*\*\r?$')
if ($questions.Count -ne 7) { throw '기획서의 제출 질문 7개를 찾을 수 없습니다.' }
$sections = for ($index = 0; $index -lt $questions.Count; $index++) {
    $number = $index + 1
    if ($questions[$index].Groups[1].Value -ne [string]$number) { throw '제출 질문 순서를 확인해 주세요.' }
    $bodyStart = $questions[$index].Index + $questions[$index].Length
    $bodyEnd = if ($index + 1 -lt $questions.Count) { $questions[$index + 1].Index } else { $proposal.Length }
    $markdown = $proposal.Substring($bodyStart, $bodyEnd - $bodyStart).Trim()
    $markdown = [regex]::Replace($markdown, '(?m)^\*\*([①②③④]) (.+)\*\*\r?$', '### $1 $2')
    $content = Convert-DocumentFragment $markdown
    $questionText = [System.Net.WebUtility]::HtmlEncode($questions[$index].Groups[2].Value)
    $sectionTitle = $sectionTitles[$index]
    @"
<section class="proposal-section" id="proposal-$number" aria-labelledby="heading-$number">
  <div class="section-heading"><span class="section-number">0$number</span><h2 id="heading-$number">$sectionTitle</h2></div>
  <p class="question">$questionText</p>
  <div class="prose">$content</div>
</section>
"@
}

$navigation = for ($index = 0; $index -lt $sectionTitles.Count; $index++) {
    $number = $index + 1
    '<a href="#proposal-{0}"><span>0{0}</span>{1}</a>' -f $number, $sectionTitles[$index]
}
$reviewHeadings = [regex]::Matches($review, '(?m)^## ([1-7])\. (.+)\r?$')
if ($reviewHeadings.Count -ne 7) { throw '검증 기록의 7개 항목을 찾을 수 없습니다.' }
$reviewSections = for ($index = 0; $index -lt $reviewHeadings.Count; $index++) {
    $bodyStart = $reviewHeadings[$index].Index + $reviewHeadings[$index].Length
    $bodyEnd = if ($index + 1 -lt $reviewHeadings.Count) { $reviewHeadings[$index + 1].Index } else { $review.Length }
    $content = Convert-DocumentFragment ($review.Substring($bodyStart, $bodyEnd - $bodyStart).Trim())
    $heading = [System.Net.WebUtility]::HtmlEncode($reviewHeadings[$index].Groups[2].Value.Trim())
    '<details class="review-item"><summary>{0}</summary><div class="prose">{1}</div></details>' -f $heading, $content
}
$reviewLead = $review.Substring(0, $reviewHeadings[0].Index)
$reviewLead = [regex]::Replace($reviewLead, '(?m)^# .+\r?\n', '').Trim()
$reviewLeadHtml = Convert-DocumentFragment $reviewLead
$coreMatch = [regex]::Match($proposal, '(?m)^\*\*핵심 제안:\*\* (.+)\r?$')
if (-not $coreMatch.Success) { throw '핵심 제안을 찾을 수 없습니다.' }
$coreHtml = Convert-DocumentFragment $coreMatch.Groups[1].Value.Trim()
$experienceMatch = [regex]::Match($proposal, '(?m)^\*\*대표 경험:\*\* “(.+?)”\s*(.+)\r?$')
if (-not $experienceMatch.Success) { throw '대표 경험을 찾을 수 없습니다.' }
$discoveryQuote = [System.Net.WebUtility]::HtmlEncode('“' + $experienceMatch.Groups[1].Value + '”')
$discoveryNote = [System.Net.WebUtility]::HtmlEncode($experienceMatch.Groups[2].Value.Trim())
$proposalHash = (Get-FileHash -LiteralPath $proposalPath -Algorithm SHA256).Hash.ToLowerInvariant()
$reviewHash = (Get-FileHash -LiteralPath $reviewPath -Algorithm SHA256).Hash.ToLowerInvariant()

$template = @'
<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="description" content="WorkMemory AI: 평소처럼 일하며 쌓인 데이터를 AI가 업무 성격에 맞게 분류·패턴화해 효율화 솔루션을 제안하고, 업무 회고와 문서 자동화를 돕는 AI 업무 비서.">
  <meta name="proposal-source-sha256" content="<!--PROPOSAL_HASH-->">
  <meta name="review-source-sha256" content="<!--REVIEW_HASH-->">
  <title>WorkMemory AI · AX 해커톤 기획서</title>
  <style>
    :root { color-scheme: light; --paper:#f6f7f3; --white:#fff; --ink:#18382f; --text:#35443e; --muted:#637269; --line:#dde5dc; --green:#14735b; --wash:#eaf2e9; --warm:#fbf4e7; --gold:#855714; }
    * { box-sizing:border-box; }
    html { scroll-behavior:smooth; scroll-padding-top:100px; }
    body { margin:0; background:var(--paper); color:var(--text); font-family:"Segoe UI","Apple SD Gothic Neo","Malgun Gothic",sans-serif; line-height:1.85; word-break:keep-all; overflow-wrap:anywhere; }
    a { color:var(--green); text-underline-offset:4px; }
    a:hover { color:var(--ink); }
    :focus-visible { outline:3px solid #be8b34; outline-offset:5px; border-radius:4px; }
    button { font:inherit; cursor:pointer; }
    .skip-link { position:fixed; top:-80px; left:16px; z-index:30; background:var(--ink); color:white; padding:10px 18px; }
    .skip-link:focus { top:10px; }
    .reading-progress { position:fixed; top:0; left:0; width:0; height:3px; z-index:25; background:var(--green); }
    .topbar { position:sticky; top:0; z-index:20; height:76px; border-bottom:1px solid var(--line); background:rgba(246,247,243,.97); }
    .topbar-inner { max-width:1380px; margin:auto; height:100%; padding:0 42px; display:flex; align-items:center; justify-content:space-between; gap:20px; }
    .brand { display:flex; align-items:center; gap:12px; color:var(--ink); text-decoration:none; font-size:19px; font-weight:750; letter-spacing:-.5px; }
    .brand-mark { width:34px; height:34px; border:1px solid var(--green); border-radius:11px; display:grid; place-items:center; font-size:13px; letter-spacing:-1px; }
    .brand small { color:var(--muted); font-size:11px; letter-spacing:1.2px; font-weight:600; padding-left:4px; }
    .top-actions { display:flex; align-items:center; gap:19px; font-size:13px; }
    .top-actions a { color:var(--muted); text-decoration:none; }
    .print-button { background:var(--ink); color:white; border:0; border-radius:8px; padding:9px 16px; line-height:1.6; font-size:13px; }
    .text-download-button { background:var(--white); color:var(--green); border:1px solid var(--line); border-radius:8px; padding:8px 14px; line-height:1.6; font-size:13px; white-space:nowrap; }
    .layout { max-width:1380px; margin:0 auto; padding:0 42px; display:grid; grid-template-columns:210px minmax(0,1fr); gap:55px; }
    .sidebar { position:sticky; top:106px; align-self:start; padding-top:54px; }
    .nav-label,.eyebrow { font-size:11px; font-weight:750; letter-spacing:2px; color:var(--muted); }
    .sidebar nav { margin-top:18px; display:grid; gap:5px; }
    .sidebar nav a { display:flex; align-items:center; gap:14px; padding:10px 12px; color:var(--muted); text-decoration:none; font-size:13px; border-radius:8px; line-height:1.5; }
    .sidebar nav a span { font-size:11px; font-variant-numeric:tabular-nums; opacity:.7; }
    .sidebar nav a:hover { background:#edf1ea; }
    .sidebar nav a[aria-current="location"] { background:var(--wash); color:var(--green); font-weight:700; }
    .sidebar-note { border-top:1px solid var(--line); margin-top:23px; padding:19px 12px 0; font-size:12px; color:var(--muted); }
    .sidebar-note a { display:block; margin-top:9px; text-decoration:none; }
    main { min-width:0; padding-bottom:75px; }
    .hero { padding:61px 0 44px; }
    .hero .eyebrow { display:flex; align-items:center; gap:10px; color:var(--green); }
    .hero .eyebrow::before { content:""; width:7px; height:7px; border-radius:50%; background:var(--green); }
    h1 { font-size:clamp(34px,4.1vw,57px); line-height:1.28; letter-spacing:-2.6px; color:var(--ink); font-weight:750; margin:21px 0 23px; }
    h1 em { color:var(--green); font-style:normal; }
    .hero-description { max-width:800px; font-size:16px; line-height:1.95; }
    .hero-description p { margin:0; }
    .tags { display:flex; flex-wrap:wrap; gap:8px; margin-top:25px; }
    .tag { font-size:11px; color:var(--green); border:1px solid #d0decf; background:#f0f5ed; border-radius:5px; padding:3px 9px; }
    .discovery { background:var(--ink); color:#edf4ec; border-radius:15px; padding:28px 31px; margin-top:36px; }
    .discovery-label { color:#b7d8c5; font-size:11px; letter-spacing:1.5px; font-weight:700; }
    .discovery blockquote { border:0; margin:12px 0 14px; padding:0; font-size:21px; font-weight:600; line-height:1.7; color:#fff; letter-spacing:-.4px; }
    .discovery small { display:block; color:#c2d4c8; font-size:11px; }
    .journey { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:12px; margin-top:16px; }
    .journey-card { border:1px solid var(--line); background:var(--white); border-radius:11px; padding:19px; }
    .journey-card span { display:block; font-size:10px; color:var(--green); font-weight:750; letter-spacing:1.1px; }
    .journey-card h3 { margin:7px 0 7px; color:var(--ink); font-size:16px; line-height:1.5; }
    .journey-card p { margin:0; font-size:12px; color:var(--muted); line-height:1.8; }
    .flow { display:flex; align-items:center; flex-wrap:wrap; gap:9px; font-size:12px; margin-top:23px; color:var(--green); }
    .flow b { font-weight:600; }
    .flow span { color:#93a493; }
    .document-download { margin-top:18px; display:flex; flex-wrap:wrap; align-items:center; gap:12px; font-size:12px; }
    .document-download a { font-weight:700; }
    .document-download span { color:var(--muted); }
    .proposal-section { padding:39px 0 40px; border-top:1px solid var(--line); }
    .section-heading { display:flex; align-items:center; gap:15px; }
    .section-number { font-size:12px; font-weight:750; color:var(--green); width:30px; height:30px; display:grid; place-items:center; background:var(--wash); border-radius:8px; flex-shrink:0; }
    h2 { margin:0; color:var(--ink); font-size:25px; line-height:1.5; letter-spacing:-.8px; }
    .question { font-size:12px; line-height:1.8; color:var(--muted); margin:10px 0 25px 45px; }
    .prose { font-size:15px; }
    .prose p { margin:0 0 17px; }
    .prose strong { color:var(--ink); font-weight:700; }
    .prose h3 { font-size:18px; color:var(--ink); margin:30px 0 15px; line-height:1.6; letter-spacing:-.3px; }
    .prose h3:first-child { margin-top:0; }
    .prose ul,.prose ol { padding-left:24px; margin:16px 0 20px; }
    .prose li { padding-left:4px; margin-bottom:11px; }
    .prose code { font-family:Consolas,monospace; font-size:.86em; padding:2px 5px; border-radius:4px; background:#edf1e9; color:#28614e; overflow-wrap:anywhere; }
    .table-wrap { overflow-x:auto; border:1px solid var(--line); border-radius:10px; margin:22px 0; background:var(--white); }
    table { width:100%; border-collapse:collapse; table-layout:auto; font-size:13px; line-height:1.85; }
    th { padding:13px 15px; text-align:left; background:#edf3e9; color:var(--ink); font-weight:700; border-bottom:1px solid var(--line); }
    td { padding:14px 15px; border-bottom:1px solid #e8ede4; vertical-align:top; min-width:115px; }
    td:first-child { font-weight:600; color:var(--ink); }
    tr:last-child td { border-bottom:0; }
    tbody tr:hover { background:#fafcf7; }
    #proposal-5 .prose { background:var(--wash); border-left:3px solid var(--green); border-radius:0 10px 10px 0; padding:25px 29px; font-size:22px; line-height:1.8; letter-spacing:-.5px; }
    #proposal-5 .prose p { margin:0; }
    .review-section { margin-top:12px; padding-top:36px; border-top:2px solid var(--ink); }
    .review-section > p { font-size:14px; color:var(--muted); margin-bottom:23px; }
    .review-intro { margin:20px 0; font-size:12px; color:var(--muted); }
    .review-intro p { margin:0; }
    .review-item { border:1px solid var(--line); border-radius:9px; margin-bottom:10px; background:var(--white); }
    .review-item summary { padding:16px 20px; color:var(--ink); font-size:14px; font-weight:650; cursor:pointer; }
    .review-item summary::marker { color:var(--green); }
    .review-item[open] summary { border-bottom:1px solid var(--line); }
    .review-item > .prose { padding:22px; font-size:14px; }
    .review-item > .prose > :last-child { margin-bottom:0; }
    .footer { border-top:1px solid var(--line); margin-top:42px; padding-top:24px; display:flex; justify-content:space-between; gap:16px; font-size:11px; color:var(--muted); }
    .footer a { color:var(--muted); }
    @media(min-width:1500px) { .layout { gap:70px; } }
    @media(max-width:1000px) {
      .layout { grid-template-columns:170px minmax(0,1fr); gap:26px; padding:0 25px; }
      .topbar-inner { padding:0 25px; }
      .sidebar nav a { padding:9px 7px; gap:9px; font-size:12px; }
      .journey { gap:9px; } .journey-card { padding:15px; }
    }
    @media(max-width:760px) {
      html { scroll-padding-top:130px; }
      .topbar { height:65px; }
      .topbar-inner { padding:0 20px; }
      .brand { font-size:17px; gap:9px; } .brand small { display:none; }
      .top-actions { gap:12px; } .top-actions .source-link { display:none; }
      .print-button,.text-download-button { padding:7px 11px; font-size:11px; }
      .layout { display:block; padding:0 21px; }
      .sidebar { top:65px; margin:0 -21px; padding:0 15px; z-index:15; background:var(--paper); border-bottom:1px solid var(--line); }
      .nav-label,.sidebar-note { display:none; }
      .sidebar nav { display:flex; overflow-x:auto; gap:5px; margin:0; padding:10px 0; scrollbar-width:thin; }
      .sidebar nav a { flex-shrink:0; font-size:11px; padding:7px 10px; gap:6px; white-space:nowrap; }
      .hero { padding:35px 0 30px; }
      h1 { font-size:39px; letter-spacing:-1.9px; }
      .hero-description { font-size:14px; }
      .discovery { padding:23px; margin-top:26px; }
      .discovery blockquote { font-size:18px; }
      .journey { grid-template-columns:1fr; }
      .journey-card { padding:17px 19px; }
      .journey-card h3 { margin:3px 0 5px; }
      .proposal-section { padding:29px 0; }
      h2 { font-size:22px; } .question { margin-left:0; font-size:11px; }
      .prose { font-size:14px; } .prose h3 { font-size:17px; }
      table { font-size:12px; } th,td { padding:12px; } td { min-width:122px; }
      #proposal-5 .prose { font-size:19px; padding:22px; }
      .review-item > .prose { padding:17px; }
      .footer { flex-direction:column; gap:7px; }
    }
    @media(max-width:480px) {
      html { scroll-padding-top:165px; }
      .topbar { height:105px; }
      .topbar-inner { flex-direction:column; align-items:flex-start; justify-content:center; gap:8px; }
      .brand-mark { display:none; }
      .sidebar { top:105px; }
    }
    @media(prefers-reduced-motion:reduce) { html { scroll-behavior:auto; } }
    @media print {
      @page { size:A4; margin:17mm 15mm; }
      html { scroll-behavior:auto; }
      body { background:white; color:#222; font-size:10pt; print-color-adjust:exact; -webkit-print-color-adjust:exact; }
      .screen-only,.topbar,.sidebar,.reading-progress,.skip-link { display:none!important; }
      .layout { display:block; max-width:none; padding:0; }
      main { padding:0; }
      .hero { padding:0 0 22px; }
      h1 { font-size:32pt; }
      .hero-description,.prose { font-size:10pt; line-height:1.75; }
      .discovery { padding:20px 24px; break-inside:avoid; }
      .discovery blockquote { font-size:15pt; }
      .journey { display:none; }
      .proposal-section { padding:22px 0; }
      .section-heading,.prose h3,.review-item summary { break-after:avoid; }
      h2 { font-size:17pt; }
      .question { font-size:9pt; margin-bottom:16px; }
      .table-wrap { overflow:visible; border-radius:0; }
      table { font-size:8.5pt; }
      th,td { padding:8px; min-width:0; }
      thead { display:table-header-group; }
      tr { break-inside:avoid; }
      p { orphans:3; widows:3; }
      #proposal-5 .prose { font-size:15pt; padding:18px 22px; }
      .review-section { break-before:page; }
      .review-item { border:0; border-bottom:1px solid var(--line); border-radius:0; }
      .review-item > .prose { padding:12px 0; font-size:9pt; }
      .review-item summary { padding:14px 0 8px; list-style:none; }
      .footer { font-size:8pt; }
      a { color:inherit; text-decoration:none; }
    }
  </style>
</head>
<body>
  <a class="skip-link" href="#main-content">본문으로 바로가기</a>
  <div class="reading-progress" aria-hidden="true"></div>
  <header class="topbar screen-only">
    <div class="topbar-inner">
      <a class="brand" href="#main-content"><span class="brand-mark" aria-hidden="true">wm</span>WorkMemory AI <small>PROPOSAL</small></a>
      <div class="top-actions"><button class="text-download-button" id="download-text" type="button" title="기획서 본문을 Word에 복사할 수 있는 순수 텍스트로 저장">텍스트 다운로드</button><button class="print-button" id="print-document" type="button">인쇄 / PDF 저장</button></div>
    </div>
  </header>
  <div class="layout">
    <aside class="sidebar screen-only" aria-label="문서 목차">
      <div class="nav-label">CONTENTS</div>
      <nav aria-label="기획서 목차"><!--NAVIGATION--></nav>
      <div class="sidebar-note">AX 해커톤 아이디어 기획서<br>일상의 기록이 업무 개선으로.<a href="#review">심사기준 검증 기록 ↗</a></div>
    </aside>
    <main id="main-content" tabindex="-1">
      <section class="hero" aria-labelledby="document-title">
        <div class="eyebrow">AX HACKATHON · WORK MEMORY</div>
        <h1 id="document-title">평소의 업무가<br><em>효율화 솔루션으로.</em></h1>
        <div class="hero-description"><!--CORE_PROPOSAL--></div>
        <div class="tags"><span class="tag">업무 자동 분류·패턴화</span><span class="tag">효율화 솔루션</span><span class="tag">업무 회고·문서 자동화</span><span class="tag">선택적 데이터 수집</span></div>
        <div class="discovery">
          <div class="discovery-label">THE MOMENT OF DISCOVERY</div>
          <blockquote><!--DISCOVERY_QUOTE--></blockquote>
          <small><!--DISCOVERY_NOTE--></small>
        </div>
        <div class="journey" aria-label="핵심 가치의 흐름">
          <article class="journey-card"><span>01 · COLLECT</span><h3>평소처럼 일하면</h3><p>허용한 범위의 데이터가 자동으로 쌓이고 업무 성격에 맞게 분류됩니다.</p></article>
          <article class="journey-card"><span>02 · SOLVE</span><h3>패턴에서 개선 방법을</h3><p>반복 흐름을 패턴화해 적용 방법을 담은 효율화 솔루션을 제안합니다.</p></article>
          <article class="journey-card"><span>03 · REUSE</span><h3>회고와 문서 자동화로</h3><p>같은 기록으로 나의 업무를 돌아보고 필요한 보고 문서의 초안을 만듭니다.</p></article>
        </div>
        <div class="flow" aria-label="해커톤 개발 목표"><b>선택 범위의 데이터 자동 축적</b><span>→</span><b>업무 분류·패턴화</b><span>→</span><b>효율화 솔루션 제안</b><span>+</span><b>업무 회고·문서 자동화</b></div>
        <div class="document-download screen-only"><a href="AX-hackathon-proposal.docx" download>Word 기획서 다운로드 (.docx)</a><span>편집 가능한 표 · 바탕체 14포인트</span></div>
      </section>
      <!--PROPOSAL_SECTIONS-->
      <section id="review" class="review-section" aria-labelledby="review-title">
        <div class="section-heading"><span class="section-number">+</span><h2 id="review-title">심사기준 기반 검증</h2></div>
        <p>기존 검증 에이전트 3개의 검토 이력과 사용자 피드백 반영 내용, 심사기준과의 연결, 남은 검증 과제입니다. 각 항목을 펼쳐 읽을 수 있습니다.</p>
        <div class="review-intro"><!--REVIEW_LEAD--></div>
        <!--REVIEW_SECTIONS-->
      </section>
      <footer class="footer"><span>WorkMemory AI · AX 해커톤 아이디어 기획서</span><span><a href="AX-hackathon-proposal.md">기획서 원문</a> · <a href="AX-hackathon-implementation-notes.md">설계 메모</a> · <a href="AX-hackathon-review.md">검증 기록</a></span></footer>
    </main>
  </div>
  <script>
    (() => {
      const links = [...document.querySelectorAll('nav a')];
      const sections = links.map(link => document.querySelector(link.getAttribute('href')));
      const progress = document.querySelector('.reading-progress');
      let scheduled = false;
      function updateReadingPosition() {
        const scrollRange = document.documentElement.scrollHeight - window.innerHeight;
        progress.style.width = (scrollRange > 0 ? Math.min(100, window.scrollY / scrollRange * 100) : 100) + '%';
        let current = 0;
        for (let index = 0; index < sections.length; index++) {
          if (sections[index].getBoundingClientRect().top <= 160) current = index;
        }
        links.forEach((link, index) => {
          if (index === current) link.setAttribute('aria-current', 'location');
          else link.removeAttribute('aria-current');
        });
        scheduled = false;
      }
      function scheduleUpdate() {
        if (!scheduled) { scheduled = true; requestAnimationFrame(updateReadingPosition); }
      }
      window.addEventListener('scroll', scheduleUpdate, { passive: true });
      window.addEventListener('resize', scheduleUpdate);
      document.querySelectorAll('details').forEach(detail => detail.addEventListener('toggle', scheduleUpdate));
      let printState = null;
      window.addEventListener('beforeprint', () => {
        if (printState) return;
        printState = [...document.querySelectorAll('details')].map(detail => [detail, detail.open]);
        printState.forEach(([detail]) => { detail.open = true; });
      });
      window.addEventListener('afterprint', () => {
        if (printState) printState.forEach(([detail, wasOpen]) => { detail.open = wasOpen; });
        printState = null;
        scheduleUpdate();
      });
      document.getElementById('print-document').addEventListener('click', () => window.print());
      document.getElementById('download-text').addEventListener('click', () => {
        const bytes = Uint8Array.from(atob('<!--PROPOSAL_TEXT_BASE64-->'), character => character.charCodeAt(0));
        const url = URL.createObjectURL(new Blob([bytes], { type: 'text/plain;charset=utf-8' }));
        const link = document.createElement('a');
        link.href = url;
        link.download = 'AX-hackathon-proposal.txt';
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      });
      updateReadingPosition();
    })();
  </script>
</body>
</html>
'@

$html = $template.Replace('<!--PROPOSAL_HASH-->', $proposalHash).Replace('<!--REVIEW_HASH-->', $reviewHash)
$html = $html.Replace('<!--NAVIGATION-->', ($navigation -join "`n"))
$html = $html.Replace('<!--CORE_PROPOSAL-->', $coreHtml)
$html = $html.Replace('<!--DISCOVERY_QUOTE-->', $discoveryQuote).Replace('<!--DISCOVERY_NOTE-->', $discoveryNote)
$html = $html.Replace('<!--PROPOSAL_SECTIONS-->', ($sections -join "`n"))
$html = $html.Replace('<!--REVIEW_LEAD-->', $reviewLeadHtml)
$html = $html.Replace('<!--REVIEW_SECTIONS-->', ($reviewSections -join "`n"))
$html = $html.Replace('<!--PROPOSAL_TEXT_BASE64-->', $proposalTextBase64)
if ($html -match '<!--[A-Z_]+-->') { throw '채워지지 않은 HTML 템플릿 항목이 있습니다.' }
[System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($OutputPath), $html, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::WriteAllBytes($textOutputPath, $proposalTextBytes)
$wordOutputPath = Join-Path (Split-Path -Parent ([System.IO.Path]::GetFullPath($OutputPath))) 'AX-hackathon-proposal.docx'
& (Join-Path $PSScriptRoot 'export-proposal-word.ps1') -OutputPath $wordOutputPath
Write-Output "HTML 생성 완료: $([System.IO.Path]::GetFullPath($OutputPath))"
Write-Output "텍스트 생성 완료: $textOutputPath"
