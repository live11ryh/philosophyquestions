$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression

function Escape-Html([string]$text) {
  return [System.Net.WebUtility]::HtmlEncode($text)
}

function Get-Slug([string]$name) {
  $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
  $slug = $base.Normalize([Text.NormalizationForm]::FormD)
  $slug = [regex]::Replace($slug, '[^\p{IsCJKUnifiedIdeographs}A-Za-z0-9]+', '-').Trim('-')
  if ([string]::IsNullOrWhiteSpace($slug)) {
    $slug = [guid]::NewGuid().ToString("N").Substring(0, 8)
  }
  return $slug
}

function Convert-Inline([string]$text, $references) {
  $encoded = Escape-Html $text
  $encoded = [regex]::Replace($encoded, '\[([^\]]+)\]\[(\d+)\]', {
    param($m)
    $label = $m.Groups[1].Value
    $num = $m.Groups[2].Value
    if ($references.Contains($num)) {
      return '<a class="source-link" href="#ref-' + $num + '">' + $label + '</a>'
    }
    return $m.Value
  })
  $encoded = [regex]::Replace($encoded, '\*\*(.+?)\*\*', '<strong>$1</strong>')
  $encoded = [regex]::Replace($encoded, '\*(.+?)\*', '<em>$1</em>')
  return $encoded
}

function Formula-Html([string]$text) {
  $out = Escape-Html $text.Trim()
  $out = $out -replace '\\times', '<span class="math-times">×</span>'
  $out = [regex]::Replace($out, '\^\{([^}]+)\}', '<sup>$1</sup>')
  $out = [regex]::Replace($out, '\^(\d+)', '<sup>$1</sup>')
  return $out
}

function New-Anchor([string]$heading, [int]$index) {
  $plain = [regex]::Replace($heading, '[#*_`]+', '')
  $anchor = [regex]::Replace($plain, '[^\p{IsCJKUnifiedIdeographs}A-Za-z0-9]+', '-').Trim('-')
  if ([string]::IsNullOrWhiteSpace($anchor)) { $anchor = "section" }
  return "s$index-$($anchor.ToLower())"
}

function Render-Table($rows, $references) {
  $html = [System.Collections.Generic.List[string]]::new()
  $html.Add('<div class="table-wrap"><table>')
  $headerDone = $false
  foreach ($row in $rows) {
    if ($row -match '^\s*\|?\s*:?-{2,}') { continue }
    $cells = $row.Trim().Trim('|').Split('|') | ForEach-Object { $_.Trim() }
    if (-not $headerDone) {
      $html.Add('<thead><tr>')
      foreach ($cell in $cells) { $html.Add('<th>' + (Convert-Inline $cell $references) + '</th>') }
      $html.Add('</tr></thead><tbody>')
      $headerDone = $true
    } else {
      $html.Add('<tr>')
      foreach ($cell in $cells) { $html.Add('<td>' + (Convert-Inline $cell $references) + '</td>') }
      $html.Add('</tr>')
    }
  }
  if ($headerDone) { $html.Add('</tbody>') }
  $html.Add('</table></div>')
  return ($html -join "`n")
}

function Render-List($rows, [bool]$ordered, $references) {
  $tag = if ($ordered) { 'ol' } else { 'ul' }
  $html = [System.Collections.Generic.List[string]]::new()
  $html.Add("<$tag>")
  foreach ($row in $rows) {
    $item = if ($ordered) { $row -replace '^\s*\d+\.\s+', '' } else { $row -replace '^\s*[*-]\s+', '' }
    $html.Add('<li>' + (Convert-Inline $item $references) + '</li>')
  }
  $html.Add("</$tag>")
  return ($html -join "`n")
}

function Get-DocxParagraphs([string]$path) {
  $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
  try {
    $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Read, $false)
    $entry = $zip.GetEntry('word/document.xml')
    $stream = $entry.Open()
    $reader = [System.IO.StreamReader]::new($stream)
    [xml]$xml = $reader.ReadToEnd()
    $ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $paragraphs = [System.Collections.Generic.List[string]]::new()
    foreach ($p in $xml.SelectNodes('//w:body/w:p', $ns)) {
      $text = (($p.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.'#text' }) -join '').Trim()
      if ($text.Length -gt 0) { $paragraphs.Add($text) }
    }
    return $paragraphs
  } finally {
    if ($reader) { $reader.Close() }
    if ($stream) { $stream.Close() }
    if ($zip) { $zip.Dispose() }
    $fs.Close()
  }
}

function Convert-DocxToHtml([System.IO.FileInfo]$file) {
  $paragraphs = Get-DocxParagraphs $file.FullName
  $references = [ordered]@{}
  $lines = [System.Collections.Generic.List[string]]::new()
  foreach ($line in $paragraphs) {
    if ($line -match '^\[(\d+)\]:\s+(\S+)\s+"([^"]+)"') {
      $url = $matches[2] -replace '\?utm_source=chatgpt\.com', ''
      $references[$matches[1]] = [ordered]@{ url = $url; title = $matches[3] }
    } else {
      $lines.Add($line)
    }
  }

  $body = [System.Collections.Generic.List[string]]::new()
  $toc = [System.Collections.Generic.List[object]]::new()
  $headingIndex = 0
  $inSection = $false
  $leadDone = $false
  $i = 0

  while ($i -lt $lines.Count) {
    $line = $lines[$i]
    if ($line -eq '---') {
      $body.Add('<hr>')
      $i++
      continue
    }
    if ($line -match '^(#{1,4})\s+(.+)$') {
      if ($inSection) { $body.Add('</section>') }
      $headingIndex++
      $level = [Math]::Min($matches[1].Length + 1, 5)
      $title = $matches[2]
      $id = New-Anchor $title $headingIndex
      $sectionClass = if ($level -le 2) { 'chapter chapter-major' } else { 'chapter chapter-minor' }
      $body.Add('<section class="' + $sectionClass + '">')
      $body.Add("<h$level id=`"$id`"><span>" + (Convert-Inline $title $references) + "</span></h$level>")
      $toc.Add([pscustomobject]@{ id = $id; title = $title; level = $level })
      $inSection = $true
      $i++
      continue
    }
    if ($line -eq '[') {
      $formulaRows = [System.Collections.Generic.List[string]]::new()
      $i++
      while ($i -lt $lines.Count -and $lines[$i] -ne ']') {
        $formulaRows.Add($lines[$i])
        $i++
      }
      if ($i -lt $lines.Count -and $lines[$i] -eq ']') { $i++ }
      $body.Add('<div class="formula">' + (Formula-Html ($formulaRows -join ' ')) + '</div>')
      continue
    }
    if ($line -match '^\|') {
      $rows = [System.Collections.Generic.List[string]]::new()
      while ($i -lt $lines.Count -and $lines[$i] -match '^\|') {
        $rows.Add($lines[$i])
        $i++
      }
      $body.Add((Render-Table $rows $references))
      continue
    }
    if ($line -match '^>\s+') {
      $quotes = [System.Collections.Generic.List[string]]::new()
      while ($i -lt $lines.Count -and $lines[$i] -match '^>\s+') {
        $quotes.Add(($lines[$i] -replace '^>\s+', ''))
        $i++
      }
      $quoteText = ($quotes | ForEach-Object { Convert-Inline $_ $references }) -join '<br>'
      $body.Add('<blockquote>' + $quoteText + '</blockquote>')
      continue
    }
    if ($line -match '^\d+\.\s+') {
      $rows = [System.Collections.Generic.List[string]]::new()
      while ($i -lt $lines.Count -and $lines[$i] -match '^\d+\.\s+') {
        $rows.Add($lines[$i])
        $i++
      }
      $body.Add((Render-List $rows $true $references))
      continue
    }
    if ($line -match '^\s*[*-]\s+') {
      $rows = [System.Collections.Generic.List[string]]::new()
      while ($i -lt $lines.Count -and $lines[$i] -match '^\s*[*-]\s+') {
        $rows.Add($lines[$i])
        $i++
      }
      $body.Add((Render-List $rows $false $references))
      continue
    }
    $converted = Convert-Inline $line $references
    if ($line -match '^#{0,3}\s*\*\*.+\*\*:?\s*$') {
      $body.Add('<p class="kicker">' + $converted + '</p>')
    } elseif (-not $leadDone -and -not $inSection) {
      $body.Add('<p class="lede">' + $converted + '</p>')
      $leadDone = $true
    } else {
      $body.Add('<p>' + $converted + '</p>')
    }
    $i++
  }
  if ($inSection) { $body.Add('</section>') }

  $tocHtml = ($toc | ForEach-Object {
    $class = if ($_.level -le 2) { ' class="toc-major"' } else { ' class="toc-minor"' }
    '<a' + $class + ' href="#' + $_.id + '">' + (Escape-Html $_.title) + '</a>'
  }) -join "`n"

  if ([string]::IsNullOrWhiteSpace($tocHtml)) {
    $tocHtml = '<span class="toc-empty">正文无显式标题</span>'
  }

  $refsHtml = [System.Collections.Generic.List[string]]::new()
  foreach ($key in $references.Keys) {
    $ref = $references[$key]
    $refsHtml.Add('<li id="ref-' + $key + '"><a href="' + (Escape-Html $ref.url) + '">' + (Escape-Html $ref.title) + '</a></li>')
  }
  $refsBlock = if ($refsHtml.Count -gt 0) { $refsHtml -join "`n" } else { '<li>原文未列出独立参考链接。</li>' }

  $title = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
  $slug = Get-Slug $file.Name
  $compositionId = 'philosophy-' + ([regex]::Replace($slug, '[^\p{IsCJKUnifiedIdeographs}A-Za-z0-9-]', '-'))
  $generatedDate = Get-Date -Format 'yyyy-MM-dd'
  $contentHtml = $body -join "`n"
  $outputPath = Join-Path $file.DirectoryName ($slug + '.html')

  $html = @"
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>$([System.Net.WebUtility]::HtmlEncode($title))</title>
  <meta name="description" content="由 $([System.Net.WebUtility]::HtmlEncode($file.Name)) 转化而来的哲学长文 HTML。">
  <style>
    :root {
      --sky: #e8f5ff;
      --mint: #e8f8ef;
      --foam: #f7fcfb;
      --ink: #10212a;
      --muted: #587174;
      --deep: #155b64;
      --green: #32856a;
      --line: rgba(16, 33, 42, 0.14);
      --panel: rgba(255, 255, 255, 0.62);
    }
    * { box-sizing: border-box; }
    html { scroll-behavior: smooth; }
    body {
      margin: 0;
      color: var(--ink);
      background:
        radial-gradient(circle at 14% 6%, rgba(255,255,255,0.78), transparent 32%),
        linear-gradient(135deg, var(--sky) 0%, var(--foam) 48%, var(--mint) 100%);
      font-family: "Noto Serif SC", "Songti SC", "SimSun", serif;
      line-height: 1.85;
    }
    a { color: var(--deep); text-decoration-thickness: 1px; text-underline-offset: 0.22em; }
    .page-shell { min-height: 100vh; }
    .hero {
      padding: clamp(28px, 5vw, 64px) clamp(20px, 5vw, 72px) 32px;
      border-bottom: 1px solid var(--line);
    }
    .hero-grid {
      max-width: 1240px;
      margin: 0 auto;
      display: grid;
      grid-template-columns: minmax(0, 1.08fr) minmax(320px, 0.92fr);
      gap: clamp(28px, 5vw, 72px);
      align-items: end;
    }
    .eyebrow {
      margin: 0 0 18px;
      font-family: "Inter", "Segoe UI", sans-serif;
      font-size: 0.78rem;
      letter-spacing: 0.16em;
      text-transform: uppercase;
      color: var(--green);
      font-weight: 800;
    }
    .hero h1 {
      margin: 0;
      font-size: clamp(2.2rem, 6.5vw, 5.8rem);
      line-height: 1.05;
      letter-spacing: 0;
      max-width: 820px;
    }
    .hero-subtitle {
      max-width: 760px;
      margin: 24px 0 0;
      color: var(--deep);
      font-size: clamp(1.04rem, 1.7vw, 1.24rem);
    }
    .hero-meta {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 26px;
      font-family: "Inter", "Segoe UI", sans-serif;
      color: var(--muted);
      font-size: 0.86rem;
    }
    .hero-meta span {
      border-top: 1px solid var(--line);
      padding-top: 8px;
      min-width: 118px;
    }
    .portrait-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
      align-items: end;
    }
    .portrait-card {
      margin: 0;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      min-width: 0;
      box-shadow: 0 18px 44px rgba(21, 91, 100, 0.08);
    }
    .portrait-card img {
      width: 100%;
      aspect-ratio: 4 / 5;
      object-fit: cover;
      display: block;
      filter: saturate(0.82) contrast(1.02);
    }
    .portrait-card figcaption {
      padding: 10px 10px 12px;
      font-family: "Inter", "Segoe UI", sans-serif;
      font-size: 0.74rem;
      line-height: 1.35;
      color: var(--muted);
    }
    .portrait-card strong {
      display: block;
      color: var(--ink);
      font-size: 0.82rem;
      margin-bottom: 2px;
    }
    .layout {
      max-width: 1240px;
      margin: 0 auto;
      display: grid;
      grid-template-columns: 260px minmax(0, 1fr);
      gap: clamp(24px, 4vw, 56px);
      padding: 32px clamp(20px, 5vw, 72px) 72px;
    }
    .toc {
      position: sticky;
      top: 18px;
      align-self: start;
      max-height: calc(100vh - 36px);
      overflow: auto;
      padding: 18px 0 18px 18px;
      border-left: 1px solid var(--line);
      font-family: "Inter", "Segoe UI", sans-serif;
    }
    .toc-title {
      margin: 0 0 14px;
      color: var(--green);
      font-size: 0.78rem;
      font-weight: 800;
      letter-spacing: 0.14em;
      text-transform: uppercase;
    }
    .toc a, .toc-empty {
      display: block;
      margin: 0 0 10px;
      color: var(--muted);
      text-decoration: none;
      font-size: 0.86rem;
      line-height: 1.35;
    }
    .toc a:hover, .toc a.active { color: var(--deep); }
    .toc a.toc-major {
      margin-top: 16px;
      color: var(--ink);
      font-weight: 800;
    }
    .toc a.toc-minor { padding-left: 10px; }
    article {
      min-width: 0;
      font-size: clamp(1rem, 1.15vw, 1.1rem);
    }
    .lede {
      margin: 0 0 28px;
      padding: 22px 24px;
      background: rgba(255, 255, 255, 0.64);
      border-left: 4px solid var(--green);
      border-radius: 0 8px 8px 0;
      color: var(--deep);
      font-size: clamp(1.08rem, 1.5vw, 1.25rem);
    }
    .chapter {
      padding: clamp(24px, 4vw, 42px) 0;
      border-top: 1px solid var(--line);
    }
    .chapter:first-of-type { border-top: 0; }
    h2, h3, h4, h5 {
      margin: 0 0 20px;
      line-height: 1.22;
      letter-spacing: 0;
      color: var(--ink);
    }
    h2 { font-size: clamp(1.62rem, 3vw, 2.8rem); }
    h3 { font-size: clamp(1.3rem, 2.1vw, 2rem); }
    h4, h5 { font-size: clamp(1.12rem, 1.6vw, 1.42rem); }
    h2 span, h3 span, h4 span, h5 span {
      background-image: linear-gradient(transparent 68%, rgba(50, 133, 106, 0.2) 68%);
    }
    .chapter p { margin: 0 0 16px; }
    .kicker {
      display: inline-block;
      margin: 4px 0 12px;
      font-family: "Inter", "Segoe UI", sans-serif;
      color: var(--green);
      font-size: 0.82rem;
      font-weight: 800;
      letter-spacing: 0.08em;
    }
    blockquote {
      margin: 24px 0;
      padding: 18px 24px 18px 28px;
      border-left: 4px solid var(--deep);
      background: rgba(255, 255, 255, 0.56);
      color: var(--deep);
      border-radius: 0 8px 8px 0;
      font-size: clamp(1.04rem, 1.45vw, 1.22rem);
      line-height: 1.7;
    }
    hr {
      border: 0;
      border-top: 1px solid var(--line);
      margin: 28px 0;
    }
    .formula {
      margin: 22px 0;
      padding: 18px 22px;
      text-align: center;
      font-family: "Cambria Math", "Times New Roman", serif;
      font-size: clamp(1.24rem, 2.5vw, 2.05rem);
      color: var(--deep);
      border-top: 1px solid var(--line);
      border-bottom: 1px solid var(--line);
      background: rgba(255, 255, 255, 0.5);
    }
    .table-wrap {
      overflow-x: auto;
      margin: 22px 0 26px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: rgba(255, 255, 255, 0.58);
    }
    table {
      width: 100%;
      border-collapse: collapse;
      min-width: 620px;
      font-size: 0.96rem;
    }
    th, td {
      padding: 12px 14px;
      border-bottom: 1px solid var(--line);
      vertical-align: top;
      text-align: left;
    }
    th {
      font-family: "Inter", "Segoe UI", sans-serif;
      color: var(--deep);
      background: rgba(50, 133, 106, 0.12);
      font-weight: 800;
    }
    tr:last-child td { border-bottom: 0; }
    ol, ul { padding-left: 1.35em; margin: 14px 0 20px; }
    li { margin: 6px 0; }
    strong { color: var(--ink); font-weight: 800; }
    em { color: var(--deep); }
    .source-link {
      font-family: "Inter", "Segoe UI", sans-serif;
      font-size: 0.86em;
      white-space: nowrap;
    }
    .closing-band {
      margin-top: 24px;
      padding: 30px clamp(20px, 5vw, 72px) 48px;
      border-top: 1px solid var(--line);
      background: rgba(255, 255, 255, 0.38);
    }
    .closing-inner {
      max-width: 1240px;
      margin: 0 auto;
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(280px, 0.8fr);
      gap: 36px;
    }
    .credits h2, .references h2 {
      margin: 0 0 16px;
      font-size: 1rem;
      font-family: "Inter", "Segoe UI", sans-serif;
      letter-spacing: 0.12em;
      text-transform: uppercase;
      color: var(--green);
    }
    .credits p, .references li {
      color: var(--muted);
      font-size: 0.92rem;
      line-height: 1.7;
    }
    .references ol { margin: 0; padding-left: 1.2em; }
    @media (max-width: 960px) {
      .hero-grid, .layout, .closing-inner { grid-template-columns: 1fr; }
      .portrait-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .toc {
        position: relative;
        top: auto;
        max-height: none;
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 4px 14px;
        padding: 16px 0;
        border-left: 0;
        border-top: 1px solid var(--line);
        border-bottom: 1px solid var(--line);
      }
      .toc-title { grid-column: 1 / -1; }
      .toc a { margin-bottom: 6px; }
    }
    @media (max-width: 560px) {
      .hero { padding-top: 28px; }
      .portrait-grid { grid-template-columns: 1fr 1fr; gap: 10px; }
      .layout { padding-top: 22px; }
      .toc { grid-template-columns: 1fr; }
      .lede { padding: 18px; }
      blockquote { padding: 16px 18px; }
      table { min-width: 560px; }
    }
    @media print {
      body { background: white; }
      .toc { display: none; }
      .layout, .hero-grid, .closing-inner { display: block; max-width: none; }
      .hero, .layout, .closing-band { padding-left: 0; padding-right: 0; }
      a { color: inherit; }
    }
  </style>
</head>
<body>
  <div class="page-shell" data-composition-id="$compositionId">
    <header class="hero">
      <div class="hero-grid">
        <div>
          <p class="eyebrow">Philosophy Notes / Blue-Green Edition</p>
          <h1>$([System.Net.WebUtility]::HtmlEncode($title))</h1>
          <p class="hero-subtitle">由 Word 文档转化为适合阅读、展示与归档的哲学长文页面。背景采用淡蓝与淡绿，保留清晰目录、章节节奏和文内参考。</p>
          <div class="hero-meta" aria-label="文档信息">
            <span>来源：$([System.Net.WebUtility]::HtmlEncode($file.Name))</span>
            <span>生成：$generatedDate</span>
            <span>版式：淡蓝 + 淡绿</span>
          </div>
        </div>
        <div class="portrait-grid" aria-label="经典哲学家肖像">
          <figure class="portrait-card">
            <img src="https://commons.wikimedia.org/wiki/Special:FilePath/Bust%20of%20the%20Philosopher%20Plato.jpg?width=900" alt="柏拉图 bust portrait">
            <figcaption><strong>柏拉图</strong>理念与爱欲</figcaption>
          </figure>
          <figure class="portrait-card">
            <img src="https://commons.wikimedia.org/wiki/Special:FilePath/Bust%20of%20Aristotle.jpg?width=900" alt="亚里士多德 bust portrait">
            <figcaption><strong>亚里士多德</strong>实体与因果</figcaption>
          </figure>
          <figure class="portrait-card">
            <img src="https://commons.wikimedia.org/wiki/Special:FilePath/Frans%20Hals%2C%20Portrait%20of%20Ren%C3%A9%20Descartes.jpg?width=900" alt="笛卡尔 portrait by Frans Hals">
            <figcaption><strong>笛卡尔</strong>主体与方法</figcaption>
          </figure>
          <figure class="portrait-card">
            <img src="https://commons.wikimedia.org/wiki/Special:FilePath/Immanuel%20Kant%20%28portrait%29.jpg?width=900" alt="康德 portrait">
            <figcaption><strong>康德</strong>先验与界限</figcaption>
          </figure>
        </div>
      </div>
    </header>

    <div class="layout">
      <nav class="toc" aria-label="章节目录">
        <p class="toc-title">目录</p>
        $tocHtml
      </nav>
      <article>
        $contentHtml
      </article>
    </div>

    <footer class="closing-band">
      <div class="closing-inner">
        <section class="credits">
          <h2>图像来源</h2>
          <p>哲学家图像来自 Wikimedia Commons 的公开文件：柏拉图、亚里士多德、笛卡尔与康德肖像。页面使用远程图片链接，便于保持 HTML 文件轻量。</p>
        </section>
        <section class="references">
          <h2>文内参考</h2>
          <ol>
            $refsBlock
          </ol>
        </section>
      </div>
    </footer>
  </div>
  <script>
    window.__timelines = window.__timelines || {};
    window.__timelines["$compositionId"] = { note: "Longform HTML prepared with a HyperFrames-style composition root." };
    const links = Array.from(document.querySelectorAll('.toc a'));
    const headings = links.map(link => document.querySelector(link.getAttribute('href'))).filter(Boolean);
    const observer = new IntersectionObserver(entries => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          links.forEach(link => link.classList.toggle('active', link.getAttribute('href') === '#' + entry.target.id));
        }
      });
    }, { rootMargin: '-20% 0px -70% 0px', threshold: 0.01 });
    headings.forEach(heading => observer.observe(heading));
  </script>
</body>
</html>
"@

  [System.IO.File]::WriteAllText($outputPath, $html, [System.Text.UTF8Encoding]::new($false))
  return [pscustomobject]@{
    Source = $file.Name
    Output = [System.IO.Path]::GetFileName($outputPath)
    Paragraphs = $paragraphs.Count
    Headings = $toc.Count
    References = $references.Count
  }
}

$targets = Get-ChildItem -Path . -Filter *.docx |
  Where-Object { $_.Name -notlike '~$*' -and $_.Name -ne 'physics-unites.docx' }

$results = foreach ($target in $targets) {
  Convert-DocxToHtml $target
}

$results | Format-Table -AutoSize
