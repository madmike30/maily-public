param([string]$OutputDirectory = '')
$ErrorActionPreference = "Stop"

# Regenerate from the actual product markup/assets on every run. Only Office and
# backend responses are mocked; no ignored output file is a test dependency.
$repoRoot = Split-Path -Parent $PSScriptRoot
$sourceRoot = Join-Path $repoRoot "services/Maily.OutlookAddin.Web/wwwroot/addin"
$fixtureRoot = Join-Path $repoRoot "tests/harness/panes"
$outputRoot = if ($OutputDirectory) { [IO.Path]::GetFullPath($OutputDirectory) } else { Join-Path $repoRoot ("output/playwright/addin-harness/" + [Guid]::NewGuid().ToString("N") + "/pages") }
if ((Test-Path -LiteralPath $outputRoot) -and (Get-ChildItem -LiteralPath $outputRoot -Force | Select-Object -First 1)) { throw "Harness output must be empty; previous artifacts are retained." }
New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null

foreach ($pane in @("taskpane", "settings", "summary", "summary-read", "search", "daily-briefing")) {
    $sourcePane = $pane
    if ($pane -eq 'summary-read') {
        # The real read entry redirects to summary.html?surface=read. Exercise
        # that actual target with the read mock, avoiding a redirect to live AI.
        $entry = Get-Content -LiteralPath (Join-Path $sourceRoot 'summary-read.html') -Raw
        if ($entry -notmatch 'summary\.html\?surface=read') { throw 'Read summary target changed; update its harness.' }
        $sourcePane = 'summary'
    }
    $html = Get-Content -LiteralPath (Join-Path $sourceRoot "$sourcePane.html") -Raw
    $mock = Get-Content -LiteralPath (Join-Path $fixtureRoot "$pane.mock.js") -Raw
    $assert = Get-Content -LiteralPath (Join-Path $fixtureRoot "$pane.assert.js") -Raw
    $html = [regex]::Replace($html, '<title>.*?</title>', '<title>PENDING</title>')
    $officeScript = '<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>'
    if (-not $html.Contains($officeScript)) { throw "Office script marker missing in $pane.html" }
    $html = $html.Replace($officeScript, "<script>`n$mock`n</script>")
    # Resolve current JS/CSS/assets without copying a stale product snapshot.
    $sourceUri = ([Uri]::new($sourceRoot + [IO.Path]::DirectorySeparatorChar)).AbsoluteUri
    $html = $html.Replace('<head>', "<head>`n  <meta http-equiv=`"Content-Security-Policy`" content=`"default-src 'none'; script-src file: 'unsafe-inline'; style-src file: 'unsafe-inline'; img-src file: data:; connect-src 'none'; form-action 'none'; frame-src 'none'`">`n  <base href=`"$sourceUri`">")
    $layoutCheck = @'
setTimeout(() => {
  const result = document.getElementById('result');
  const horizontalOverflow = document.documentElement.scrollWidth > window.innerWidth + 1;
  const hiddenLeaked = Array.from(document.querySelectorAll('[hidden]')).some(el => el.getClientRects().length > 0);
  let report;
  try { report = JSON.parse(result.textContent); } catch { return; }
  report.layout = { horizontalOverflow, hiddenLeaked, width: window.innerWidth };
  if (horizontalOverflow || hiddenLeaked) document.title = 'FAIL';
  result.textContent = JSON.stringify(report, null, 2);
  if (window.parent !== window) window.parent.postMessage({ mailyHarness: true, title: document.title, report }, '*');
}, 4700);
'@
    $html = $html.Replace('</body>', "<pre id=`"result`" hidden></pre>`n<script>`n$assert`n$layoutCheck`n</script>`n</body>")
    [IO.File]::WriteAllText((Join-Path $outputRoot "$pane.html"), $html, [Text.UTF8Encoding]::new($false))
}
