param([string] $BrowserPath = $env:MAILY_TEST_BROWSER, [string] $OutputDirectory = '')

$ErrorActionPreference = "Stop"

Set-Location (Join-Path $PSScriptRoot "..")

$chromeCandidates = @(
    $BrowserPath,
    "C:\Program Files\Google\Chrome\Application\chrome.exe",
    "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
    "/usr/bin/google-chrome"
)

$browserPath = $chromeCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
if (-not $browserPath) {
    throw "Could not find Chromium, Chrome or Edge. Supply -BrowserPath or MAILY_TEST_BROWSER."
}

$harnessPath = Join-Path (Get-Location) "tests/harness/addin-compose-segmentation.html"
if (-not (Test-Path $harnessPath)) {
    throw "Harness page not found: $harnessPath"
}

$outputDir = if ($OutputDirectory) { [IO.Path]::GetFullPath($OutputDirectory) } else { Join-Path (Get-Location) ("output/playwright/compose-segmentation/" + [Guid]::NewGuid().ToString("N")) }
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$domPath = Join-Path $outputDir "compose-segmentation-dom.html"
$stderrPath = Join-Path $outputDir "compose-segmentation.stderr.txt"
$summaryPath = Join-Path $outputDir "summary.json"

if ((Test-Path $domPath) -or (Test-Path $stderrPath)) { throw "Refusing to overwrite existing browser logs." }

$fileUrl = [System.Uri]::new($harnessPath).AbsoluteUri
$arguments = @(
    "--user-data-dir=$(Join-Path $outputDir 'browser-profile')",
    "--headless=new",
    "--disable-gpu",
    "--allow-file-access-from-files",
    "--virtual-time-budget=5000",
    "--dump-dom",
    $fileUrl
)

$start = @{
    FilePath = $browserPath
    ArgumentList = ($arguments | ForEach-Object { '"' + $_.Replace('"', '\"') + '"' })
    Wait = $true
    PassThru = $true
    RedirectStandardOutput = $domPath
    RedirectStandardError = $stderrPath
}
if ([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT) { $start.WindowStyle = 'Hidden' }
$process = Start-Process @start

$html = Get-Content -LiteralPath $domPath -Raw
$titleMatch = [regex]::Match($html, "<title>(?<value>.*?)</title>", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
$title = if ($titleMatch.Success) { [System.Net.WebUtility]::HtmlDecode($titleMatch.Groups["value"].Value) } else { "" }
$resultMatch = [regex]::Match($html, "<pre[^>]*id=['""]result['""][^>]*>(?<value>.*?)</pre>", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
$resultJson = if ($resultMatch.Success) { [System.Net.WebUtility]::HtmlDecode($resultMatch.Groups["value"].Value) } else { "" }

$parsed = $null
if ($resultJson) {
    $parsed = $resultJson | ConvertFrom-Json
}

$summary = [pscustomobject]@{
    browserPath = $browserPath
    title = $title
    exitCode = $process.ExitCode
    parsed = $parsed
    stderrPath = $stderrPath
    domPath = $domPath
}

$summary | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
$summary | ConvertTo-Json -Depth 10

if ($process.ExitCode -ne 0 -or $title -ne "PASS" -or -not $parsed -or ($parsed.errors | Measure-Object).Count -gt 0) {
    exit 1
}
