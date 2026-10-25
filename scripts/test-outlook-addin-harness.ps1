param(
    [string] $BrowserPath = $env:MAILY_TEST_BROWSER,
    [string] $CaseName = '',
    [string] $OutputDirectory = '',
    [ValidateRange(320, 1920)][int] $ViewportWidth = 390
)

$ErrorActionPreference = "Stop"

Set-Location (Join-Path $PSScriptRoot "..")

$onWindows = [Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT
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

$outputDir = if ($OutputDirectory) { [IO.Path]::GetFullPath($OutputDirectory) } else { Join-Path (Get-Location) ("output/playwright/addin-harness/" + [Guid]::NewGuid().ToString("N")) }
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
& (Join-Path $PSScriptRoot "New-MailyPaneHarness.ps1") -OutputDirectory (Join-Path $outputDir "pages")

function Invoke-HeadlessBrowser {
    param(
        [string[]] $Arguments,
        [string] $StdOutPath,
        [string] $StdErrPath
    )

    if ((Test-Path $StdOutPath) -or (Test-Path $StdErrPath)) { throw "Refusing to overwrite existing browser logs." }

    $Arguments = @("--user-data-dir=$(Join-Path $outputDir ('browser-profile-' + [Guid]::NewGuid().ToString('N')))") + $Arguments
    $start = @{
        FilePath = $browserPath
        ArgumentList = ($Arguments | ForEach-Object { '"' + $_.Replace('"', '\"') + '"' })
        Wait = $true
        PassThru = $true
        RedirectStandardOutput = $StdOutPath
        RedirectStandardError = $StdErrPath
    }
    if ($onWindows) { $start.WindowStyle = 'Hidden' }
    $process = Start-Process @start

    return $process.ExitCode
}

function Get-TitleFromHtml {
    param([string] $Html)

    $match = [regex]::Match($Html, "<title>(?<value>.*?)</title>", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $match.Success) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlDecode($match.Groups["value"].Value)
}

function Get-PreContent {
    param(
        [string] $Html,
        [string] $ElementId
    )

    $pattern = "<pre[^>]*id=['""]$([regex]::Escape($ElementId))['""][^>]*>(?<value>.*?)</pre>"
    $match = [regex]::Match($Html, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $match.Success) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlDecode($match.Groups["value"].Value)
}

$cases = @(
    @{
        Name = "taskpane-failed-regenerate"
        HtmlPath = Join-Path $outputDir "pages/taskpane.html"
        Query = "?scenario=failed-regenerate"
        DomPath = Join-Path $outputDir "taskpane-failed-regenerate-dom.html"
        ErrPath = Join-Path $outputDir "taskpane-failed-regenerate-dom.stderr.txt"
        ScreenshotPath = Join-Path $outputDir "taskpane-failed-regenerate.png"
    },
    @{
        Name = "taskpane-english"
        HtmlPath = Join-Path $outputDir "pages/taskpane.html"
        Query = "?scenario=english"
        DomPath = Join-Path $outputDir "taskpane-dom.html"
        ErrPath = Join-Path $outputDir "taskpane-dom.stderr.txt"
        ScreenshotPath = Join-Path $outputDir "taskpane.png"
    },
    @{
        Name = "taskpane-hebrew"
        HtmlPath = Join-Path $outputDir "pages/taskpane.html"
        Query = "?scenario=hebrew"
        DomPath = Join-Path $outputDir "taskpane-hebrew-dom.html"
        ErrPath = Join-Path $outputDir "taskpane-hebrew-dom.stderr.txt"
        ScreenshotPath = Join-Path $outputDir "taskpane-hebrew.png"
    },
    @{
        Name = "taskpane-auto-hebrew"
        HtmlPath = Join-Path $outputDir "pages/taskpane.html"
        Query = "?scenario=auto-hebrew"
        DomPath = Join-Path $outputDir "taskpane-auto-hebrew-dom.html"
        ErrPath = Join-Path $outputDir "taskpane-auto-hebrew-dom.stderr.txt"
        ScreenshotPath = Join-Path $outputDir "taskpane-auto-hebrew.png"
    },
    @{
        Name = "settings"
        HtmlPath = Join-Path $outputDir "pages/settings.html"
        Query = ""
        DomPath = Join-Path $outputDir "settings-dom.html"
        ErrPath = Join-Path $outputDir "settings-dom.stderr.txt"
        ScreenshotPath = Join-Path $outputDir "settings.png"
    },
    @{
        Name = "summary"
        HtmlPath = Join-Path $outputDir "pages/summary.html"
        Query = ""
        DomPath = Join-Path $outputDir "summary-dom.html"
        ErrPath = Join-Path $outputDir "summary-dom.stderr.txt"
        ScreenshotPath = Join-Path $outputDir "summary.png"
    },
    @{
        Name = "summary-hebrew"
        HtmlPath = Join-Path $outputDir "pages/summary.html"
        Query = "?scenario=hebrew"
        DomPath = Join-Path $outputDir "summary-hebrew-dom.html"
        ErrPath = Join-Path $outputDir "summary-hebrew-dom.stderr.txt"
        ScreenshotPath = Join-Path $outputDir "summary-hebrew.png"
    },
    @{
        Name = "summary-read"
        HtmlPath = Join-Path $outputDir "pages/summary-read.html"
        Query = ""
        DomPath = Join-Path $outputDir "summary-read-dom.html"
        ErrPath = Join-Path $outputDir "summary-read-dom.stderr.txt"
        ScreenshotPath = Join-Path $outputDir "summary-read.png"
    },
    @{
        Name = "search"
        HtmlPath = Join-Path $outputDir "pages/search.html"
        Query = ""
        DomPath = Join-Path $outputDir "search-dom.html"
        ErrPath = Join-Path $outputDir "search-dom.stderr.txt"
        ScreenshotPath = Join-Path $outputDir "search.png"
    },
    @{
        Name = "daily-briefing"
        HtmlPath = Join-Path $outputDir "pages/daily-briefing.html"
        Query = ""
        DomPath = Join-Path $outputDir "daily-briefing-dom.html"
        ErrPath = Join-Path $outputDir "daily-briefing-dom.stderr.txt"
        ScreenshotPath = Join-Path $outputDir "daily-briefing.png"
    }
)

if ($CaseName) {
    $cases = @($cases | Where-Object { $_.Name -eq $CaseName })
    if ($cases.Count -eq 0) { throw "Unknown harness case: $CaseName" }
}

$results = @()

foreach ($case in $cases) {
    if (-not (Test-Path $case.HtmlPath)) {
        throw "Harness page not found: $($case.HtmlPath)"
    }

    $querySuffix = if ($null -ne $case.Query) { [string]$case.Query } else { "" }
    $fileUrl = ([System.Uri]::new($case.HtmlPath).AbsoluteUri) + $querySuffix
    # Chromium's headless top-level window has a minimum width of 500px.
    # An actual child viewport exercises Outlook-sized CSS/media queries.
    $framePath = Join-Path $outputDir ($case.Name + '-viewport.html')
    $frameHtml = @"
<!doctype html><html><head><meta charset="utf-8"><title>PENDING</title>
<style>html,body{margin:0;background:#f5f5f7}iframe{display:block;border:0;width:${ViewportWidth}px;height:100vh}</style></head>
<body><iframe id="pane" title="Synthetic Outlook pane" src="$fileUrl"></iframe><pre id="result" hidden></pre>
<script>window.addEventListener('message',event=>{if(event.source!==document.getElementById('pane').contentWindow||!event.data?.mailyHarness)return;document.title=event.data.title;document.getElementById('result').textContent=JSON.stringify(event.data.report);});</script></body></html>
"@
    [IO.File]::WriteAllText($framePath, $frameHtml, [Text.UTF8Encoding]::new($false))
    $fileUrl = ([System.Uri]::new($framePath)).AbsoluteUri

    $domArgs = @(
        "--window-size=$ViewportWidth,1000",
        "--headless=new",
        "--disable-gpu",
        "--allow-file-access-from-files",
        "--virtual-time-budget=5000",
        "--dump-dom",
        $fileUrl
    )

    $domExitCode = Invoke-HeadlessBrowser -Arguments $domArgs -StdOutPath $case.DomPath -StdErrPath $case.ErrPath
    $html = Get-Content -LiteralPath $case.DomPath -Raw
    $title = Get-TitleFromHtml -Html $html
    $resultJson = Get-PreContent -Html $html -ElementId "result"
    $parsed = $null
    if ($resultJson) {
        try {
            $parsed = $resultJson | ConvertFrom-Json
        } catch {
            $parsed = [pscustomobject]@{
                parseError = $_.Exception.Message
                rawResult = $resultJson
            }
        }
    }

    $shotArgs = @(
        "--headless=new",
        "--disable-gpu",
        "--allow-file-access-from-files",
        "--virtual-time-budget=5000",
        "--window-size=$ViewportWidth,1000",
        "--screenshot=$($case.ScreenshotPath)",
        $fileUrl
    )

    $shotErrPath = [System.IO.Path]::ChangeExtension($case.ErrPath, ".screenshot.stderr.txt")
    $shotOutPath = [System.IO.Path]::ChangeExtension($case.DomPath, ".screenshot.stdout.txt")
    $screenshotExitCode = Invoke-HeadlessBrowser -Arguments $shotArgs -StdOutPath $shotOutPath -StdErrPath $shotErrPath

    $errors = @()
    if ($domExitCode -ne 0) {
        $errors += "DOM capture exited with code $domExitCode."
    }
    if ($screenshotExitCode -ne 0) {
        $errors += "Screenshot capture exited with code $screenshotExitCode."
    }
    if ($title -ne "PASS") {
        $errors += "Expected PASS title but saw '$title'."
    }
    if ($parsed -and $parsed.parseError) {
        $errors += "Could not parse harness JSON result: $($parsed.parseError)"
    }
    if (-not $parsed) {
        $errors += "Harness did not produce a JSON result."
    }
    if ($parsed -and $parsed.errors -and $parsed.errors.Count -gt 0) {
        $errors += "Harness reported startup errors: $($parsed.errors -join '; ')"
    }

    $results += [pscustomobject]@{
        name = $case.Name
        title = $title
        domPath = $case.DomPath
        screenshotPath = $case.ScreenshotPath
        result = $parsed
        errors = $errors
    }
}

$summaryPath = Join-Path $outputDir $(if ($CaseName) { "$CaseName-summary.json" } else { 'summary.json' })
$summary = [pscustomobject]@{
    browserPath = $browserPath
    generatedAtUtc = [DateTime]::UtcNow.ToString("o")
    results = $results
}

$summary | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
$summary | ConvertTo-Json -Depth 10

if ($results | Where-Object { $_.errors.Count -gt 0 }) {
    exit 1
}
