param([Parameter(Mandatory)][string]$DotnetPath)
$ErrorActionPreference = 'Stop'
$project = Join-Path $PSScriptRoot '../../tools/Maily.Generation.TestRunner/Maily.Generation.TestRunner.csproj'
$artifactRoot = Join-Path $PSScriptRoot ('../../output/runner-diagnostics/' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null
& $DotnetPath build $project --nologo *> (Join-Path $artifactRoot 'build.log')
if ($LASTEXITCODE) { throw "Runner build failed; see $artifactRoot/build.log" }
function Invoke-Runner([string[]]$Arguments) {
    $lines = & $DotnetPath run --no-build --no-restore --project $project -- @Arguments 2>&1
    $exitCode = $LASTEXITCODE
    $lines | Set-Content -LiteralPath (Join-Path $artifactRoot ([Guid]::NewGuid().ToString('N') + '.log')) -Encoding utf8NoBOM
    return @{ Code = $exitCode; Text = ($lines -join "`n") }
}
$result = Invoke-Runner @('--filter', 'keeps ordinary reply text')
if ($result.Code -ne 0 -or $result.Text -notmatch 'RESULT: 1 passed, 0 failed') { throw 'Filtered run must report exactly one passing test.' }
$result = Invoke-Runner @('--filter', 'synthetic-no-matching-case-7184')
if ($result.Code -eq 0 -or $result.Text -notmatch 'No tests matched') { throw 'Empty filter must fail instead of reporting green.' }
$result = Invoke-Runner @('--diagnostics-self-test')
if ($result.Code -eq 0 -or $result.Text -notmatch 'Synthetic inner failure' -or $result.Text -notmatch ' at ' -or $result.Text -notmatch 'RESULT: 0 passed, 1 failed') { throw 'Synthetic failure must include inner exception, stack trace and failed count.' }
Write-Output 'PASS runner diagnostics: filtering, empty selection, full synthetic exception and counts'
