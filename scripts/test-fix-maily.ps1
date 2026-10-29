param([string] $DotnetPath)
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($DotnetPath)) {
    $localSdk = Join-Path $PSScriptRoot '../output/fix-tools/dotnet/dotnet.exe'
    if (Test-Path -LiteralPath $localSdk -PathType Leaf) { $DotnetPath = $localSdk }
    else { $DotnetPath = (Get-Command dotnet -CommandType Application -ErrorAction Stop).Source }
}
$DotnetPath = (Get-Item -LiteralPath $DotnetPath -ErrorAction Stop).FullName
$env:DOTNET_ROOT = Split-Path $DotnetPath
Set-Location (Join-Path $PSScriptRoot '..')
$runRoot = Join-Path (Get-Location) ('output/fix-20261023/suite-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $runRoot | Out-Null
$results = [Collections.Generic.List[object]]::new()
function Invoke-Check([string] $CheckName, [scriptblock] $Command) {
    $log = Join-Path $runRoot ($CheckName + '.log')
    $global:LASTEXITCODE = 0
    try {
        & $Command *> $log
        $code = $LASTEXITCODE
    } catch {
        $_ | Out-File -LiteralPath $log -Append -Encoding utf8
        $code = 1
    }
    $results.Add([pscustomobject]@{ name=$CheckName; passed=($code -eq 0); exitCode=$code; log=$log })
    Write-Output "$CheckName : exit $code"
}
foreach ($project in @('services/Maily.OutlookAddin.Web','services/Maily.OutlookAddin.Supervisor','tools/Maily.OutlookAddin.Setup')) {
    $name = Split-Path $project -Leaf
    Invoke-Check "build-$name" { & $DotnetPath build "$project/$name.csproj" --nologo }
}
Invoke-Check 'generation' { & $DotnetPath run --project tools/Maily.Generation.TestRunner/Maily.Generation.TestRunner.csproj }
Invoke-Check 'runner-diagnostics' { & pwsh -NoProfile -File tests/runner-diagnostics/run.ps1 -DotnetPath $DotnetPath }
Invoke-Check 'backend-bootstrap' { & $DotnetPath run --project tests/Maily.Bootstrap.Tests/Maily.Bootstrap.Tests.csproj }
Invoke-Check 'temporary-paths' { & $DotnetPath run --project tests/Maily.TempFile.Tests/Maily.TempFile.Tests.csproj }
Invoke-Check 'settings-safety' { & node tests/settings-safety/settings.mjs }
Invoke-Check 'settings-tour' { & node tests/settings-safety/tour.mjs }
Invoke-Check 'automatic-language' { & node tests/language/automatic.mjs }
Invoke-Check 'summary-reply' { & node tests/summary-reply/handoff.mjs }
Invoke-Check 'outlook-connect' { & node tests/outlook-connect/connect.cjs }
Invoke-Check 'connector-package-size' { & node tests/outlook-connect/size.cjs }
Invoke-Check 'pane-key' { & node tests/Maily.Bootstrap.Tests/pane-key.mjs }
Invoke-Check 'segmentation' { & pwsh -NoProfile -File scripts/test-outlook-compose-segmentation.ps1 }
Invoke-Check 'panes' { & pwsh -NoProfile -File scripts/test-outlook-addin-harness.ps1 }
Invoke-Check 'manifest' { & pwsh -NoProfile -File scripts/test-hosted-manifest.ps1 }
# Issue-specific mocked tests are registered here as their worktrees merge.
Invoke-Check 'office-stale-body' { & node tests/office-safety/stale-body.mjs }
Invoke-Check 'office-undo' { & node tests/office-safety/undo.mjs }
Invoke-Check 'compose-review' { & node tests/taskpane-release/run.mjs }
Invoke-Check 'cli-quoting' { & $DotnetPath run --project tests/Maily.CliQuoting.Tests/Maily.CliQuoting.Tests.csproj }
Invoke-Check 'native-draft' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-native-draft-safety.ps1 }
Invoke-Check 'provider-privacy' { & $DotnetPath run --project tests/Maily.Privacy.Tests/Maily.Privacy.Tests.csproj }
Invoke-Check 'model-catalog' { & $DotnetPath run --project tests/Maily.ModelCatalog.Tests/Maily.ModelCatalog.Tests.csproj }
Invoke-Check 'installed-app-discovery' { & $DotnetPath run --project tests/Maily.InstalledApp.Tests/Maily.InstalledApp.Tests.csproj }
Invoke-Check 'telemetry-privacy' { & node tests/Maily.Privacy.Tests/telemetry.mjs }
Invoke-Check 'backend-privacy' { & node tests/Maily.Privacy.Tests/backend-contracts.mjs }
Invoke-Check 'utf8' { & $DotnetPath run --project tests/Maily.Utf8.Tests/Maily.Utf8.Tests.csproj }
Invoke-Check 'bridge-lifetime' { & $DotnetPath run --project tests/Maily.Bridge.Lifetime.Tests/Maily.Bridge.Lifetime.Tests.csproj -- (Join-Path $runRoot 'lifetime-artifacts') }
Invoke-Check 'maui-mailbox' { & $DotnetPath run --project tests/Maily.Maui.Mailbox.Tests/Maily.Maui.Mailbox.Tests.csproj }
Invoke-Check 'mailbox-isolation' { & pwsh -NoProfile -File scripts/test-mailbox-isolation-23.ps1 }
Invoke-Check 'bridge-arrays' { & pwsh -NoProfile -File scripts/test-bridge-array-contract-23.ps1 }
Invoke-Check 'stdin-cancellation' { & $DotnetPath run --project tests/Maily.Stdin.Tests/Maily.Stdin.Tests.csproj }
Invoke-Check 'cache-metadata' { & $DotnetPath run --project tests/Maily.Cache.Tests/Maily.Cache.Tests.csproj }
Invoke-Check 'search-coverage' { & $DotnetPath run --project tests/Maily.SearchCoverage.Tests/Maily.SearchCoverage.Tests.csproj }
Invoke-Check 'search-coverage-ui' { & node tests/Maily.SearchCoverage.Tests/search-ui.mjs }
Invoke-Check 'release-folder-mocks' { & pwsh -NoProfile -File scripts/test-outlook-release-folder-mocks.ps1 }
Invoke-Check 'search-smoke-scope' { & pwsh -NoProfile -File scripts/test-outlook-search-smoke-mocks.ps1 }
Invoke-Check 'installer-lifecycle' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-installer-lifecycle.ps1 }
Invoke-Check 'installed-desktop-runtime' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-installed-desktop-runtime.ps1 }
Invoke-Check 'hosted-uninstall' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-hosted-uninstall.ps1 }
Invoke-Check 'uninstall-completion' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-uninstall-completion.ps1 }
Invoke-Check 'installer-mailbox-activation' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-installer-mailbox-activation.ps1 }
Invoke-Check 'installer-asset-upgrade' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-installer-asset-upgrade.ps1 }
Invoke-Check 'package-retention' { & pwsh -NoProfile -File scripts/test-package-retention.ps1 }
Invoke-Check 'package-runtime-sharing' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-package-runtime-dedup.ps1 }
Invoke-Check 'package-entry-local' { & pwsh -NoProfile -File scripts/test-installer-package-entry.ps1 -PackageMode Local }
Invoke-Check 'package-entry-bundled' { & pwsh -NoProfile -File scripts/test-installer-package-entry.ps1 -PackageMode Local -AiRuntimeMode Bundled }
Invoke-Check 'package-entry-hosted' { & pwsh -NoProfile -File scripts/test-installer-package-entry.ps1 -PackageMode Hosted }
Invoke-Check 'installer-timeout' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-installer-timeout.ps1 }
Invoke-Check 'installer-lock' { & $DotnetPath run --project tests/Maily.Setup.Lock.Tests/Maily.Setup.Lock.Tests.csproj -- (Join-Path $runRoot 'setup-lock-artifacts') }
Invoke-Check 'outlook-shutdown' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-outlook-shutdown.ps1 }
Invoke-Check 'outlook-window-close' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-outlook-shutdown-windows.ps1 }
Invoke-Check 'uninstall-shutdown-guard' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-uninstall-shutdown-guard.ps1 }
Invoke-Check 'setup-connection' { & $DotnetPath run --project tests/Maily.Setup.Connection.Tests/Maily.Setup.Connection.Tests.csproj }
Invoke-Check 'setup-shutdown' { & $DotnetPath run --project tests/Maily.Setup.Shutdown.Tests/Maily.Setup.Shutdown.Tests.csproj }
Invoke-Check 'installer-flow' { & $DotnetPath run --project tests/Maily.Setup.Flow.Tests/Maily.Setup.Flow.Tests.csproj }
Invoke-Check 'installer-render' { & $DotnetPath run --project tests/Maily.Setup.Render.Tests/Maily.Setup.Render.Tests.csproj -- (Join-Path $runRoot 'installer-render') }
Invoke-Check 'fixture-safety' { & node tests/fixture-safety/scan.mjs }
Invoke-Check 'bridge-logging' { & powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-bridge-logging.ps1 }
Invoke-Check 'source-utf8' {
    foreach ($changedPath in @(& git diff main --name-only)) {
        if (-not (Test-Path -LiteralPath $changedPath -PathType Leaf)) { continue }
        $bytes = [IO.File]::ReadAllBytes((Join-Path (Get-Location) $changedPath))
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) { throw "UTF-8 BOM detected: $changedPath" }
        $null = [Text.UTF8Encoding]::new($false, $true).GetString($bytes)
    }
    Write-Output 'PASS changed source files are UTF-8 without BOM'
}
Invoke-Check 'diff-check' { & git diff --check }
$failed = @($results | Where-Object { -not $_.passed }).Count
$summary = [pscustomobject]@{ passed=$results.Count-$failed; failed=$failed; results=$results }
[IO.File]::WriteAllText((Join-Path $runRoot 'summary.json'), ($summary | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
Write-Output "RESULT: $($summary.passed) passed, $failed failed. Artifacts: $runRoot"
if ($failed) { exit 1 }
