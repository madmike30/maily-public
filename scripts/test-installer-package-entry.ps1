param(
  [string]$BuilderPath = (Join-Path $PSScriptRoot 'build-maily-outlook-installer.ps1'),
  [ValidateSet('Local','Hosted')][string]$PackageMode = 'Local',
  [ValidateSet('InstalledApp','Bundled')][string]$AiRuntimeMode = 'InstalledApp'
)
$ErrorActionPreference = 'Stop'
$run = Join-Path $PSScriptRoot ('../output/package-entry-' + [guid]::NewGuid().ToString('N'))
$fixture = Join-Path $run 'fixture'
foreach ($directory in @('scripts', 'addin', 'docs', 'smoke_data/prompts')) {
  New-Item -ItemType Directory -Path (Join-Path $fixture $directory) -Force | Out-Null
}
function Write-Fixture([string]$Relative, [string]$Content) {
  [IO.File]::WriteAllText((Join-Path $fixture $Relative), $Content, [Text.UTF8Encoding]::new($false))
}
Copy-Item -LiteralPath $BuilderPath -Destination (Join-Path $fixture 'scripts/build.ps1')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'MailyPackageRuntime.ps1') -Destination (Join-Path $fixture 'scripts/MailyPackageRuntime.ps1')
Write-Fixture 'addin/manifest.xml' '<OfficeApp><Id>683ea353-6d71-4315-afd4-147e968605a9</Id><Version>1.3.1.0</Version></OfficeApp>'
foreach ($document in @('DISCLAIMER.md', 'RELEASE_NOTES_1.3.1.md', 'INSTALLER.md')) { Write-Fixture "docs/$document" "Synthetic document $document" }
foreach ($license in @('LICENSE', 'LICENSE.md')) { Write-Fixture $license "Synthetic license $license" }
Write-Fixture 'smoke_data/prompts/prompt-system.md' 'Synthetic prompt.'
foreach ($script in @('outlook-bridge.ps1', 'Install-MailyOutlookAddinPackage.ps1', 'Uninstall-MailyOutlookAddinPackage.ps1', 'OutlookShutdown.ps1')) { Write-Fixture "scripts/$script" "throw 'Synthetic payload must never execute.'" }
foreach ($script in @('connect-outlook.cjs', 'outlook-credential-preflight.cjs')) { Write-Fixture "scripts/$script" "throw new Error('Synthetic payload must never execute');" }
Write-Fixture 'scripts/Bundle-MailyOutlookConnector.ps1' @'
param([string]$Destination, [string]$Runtime)
New-Item -ItemType Directory -Path $Destination -Force | Out-Null
[IO.File]::WriteAllText((Join-Path $Destination 'node.exe'), 'Synthetic runtime: never run.')
'@
Write-Fixture 'scripts/codex-runtime.lock.json' '{"version":"0.151.0"}'
Write-Fixture 'scripts/New-MailyHostedManifest.ps1' @'
param([string]$BaseUrl, [string]$SourcePath, [string]$DestinationPath)
Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath
'@
Write-Fixture 'scripts/Bundle-MailyCodexRuntime.ps1' @'
param([string]$Destination, [string]$Runtime)
New-Item -ItemType Directory -Path $Destination -Force | Out-Null
foreach ($name in @('codex.exe', 'LICENSE', 'NOTICE', 'codex-runtime.lock.json')) {
  [IO.File]::WriteAllText((Join-Path $Destination $name), 'Synthetic runtime fixture.')
}
'@
function dotnet {
  $publish = [array]$args
  if ($publish[0] -ne 'publish') { throw 'Only synthetic publish is allowed.' }
  $destination = $publish[[array]::IndexOf($publish, '-o') + 1]
  $project = [IO.Path]::GetFileNameWithoutExtension($publish[1])
  New-Item -ItemType Directory -Path $destination -Force | Out-Null
  [IO.File]::WriteAllText((Join-Path $destination ($project + '.exe')), 'Synthetic executable: never run.')
  $global:LASTEXITCODE = 0
}
& (Join-Path $fixture 'scripts/build.ps1') -OutputRoot (Join-Path $run 'build') -PackageMode $PackageMode -AiRuntimeMode $AiRuntimeMode -HostedBaseUrl 'https://maily.example.invalid' *> (Join-Path $run 'build.log')
$packageName = if ($PackageMode -eq 'Hosted') { 'MailyOutlookAddinHosted-1.3.1.0' } else { 'MailyOutlookAddin-1.3.1.0' }
$package = Join-Path $run "build/$packageName"
$publicSummary = Get-Content (Join-Path $package 'package-summary.json') -Raw | ConvertFrom-Json
foreach ($value in @($publicSummary.packageRoot, $publicSummary.setupExecutable, $publicSummary.zipPath)) {
  if ([IO.Path]::IsPathRooted($value)) { throw 'Package metadata must not disclose builder filesystem paths.' }
}
$launchers = @(Get-ChildItem -LiteralPath $package -File | Where-Object Extension -in @('.exe', '.bat', '.cmd', '.ps1'))
if ($launchers.Count -ne 1 -or $launchers[0].Name -ne 'MailyOutlookAddinSetup.exe') { throw 'Recipients must have exactly one launch entry: MailyOutlookAddinSetup.exe.' }
$start = [IO.File]::ReadAllText((Join-Path $package 'START HERE.txt'))
if ($start -notmatch 'MailyOutlookAddinSetup.exe' -or $start -match '\.bat|manifest|\.NET|Codex CLI') { throw 'Start instructions must name one installer without developer recovery steps.' }
foreach ($name in @('DISCLAIMER.md', 'RELEASE_NOTES_1.3.1.md', 'INSTALLER.md', 'LICENSE', 'LICENSE.md')) {
  if (-not (Test-Path -LiteralPath (Join-Path $package "docs/$name"))) { throw "Missing disclosure: $name" }
  if (Test-Path -LiteralPath (Join-Path $package $name)) { throw "Reference document clutters package root: $name" }
}
$supportFiles = @('setup/install.ps1', 'setup/uninstall.ps1', 'setup/OutlookShutdown.ps1', 'payload/addin/manifest.xml')
$supportFiles += @('payload/runtime/atk/node.exe', 'payload/scripts/connect-outlook.cjs', 'payload/scripts/outlook-credential-preflight.cjs')
if ($PackageMode -eq 'Local' -and $AiRuntimeMode -eq 'Bundled') { $supportFiles += @('payload/runtime/codex/codex.exe', 'third-party/codex/NOTICE') }
if (([IO.File]::ReadAllText((Join-Path $package 'ai-runtime-mode.txt'))).Trim() -ne $AiRuntimeMode) { throw 'Runtime mode must be explicit in the package.' }
if ($publicSummary.aiRuntimeMode -ne $AiRuntimeMode) { throw 'Runtime mode metadata mismatch.' }
if ($AiRuntimeMode -eq 'InstalledApp' -or $PackageMode -eq 'Hosted') {
  if (Test-Path (Join-Path $package 'payload/runtime/codex')) { throw 'Existing-app package must not carry another Codex runtime.' }
  if ($publicSummary.bundledCodexVersion) { throw 'Metadata must not claim a bundled runtime.' }
}
foreach ($relative in $supportFiles) {
  if (-not (Test-Path -LiteralPath (Join-Path $package $relative))) { throw "Required support file missing: $relative" }
}
$inventory = Get-Content -LiteralPath (Join-Path $package 'SHA256SUMS.json') -Raw | ConvertFrom-Json
foreach ($entry in $inventory) {
  if ((Get-FileHash -LiteralPath (Join-Path $package $entry.path) -Algorithm SHA256).Hash -ne $entry.sha256) { throw "Package checksum mismatch: $($entry.path)" }
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead((Join-Path $run "build/$packageName.zip"))
try {
  $names = @($zip.Entries.FullName | ForEach-Object { $_.Replace('\', '/') })
  if ('START HERE.txt' -notin $names -or 'docs/DISCLAIMER.md' -notin $names -or @($names | Where-Object { $_ -match '^[^/]+\.(bat|cmd|ps1)$' }).Count -gt 0) { throw 'Distributed archive does not preserve the simplified entry layout.' }
} finally { $zip.Dispose() }
"PASS package entry ($PackageMode): one installer, concise instructions, retained disclosures/support files, checksum inventory and ZIP layout. Synthetic publishing only. Artifacts: $run"
