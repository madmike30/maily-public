param(
  [ValidateSet("Debug", "Release")]
  [string]$Configuration = "Release",
  [string]$Runtime = "win-x64",
  [ValidateSet("Hosted", "Local")]
  [string]$PackageMode = "Local",
  [ValidateSet("InstalledApp", "Bundled")]
  [string]$AiRuntimeMode = "InstalledApp",
  [string]$HostedBaseUrl = "",
  [string]$OutputRoot = "",
  [switch]$NoZip
)

$ErrorActionPreference = "Stop"

function Set-MailyUtf8Content {
  [CmdletBinding()] param([Parameter(Mandatory=$true)][string]$LiteralPath, [Parameter(ValueFromPipeline=$true)][AllowEmptyString()][string]$Value)
  begin { $lines = New-Object 'System.Collections.Generic.List[string]' }
  process { $lines.Add($Value) }
  end { [IO.File]::WriteAllText($LiteralPath, (($lines -join [Environment]::NewLine) + [Environment]::NewLine), (New-Object Text.UTF8Encoding($false))) }
}


$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
  $OutputRoot = Join-Path $repoRoot ("dist/" + [Guid]::NewGuid().ToString("N"))
}
$OutputRoot = [System.IO.Path]::GetFullPath($OutputRoot)

$manifestSource = Join-Path $repoRoot "addin\manifest.xml"
if (-not (Test-Path -LiteralPath $manifestSource)) {
  throw "Manifest not found at $manifestSource."
}

$manifestXml = [xml](Get-Content -LiteralPath $manifestSource -Raw)
$manifestVersion = [string]$manifestXml.OfficeApp.Version
if ([string]::IsNullOrWhiteSpace($manifestVersion)) {
  $manifestVersion = "0.0.0.0"
}

$isHostedPackage = $PackageMode -eq "Hosted"
if ($isHostedPackage -and [string]::IsNullOrWhiteSpace($HostedBaseUrl)) {
  throw "HostedBaseUrl is required when PackageMode is Hosted."
}

$packageName = if ($isHostedPackage) { "MailyOutlookAddinHosted-$manifestVersion" } else { "MailyOutlookAddin-$manifestVersion" }
$packageRoot = Join-Path $OutputRoot $packageName
$payloadRoot = Join-Path $packageRoot "payload"
$webProject = Join-Path $repoRoot "services\Maily.OutlookAddin.Web\Maily.OutlookAddin.Web.csproj"
$supervisorProject = Join-Path $repoRoot "services\Maily.OutlookAddin.Supervisor\Maily.OutlookAddin.Supervisor.csproj"
$setupProject = Join-Path $repoRoot "tools\Maily.OutlookAddin.Setup\Maily.OutlookAddin.Setup.csproj"
$webPublishDir = Join-Path $payloadRoot "app\Maily.OutlookAddin.Web"
$supervisorPublishDir = Join-Path $payloadRoot "app\Maily.OutlookAddin.Supervisor"
$setupPublishDir = Join-Path $OutputRoot ($packageName + "-setup-publish")
$setupExeName = "MailyOutlookAddinSetup.exe"
$summaryPath = Join-Path $packageRoot "package-summary.json"
$packageModePath = Join-Path $packageRoot "package-mode.txt"

function Assert-ChildPath {
  param(
    [Parameter(Mandatory = $true)][string]$Parent,
    [Parameter(Mandatory = $true)][string]$Child
  )

  $separator = [System.IO.Path]::DirectorySeparatorChar
  $parentFull = [System.IO.Path]::GetFullPath($Parent).TrimEnd($separator)
  $childFull = [System.IO.Path]::GetFullPath($Child)
  if ($childFull -ne $parentFull -and -not $childFull.StartsWith($parentFull + $separator, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to remove '$childFull' because it is outside '$parentFull'."
  }
}

function Invoke-DotnetPublish {
  param(
    [Parameter(Mandatory = $true)][string]$ProjectPath,
    [Parameter(Mandatory = $true)][string]$PublishDir
  )

  New-Item -ItemType Directory -Force -Path $PublishDir | Out-Null
  & dotnet publish $ProjectPath `
    -c $Configuration `
    -r $Runtime `
    --self-contained true `
    -o $PublishDir `
    /nodeReuse:false `
    -p:Version=$manifestVersion `
    -p:FileVersion=$manifestVersion `
    -p:PublishSingleFile=false `
    -p:EnableWindowsTargeting=true `
    -p:PublishReadyToRun=true `
    -p:DebugType=None `
    -p:DebugSymbols=false
  if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed for $ProjectPath."
  }
}

function Invoke-DotnetPublishSetup {
  param(
    [Parameter(Mandatory = $true)][string]$ProjectPath,
    [Parameter(Mandatory = $true)][string]$PublishDir
  )

  New-Item -ItemType Directory -Force -Path $PublishDir | Out-Null
  & dotnet publish $ProjectPath `
    -c $Configuration `
    -r $Runtime `
    --self-contained true `
    -o $PublishDir `
    /nodeReuse:false `
    -p:EnableWindowsTargeting=true `
    -p:Version=$manifestVersion `
    -p:FileVersion=$manifestVersion `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:PublishReadyToRun=true `
    -p:DebugType=None `
    -p:DebugSymbols=false
  if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed for $ProjectPath."
  }
}

New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null
if (Test-Path -LiteralPath $packageRoot) {
  throw "Package output already exists: $packageRoot. Choose a new output directory."
}

if ((Test-Path -LiteralPath $setupPublishDir) -or (Test-Path -LiteralPath (Join-Path $OutputRoot "$packageName.zip"))) { throw "Build output already exists. Choose a new output directory." }
New-Item -ItemType Directory -Force -Path $payloadRoot | Out-Null

if ($isHostedPackage) {
  Write-Host "Building hosted Outlook add-in package. Local host and supervisor payloads will not be included."
}
else {
  Write-Host "Publishing self-contained Maily Outlook add-in host..."
  Invoke-DotnetPublish -ProjectPath $webProject -PublishDir $webPublishDir

  Write-Host "Publishing self-contained hidden supervisor..."
  Invoke-DotnetPublish -ProjectPath $supervisorProject -PublishDir $supervisorPublishDir

  . (Join-Path $PSScriptRoot 'MailyPackageRuntime.ps1')
  $runtimeSharing = Compress-MailyPackageRuntime -AppRoot (Join-Path $payloadRoot 'app') -EvidenceRoot (Join-Path $OutputRoot ($packageName + '-shared-runtime-evidence'))
  Write-Host "Shared runtime package files: $($runtimeSharing.files); bytes avoided: $($runtimeSharing.bytes)."

  if ($AiRuntimeMode -eq 'Bundled') {
    & (Join-Path $PSScriptRoot 'Bundle-MailyCodexRuntime.ps1') -Destination (Join-Path $payloadRoot 'runtime/codex') -Runtime $Runtime
  }
}

& (Join-Path $PSScriptRoot 'Bundle-MailyOutlookConnector.ps1') -Destination (Join-Path $payloadRoot 'runtime/atk') -Runtime $Runtime

Write-Host "Publishing self-contained GUI installer..."
Invoke-DotnetPublishSetup -ProjectPath $setupProject -PublishDir $setupPublishDir
$setupSource = Join-Path $setupPublishDir "Maily.OutlookAddin.Setup.exe"
if (-not (Test-Path -LiteralPath $setupSource)) {
  throw "GUI setup executable was not published at $setupSource."
}
Copy-Item -LiteralPath $setupSource -Destination (Join-Path $packageRoot $setupExeName) -Force
# Retain setup publish evidence outside the distributable package.

Write-Host "Copying package payload..."
New-Item -ItemType Directory -Force -Path (Join-Path $payloadRoot "addin") | Out-Null
$packageManifestPath = Join-Path $payloadRoot "addin\manifest.xml"
if ($isHostedPackage) {
  $manifestGenerator = Join-Path $repoRoot "scripts\New-MailyHostedManifest.ps1"
  if (-not (Test-Path -LiteralPath $manifestGenerator)) {
    throw "Hosted manifest generator not found at $manifestGenerator."
  }

  & $manifestGenerator -BaseUrl $HostedBaseUrl -SourcePath $manifestSource -DestinationPath $packageManifestPath | Out-Host
}
else {
  Copy-Item -LiteralPath $manifestSource -Destination $packageManifestPath -Force

  New-Item -ItemType Directory -Force -Path (Join-Path $payloadRoot "scripts") | Out-Null
  Copy-Item -LiteralPath (Join-Path $repoRoot "scripts\outlook-bridge.ps1") -Destination (Join-Path $payloadRoot "scripts\outlook-bridge.ps1") -Force

  New-Item -ItemType Directory -Force -Path (Join-Path $payloadRoot "smoke_data") | Out-Null
  Copy-Item -LiteralPath (Join-Path $repoRoot "smoke_data\prompts") -Destination (Join-Path $payloadRoot "smoke_data\prompts") -Recurse -Force
}

New-Item -ItemType Directory -Force -Path (Join-Path $payloadRoot 'scripts') | Out-Null
foreach ($connectorScript in @('connect-outlook.cjs', 'outlook-credential-preflight.cjs')) {
  Copy-Item -LiteralPath (Join-Path $PSScriptRoot $connectorScript) -Destination (Join-Path $payloadRoot "scripts/$connectorScript")
}
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "setup") | Out-Null
Copy-Item -LiteralPath (Join-Path $repoRoot "scripts\Install-MailyOutlookAddinPackage.ps1") -Destination (Join-Path $packageRoot "setup\install.ps1") -Force
Copy-Item -LiteralPath (Join-Path $repoRoot "scripts\Uninstall-MailyOutlookAddinPackage.ps1") -Destination (Join-Path $packageRoot "setup\uninstall.ps1") -Force
Copy-Item -LiteralPath (Join-Path $repoRoot "scripts\OutlookShutdown.ps1") -Destination (Join-Path $packageRoot "setup\OutlookShutdown.ps1") -Force
Copy-Item -LiteralPath (Join-Path $repoRoot 'scripts/MailyPackageRuntime.ps1') -Destination (Join-Path $packageRoot 'setup/MailyPackageRuntime.ps1')
Set-Content -LiteralPath $packageModePath -Value $PackageMode -Encoding ASCII
Set-Content -LiteralPath (Join-Path $packageRoot 'ai-runtime-mode.txt') -Value $AiRuntimeMode -Encoding ASCII
$documentsRoot = Join-Path $packageRoot 'docs'
New-Item -ItemType Directory -Force -Path $documentsRoot | Out-Null
$releaseNotesVersion = ([version]$manifestVersion).ToString(3)
foreach ($document in @('DISCLAIMER.md', "RELEASE_NOTES_$releaseNotesVersion.md", 'INSTALLER.md')) {
  Copy-Item -LiteralPath (Join-Path $repoRoot "docs/$document") -Destination (Join-Path $documentsRoot $document) -Force
}
foreach ($license in @('LICENSE', 'LICENSE.md')) {
  Copy-Item -LiteralPath (Join-Path $repoRoot $license) -Destination (Join-Path $documentsRoot $license) -Force
}
if (-not $isHostedPackage -and $AiRuntimeMode -eq 'Bundled') {
  $noticesRoot = Join-Path $packageRoot 'third-party/codex'
  New-Item -ItemType Directory -Force -Path $noticesRoot | Out-Null
  foreach ($notice in @('LICENSE', 'NOTICE', 'codex-runtime.lock.json')) {
    Copy-Item -LiteralPath (Join-Path $payloadRoot "runtime/codex/$notice") -Destination (Join-Path $noticesRoot $notice) -Force
  }
}

# Keep one executable entry point. Support paths remain stable for installation,
# package integrity and Windows Installed Apps removal.
$connection = if ($isHostedPackage) {
  'Your computer must be able to connect to the Maily service supplied by your pilot organizer.'
} elseif ($AiRuntimeMode -eq 'InstalledApp') {
  'You need the ChatGPT desktop app with its Codex features installed and signed in. Setup checks and reuses that app; it does not install another AI runtime.'
} else {
  'You need a ChatGPT account. Setup will ask you to sign in only if needed. No separate software installation is required.'
}
$startHere = @"
INSTALL MAILY

1. Extract the entire ZIP into a folder. Keep its files and folders together.
2. Open MailyOutlookAddinSetup.exe.
3. Follow the setup window. Save your drafts before allowing it to restart Outlook.

You need classic Outlook for Windows, already signed in.
$connection

If setup pauses, follow the next step shown in that window. Installation is not
complete until Outlook is connected. If your organization does not allow this
add-in, contact your pilot organizer before continuing.

AI can make mistakes. Review all output and actions before use. You are responsible
for your use of Maily and for following your organization's data and AI policies.
Full terms and release information are in the docs folder.

To remove Maily, open Windows Settings > Apps > Installed apps > Maily Outlook
Add-in > Uninstall. You do not need this downloaded folder to uninstall it.
"@
$startHere | Set-MailyUtf8Content -LiteralPath (Join-Path $packageRoot 'START HERE.txt')
$packageFiles = Get-ChildItem -LiteralPath $packageRoot -Recurse -File
$recipientPrerequisites = if ($isHostedPackage) {
  @(
    "Windows",
    "Microsoft Outlook desktop signed in",
    "Network access to hosted Maily URL"
  )
}
else {
  @(
    "Windows",
    "Microsoft Outlook desktop signed in",
    $(if ($AiRuntimeMode -eq 'InstalledApp') { "ChatGPT desktop app with Codex runtime installed and signed in" } else { "Eligible ChatGPT account; setup guides sign-in using the bundled Codex CLI" })
  )
}
$recipientDoesNotRequire = if ($isHostedPackage) {
  @(
    "Maily source repo",
    ".NET SDK or .NET runtime",
    "Node.js or npm",
    "Office add-in developer tooling",
    "Local Codex configuration",
    "Local web payload or Outlook bridge",
    "Supervisor, scheduled task, Run key, Windows service, or localhost certificates",
    "Administrator rights for the default per-user install"
  )
}
else {
  @(
    "Maily source repo",
    ".NET SDK or .NET runtime",
    "Node.js or npm",
    "Office add-in developer tooling",
    "Administrator rights for the default per-user install"
  )
}
$summary = [pscustomobject]@{
  generatedAtUtc = [DateTime]::UtcNow.ToString("o")
  packageRoot = '.'
  packageName = $packageName
  packageMode = $PackageMode
  aiRuntimeMode = $AiRuntimeMode
  hostedBaseUrl = if ($isHostedPackage) { $HostedBaseUrl } else { "" }
  manifestVersion = $manifestVersion
  runtime = $Runtime
  configuration = $Configuration
  fileCount = $packageFiles.Count
  totalSizeBytes = ($packageFiles | Measure-Object -Property Length -Sum).Sum
  setupExecutable = $setupExeName
  recipientPrerequisites = $recipientPrerequisites
  recipientDoesNotRequire = $recipientDoesNotRequire
  bundledCodexVersion = if ($isHostedPackage -or $AiRuntimeMode -eq 'InstalledApp') { $null } else { (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'codex-runtime.lock.json') -Raw | ConvertFrom-Json).version }
}
if (-not $isHostedPackage) { $summary.recipientDoesNotRequire += 'Separate Codex CLI installation' }
$zipPath = Join-Path $OutputRoot "$packageName.zip"
if (-not $NoZip) { $summary | Add-Member -NotePropertyName zipPath -NotePropertyValue "$packageName.zip" }
$summary | ConvertTo-Json -Depth 8 | Set-MailyUtf8Content -LiteralPath $summaryPath

# Hash all packaged files after the final summary write. The inventory excludes
# itself; the archive has its own sidecar to avoid circular hashes.
$hashes = @(Get-ChildItem -LiteralPath $packageRoot -Recurse -File | Sort-Object FullName | ForEach-Object {
  [pscustomobject]@{
    path = $_.FullName.Substring($packageRoot.Length + 1).Replace('\', '/')
    bytes = $_.Length
    sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
  }
})
$hashes | ConvertTo-Json -Depth 4 | Set-MailyUtf8Content -LiteralPath (Join-Path $packageRoot 'SHA256SUMS.json')

if (-not $NoZip) {
  if (Test-Path -LiteralPath $zipPath) {
    throw "Archive already exists: $zipPath. Choose a new output directory."
  }

  Write-Host "Creating zip package..."
  Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zipPath -Force
  $archiveHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
  "$archiveHash  $packageName.zip" | Set-Content -LiteralPath "$zipPath.sha256" -Encoding ASCII
}

$summary | ConvertTo-Json -Depth 8
