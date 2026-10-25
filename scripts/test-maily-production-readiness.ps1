param(
  [string]$InstallRoot = "",
  [string]$MailboxAddress = "",
  [string]$ProfileId = "default",
  [switch]$RequireDotNetSdk,
  [switch]$RequireOutlookRegistration
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$targetRoot = if ([string]::IsNullOrWhiteSpace($InstallRoot)) {
  $repoRoot
}
else {
  [System.IO.Path]::GetFullPath($InstallRoot)
}
$isInstalledTarget = -not [string]::IsNullOrWhiteSpace($InstallRoot)
$runId = Get-Date -Format "yyyyMMdd-HHmmss"
$outputDir = if ($isInstalledTarget) {
  Join-Path $targetRoot "data\production-readiness\$runId"
}
else {
  Join-Path $repoRoot "output\production-readiness\$runId"
}
$summaryPath = Join-Path $outputDir "summary.json"
$healthUrl = "https://localhost:7097/api/health"
$readinessUrl = "https://localhost:7097/api/outlook/readiness"
$settingsUrl = "https://localhost:7097/api/outlook/settings"
$addinId = "683ea353-6d71-4315-afd4-147e968605a9"
$developerRootPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$developerKeyPath = "$developerRootPath\$addinId"
$sideloadTitleKeyPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer\OutlookSideloadManifestPath"

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

function New-Check {
  param(
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][bool]$Ok,
    [string]$Detail = "",
    [object]$Data = $null
  )

  [pscustomobject]@{
    name = $Name
    ok = $Ok
    detail = $Detail
    data = $Data
  }
}

function Invoke-JsonGet {
  param([Parameter(Mandatory = $true)][string]$Uri)

  $response = Invoke-WebRequest -UseBasicParsing -Uri $Uri -TimeoutSec 20
  return $response.Content | ConvertFrom-Json
}

function Test-HttpOk {
  param([Parameter(Mandatory = $true)][string]$Uri)

  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri $Uri -TimeoutSec 10
    return $response.StatusCode -eq 200
  }
  catch {
    return $false
  }
}

function Get-CommandText {
  param([string]$Command)

  try {
    return [string](& cmd /c $Command 2>&1)
  }
  catch {
    return $_.Exception.Message
  }
}

function Resolve-NpxCommand {
  $cmd = Get-Command npx.cmd -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($cmd) {
    return $cmd.Source
  }

  $plain = Get-Command npx -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($plain) {
    return $plain.Source
  }

  return ""
}

function Test-OutlookRegistration {
  param([Parameter(Mandatory = $true)][string]$ManifestPath)

  $developerManifestPath = ""
  try {
    $developerManifestPath = [string](Get-ItemProperty -Path $developerRootPath -Name $addinId -ErrorAction Stop | Select-Object -ExpandProperty $addinId)
  }
  catch {
    $developerManifestPath = ""
  }
  $developerKeyPresent = -not [string]::IsNullOrWhiteSpace($developerManifestPath)
  $developerKeyMatches = $developerKeyPresent -and ([System.IO.Path]::GetFullPath($developerManifestPath) -eq [System.IO.Path]::GetFullPath($ManifestPath))
  $sideloadTitle = ""
  try {
    $sideloadTitle = [string](Get-ItemProperty -Path $sideloadTitleKeyPath -ErrorAction Stop | Select-Object -ExpandProperty TitleId)
  }
  catch {
    $sideloadTitle = ""
  }

  $npxPath = Resolve-NpxCommand
  $registeredOutput = ""
  $registeredWithTool = $false
  if (-not [string]::IsNullOrWhiteSpace($npxPath)) {
    try {
      $registeredOutput = [string](& $npxPath --yes office-addin-dev-settings registered 2>&1)
      $registeredWithTool = $registeredOutput -match [regex]::Escape($addinId) -and
        $registeredOutput -match [regex]::Escape($ManifestPath)
    }
    catch {
      $registeredOutput = $_.Exception.Message
    }
  }

  return [pscustomobject]@{
    ok = $developerKeyMatches -or $registeredWithTool -or -not [string]::IsNullOrWhiteSpace($sideloadTitle)
    developerKeyPresent = $developerKeyPresent
    developerKeyMatches = $developerKeyMatches
    developerRootPath = $developerRootPath
    developerValueName = $addinId
    developerManifestPath = $developerManifestPath
    sideloadTitleKeyPresent = -not [string]::IsNullOrWhiteSpace($sideloadTitle)
    sideloadTitle = $sideloadTitle
    npxPath = $npxPath
    registeredWithTool = $registeredWithTool
    registeredOutput = $registeredOutput
  }
}

$checks = @()

$manifestPath = Join-Path $targetRoot "addin\manifest.xml"
$manifestVersion = ""
if (Test-Path $manifestPath) {
  try {
    $manifestVersion = ([xml](Get-Content -LiteralPath $manifestPath -Raw)).OfficeApp.Version
  }
  catch {
    $manifestVersion = ""
  }
}
$checks += New-Check -Name "manifest" -Ok (Test-Path $manifestPath) -Detail "Manifest version $manifestVersion" -Data @{ path = $manifestPath; version = $manifestVersion }

if ($RequireDotNetSdk -or -not $isInstalledTarget) {
  $dotnetSdks = Get-CommandText "dotnet --list-sdks"
  $checks += New-Check -Name "dotnet-sdk" -Ok ($dotnetSdks -match "^10\.") -Detail "Developer/source checks require .NET SDK 10." -Data @{ sdks = $dotnetSdks }
}
else {
  $hostExe = Join-Path $targetRoot "app\Maily.OutlookAddin.Web\Maily.OutlookAddin.Web.exe"
  $supervisorExe = Join-Path $targetRoot "app\Maily.OutlookAddin.Supervisor\Maily.OutlookAddin.Supervisor.exe"
  $checks += New-Check -Name "self-contained-runtime" -Ok ((Test-Path -LiteralPath $hostExe) -and (Test-Path -LiteralPath $supervisorExe)) -Detail "Installed package carries self-contained host and supervisor binaries." -Data @{
    hostExe = $hostExe
    supervisorExe = $supervisorExe
  }
}

$installStatePath = Join-Path $targetRoot "data\install-state.json"
$installState = $null
if ($isInstalledTarget -and (Test-Path -LiteralPath $installStatePath)) {
  try {
    $installState = Get-Content -LiteralPath $installStatePath -Raw | ConvertFrom-Json
  }
  catch {
    $installState = $null
  }
}
$codexHome = if ($installState -and ($installState.PSObject.Properties.Name -contains "codexHome") -and -not [string]::IsNullOrWhiteSpace([string]$installState.codexHome)) {
  [string]$installState.codexHome
}
elseif (-not [string]::IsNullOrWhiteSpace($env:CODEX_HOME)) {
  $env:CODEX_HOME
}
else {
  Join-Path $env:USERPROFILE ".codex"
}
$codexExecutablePath = if ($installState -and ($installState.PSObject.Properties.Name -contains "codexExecutablePath") -and -not [string]::IsNullOrWhiteSpace([string]$installState.codexExecutablePath)) {
  [string]$installState.codexExecutablePath
}
else {
  "codex"
}
$codexCommandFound = $false
$codexCommandPath = ""
$resolvedCodexExecutable = $codexExecutablePath
if ([System.IO.Path]::IsPathRooted($codexExecutablePath) -or $codexExecutablePath.Contains("\") -or $codexExecutablePath.Contains("/")) {
  $codexCommandFound = Test-Path -LiteralPath $codexExecutablePath
  if ($codexCommandFound) {
    $codexCommandPath = [System.IO.Path]::GetFullPath($codexExecutablePath)
    $resolvedCodexExecutable = $codexCommandPath
  }
}
else {
  $codexCommand = Get-Command $codexExecutablePath -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($codexCommand) {
    $codexCommandFound = $true
    $codexCommandPath = $codexCommand.Source
    $resolvedCodexExecutable = $codexCommand.Source
  }
}
$loginStatusOutput = ""
$loginStatusExitCode = -1
$loginStatusOk = $false
if ($codexCommandFound) {
  $previousCodexHome = $env:CODEX_HOME
  try {
    $env:CODEX_HOME = $codexHome
    $rawLoginStatusOutput = & $resolvedCodexExecutable login status 2>&1
    $loginStatusOutput = ($rawLoginStatusOutput | Out-String).Trim()
    $loginStatusExitCode = $LASTEXITCODE
    # The Codex desktop launcher can leave Windows PowerShell's LASTEXITCODE at -1
    # even when `login status` succeeds, so use its explicit status line.
    $loginStatusOk = $loginStatusOutput -match "(?im)^\s*Logged in(?:\s|$)"
  }
  catch {
    $loginStatusOutput = $_.Exception.Message
  }
  finally {
    $env:CODEX_HOME = $previousCodexHome
  }
}
$loginStatusOk = $codexCommandFound -and ($loginStatusOutput -match "(?im)^\s*Logged in(?:\s|$)")
if ($loginStatusOutput.Length -gt 1200) {
  $loginStatusOutput = $loginStatusOutput.Substring(0, 1200)
}
$codexDetail = if ($codexCommandFound -and $loginStatusOk) {
  "Codex CLI is available and signed in with ChatGPT."
}
elseif (-not $codexCommandFound) {
  "Codex CLI was not found."
}
else {
  "Run or open codex login and complete ChatGPT sign-in."
}
$checks += New-Check -Name "codex-login" -Ok ($codexCommandFound -and $loginStatusOk) -Detail $codexDetail -Data @{
  codexHome = $codexHome
  codexExecutablePath = $codexExecutablePath
  codexCommandPath = $codexCommandPath
  codexCommandFound = $codexCommandFound
  loginStatusOk = $loginStatusOk
  loginStatusExitCode = $loginStatusExitCode
  loginStatusOutput = $loginStatusOutput
}

$checks += New-Check -Name "host-health" -Ok (Test-HttpOk $healthUrl) -Detail "Local backend health endpoint."

try {
  $readiness = Invoke-JsonGet $readinessUrl
  $checks += New-Check -Name "host-readiness" -Ok ([bool]$readiness.ready) -Detail "Backend can resolve repo, prompts, bridge, settings, and Codex CLI login readiness." -Data $readiness
}
catch {
  $checks += New-Check -Name "host-readiness" -Ok $false -Detail $_.Exception.Message
}

$taskAction = $null
try {
  $taskAction = Get-ScheduledTask -TaskName "MailyOutlookAddinHostWatcher" -ErrorAction Stop | Select-Object -ExpandProperty Actions
  $checks += New-Check -Name "startup-task" -Ok $false -Detail "Legacy scheduled task should not exist." -Data $taskAction
}
catch {
  $checks += New-Check -Name "startup-task" -Ok $true -Detail "Legacy scheduled task is absent."
}

$startupShortcutPath = Join-Path ([Environment]::GetFolderPath("Startup")) "Maily Outlook Add-in Host.lnk"
try {
  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($startupShortcutPath)
  $expectedTarget = if ($isInstalledTarget) {
    Join-Path $targetRoot "app\Maily.OutlookAddin.Supervisor\Maily.OutlookAddin.Supervisor.exe"
  }
  else {
    Join-Path $targetRoot "output\startup\Maily.OutlookAddin.Supervisor\Maily.OutlookAddin.Supervisor.exe"
  }
  $checks += New-Check -Name "startup-shortcut" -Ok ((Test-Path -LiteralPath $startupShortcutPath) -and ($shortcut.TargetPath -eq $expectedTarget)) -Detail "Current-user Startup shortcut points to the Maily supervisor." -Data @{ path = $startupShortcutPath; target = $shortcut.TargetPath }
}
catch {
  $checks += New-Check -Name "startup-shortcut" -Ok $false -Detail $_.Exception.Message
}

$mailyProcesses = @(Get-Process -Name Maily.OutlookAddin.Supervisor,Maily.OutlookAddin.Web -ErrorAction SilentlyContinue | Select-Object ProcessName,Id,MainWindowTitle,@{Name="WorkingSetMB";Expression={[math]::Round($_.WorkingSet64 / 1MB, 1)}})
$checks += New-Check -Name "invisible-processes" -Ok (($mailyProcesses.Count -ge 1) -and -not ($mailyProcesses | Where-Object { -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle) })) -Detail "Maily processes have no visible window title." -Data $mailyProcesses

$legacyWatcher = @(
  Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe' OR Name = 'pwsh.exe'" |
    Where-Object { $_.CommandLine -like "*watch-outlook-and-start-maily-addin.ps1*" }
)
$checks += New-Check -Name "legacy-watcher" -Ok ($legacyWatcher.Count -eq 0) -Detail "No old PowerShell watcher should be running." -Data $legacyWatcher

$registrationState = Test-OutlookRegistration -ManifestPath $manifestPath
$checks += New-Check -Name "outlook-registration" -Ok ((-not $RequireOutlookRegistration) -or [bool]$registrationState.ok) -Detail "Outlook add-in registration through dev settings or sideload title state." -Data $registrationState

if (-not [string]::IsNullOrWhiteSpace($MailboxAddress)) {
  try {
    $settings = Invoke-JsonGet ("{0}?mailboxAddress={1}&profileId={2}" -f $settingsUrl, [System.Uri]::EscapeDataString($MailboxAddress), [System.Uri]::EscapeDataString($ProfileId))
    $styleOk = $settings.style.sampleCount -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$settings.style.summary)
    $checks += New-Check -Name "personal-style" -Ok $styleOk -Detail "Personal style profile should have refreshed Sent Items evidence." -Data $settings.style
  }
  catch {
    $checks += New-Check -Name "personal-style" -Ok $false -Detail $_.Exception.Message
  }
}

$summary = [pscustomobject]@{
  runId = $runId
  generatedAtUtc = [DateTime]::UtcNow.ToString("o")
  targetRoot = $targetRoot
  installedTarget = $isInstalledTarget
  outputDir = $outputDir
  checks = $checks
}

$summary | ConvertTo-Json -Depth 12 | Set-Content -Path $summaryPath -Encoding UTF8
$summary | ConvertTo-Json -Depth 12

if ($checks | Where-Object { -not $_.ok }) {
  exit 1
}
