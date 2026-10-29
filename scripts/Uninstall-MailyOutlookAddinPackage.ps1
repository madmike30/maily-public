param(
  [string]$InstallRoot = "",
  [switch]$KeepUserData = $true,
  [switch]$KeepLogs = $true,
  [switch]$KeepCertificate,
  [switch]$RemoveTrustedCertificate,
  [switch]$CloseOutlook,
  [ValidateSet("Auto", "Hosted", "Local")]
  [string]$PackageMode = "Auto",
  [switch]$Quiet,
  [switch]$Force
)

$ErrorActionPreference = "Stop"

function Set-MailyUtf8Content {
  [CmdletBinding()] param([Parameter(Mandatory=$true)][string]$LiteralPath, [Parameter(ValueFromPipeline=$true)][AllowEmptyString()][string]$Value)
  begin { $lines = New-Object 'System.Collections.Generic.List[string]' }
  process { $lines.Add($Value) }
  end { [IO.File]::WriteAllText($LiteralPath, (($lines -join [Environment]::NewLine) + [Environment]::NewLine), (New-Object Text.UTF8Encoding($false))) }
}


if ([string]::IsNullOrWhiteSpace($InstallRoot)) {
  if ($PSScriptRoot -and (Test-Path -LiteralPath (Join-Path $PSScriptRoot "app"))) {
    $InstallRoot = $PSScriptRoot
  }
  elseif ($PSScriptRoot -and (Test-Path -LiteralPath (Join-Path $PSScriptRoot "addin"))) {
    $InstallRoot = $PSScriptRoot
  }
  elseif ($PSScriptRoot -and (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) "app"))) {
    $InstallRoot = Split-Path -Parent $PSScriptRoot
  }
  elseif ($PSScriptRoot -and (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) "addin"))) {
    $InstallRoot = Split-Path -Parent $PSScriptRoot
  }
  else {
    $InstallRoot = Join-Path $env:LOCALAPPDATA "Maily\OutlookAddin"
  }
}

$installRootPath = [System.IO.Path]::GetFullPath($InstallRoot)
$runId = Get-Date -Format "yyyyMMdd-HHmmss"
$steps = New-Object System.Collections.Generic.List[object]
$developerKeyPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$runKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$runKeyName = "MailyOutlookAddinHostStartup"
$startupShortcutName = "Maily Outlook Add-in Host.lnk"
$taskName = "MailyOutlookAddinHostWatcher"
$uninstallKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\MailyOutlookAddin"
$packageModeFileName = "package-mode.txt"

function Write-UninstallerHost {
  param([string]$Message)

  if (-not $Quiet) {
    Write-Host $Message
  }
}

function Get-MailyRemovalOutcome {
  param([AllowEmptyCollection()][object[]]$Steps = @())
  $failed = @($Steps | Where-Object { $_.status -eq 'failed' }).Count -gt 0
  $filesRemoved = @($Steps | Where-Object { $_.name -eq 'install-files' -and $_.status -eq 'passed' }).Count -gt 0
  $localComplete = $filesRemoved -and -not $failed
  # Local registry/cache cleanup cannot prove that a mailbox registration was removed.
  [pscustomobject]@{
    exitCode = if ($localComplete) { 22 } else { 1 }
    status = if ($localComplete) { 'mailbox-removal-unverified' } else { 'failed' }
    localRemovalComplete = $localComplete
    mailboxRemovalVerified = $false
  }
}

function Add-Step {
  param(
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][ValidateSet("passed", "failed", "warning", "skipped")][string]$Status,
    [string]$Detail = "",
    [object]$Data = $null
  )

  $steps.Add([pscustomobject]@{
    name = $Name
    status = $Status
    detail = $Detail
    data = $Data
  }) | Out-Null
}

function Invoke-UninstallerStep {
  param(
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock
  )

  try {
    $data = & $ScriptBlock
    Add-Step -Name $Name -Status "passed" -Detail "Completed." -Data $data
    return $data
  }
  catch {
    Add-Step -Name $Name -Status "failed" -Detail $_.Exception.Message
    if (-not $Force -or $Name -eq "outlook-close") {
      throw
    }
  }
}

function Assert-ChildPath {
  param([Parameter(Mandatory = $true)][string]$Parent, [Parameter(Mandatory = $true)][string]$Child)
  $parentFull = [IO.Path]::GetFullPath($Parent).TrimEnd([char[]]'\/')
  $childFull = [IO.Path]::GetFullPath($Child)
  if ($childFull -ne $parentFull -and -not $childFull.StartsWith($parentFull + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to operate outside '$parentFull'."
  }
}

function Assert-SafeInstallRoot {
  param([Parameter(Mandatory = $true)][string]$Path)

  $full = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
  $localAppData = [System.IO.Path]::GetFullPath($env:LOCALAPPDATA).TrimEnd('\')
  $expectedDefault = [System.IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA "Maily\OutlookAddin")).TrimEnd('\')

  if ($full -eq $expectedDefault) {
    return
  }

  if (-not $full.StartsWith($localAppData + "\", [System.StringComparison]::OrdinalIgnoreCase) -and -not $Force) {
    throw "Refusing to remove custom install root '$full' without -Force."
  }
}

function Read-InstallState {
  $statePath = Join-Path $installRootPath "data\install-state.json"
  if (-not (Test-Path -LiteralPath $statePath)) {
    return [pscustomobject]@{
      packageMode = ""
      addinId = ""
      manifestPath = Join-Path $installRootPath "addin\manifest.xml"
      certificateThumbprint = ""
      statePath = $statePath
    }
  }

  $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
  if (-not ($state.PSObject.Properties.Name -contains "statePath")) {
    $state | Add-Member -NotePropertyName statePath -NotePropertyValue $statePath
  }

  return $state
}

function Resolve-PackageMode {
  param([object]$State = $null)

  if ($PackageMode -ne "Auto") {
    return $PackageMode
  }

  $modePath = Join-Path $installRootPath $packageModeFileName
  if (Test-Path -LiteralPath $modePath) {
    $modeText = (Get-Content -LiteralPath $modePath -Raw).Trim()
    if ($modeText -eq "Hosted" -or $modeText -eq "Local") {
      return $modeText
    }
  }

  if ($null -ne $State) {
    $stateMode = [string]$State.packageMode
    if ($stateMode -eq "Hosted" -or $stateMode -eq "Local") {
      return $stateMode
    }
  }

  if (Test-Path -LiteralPath (Join-Path $installRootPath "app\Maily.OutlookAddin.Web\Maily.OutlookAddin.Web.exe")) {
    return "Local"
  }

  return "Hosted"
}

function Resolve-ManifestId {
  param([string]$ManifestPath)

  if (-not (Test-Path -LiteralPath $ManifestPath)) {
    return ""
  }

  try {
    $xml = [xml](Get-Content -LiteralPath $ManifestPath -Raw)
    return [string]$xml.OfficeApp.Id
  }
  catch {
    return ""
  }
}

function Stop-MailyProcesses {
  $stopped = New-Object System.Collections.Generic.List[object]
  # Only the runtime inside this install belongs to Maily; leave shared Codex/ChatGPT sessions alone.
  $runtimeExecutable = [System.IO.Path]::GetFullPath((Join-Path $installRootPath "runtime\codex\codex.exe"))
  Get-Process -Name codex -ErrorAction SilentlyContinue | Where-Object {
    try { [string]::Equals($_.Path, $runtimeExecutable, [StringComparison]::OrdinalIgnoreCase) } catch { $false }
  } | ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction Stop }
  $ownedHosts = @((Join-Path $installRootPath 'app\Maily.OutlookAddin.Supervisor\Maily.OutlookAddin.Supervisor.exe'), (Join-Path $installRootPath 'app\Maily.OutlookAddin.Web\Maily.OutlookAddin.Web.exe'))
  Get-Process -Name Maily.OutlookAddin.Supervisor,Maily.OutlookAddin.Web -ErrorAction SilentlyContinue | Where-Object {
    try { $ownedHosts -contains $_.Path } catch { $false }
  } | ForEach-Object {
    try {
      $stopped.Add([pscustomobject]@{ processName = $_.ProcessName; id = $_.Id }) | Out-Null
      Stop-Process -Id $_.Id -Force -ErrorAction Stop
    }
    catch {
    }
  }

  Start-Sleep -Milliseconds 500
  return $stopped
}

function Test-OutlookRunning {
  return $null -ne (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
}

function Close-OutlookForSetup {
  $helper = Join-Path $PSScriptRoot 'OutlookShutdown.ps1'
  if (-not (Test-Path -LiteralPath $helper -PathType Leaf)) { throw 'The Outlook shutdown helper is missing. Extract the complete Maily package and retry.' }
  . $helper
  Invoke-MailyOutlookShutdown -HelperPath $helper
}
function Unregister-MailyStartup {
  $removed = @{
    runKey = $false
    scheduledTask = $false
    startupShortcut = $false
  }

  if (Get-ItemProperty -Path $runKeyPath -Name $runKeyName -ErrorAction SilentlyContinue) {
    Remove-ItemProperty -Path $runKeyPath -Name $runKeyName -ErrorAction SilentlyContinue
    $removed.runKey = $true
  }

  $startupDirectory = [Environment]::GetFolderPath("Startup")
  if (-not [string]::IsNullOrWhiteSpace($startupDirectory)) {
    $startupShortcutPath = Join-Path $startupDirectory $startupShortcutName
    if (Test-Path -LiteralPath $startupShortcutPath) {
      Remove-Item -LiteralPath $startupShortcutPath -Force -ErrorAction SilentlyContinue
      $removed.startupShortcut = $true
    }
  }

  if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    $removed.scheduledTask = $true
  }

  [pscustomobject]$removed
}

function Remove-MailyCacheFiles {
  param([Parameter(Mandatory = $true)][string]$CacheRoot, [Parameter(Mandatory = $true)][string]$AddinId)
  $guid = [Guid]::Empty
  if (-not [Guid]::TryParse($AddinId, [ref]$guid)) { throw 'A valid Maily add-in ID is required for cache cleanup.' }
  if (-not (Test-Path -LiteralPath $CacheRoot)) { return 0 }
  $root = Get-Item -LiteralPath $CacheRoot -Force
  if (($root.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Refusing to clean a redirected Office cache root.' }
  $pattern = '(?i)(?<![a-f0-9])(?:' + [regex]::Escape($guid.ToString('D')) + '|' + $guid.ToString('N') + ')(?![a-f0-9])'
  $pending = New-Object 'System.Collections.Generic.Queue[string]'
  $matchesToRemove = New-Object 'System.Collections.Generic.List[string]'
  $pending.Enqueue($root.FullName)
  while ($pending.Count -gt 0) {
    foreach ($item in Get-ChildItem -LiteralPath $pending.Dequeue() -Force -ErrorAction Stop) {
      Assert-ChildPath -Parent $root.FullName -Child $item.FullName
      if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        if ($item.FullName -match $pattern) { throw "Maily cache entry is a reparse point: $($item.FullName)" }
        continue
      }
      if ($item.PSIsContainer) { $pending.Enqueue($item.FullName) }
      if ($item.Name -match $pattern) { $matchesToRemove.Add($item.FullName) | Out-Null }
    }
  }
  $removed = 0
  foreach ($path in ($matchesToRemove | Sort-Object Length -Descending)) {
    if (-not (Test-Path -LiteralPath $path)) { continue }
    Assert-ChildPath -Parent $root.FullName -Child $path
    try { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop; $removed++ }
    catch { throw "Could not clear Maily cache '$path'. Close Outlook and retry. $($_.Exception.Message)" }
  }
  return $removed
}

function Remove-MailyOutlookCache {
  param([Parameter(Mandatory = $true)][string]$AddinId)
  if (Test-OutlookRunning) { throw 'Close Outlook before refreshing Maily cache.' }
  $guid = [Guid]::Empty
  if (-not [Guid]::TryParse($AddinId, [ref]$guid)) { throw 'A valid Maily add-in ID is required for cache cleanup.' }
  $removed = 0
  $compactId = $guid.ToString('N')
  $canonicalId = $guid.ToString('D')
  $customUiValidationCachePath = 'HKCU:\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache'
  if (Test-Path -Path $customUiValidationCachePath) {
    $properties = (Get-ItemProperty -Path $customUiValidationCachePath -ErrorAction Stop).PSObject.Properties |
      Where-Object { $_.Name -like "$canonicalId*" -or $_.Name -like "$compactId*" }
    foreach ($property in @($properties)) {
      Remove-ItemProperty -Path $customUiValidationCachePath -Name $property.Name -ErrorAction Stop
      $removed++
    }
  }
  $removed += Remove-MailyCacheFiles -CacheRoot (Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef') -AddinId $canonicalId
  $removed += Remove-MailyCacheFiles -CacheRoot (Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook\RoamCache') -AddinId $canonicalId
  return $removed
}

function Unregister-MailyOutlookAddin {
  param(
    [Parameter(Mandatory = $true)][object]$State,
    [switch]$SkipCacheCleanup
  )

  $addinId = [string]$State.addinId
  if ([string]::IsNullOrWhiteSpace($addinId)) {
    $addinId = Resolve-ManifestId -ManifestPath ([string]$State.manifestPath)
  }

  $removedRegistration = $false
  if (-not [string]::IsNullOrWhiteSpace($addinId) -and (Test-Path -Path $developerKeyPath)) {
    $currentValue = ""
    try {
      $currentValue = [string](Get-ItemProperty -Path $developerKeyPath -Name $addinId -ErrorAction Stop | Select-Object -ExpandProperty $addinId)
    }
    catch {
      $currentValue = ""
    }

    $manifestPath = [string]$State.manifestPath
    if ($Force -or [string]::IsNullOrWhiteSpace($currentValue) -or $currentValue -eq $manifestPath -or $currentValue.StartsWith($installRootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
      Remove-ItemProperty -Path $developerKeyPath -Name $addinId -ErrorAction SilentlyContinue
      $removedRegistration = $true
    }

    New-ItemProperty -Path $developerKeyPath -Name "RefreshAddins" -PropertyType DWord -Value 1 -Force | Out-Null
  }

  [pscustomobject]@{
    addinId = $addinId
    removedRegistration = $removedRegistration
    removedCacheItems = if ($SkipCacheCleanup -or [string]::IsNullOrWhiteSpace($addinId)) { 0 } else { Remove-MailyOutlookCache -AddinId $addinId }
  }
}

function Test-MailyCertificateOwnership {
  param([AllowNull()][object]$Certificate, [string]$ExpectedThumbprint)
  # A saved thumbprint is a reference, not proof of ownership: older installers
  # could reuse another application's trusted localhost certificate.
  return $null -ne $Certificate -and
    -not [string]::IsNullOrWhiteSpace($ExpectedThumbprint) -and
    [string]::Equals([string]$Certificate.Thumbprint, $ExpectedThumbprint, [StringComparison]::OrdinalIgnoreCase) -and
    [string]::Equals([string]$Certificate.FriendlyName, 'Maily Outlook Add-in Localhost', [StringComparison]::Ordinal) -and
    [string]::Equals([string]$Certificate.Subject, 'CN=localhost', [StringComparison]::OrdinalIgnoreCase) -and
    [bool]$Certificate.HasPrivateKey
}

function Remove-MailyCertificate {
  param([Parameter(Mandatory = $true)][object]$State)

  $thumbprint = [string]$State.certificateThumbprint
  if ([string]::IsNullOrWhiteSpace($thumbprint)) {
    return [pscustomobject]@{ removed = 0; thumbprint = "" }
  }

  $ownedCertificate = $false
  $personalStore = New-Object System.Security.Cryptography.X509Certificates.X509Store(
    'My', [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
  try {
    $personalStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
    $personalMatches = $personalStore.Certificates.Find(
      [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $thumbprint, $false)
    foreach ($candidate in $personalMatches) {
      if (Test-MailyCertificateOwnership -Certificate $candidate -ExpectedThumbprint $thumbprint) { $ownedCertificate = $true }
    }
  }
  finally { $personalStore.Close() }
  if (-not $ownedCertificate) {
    Add-Step -Name 'shared-localhost-certificate' -Status skipped -Detail 'Retained the shared or unowned localhost certificate in all Windows stores; Maily ownership could not be established.'
    return [pscustomobject]@{ removed = 0; thumbprint = $thumbprint; skippedTrustedRoot = $true; retainedSharedCertificate = $true }
  }

  $removed = 0
  $skippedTrustedRoot = $false
  $storeNames = New-Object System.Collections.Generic.List[string]
  $storeNames.Add("My") | Out-Null
  $storeNames.Add("TrustedPeople") | Out-Null
  if (-not $RemoveTrustedCertificate) {
    $skippedTrustedRoot = $true
  }

  foreach ($storeName in $storeNames) {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
      $storeName,
      [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    try {
      $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
      $matches = $store.Certificates.Find(
        [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,
        $thumbprint,
        $false)
      foreach ($cert in $matches) {
        $store.Remove($cert)
        $removed++
      }
    }
    finally {
      $store.Close()
    }
  }

  if ($RemoveTrustedCertificate) {
    $certutil = Join-Path $env:WINDIR "System32\certutil.exe"
    $certutilOutput = & $certutil -user -delstore Root $thumbprint 2>&1
    if ($LASTEXITCODE -eq 0) {
      $removed++
    }
    elseif (($certutilOutput -join [Environment]::NewLine) -notmatch "(?i)cannot find|not found|0x80092004") {
      throw "certutil failed to remove the localhost certificate from CurrentUser Root. $certutilOutput"
    }
  }

  [pscustomobject]@{
    removed = $removed
    thumbprint = $thumbprint
    skippedTrustedRoot = $skippedTrustedRoot
    retainedSharedCertificate = $false
  }
}

function Remove-InstallFiles {
  Assert-SafeInstallRoot -Path $installRootPath
  if (-not (Test-Path -LiteralPath $installRootPath)) {
    return [pscustomobject]@{ removed = $false; path = $installRootPath }
  }

  if ($isHostedPackage) {
    $removedNames = New-Object System.Collections.Generic.List[string]
    foreach ($name in @("runtime", "scripts", "addin", "setup", "MailyOutlookAddinSetup.exe", $packageModeFileName, "ai-runtime-mode.txt", "package-summary.json")) {
      $target = Join-Path $installRootPath $name
      Assert-ChildPath -Parent $installRootPath -Child $target
      if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force
        $removedNames.Add($name) | Out-Null
      }
    }

    if (-not $KeepUserData) {
      $data = Join-Path $installRootPath "data"
      if (Test-Path -LiteralPath $data) {
        Remove-Item -LiteralPath $data -Recurse -Force
        $removedNames.Add("data") | Out-Null
      }
    }

    try {
      $remaining = @(Get-ChildItem -LiteralPath $installRootPath -Force -ErrorAction SilentlyContinue)
      if ($remaining.Count -eq 0) {
        Remove-Item -LiteralPath $installRootPath -Force
      }
    }
    catch {
    }

    return [pscustomobject]@{
      removed = $true
      path = $installRootPath
      removedNames = $removedNames
    }
  }

  if (-not $KeepUserData -and -not $KeepLogs) {
    Remove-Item -LiteralPath $installRootPath -Recurse -Force
    return [pscustomobject]@{ removed = $true; path = $installRootPath }
  }

  $removedNames = New-Object System.Collections.Generic.List[string]
  foreach ($name in @("app", "runtime", "addin", "scripts", "smoke_data", "certs", "setup", "uninstall.ps1", "MailyOutlookAddinSetup.exe", "package-mode.txt", "ai-runtime-mode.txt", "package-summary.json")) {
    $target = Join-Path $installRootPath $name
    Assert-ChildPath -Parent $installRootPath -Child $target
    if (Test-Path -LiteralPath $target) {
      Remove-Item -LiteralPath $target -Recurse -Force
      $removedNames.Add($name) | Out-Null
    }
  }

  if (-not $KeepUserData) {
    $data = Join-Path $installRootPath "data"
    if (Test-Path -LiteralPath $data) {
      Remove-Item -LiteralPath $data -Recurse -Force
      $removedNames.Add("data") | Out-Null
    }
  }

  if (-not $KeepLogs) {
    $logs = Join-Path $installRootPath "logs"
    if (Test-Path -LiteralPath $logs) {
      Remove-Item -LiteralPath $logs -Recurse -Force
      $removedNames.Add("logs") | Out-Null
    }
  }

  [pscustomobject]@{
    removed = $true
    path = $installRootPath
    removedNames = $removedNames
  }
}

$resolvedPackageMode = Resolve-PackageMode
$isHostedPackage = $resolvedPackageMode -eq "Hosted"

try {
  Write-UninstallerHost "Uninstalling Maily Outlook Add-in ($resolvedPackageMode package) from $installRootPath"
  $state = Invoke-UninstallerStep -Name "read-install-state" -ScriptBlock {
    Read-InstallState
  }
  $resolvedPackageMode = Resolve-PackageMode -State $state
  $isHostedPackage = $resolvedPackageMode -eq "Hosted"
  Add-Step -Name "package-mode" -Status "passed" -Detail "Resolved $resolvedPackageMode uninstall mode."

  if ($CloseOutlook) {
    Invoke-UninstallerStep -Name "outlook-close" -ScriptBlock {
      Close-OutlookForSetup
    } | Out-Null
  }
  elseif (Test-OutlookRunning) {
    Add-Step -Name "outlook-close" -Status "warning" -Detail "Outlook is running. It may keep showing cached Maily ribbon buttons until Outlook is restarted."
  }
  else {
    Add-Step -Name "outlook-close" -Status "passed" -Detail "Outlook is not running."
  }

  if ($isHostedPackage) {
    Add-Step -Name "startup-registration" -Status "skipped" -Detail "Hosted package does not manage supervisor startup, scheduled task, Run key, or Windows service."
  }
  else {
    Invoke-UninstallerStep -Name "startup-registration" -ScriptBlock {
      Unregister-MailyStartup
    } | Out-Null
  }

  if ($isHostedPackage) {
    Add-Step -Name "stop-processes" -Status "skipped" -Detail "Hosted package does not manage local Maily processes."
  }
  else {
    Invoke-UninstallerStep -Name "stop-processes" -ScriptBlock {
      Stop-MailyProcesses
    } | Out-Null
  }

  Invoke-UninstallerStep -Name "outlook-registration" -ScriptBlock {
    Unregister-MailyOutlookAddin -State $state
  } | Out-Null

  if ($isHostedPackage) {
    Add-Step -Name "localhost-certificate" -Status "skipped" -Detail "Hosted package does not install localhost certificates."
  }
  elseif ($KeepCertificate) {
    Add-Step -Name "localhost-certificate" -Status "skipped" -Detail "Kept by request."
  }
  else {
    Invoke-UninstallerStep -Name "localhost-certificate" -ScriptBlock {
      Remove-MailyCertificate -State $state
    } | Out-Null
  }

  Invoke-UninstallerStep -Name "windows-uninstall-entry" -ScriptBlock {
    if (Test-Path -Path $uninstallKeyPath) {
      Remove-Item -Path $uninstallKeyPath -Recurse -Force
      return @{ removed = $true; path = $uninstallKeyPath }
    }

    @{ removed = $false; path = $uninstallKeyPath }
  } | Out-Null

  Invoke-UninstallerStep -Name "install-files" -ScriptBlock {
    Remove-InstallFiles
  } | Out-Null
  Add-Step -Name 'mailbox-removal' -Status warning -Detail 'Maily was removed from this PC. Outlook may still have its add-in. Open https://aka.ms/olksideload and remove Maily from My add-ins if present. If your organization installed it, contact your administrator. Local setup cannot verify mailbox removal.'
}
catch {
  if (-not ($steps | Where-Object { $_.status -eq "failed" })) {
    Add-Step -Name "uninstaller" -Status "failed" -Detail $_.Exception.Message
  }
}
finally {
  try {
    $failedSteps = @($steps | Where-Object { $_.status -eq "failed" })
    $warningSteps = @($steps | Where-Object { $_.status -eq "warning" })
    $removalOutcome = Get-MailyRemovalOutcome -Steps $steps.ToArray()
    $summaryPath = Join-Path $env:TEMP "MailyOutlookAddin-uninstall-$runId.json"
    $summary = [pscustomobject]@{
      runId = $runId
      generatedAtUtc = [DateTime]::UtcNow.ToString("o")
      packageMode = $resolvedPackageMode
      installRoot = $installRootPath
      summaryPath = $summaryPath
      status = $removalOutcome.status
      localRemovalComplete = $removalOutcome.localRemovalComplete
      mailboxRemovalVerified = $removalOutcome.mailboxRemovalVerified
      error = if ($failedSteps.Count -gt 0) { [string]$failedSteps[0].detail } else { "" }
      failureCount = $failedSteps.Count
      warningCount = $warningSteps.Count
      steps = $steps
    }
    $summary | ConvertTo-Json -Depth 12 | Set-MailyUtf8Content -LiteralPath $summaryPath
    if (-not $Quiet) {
      $summary | ConvertTo-Json -Depth 12
    }
  }
  catch {
  }
}

exit (Get-MailyRemovalOutcome -Steps $steps.ToArray()).exitCode
