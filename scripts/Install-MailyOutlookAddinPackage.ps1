param(
  [string]$InstallRoot = "",
  [string]$CodexHome = "",
  [string]$CodexExecutablePath = "codex",
  [string]$MailboxAddress = "",
  [string]$ProfileId = "default",
  [ValidateSet("Auto", "Hosted", "Local")]
  [string]$PackageMode = "Auto",
  [switch]$SkipOutlookRegistration,
  [switch]$SkipStartupRegistration,
  [switch]$SkipVerification,
  [switch]$SkipStyleRefresh,
  [switch]$CloseOutlook,
  [switch]$TrustLocalhostRoot,
  [switch]$VerifyOutlookActivation,
  [switch]$Quiet
)

$ErrorActionPreference = "Stop"

function Set-MailyUtf8Content {
  [CmdletBinding()] param([Parameter(Mandatory=$true)][string]$LiteralPath, [Parameter(ValueFromPipeline=$true)][AllowEmptyString()][string]$Value)
  begin { $lines = New-Object 'System.Collections.Generic.List[string]' }
  process { $lines.Add($Value) }
  end { [IO.File]::WriteAllText($LiteralPath, (($lines -join [Environment]::NewLine) + [Environment]::NewLine), (New-Object Text.UTF8Encoding($false))) }
}

$ConfirmPreference = "None"

$packageRoot = $PSScriptRoot
if (-not (Test-Path -LiteralPath (Join-Path $packageRoot "payload"))) {
  $parentPackageRoot = Split-Path -Parent $PSScriptRoot
  if (-not [string]::IsNullOrWhiteSpace($parentPackageRoot) -and (Test-Path -LiteralPath (Join-Path $parentPackageRoot "payload"))) {
    $packageRoot = $parentPackageRoot
  }
}
$payloadRoot = Join-Path $packageRoot "payload"
if ([string]::IsNullOrWhiteSpace($InstallRoot)) {
  $InstallRoot = Join-Path $env:LOCALAPPDATA "Maily\OutlookAddin"
}

$installRootPath = [System.IO.Path]::GetFullPath($InstallRoot)
$runId = Get-Date -Format "yyyyMMdd-HHmmss"
$steps = New-Object System.Collections.Generic.List[object]
$liveLogPath = Join-Path $installRootPath "logs\install-$runId.log"
$healthUrl = "https://localhost:7097/api/health"
$readinessUrl = "https://localhost:7097/api/outlook/readiness"
$settingsUrl = "https://localhost:7097/api/outlook/settings"
$developerKeyPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$runKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$runKeyName = "MailyOutlookAddinHostStartup"
$startupShortcutName = "Maily Outlook Add-in Host.lnk"
$taskName = "MailyOutlookAddinHostWatcher"
$uninstallKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\MailyOutlookAddin"
$setupExeName = "MailyOutlookAddinSetup.exe"
$packageModeFileName = "package-mode.txt"
$resolvedCodexExecutablePath = ""
$upgradeSnapshot = $null
$applicationMutationStarted = $false
$outlookActivationRequired = $false
$outlookRestartVerificationRequired = $false
$outlookSupportRequired = $false
$script:installationPayloadReady = $false
$rollbackStatus = "not-needed"

function Write-InstallerHost {
  param([string]$Message)

  if (-not $Quiet) {
    Write-Host $Message
  }
}

function Write-InstallerLog {
  param([string]$Message)

  try {
    Assert-MailyUnredirectedPath -Path $liveLogPath
    $logDir = Split-Path -Parent $liveLogPath
    New-Item -ItemType Directory -Force -Path $logDir | Out-Null
    [IO.File]::AppendAllText($liveLogPath, (("{0}`t{1}" -f [DateTime]::UtcNow.ToString("o"), $Message) + [Environment]::NewLine), (New-Object Text.UTF8Encoding($false)))
  }
  catch {
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
  Write-InstallerLog ("{0}`t{1}`t{2}" -f $Status, $Name, $Detail)
  Write-Output ("MAILY_STEP|{0}|{1}|{2}" -f $Name, $Status, ($Detail -replace '[\r\n|]', ' ')) | Out-Host
}

function Resolve-PackageMode {
  if ($PackageMode -ne "Auto") {
    return $PackageMode
  }

  $modePath = Join-Path $packageRoot $packageModeFileName
  if (Test-Path -LiteralPath $modePath) {
    $modeText = (Get-Content -LiteralPath $modePath -Raw).Trim()
    if ($modeText -eq "Hosted" -or $modeText -eq "Local") {
      return $modeText
    }
  }

  if (Test-Path -LiteralPath (Join-Path $payloadRoot "app\Maily.OutlookAddin.Web\Maily.OutlookAddin.Web.exe")) {
    return "Local"
  }

  return "Hosted"
}

function Invoke-InstallerStep {
  param(
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock
  )

  try {
    Write-Output ("MAILY_STEP|{0}|running|Working..." -f $Name) | Out-Host
    $data = & $ScriptBlock
    Add-Step -Name $Name -Status "passed" -Detail "Completed." -Data $data
    return $data
  }
  catch {
    Add-Step -Name $Name -Status "failed" -Detail $_.Exception.Message
    throw
  }
}

function Assert-ChildPath {
  param(
    [Parameter(Mandatory = $true)][string]$Parent,
    [Parameter(Mandatory = $true)][string]$Child
  )

  $parentFull = [System.IO.Path]::GetFullPath($Parent).TrimEnd([char[]]'\/')
  $childFull = [System.IO.Path]::GetFullPath($Child)
  if ($childFull -ne $parentFull -and -not $childFull.StartsWith($parentFull + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to operate on '$childFull' because it is outside '$parentFull'."
  }
}

function Resolve-MailyManifest {
  param([Parameter(Mandatory = $true)][string]$ManifestPath)

  if (-not (Test-Path -LiteralPath $ManifestPath)) {
    throw "Manifest not found at $ManifestPath."
  }

  $xml = [xml](Get-Content -LiteralPath $ManifestPath -Raw)
  $id = [string]$xml.OfficeApp.Id
  $version = [string]$xml.OfficeApp.Version
  if ([string]::IsNullOrWhiteSpace($id)) {
    throw "Manifest id could not be read from $ManifestPath."
  }

  [pscustomobject]@{
    Id = $id
    Version = if ([string]::IsNullOrWhiteSpace($version)) { "0.0.0.0" } else { $version }
    Path = $ManifestPath
  }
}

function Resolve-CodexHome {
  if (-not [string]::IsNullOrWhiteSpace($CodexHome)) {
    return [System.IO.Path]::GetFullPath($CodexHome)
  }

  if (-not [string]::IsNullOrWhiteSpace($env:CODEX_HOME)) {
    return [System.IO.Path]::GetFullPath($env:CODEX_HOME)
  }

  return ""
}

function Test-CodexExecutablePathValue {
  param([Parameter(Mandatory = $true)][string]$ExecutablePath)

  return [System.IO.Path]::IsPathRooted($ExecutablePath) -or
    $ExecutablePath.Contains("\") -or
    $ExecutablePath.Contains("/")
}

function Quote-CodexCommandArgument {
  param([Parameter(Mandatory = $true)][string]$Value)

  return "`"" + $Value.Replace("`"", "\`"") + "`""
}

function Invoke-CodexProcess {
  param(
    [Parameter(Mandatory = $true)][string]$ExecutablePath,
    [Parameter(Mandatory = $true)][string[]]$Arguments,
    [string]$ResolvedCodexHome = ""
  )

  try {
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    if ($ExecutablePath.EndsWith(".cmd", [System.StringComparison]::OrdinalIgnoreCase) -or
        $ExecutablePath.EndsWith(".bat", [System.StringComparison]::OrdinalIgnoreCase)) {
      $startInfo.FileName = if ([string]::IsNullOrWhiteSpace($env:COMSPEC)) { "cmd.exe" } else { $env:COMSPEC }
      $startInfo.Arguments = "/d /c " + ((@($ExecutablePath) + $Arguments | ForEach-Object { Quote-CodexCommandArgument -Value $_ }) -join " ")
    }
    else {
      $startInfo.FileName = $ExecutablePath
      $startInfo.Arguments = ($Arguments | ForEach-Object { Quote-CodexCommandArgument -Value $_ }) -join " "
    }

    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true
    if ([string]::IsNullOrWhiteSpace($ResolvedCodexHome)) {
      $startInfo.Environment.Remove("CODEX_HOME") | Out-Null
    }
    else {
      $startInfo.Environment["CODEX_HOME"] = $ResolvedCodexHome
    }
    $startInfo.Environment["NO_COLOR"] = "1"

    $process = [System.Diagnostics.Process]::Start($startInfo)
    if ($null -eq $process) {
      return [pscustomobject]@{ started = $false; exitCode = -1; stdout = ""; stderr = ""; error = "Codex CLI did not start." }
    }

    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit(10000)) {
      try {
        # Setup runs Windows PowerShell 5.1: Process.Kill(Boolean) is unavailable.
        # Terminate only the process tree created by this probe.
        & "$env:SystemRoot\System32\taskkill.exe" /PID $process.Id /T /F | Out-Null
        if ($LASTEXITCODE -ne 0 -and -not $process.HasExited) { $process.Kill() }
      }
      catch {
      }
      $process.WaitForExit(1000) | Out-Null
      return [pscustomobject]@{
        started = $true
        exitCode = -1
        stdout = if ($stdoutTask.Status -eq 'RanToCompletion') { $stdoutTask.Result } else { "" }
        stderr = if ($stderrTask.Status -eq 'RanToCompletion') { $stderrTask.Result } else { "" }
        error = "Codex CLI timed out."
      }
    }

    [pscustomobject]@{
      started = $true
      exitCode = $process.ExitCode
      stdout = $stdoutTask.Result
      stderr = $stderrTask.Result
      error = ""
    }
  }
  catch {
    [pscustomobject]@{
      started = $false
      exitCode = -1
      stdout = ""
      stderr = ""
      error = $_.Exception.Message
    }
  }
}

function Test-CodexSetup {
  param(
    [string]$ResolvedCodexHome = "",
    [string]$ExecutablePath = "codex"
  )

  $resolvedExecutable = $ExecutablePath
  $codexCommandFound = $false
  $codexCommandPath = ""
  if ([System.IO.Path]::IsPathRooted($ExecutablePath) -or $ExecutablePath.Contains("\") -or $ExecutablePath.Contains("/")) {
    $codexCommandFound = Test-Path -LiteralPath $ExecutablePath
    if ($codexCommandFound) {
      $codexCommandPath = [System.IO.Path]::GetFullPath($ExecutablePath)
      $resolvedExecutable = $codexCommandPath
    }
  }
  else {
    $command = Get-Command $ExecutablePath -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) {
      $codexCommandFound = $true
      $codexCommandPath = $command.Source
      $resolvedExecutable = $command.Source
    }
  }

  $loginStatusOutput = ""
  $loginStatusExitCode = -1
  $loginStatusOk = $false
  if ($codexCommandFound) {
    $loginStatus = Invoke-CodexProcess -ExecutablePath $resolvedExecutable -Arguments @("login", "status") -ResolvedCodexHome $ResolvedCodexHome
    $loginStatusOutput = (@($loginStatus.stdout, $loginStatus.stderr, $loginStatus.error) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine
    $loginStatusOutput = $loginStatusOutput.Trim()
    $loginStatusExitCode = [int]$loginStatus.exitCode
    $loginStatusOk = [bool]$loginStatus.started -and ($loginStatusExitCode -eq 0) -and ($loginStatusOutput -match "(?i)logged in") -and ($loginStatusOutput -match "(?i)chatgpt")
  }

  if ($loginStatusOutput.Length -gt 1200) {
    $loginStatusOutput = $loginStatusOutput.Substring(0, 1200)
  }

  $ok = $codexCommandFound -and $loginStatusOk
  $accessDenied = $codexCommandFound -and ($loginStatusOutput -match "(?i)access is denied")
  $message = if (-not $codexCommandFound) {
    "Maily's AI runtime was not found. Obtain the complete Maily installer package; no separate Codex installation is required."
  }
  elseif ($accessDenied) {
    "Windows denied execution of the AI runtime. Obtain a complete Maily package and check your organization's application policy."
  }
  elseif (-not $loginStatusOk) {
    "ChatGPT sign-in is not complete. Open MailyOutlookAddinSetup.exe and click Install for guided browser sign-in."
  }
  else {
    "Codex CLI is ready and signed in with ChatGPT."
  }

  [pscustomobject]@{
    ok = $ok
    codexHome = $ResolvedCodexHome
    codexExecutablePath = $ExecutablePath
    codexCommandPath = $codexCommandPath
    codexCommandFound = $codexCommandFound
    loginStatusOk = $loginStatusOk
    loginStatusExitCode = $loginStatusExitCode
    loginStatusOutput = if ($loginStatusOk) { "ChatGPT sign-in verified." } else { "ChatGPT sign-in unavailable." }
    accessDenied = $accessDenied
    message = $message
  }
}

function Add-CodexCandidate {
  param(
    [System.Collections.Generic.List[object]]$Candidates,
    [string]$Path,
    [string]$Source
  )

  if ($null -eq $Candidates) {
    return
  }

  if ([string]::IsNullOrWhiteSpace($Path)) {
    return
  }

  # The bundled runtime is a verified native executable.
  if ([IO.Path]::GetExtension($Path) -in @('.cmd', '.bat')) { return }

  $Candidates.Add([pscustomobject]@{
    path = $Path
    source = $Source
  }) | Out-Null
}

function Get-CodexCliCandidates {
  $bundledRuntime = Join-Path $payloadRoot "runtime\codex\codex.exe"
  if (Test-Path -LiteralPath $bundledRuntime -PathType Leaf) {
    return @([pscustomobject]@{ path = $bundledRuntime; source = "bundled" })
  }
  $candidates = New-Object System.Collections.Generic.List[object]
  if (Test-CodexExecutablePathValue -ExecutablePath $CodexExecutablePath) {
    Add-CodexCandidate -Candidates $candidates -Path $CodexExecutablePath -Source "explicit"
  }
  else {
    foreach ($name in @($CodexExecutablePath, "codex", "codex.exe", "codex.cmd")) {
      if ([string]::IsNullOrWhiteSpace($name)) {
        continue
      }

      Get-Command $name -All -ErrorAction SilentlyContinue | ForEach-Object {
        Add-CodexCandidate -Candidates $candidates -Path $_.Source -Source "path:$name"
      }
    }
  }

  $knownPaths = @(
    (Join-Path $env:APPDATA "npm\codex.cmd"),
    (Join-Path $env:APPDATA "npm\codex.exe"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\codex.exe"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\codex"),
    (Join-Path $env:LOCALAPPDATA "OpenAI\Codex\bin\codex.exe"),
    (Join-Path $env:LOCALAPPDATA "OpenAI\Codex\bin\codex"),
    (Join-Path $env:LOCALAPPDATA "Programs\Codex\codex.exe"),
    (Join-Path $env:LOCALAPPDATA "Programs\OpenAI Codex\codex.exe"),
    (Join-Path $env:ProgramFiles "OpenAI\Codex\codex.exe")
  )

  foreach ($path in $knownPaths) {
    Add-CodexCandidate -Candidates $candidates -Path $path -Source "known-path"
  }

  $openAiCodexBin = Join-Path $env:LOCALAPPDATA "OpenAI\Codex\bin"
  if (Test-Path -LiteralPath $openAiCodexBin) {
    Get-ChildItem -LiteralPath $openAiCodexBin -Recurse -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -in @("codex.exe", "codex.cmd", "codex.bat", "codex") } |
      Sort-Object LastWriteTimeUtc -Descending |
      ForEach-Object {
        Add-CodexCandidate -Candidates $candidates -Path $_.FullName -Source "openai-codex-bin"
      }
  }

  $seen = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
  $unique = New-Object System.Collections.Generic.List[object]
  foreach ($candidate in $candidates) {
    $candidatePath = [string]$candidate.path
    if ([string]::IsNullOrWhiteSpace($candidatePath) -or -not $seen.Add($candidatePath)) {
      continue
    }

    $unique.Add($candidate) | Out-Null
  }

  return $unique
}

function Resolve-RunnableCodexCli {
  param([string]$ResolvedCodexHome = "")

  $attempts = New-Object System.Collections.Generic.List[object]
  foreach ($candidate in Get-CodexCliCandidates) {
    $path = [string]$candidate.path
    $check = Test-CodexSetup -ResolvedCodexHome $ResolvedCodexHome -ExecutablePath $path
    $attempts.Add([pscustomobject]@{
      path = $path
      source = [string]$candidate.source
      ok = [bool]$check.ok
      commandFound = [bool]$check.codexCommandFound
      commandPath = [string]$check.codexCommandPath
      accessDenied = [bool]$check.accessDenied
      loginStatusOk = [bool]$check.loginStatusOk
      loginStatusExitCode = [int]$check.loginStatusExitCode
      message = [string]$check.message
    }) | Out-Null

    if ($check.ok) {
      $check | Add-Member -NotePropertyName attemptedCandidates -NotePropertyValue $attempts
      return $check
    }
  }

  $last = if ($attempts.Count -gt 0) { $attempts[$attempts.Count - 1] } else { $null }
  $message = if ($attempts.Count -eq 0) {
    "Maily's AI runtime is missing. Obtain the complete Maily installer package."
  }
  elseif (@($attempts | Where-Object { $_.accessDenied }).Count -gt 0) {
    "The AI runtime could not start. Obtain a complete Maily package and check your organization's application policy."
  }
  else {
    "ChatGPT sign-in is not complete. Open MailyOutlookAddinSetup.exe and click Install for guided browser sign-in."
  }

  [pscustomobject]@{
    ok = $false
    codexHome = $ResolvedCodexHome
    codexExecutablePath = $CodexExecutablePath
    codexCommandPath = if ($null -eq $last) { "" } else { [string]$last.commandPath }
    codexCommandFound = $attempts.Count -gt 0
    loginStatusOk = $false
    loginStatusExitCode = if ($null -eq $last) { -1 } else { [int]$last.loginStatusExitCode }
    loginStatusOutput = ""
    accessDenied = @($attempts | Where-Object { $_.accessDenied }).Count -gt 0
    attemptedCandidates = $attempts
    message = $message
  }
}

function Test-PathListContainsDirectory {
  param(
    [string]$PathList,
    [Parameter(Mandatory = $true)][string]$DirectoryPath
  )

  if ([string]::IsNullOrWhiteSpace($PathList) -or [string]::IsNullOrWhiteSpace($DirectoryPath)) {
    return $false
  }

  $target = [System.IO.Path]::GetFullPath($DirectoryPath).TrimEnd('\')
  foreach ($entry in $PathList.Split([System.IO.Path]::PathSeparator, [System.StringSplitOptions]::RemoveEmptyEntries)) {
    try {
      $entryPath = [System.IO.Path]::GetFullPath($entry.Trim()).TrimEnd('\')
      if ($entryPath.Equals($target, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
      }
    }
    catch {
    }
  }

  return $false
}

function Ensure-CodexCliDirectoryOnUserPath {
  param([Parameter(Mandatory = $true)][string]$ResolvedExecutablePath)

  if ([string]::IsNullOrWhiteSpace($ResolvedExecutablePath) -or -not (Test-CodexExecutablePathValue -ExecutablePath $ResolvedExecutablePath)) {
    return [pscustomobject]@{
      changed = $false
      skipped = $true
      reason = "Resolved Codex CLI is a command name, not a filesystem path."
      directory = ""
    }
  }

  $directory = Split-Path -Parent ([System.IO.Path]::GetFullPath($ResolvedExecutablePath))
  if ([string]::IsNullOrWhiteSpace($directory)) {
    return [pscustomobject]@{
      changed = $false
      skipped = $true
      reason = "Resolved Codex CLI has no parent directory."
      directory = ""
    }
  }

  $processPath = [Environment]::GetEnvironmentVariable("Path", "Process")
  $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
  $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
  if ((Test-PathListContainsDirectory -PathList $processPath -DirectoryPath $directory) -or
      (Test-PathListContainsDirectory -PathList $userPath -DirectoryPath $directory) -or
      (Test-PathListContainsDirectory -PathList $machinePath -DirectoryPath $directory)) {
    return [pscustomobject]@{
      changed = $false
      skipped = $true
      reason = "Codex CLI directory is already on PATH."
      directory = $directory
    }
  }

  $newUserPath = if ([string]::IsNullOrWhiteSpace($userPath)) {
    $directory
  }
  else {
    $userPath.TrimEnd(';') + ";" + $directory
  }
  [Environment]::SetEnvironmentVariable("Path", $newUserPath, "User")
  $env:Path = if ([string]::IsNullOrWhiteSpace($processPath)) {
    $directory
  }
  else {
    $processPath.TrimEnd(';') + ";" + $directory
  }

  [pscustomobject]@{
    changed = $true
    skipped = $false
    reason = "Added Codex CLI directory to the current-user PATH."
    directory = $directory
  }
}

function Resolve-OutlookExe {
  $registryPaths = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
  )

  foreach ($path in $registryPaths) {
    try {
      $item = Get-Item -Path $path -ErrorAction Stop
      $value = [string]$item.GetValue("")
      if (-not [string]::IsNullOrWhiteSpace($value) -and (Test-Path -LiteralPath $value)) {
        return $value
      }
    }
    catch {
    }
  }

  $command = Get-Command OUTLOOK.EXE -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($command) {
    return $command.Source
  }

  return ""
}

function Stop-MailyProcesses {
  $stopped = New-Object System.Collections.Generic.List[object]
  $ownedExecutables = @(
    (Join-Path $installRootPath 'app\Maily.OutlookAddin.Supervisor\Maily.OutlookAddin.Supervisor.exe'),
    (Join-Path $installRootPath 'app\Maily.OutlookAddin.Web\Maily.OutlookAddin.Web.exe'),
    (Join-Path $installRootPath 'runtime\codex\codex.exe')
  )
  Get-Process -Name Maily.OutlookAddin.Supervisor,Maily.OutlookAddin.Web,codex -ErrorAction SilentlyContinue | Where-Object {
    try { $ownedExecutables -contains $_.Path } catch { $false }
  } | Sort-Object @{ Expression = { if ($_.ProcessName -eq 'Maily.OutlookAddin.Supervisor') { 0 } elseif ($_.ProcessName -eq 'Maily.OutlookAddin.Web') { 1 } else { 2 } } } | ForEach-Object {
    try {
      $stopped.Add([pscustomobject]@{ processName = $_.ProcessName; id = $_.Id }) | Out-Null
      Stop-Process -Id $_.Id -Force -ErrorAction Stop
    }
    catch { throw "Could not stop installed Maily process $($_.Exception.Message)" }
  }

  Start-Sleep -Milliseconds 500
  return $stopped
}

function Close-OutlookForSetup {
  $helper = Join-Path $PSScriptRoot 'OutlookShutdown.ps1'
  if (-not (Test-Path -LiteralPath $helper -PathType Leaf)) { throw 'The Outlook shutdown helper is missing. Extract the complete Maily package and retry.' }
  . $helper
  Invoke-MailyOutlookShutdown -HelperPath $helper
}
function Copy-PayloadDirectory {
  param([Parameter(Mandatory = $true)][string]$Name)

  $source = Join-Path $payloadRoot $Name
  $target = Join-Path $installRootPath $Name
  if (-not (Test-Path -LiteralPath $source)) {
    throw "Package payload directory missing: $source."
  }

  Assert-ChildPath -Parent $installRootPath -Child $target
  if (Test-Path -LiteralPath $target) {
    Remove-Item -LiteralPath $target -Recurse -Force
  }

  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $target) | Out-Null
  Copy-Item -LiteralPath $source -Destination $target -Recurse -Force
}

function Get-MailyPlainFiles {
  param([Parameter(Mandatory = $true)][string]$Root)
  $rootItem = Get-Item -LiteralPath $Root -Force
  if (($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Reparse points are not permitted: $Root" }
  $pending = New-Object 'System.Collections.Generic.Queue[string]'
  $pending.Enqueue($rootItem.FullName)
  while ($pending.Count -gt 0) {
    foreach ($item in Get-ChildItem -LiteralPath $pending.Dequeue() -Force -ErrorAction Stop) {
      Assert-ChildPath -Parent $Root -Child $item.FullName
      if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Reparse points are not permitted: $($item.FullName)" }
      if ($item.PSIsContainer) { $pending.Enqueue($item.FullName) } else { $item }
    }
  }
}

function Test-MailyPackageIntegrity {
  $inventoryPath = Join-Path $packageRoot 'SHA256SUMS.json'
  if (-not (Test-Path -LiteralPath $inventoryPath -PathType Leaf)) { throw 'Package integrity inventory is missing. Download or build the complete installer.' }
  # Windows PowerShell 5.1 emits a JSON array as one pipeline object. Wrapping
  # the pipeline in @() nests it and concatenates all paths during validation.
  $inventory = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($inventoryPath))
  if ($null -eq $inventory -or @($inventory).Count -eq 0) { throw 'Package integrity inventory is empty.' }
  $listed = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  $plainFiles = @(Get-MailyPlainFiles -Root $packageRoot)
  foreach ($entry in $inventory) {
    if ($entry -is [Array] -or $entry.path -isnot [string] -or $entry.sha256 -isnot [string]) { throw 'Invalid package inventory entry.' }
    $relative = [string]$entry.path
    if ([IO.Path]::IsPathRooted($relative) -or [string]::IsNullOrWhiteSpace($relative)) { throw 'Invalid package inventory path.' }
    $file = [IO.Path]::GetFullPath((Join-Path $packageRoot $relative))
    Assert-ChildPath -Parent $packageRoot -Child $file
    if (-not $listed.Add($file)) { throw "Duplicate package inventory entry: $relative" }
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Package file missing: $relative" }
    if ((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash -ne [string]$entry.sha256) { throw "Package file failed integrity verification: $relative" }
  }
  foreach ($file in $plainFiles) {
    if ($file.FullName -eq $inventoryPath) { continue }
    if (-not $listed.Contains($file.FullName)) { throw "Package inventory omitted a file: $($file.Name)" }
  }
  if (Test-Path -LiteralPath (Join-Path $payloadRoot 'app/runtime-file-map.json')) {
    . (Join-Path $PSScriptRoot 'MailyPackageRuntime.ps1')
    Test-MailyPackageRuntime -AppRoot (Join-Path $payloadRoot 'app') | Out-Null
  }
  return @{ verifiedFiles = @($inventory).Count }
}

function Get-MailyManagedNames {
  # User profiles and logs are never removed as part of replacing an application version.
  return @('app', 'runtime', 'addin', 'scripts', 'smoke_data', 'setup', 'MailyOutlookAddinSetup.exe', 'package-mode.txt', 'uninstall.ps1', 'package-summary.json')
}

function Get-MailyPathAttributes {
  param([string]$Path)
  if ([IO.Directory]::Exists($Path) -or [IO.File]::Exists($Path)) { return [IO.File]::GetAttributes($Path) }
  return $null
}

function Assert-MailyUnredirectedPath {
  param([Parameter(Mandatory = $true)][string]$Path)
  $current = [IO.Path]::GetFullPath($Path)
  while (-not [string]::IsNullOrWhiteSpace($current)) {
    $attributes = Get-MailyPathAttributes -Path $current
    if ($null -ne $attributes -and ($attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
      throw "Refusing to operate through redirected installed path: $current"
    }
    $current = [IO.Path]::GetDirectoryName($current)
  }
}
function Backup-MailyInstallation {
  Assert-MailyUnredirectedPath -Path (Join-Path $installRootPath 'data\upgrade-backups')
  $previousManifest = Join-Path $installRootPath 'addin\manifest.xml'
  $existing = Test-Path -LiteralPath $previousManifest -PathType Leaf
  $previousVersion = if ($existing) { (Resolve-MailyManifest -ManifestPath $previousManifest).Version } else { '' }
  $backupRoot = Join-Path $installRootPath ("data\upgrade-backups\" + $runId + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
  Assert-ChildPath -Parent $installRootPath -Child $backupRoot
  New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
  $existingNames = New-Object System.Collections.Generic.List[string]
  foreach ($name in @((Get-MailyManagedNames) + @('certs', 'data\addin', 'data\install-state.json'))) {
    $source = Join-Path $installRootPath $name
    Assert-MailyUnredirectedPath -Path $source
    if (-not (Test-Path -LiteralPath $source)) { continue }
    $sourceItem = Get-Item -LiteralPath $source -Force
    if (($sourceItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Refusing to back up redirected installed path: $source" }
    if ($sourceItem.PSIsContainer) { Get-MailyPlainFiles -Root $source | Out-Null }
    $destination = Join-Path $backupRoot $name
    New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force | Out-Null
    Copy-Item -LiteralPath $source -Destination $destination -Recurse -Force
    $existingNames.Add($name) | Out-Null
  }
  $snapshot = [pscustomobject]@{ exists = $existing; previousVersion = $previousVersion; backupRoot = $backupRoot; names = @($existingNames); transactionId = [Guid]::NewGuid().ToString("N"); installRoot = $installRootPath }
  $snapshot | ConvertTo-Json -Depth 5 | Set-MailyUtf8Content -LiteralPath (Join-Path $backupRoot 'upgrade-snapshot.json')
  return $snapshot
}

function Restore-MailyInstallationFiles {
  param([Parameter(Mandatory = $true)][object]$Snapshot)
  Assert-ChildPath -Parent $installRootPath -Child $Snapshot.backupRoot
  Assert-MailyUnredirectedPath -Path $Snapshot.backupRoot
  $markerPath = Join-Path $Snapshot.backupRoot 'upgrade-snapshot.json'
  if (-not (Test-Path -LiteralPath $markerPath -PathType Leaf)) { throw 'Owned installation transaction marker is missing.' }
  $marker = [IO.File]::ReadAllText($markerPath) | ConvertFrom-Json
  if ([string]::IsNullOrWhiteSpace([string]$Snapshot.transactionId) -or $marker.transactionId -ne $Snapshot.transactionId -or $marker.installRoot -ne $installRootPath -or $marker.backupRoot -ne $Snapshot.backupRoot) { throw 'Installation transaction marker does not match this rollback.' }
  $quarantineRoot = Join-Path $Snapshot.backupRoot ('failed-attempt-' + [Guid]::NewGuid().ToString('N'))
  foreach ($name in @((Get-MailyManagedNames) + @('certs', 'data\addin', 'data\install-state.json'))) {
    $target = Join-Path $installRootPath $name
    Assert-ChildPath -Parent $installRootPath -Child $target
    Assert-MailyUnredirectedPath -Path $target
    if (Test-Path -LiteralPath $target) {
      $quarantined = Join-Path $quarantineRoot $name
      Assert-MailyUnredirectedPath -Path $quarantined
      New-Item -ItemType Directory -Path (Split-Path -Parent $quarantined) -Force | Out-Null
      Move-Item -LiteralPath $target -Destination $quarantined
    }
    if ($Snapshot.names -contains $name) {
      $source = Join-Path $Snapshot.backupRoot $name
      Assert-MailyUnredirectedPath -Path $source
      if ((Get-Item -LiteralPath $source).PSIsContainer) { Get-MailyPlainFiles -Root $source | Out-Null }
      New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
      Copy-Item -LiteralPath $source -Destination $target -Recurse -Force
    }
  }
}

function Restore-MailyFailedInstallation {
  param([Parameter(Mandatory = $true)][object]$Snapshot)
  Stop-MailyProcesses | Out-Null
  if (-not $Snapshot.exists -and $script:installationPayloadReady) {
    # Registration may already reference these complete files. Keep the owned
    # payload so retry/uninstall remains available; the installer still fails.
    return 'retained-for-retry'
  }
  Restore-MailyInstallationFiles -Snapshot $Snapshot
  if ($Snapshot.exists) {
    $restoredManifest = Resolve-MailyManifest -ManifestPath (Join-Path $installRootPath 'addin\manifest.xml')
    if (-not $SkipOutlookRegistration) { Register-MailyOutlookAddin -Manifest $restoredManifest | Out-Null }
    Register-UninstallEntry -Manifest $restoredManifest | Out-Null
    if (-not $isHostedPackage -and -not $SkipStartupRegistration) { Register-MailyStartup | Out-Null }
  }
  return 'restored'
}
function Update-LegacyMailyProfiles {
  param([Parameter(Mandatory = $true)][string]$SettingsPath)
  if (-not (Test-Path -LiteralPath $SettingsPath -PathType Leaf)) { return @{ changedProfiles = 0 } }
  $settings = [IO.File]::ReadAllText($SettingsPath) | ConvertFrom-Json
  $changed = 0
  if ($null -ne $settings.items) {
    foreach ($entry in $settings.items.PSObject.Properties) {
      if ($null -ne $entry.Value.settings -and [string]$entry.Value.settings.model -eq 'gpt-5.4-mini') {
        $entry.Value.settings.model = ''
        $changed++
      }
    }
  }
  if ($changed -eq 0) { return @{ changedProfiles = 0 } }
  $backup = $SettingsPath + '.pre-1.3-' + $runId + '.bak'
  Copy-Item -LiteralPath $SettingsPath -Destination $backup -Force
  $temporary = $SettingsPath + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
  try {
    $settings | ConvertTo-Json -Depth 100 | Set-MailyUtf8Content -LiteralPath $temporary
    # Replace only after complete serialization, retaining the exact original file alongside it.
    Move-Item -LiteralPath $temporary -Destination $SettingsPath -Force
  }
  finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
  return @{ changedProfiles = $changed; backupPath = $backup; model = '' }
}

function New-RandomPassword {
  $bytes = New-Object byte[] 32
  $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
  try {
    $rng.GetBytes($bytes)
  }
  finally {
    $rng.Dispose()
  }

  [Convert]::ToBase64String($bytes)
}

function Remove-StaleMailyPrivateCertificates {
  param([Parameter(Mandatory = $true)][string]$KeptThumbprint)

  $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
    "My",
    [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
  try {
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    foreach ($cert in @($store.Certificates | Where-Object {
      $_.FriendlyName -eq "Maily Outlook Add-in Localhost" -and
      $_.Thumbprint -ne $KeptThumbprint
    })) {
      $store.Remove($cert)
    }
  }
  finally {
    $store.Close()
  }
}

function Get-MailyTrustedLocalhostThumbprints {
  $thumbprints = New-Object System.Collections.Generic.List[string]
  foreach ($storeName in @("TrustedPeople", "Root")) {
    Get-ChildItem "Cert:\CurrentUser\$storeName" -ErrorAction SilentlyContinue |
      Where-Object { $_.Subject -eq "CN=localhost" -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
      ForEach-Object { $thumbprints.Add($_.Thumbprint) | Out-Null }
  }

  return @($thumbprints | Select-Object -Unique)
}

function Add-MailyTrustedLocalhostCertificate {
  param(
    [Parameter(Mandatory = $true)][string]$CertificatePath,
    [Parameter(Mandatory = $true)][string]$FriendlyName
  )

  $publicCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
  try {
    $publicCertificate.FriendlyName = $FriendlyName
  }
  catch {
  }
  $trustedPeopleStore = New-Object System.Security.Cryptography.X509Certificates.X509Store(
    "TrustedPeople",
    [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
  try {
    $trustedPeopleStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $existing = $trustedPeopleStore.Certificates.Find(
      [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,
      $publicCertificate.Thumbprint,
      $false)
    if ($existing.Count -eq 0) {
      $trustedPeopleStore.Add($publicCertificate)
    }
  }
  finally {
    $trustedPeopleStore.Close()
  }

  # Adding to CurrentUser\Root can show a Windows security warning. Keep setup
  # prompt-free by trusting the exact localhost leaf in CurrentUser\TrustedPeople.
  $publicCertificate.Dispose()
}

function Add-MailyTrustedRootCertificate {
  param(
    [Parameter(Mandatory = $true)][string]$CertificatePath,
    [Parameter(Mandatory = $true)][string]$FriendlyName
  )

  $publicCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
  try {
    $publicCertificate.FriendlyName = $FriendlyName
  }
  catch {
  }

  $rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store(
    "Root",
    [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
  try {
    $rootStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
    $existingRoot = $rootStore.Certificates.Find(
      [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint,
      $publicCertificate.Thumbprint,
      $false)
    if ($existingRoot.Count -eq 0) {
      $certutil = Join-Path $env:WINDIR "System32\certutil.exe"
      $certutilOutput = & $certutil -user -f -addstore Root $CertificatePath 2>&1
      if ($LASTEXITCODE -ne 0) {
        throw "certutil failed to trust the localhost certificate. $certutilOutput"
      }
    }
  }
  finally {
    $rootStore.Close()
    $publicCertificate.Dispose()
  }
}

function Ensure-LocalhostCertificate {
  param([Parameter(Mandatory = $true)][string]$CertificateDirectory)

  New-Item -ItemType Directory -Force -Path $CertificateDirectory | Out-Null
  $friendlyName = "Maily Outlook Add-in Localhost"
  $trustedThumbprints = Get-MailyTrustedLocalhostThumbprints
  $cert = Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue |
    Where-Object {
      $_.Subject -eq "CN=localhost" -and
      $_.HasPrivateKey -and
      $_.NotAfter -gt (Get-Date).AddDays(30) -and
      $trustedThumbprints -contains $_.Thumbprint
    } |
    Sort-Object `
      @{ Expression = { if ($_.FriendlyName -eq $friendlyName) { 1 } else { 0 } }; Descending = $true },
      NotAfter -Descending |
    Select-Object -First 1

  $reusedTrustedCertificate = $null -ne $cert

  if (-not $cert) {
    $cert = Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue |
      Where-Object {
        $_.FriendlyName -eq $friendlyName -and
        $_.HasPrivateKey -and
        $_.NotAfter -gt (Get-Date).AddDays(30)
      } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1
  }

  if (-not $cert) {
    $cert = New-SelfSignedCertificate `
      -DnsName "localhost" `
      -CertStoreLocation "Cert:\CurrentUser\My" `
      -FriendlyName $friendlyName `
      -KeyExportPolicy Exportable `
      -KeyLength 2048 `
      -KeyAlgorithm RSA `
      -HashAlgorithm SHA256 `
      -NotAfter (Get-Date).AddYears(2)
  }

  $cerPath = Join-Path $CertificateDirectory "localhost.cer"
  $pfxPath = Join-Path $CertificateDirectory "localhost.pfx"
  Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null

  $trustedPeople = Get-ChildItem Cert:\CurrentUser\TrustedPeople -ErrorAction SilentlyContinue |
    Where-Object { $_.Thumbprint -eq $cert.Thumbprint } |
    Select-Object -First 1
  if (-not $trustedPeople) {
    Add-MailyTrustedLocalhostCertificate -CertificatePath $cerPath -FriendlyName $friendlyName
  }

  $trustedRoot = Get-ChildItem Cert:\CurrentUser\Root -ErrorAction SilentlyContinue |
    Where-Object { $_.Thumbprint -eq $cert.Thumbprint } |
    Select-Object -First 1
  if (-not $trustedRoot) {
    if (-not $TrustLocalhostRoot) {
      throw "Maily needs Windows to trust its localhost certificate before Outlook can open the add-in over HTTPS. Run the interactive setup UI and approve the Windows certificate trust prompt."
    }

    Add-MailyTrustedRootCertificate -CertificatePath $cerPath -FriendlyName $friendlyName
  }

  Remove-StaleMailyPrivateCertificates -KeptThumbprint $cert.Thumbprint

  $password = New-RandomPassword
  $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
  Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $securePassword -Force | Out-Null

  [pscustomobject]@{
    thumbprint = $cert.Thumbprint
    pfxPath = $pfxPath
    certificatePath = $cerPath
    password = $password
    notAfter = $cert.NotAfter.ToUniversalTime().ToString("o")
    reusedTrustedCertificate = $reusedTrustedCertificate
  }
}

function Write-InstalledConfiguration {
  param(
    [string]$ResolvedCodexHome = "",
    [Parameter(Mandatory = $true)][object]$Certificate
  )

  $webDir = Join-Path $installRootPath "app\Maily.OutlookAddin.Web"
  $supervisorDir = Join-Path $installRootPath "app\Maily.OutlookAddin.Supervisor"
  $hostExe = Join-Path $webDir "Maily.OutlookAddin.Web.exe"
  $supervisorConfigPath = Join-Path $supervisorDir "appsettings.json"
  $webConfigPath = Join-Path $webDir "appsettings.Production.json"
  $logsDir = Join-Path $installRootPath "logs"
  $dataDir = Join-Path $installRootPath "data\addin"

  New-Item -ItemType Directory -Force -Path $logsDir,$dataDir | Out-Null

  @{
    MailySupervisor = @{
      ServiceName = "MailyOutlookAddinSupervisor"
      HostServiceName = "MailyOutlookAddinHost"
      HealthUrl = $healthUrl
      PollSeconds = 5
      ActivePollSeconds = 5
      IdlePollSeconds = 3
      HealthCheckSeconds = 10
      StopAfterOutlookExitSeconds = 10
      StartupTimeoutSeconds = 25
      LogRetentionDays = 31
      PreferWindowsService = $false
      StopHostWhenOutlookExits = $true
      ExitWhenOutlookExits = $false
      ExitWhenOutlookNotRunningOnStartSeconds = 0
      RepoRoot = $installRootPath
      HostExecutablePath = $hostExe
      HostWorkingDirectory = $webDir
      LogDirectory = $logsDir
    }
  } | ConvertTo-Json -Depth 8 | Set-MailyUtf8Content -LiteralPath $supervisorConfigPath

  @{
    Logging = @{
      LogLevel = @{
        Default = "Warning"
        "Microsoft.AspNetCore" = "Warning"
      }
    }
    MailyHosting = @{
      ServiceName = "MailyOutlookAddinHost"
      HttpsUrl = "https://localhost:7097"
      HttpUrl = "http://localhost:5289"
      LogDirectory = $logsDir
      CertificatePath = $Certificate.pfxPath
      CertificatePassword = $Certificate.password
      RequireOutlookBridge = $true
      EnableHttp = $false
      EnableRequestLog = $false
      EnableEventLog = $false
      ReadinessCacheSeconds = 30
    }
    MailyGeneration = @{
      RepoRoot = $installRootPath
      PromptsRoot = (Join-Path $installRootPath "smoke_data\prompts")
      SettingsFilePath = (Join-Path $dataDir "addin_profile_settings.json")
      CodexHome = $ResolvedCodexHome
      CodexExecutablePath = $resolvedCodexExecutablePath
      Model = ""
      PromptVersion = "v1"
      BackendApiKey = ""
    }
    AllowedHosts = "*"
  } | ConvertTo-Json -Depth 8 | Set-MailyUtf8Content -LiteralPath $webConfigPath

  [pscustomobject]@{
    webConfigPath = $webConfigPath
    supervisorConfigPath = $supervisorConfigPath
    logsDir = $logsDir
    dataDir = $dataDir
  }
}

function Register-MailyStartup {
  $supervisorDir = Join-Path $installRootPath "app\Maily.OutlookAddin.Supervisor"
  $supervisorExe = Join-Path $supervisorDir "Maily.OutlookAddin.Supervisor.exe"
  if (-not (Test-Path -LiteralPath $supervisorExe)) {
    throw "Supervisor executable missing at $supervisorExe."
  }

  $startupDirectory = [Environment]::GetFolderPath("Startup")
  if ([string]::IsNullOrWhiteSpace($startupDirectory)) {
    throw "Current-user Startup folder could not be resolved."
  }

  New-Item -ItemType Directory -Force -Path $startupDirectory | Out-Null
  $startupShortcutPath = Join-Path $startupDirectory $startupShortcutName
  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($startupShortcutPath)
  $shortcut.TargetPath = $supervisorExe
  $shortcut.WorkingDirectory = $supervisorDir
  $shortcut.WindowStyle = 7
  $shortcut.Description = "Starts the local Maily Outlook add-in host when you sign in."
  $shortcut.Save()

  if (-not (Test-Path -LiteralPath $startupShortcutPath)) {
    throw "Startup shortcut did not persist."
  }

  if (Get-ItemProperty -Path $runKeyPath -Name $runKeyName -ErrorAction SilentlyContinue) {
    Remove-ItemProperty -Path $runKeyPath -Name $runKeyName -ErrorAction SilentlyContinue
  }

  $removedLegacyTask = $false
  $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
  if ($existingTask) {
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    $removedLegacyTask = $true
  }

  Start-Process -FilePath $supervisorExe -WorkingDirectory $supervisorDir -WindowStyle Hidden

  [pscustomobject]@{
    runKey = $runKeyName
    taskName = ""
    supervisorExe = $supervisorExe
    startupShortcut = $startupShortcutPath
    removedLegacyTask = $removedLegacyTask
    mode = "current-user-startup-folder-session-watcher"
  }
}

function Start-MailySupervisor {
  $supervisorDir = Join-Path $installRootPath "app\Maily.OutlookAddin.Supervisor"
  $supervisorExe = Join-Path $supervisorDir "Maily.OutlookAddin.Supervisor.exe"
  if (-not (Test-Path -LiteralPath $supervisorExe)) {
    throw "Supervisor executable missing at $supervisorExe."
  }

  Start-Process -FilePath $supervisorExe -WorkingDirectory $supervisorDir -WindowStyle Hidden

  [pscustomobject]@{
    supervisorExe = $supervisorExe
    persistentStartup = $false
  }
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

function Get-MailyManifestSignature {
  param([Parameter(Mandatory=$true)][string]$Path)
  $settings = New-Object System.Xml.XmlReaderSettings
  $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
  $settings.XmlResolver = $null
  $reader = [System.Xml.XmlReader]::Create($Path, $settings)
  try {
    $document = New-Object System.Xml.XmlDocument
    $document.XmlResolver = $null
    $document.Load($reader)
    $id = $document.SelectSingleNode('/*/*[local-name()="Id"]').InnerText.Trim()
    $version = $document.SelectSingleNode('/*/*[local-name()="Version"]').InnerText.Trim()
    $urls = @($document.SelectNodes('//@DefaultValue') | ForEach-Object {
      if ($_.Value -match '^https?://') {
        $uri = [Uri]$_.Value
        # Local pane HTML is served no-store and references revisioned assets.
        # Updating only its cache tag does not change mailbox registration.
        # Hosted URLs and all other query parameters remain exact comparisons.
        if ($uri.Scheme -eq 'https' -and $uri.IsLoopback -and $uri.Port -eq 7097 -and $uri.AbsolutePath -match '^/addin/[^/]+\.html$') {
          $builder = New-Object System.UriBuilder($uri)
          $builder.Query = (@($uri.Query.TrimStart('?').Split('&') | Where-Object { $_ -and $_ -cnotmatch '^v=' }) -join '&')
          $builder.Uri.AbsoluteUri
        }
        else { $_.Value }
      }
    } | Sort-Object -Unique)
    if (-not $id -or -not $version -or $urls.Count -eq 0) { throw 'Incomplete Outlook manifest.' }
    return ($id.ToLowerInvariant() + "`n" + $version + "`n" + ($urls -join "`n"))
  }
  finally { $reader.Dispose() }
}

function Get-MailyOutlookActivation {
  param([Parameter(Mandatory=$true)][object]$Manifest)
  $expected = Get-MailyManifestSignature -Path $Manifest.Path
  $copies = @(Get-MailyMailboxInstalledAddinCache -AddinId $Manifest.Id)
  # Absence is inconclusive: Outlook can download a mailbox copy again after local cache cleanup.
  if ($copies.Count -eq 0) { return [pscustomobject]@{ status='activation-required'; detail='Setup has not yet confirmed Maily in Outlook. Outlook may still be syncing, or the connection may not have been set up.' } }
  # Unknown ownership must win over stale user evidence, regardless of enumeration order.
  # Cache formats are not a supported tenant-management API: do not infer administrator rights.
  if (@($copies | Where-Object { $_.installedBy -ne 'User' }).Count -gt 0) {
    return [pscustomobject]@{ status='needs-support'; detail='Setup cannot confirm who manages this Outlook connection. It may be managed by your organization. Contact your administrator or Maily support before changing it.' }
  }
  foreach ($copy in $copies) {
    try {
      if ((Get-MailyManifestSignature -Path $copy.manifestPath) -cne $expected) { throw 'Manifest differs.' }
    }
    catch { return [pscustomobject]@{ status='repair-required'; detail='Outlook has an outdated or unreadable Maily mailbox registration. Replacement has not started.' } }
  }
  return [pscustomobject]@{ status='verified'; detail='Outlook cached mailbox registration matches this package identity, version and endpoints. Local frontend cache tags may differ; updated HTML loads the current assets. Open a new draft to confirm the commands are visible.' }
}

function Get-MailyActivationSignal {
  param([Parameter(Mandatory=$true)][ValidateSet('verified','activation-required','repair-required','needs-support')][string]$Status)
  return ('MAILY_ACTIVATION|' + $Status)
}

function Get-MailyOutlookActivationGuidance {
  param([Parameter(Mandatory=$true)][string]$Status, [Parameter(Mandatory=$true)][string]$ManifestPath)
  switch ($Status) {
    'repair-required' { return "An older user-installed Maily connection was found. Open https://aka.ms/olksideload and replace only that Maily entry with this package's file: '$ManifestPath'. Return to setup after Outlook finishes syncing. Your organization may require approval." }
    'activation-required' { return "Open Outlook and let it finish syncing, then return to setup to check the connection. If you already added Maily, do not add it again. Classic Outlook can take up to 24 hours to show a manually added connection. If you have never connected Maily, use setup's guided connection step." }
    'verified' { return 'Open Outlook and let it finish syncing, then return to setup to check the connection.' }
    default { return 'Contact your administrator or Maily support to confirm how this Outlook connection is managed before changing it.' }
  }
}
function Get-MailyMailboxInstalledAddinCache {
  param([Parameter(Mandatory = $true)][string]$AddinId, [string]$CacheRoot = '')

  $cacheMatches = New-Object System.Collections.Generic.List[object]
  $wefRoot = if ($CacheRoot) { $CacheRoot } else { Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef" }
  if (-not (Test-Path -LiteralPath $wefRoot)) {
    return @()
  }

  Get-ChildItem -LiteralPath $wefRoot -Recurse -Force -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -eq "$AddinId.metadata" -or $_.Name -like "$AddinId`_*metadata" } |
    ForEach-Object {
      try {
        $settings = New-Object System.Xml.XmlReaderSettings
        $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $reader = [System.Xml.XmlReader]::Create($_.FullName, $settings)
        try {
          $metadata = New-Object System.Xml.XmlDocument
          $metadata.XmlResolver = $null
          $metadata.Load($reader)
          $installedBy = $metadata.SelectSingleNode('//*[local-name()="InstalledBy"]').InnerText.Trim()
        }
        finally { $reader.Dispose() }
        if ($installedBy -ne 'User') {
          # This installer verifies the user-installed pilot workflow only. Other registration
          # sources remain unverified instead of inventing tenant deployment semantics.
          $cacheMatches.Add([pscustomobject]@{ manifestPath = ''; installedBy = 'Unknown' }) | Out-Null
          return
        }

        $manifestPath = $_.FullName -replace "\.metadata$", ""
        $manifestVersion = ""
        if (Test-Path -LiteralPath $manifestPath) {
          $manifestText = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction SilentlyContinue
          if ($manifestText -match "<Version>([^<]+)</Version>") {
            $manifestVersion = $Matches[1]
          }
        }

        $cacheMatches.Add([pscustomobject]@{
          metadataPath = $_.FullName
          manifestPath = $manifestPath
          manifestVersion = $manifestVersion
          installedBy = "User"
        }) | Out-Null
      }
      catch {
        # A known Maily metadata file that cannot be read is not evidence of absence.
        $cacheMatches.Add([pscustomobject]@{ manifestPath = ''; installedBy = 'Unknown' }) | Out-Null
      }
    }

  return $cacheMatches.ToArray()
}

function Register-MailyOutlookAddin {
  param(
    [Parameter(Mandatory = $true)][object]$Manifest,
    [switch]$SkipCacheCleanup
  )

  New-Item -Path $developerKeyPath -Force | Out-Null
  New-ItemProperty -Path $developerKeyPath -Name $Manifest.Id -PropertyType String -Value $Manifest.Path -Force | Out-Null
  New-ItemProperty -Path $developerKeyPath -Name "RefreshAddins" -PropertyType DWord -Value 1 -Force | Out-Null
  $removedCacheItems = if ($SkipCacheCleanup) { 0 } else { Remove-MailyOutlookCache -AddinId $Manifest.Id }
  $registeredPath = [string](Get-ItemProperty -Path $developerKeyPath -Name $Manifest.Id -ErrorAction Stop).($Manifest.Id)
  if ([IO.Path]::GetFullPath($registeredPath) -ne [IO.Path]::GetFullPath($Manifest.Path)) { throw 'Outlook manifest registration did not match the installed path.' }
  $registeredManifest = Resolve-MailyManifest -ManifestPath $registeredPath
  if ($registeredManifest.Id -ne $Manifest.Id -or $registeredManifest.Version -ne $Manifest.Version) { throw 'Outlook registered manifest version verification failed.' }

  [pscustomobject]@{
    developerKeyPath = $developerKeyPath
    addinId = $Manifest.Id
    manifestPath = $Manifest.Path
    manifestVersion = $registeredManifest.Version
    removedCacheItems = $removedCacheItems
  }
}

function Register-UninstallEntry {
  param([Parameter(Mandatory = $true)][object]$Manifest)

  $setupExe = Join-Path $installRootPath $setupExeName
  if (-not (Test-Path -LiteralPath $setupExe)) {
    throw "Installed GUI setup executable is missing at $setupExe."
  }

  $uninstallCommand = "`"$setupExe`" --uninstall"
  $quietUninstallCommand = "`"$setupExe`" --uninstall --quiet"
  $displayIcon = $setupExe
  $estimatedSizeKb = 0
  try {
    $estimatedSizeKb = [int]((Get-ChildItem -LiteralPath $installRootPath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum / 1KB)
  }
  catch {
    $estimatedSizeKb = 0
  }

  New-Item -Path $uninstallKeyPath -Force | Out-Null
  New-ItemProperty -Path $uninstallKeyPath -Name "DisplayName" -PropertyType String -Value "Maily Outlook Add-in" -Force | Out-Null
  New-ItemProperty -Path $uninstallKeyPath -Name "DisplayVersion" -PropertyType String -Value $Manifest.Version -Force | Out-Null
  New-ItemProperty -Path $uninstallKeyPath -Name "Publisher" -PropertyType String -Value "Maily" -Force | Out-Null
  New-ItemProperty -Path $uninstallKeyPath -Name "InstallLocation" -PropertyType String -Value $installRootPath -Force | Out-Null
  New-ItemProperty -Path $uninstallKeyPath -Name "DisplayIcon" -PropertyType String -Value $displayIcon -Force | Out-Null
  New-ItemProperty -Path $uninstallKeyPath -Name "UninstallString" -PropertyType String -Value $uninstallCommand -Force | Out-Null
  New-ItemProperty -Path $uninstallKeyPath -Name "QuietUninstallString" -PropertyType String -Value $quietUninstallCommand -Force | Out-Null
  New-ItemProperty -Path $uninstallKeyPath -Name "NoModify" -PropertyType DWord -Value 1 -Force | Out-Null
  New-ItemProperty -Path $uninstallKeyPath -Name "NoRepair" -PropertyType DWord -Value 1 -Force | Out-Null
  if ($estimatedSizeKb -gt 0) {
    New-ItemProperty -Path $uninstallKeyPath -Name "EstimatedSize" -PropertyType DWord -Value $estimatedSizeKb -Force | Out-Null
  }

  [pscustomobject]@{
    uninstallKeyPath = $uninstallKeyPath
    uninstallString = $uninstallCommand
    quietUninstallString = $quietUninstallCommand
    setupExe = $setupExe
    estimatedSizeKb = $estimatedSizeKb
  }
}

function Test-OutlookRunning {
  return $null -ne (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
}

function Test-HostReady {
  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri $healthUrl -TimeoutSec 5
    return $response.StatusCode -eq 200
  }
  catch {
    return $false
  }
}

function Wait-HostReady {
  param([int]$TimeoutSeconds = 45)

  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  while ((Get-Date) -lt $deadline) {
    if (Test-HostReady) {
      return $true
    }

    Start-Sleep -Seconds 1
  }

  return $false
}

function Start-HostForVerification {
  $webDir = Join-Path $installRootPath "app\Maily.OutlookAddin.Web"
  $hostExe = Join-Path $webDir "Maily.OutlookAddin.Web.exe"
  if (-not (Test-Path -LiteralPath $hostExe)) {
    throw "Host executable missing at $hostExe."
  }

  if (Test-HostReady) {
    return $false
  }

  Start-Process -FilePath $hostExe -WorkingDirectory $webDir -WindowStyle Hidden
  return $true
}

function Stop-HostIfInstallerStartedIt {
  param([bool]$StartedHost)

  if (-not $StartedHost -or (Test-OutlookRunning)) {
    return
  }

  $installedHost = Join-Path $installRootPath 'app\Maily.OutlookAddin.Web\Maily.OutlookAddin.Web.exe'
  Get-Process -Name Maily.OutlookAddin.Web -ErrorAction SilentlyContinue | Where-Object {
    try { $_.Path -eq $installedHost } catch { $false }
  } | ForEach-Object {
    try {
      Stop-Process -Id $_.Id -Force -ErrorAction Stop
    }
    catch {
    }
  }
}

function Invoke-JsonGet {
  param([Parameter(Mandatory = $true)][string]$Uri)

  $response = Invoke-WebRequest -UseBasicParsing -Uri $Uri -TimeoutSec 30
  return $response.Content | ConvertFrom-Json
}

function Invoke-JsonPost {
  param([Parameter(Mandatory = $true)][string]$Uri)

  $response = Invoke-WebRequest -UseBasicParsing -Uri $Uri -Method Post -TimeoutSec 240
  return $response.Content | ConvertFrom-Json
}

function Write-InstallState {
  param(
    [Parameter(Mandatory = $true)][object]$Manifest,
    [object]$Certificate = $null,
    [string]$ResolvedCodexHome = ""
  )

  $statePath = Join-Path $installRootPath "data\install-state.json"
  $state = [pscustomobject]@{
    installedAtUtc = [DateTime]::UtcNow.ToString("o")
    packageMode = $resolvedPackageMode
    installRoot = $installRootPath
    addinId = $Manifest.Id
    version = $Manifest.Version
    manifestPath = $Manifest.Path
    codexHome = $ResolvedCodexHome
    codexExecutablePath = if ($resolvedPackageMode -eq "Local") { $resolvedCodexExecutablePath } else { "" }
    certificateThumbprint = if ($null -eq $Certificate) { "" } else { $Certificate.thumbprint }
    certificatePath = if ($null -eq $Certificate) { "" } else { $Certificate.certificatePath }
    pfxPath = if ($null -eq $Certificate) { "" } else { $Certificate.pfxPath }
    runKeyName = if ($resolvedPackageMode -eq "Local" -and -not $SkipStartupRegistration) { $runKeyName } else { "" }
    taskName = ""
  }

  $state | ConvertTo-Json -Depth 8 | Set-MailyUtf8Content -LiteralPath $statePath
  return $statePath
}

$resolvedPackageMode = Resolve-PackageMode
$isHostedPackage = $resolvedPackageMode -eq "Hosted"

try {
  Write-InstallerHost "Installing Maily Outlook Add-in ($resolvedPackageMode package) to $installRootPath"

  Invoke-InstallerStep -Name "package-payload" -ScriptBlock {
    if (-not (Test-Path -LiteralPath $payloadRoot)) {
      throw "Package payload was not found at $payloadRoot. Build the installer package first."
    }

    $required = if ($isHostedPackage) {
      @("addin\manifest.xml")
    }
    else {
      @(
        "app\Maily.OutlookAddin.Web\Maily.OutlookAddin.Web.exe",
        "app\Maily.OutlookAddin.Supervisor\Maily.OutlookAddin.Supervisor.exe",
        "runtime\codex\codex.exe",
        "addin\manifest.xml",
        "scripts\outlook-bridge.ps1",
        "smoke_data\prompts\prompt-system.md",
        "smoke_data\prompts\prompt-iterate-on-draft.md",
        "smoke_data\prompts\email-samples.md"
      )
    }
    $missing = @($required | Where-Object { -not (Test-Path -LiteralPath (Join-Path $payloadRoot $_)) })
    if ($missing.Count -gt 0) {
      throw "Package is missing required payload files: $($missing -join ', ')"
    }

    $setupExecutable = Join-Path $packageRoot $setupExeName
    if (-not (Test-Path -LiteralPath $setupExecutable)) {
      throw "Package is missing the GUI setup executable: $setupExecutable"
    }

    @{ root = $payloadRoot; packageMode = $resolvedPackageMode; requiredCount = $required.Count; setupExecutable = $setupExecutable }
  } | Out-Null

  Invoke-InstallerStep -Name "package-integrity" -ScriptBlock { Test-MailyPackageIntegrity } | Out-Null
  $incomingManifest = Invoke-InstallerStep -Name "upgrade-preflight" -ScriptBlock {
    $packageFull = [IO.Path]::GetFullPath($packageRoot).TrimEnd([char[]]'\/')
    $installFull = $installRootPath.TrimEnd([char[]]'\/')
    if ($installFull -eq [IO.Path]::GetPathRoot($installRootPath).TrimEnd([char[]]'\/')) { throw 'A drive root cannot be used as the Maily install directory.' }
    if ($packageFull -eq $installFull -or $packageFull.StartsWith($installFull + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
      throw 'Extract the installer outside the existing Maily installation before upgrading.'
    }
    Assert-MailyUnredirectedPath -Path (Join-Path $installRootPath 'data\upgrade-backups')
    $candidate = Resolve-MailyManifest -ManifestPath (Join-Path $payloadRoot 'addin\manifest.xml')
    $oldManifest = Join-Path $installRootPath 'addin\manifest.xml'
    if (Test-Path -LiteralPath $oldManifest) {
      $previous = Resolve-MailyManifest -ManifestPath $oldManifest
      if ($previous.Id -ne $candidate.Id) { throw 'The target directory contains a different add-in. Choose the existing Maily install directory.' }
      $modeFile = Join-Path $installRootPath 'package-mode.txt'
      if ((Test-Path -LiteralPath $modeFile) -and (Get-Content -LiteralPath $modeFile -Raw).Trim() -ne $resolvedPackageMode) {
        throw 'Changing between hosted and local installations requires removing the previous mode first; settings can be preserved.'
      }
    }
    elseif (@(Get-MailyManagedNames | Where-Object { Test-Path -LiteralPath (Join-Path $installRootPath $_) }).Count -gt 0) {
      throw 'The target contains application files without a recognizable Maily manifest. Choose an empty directory or restore the previous installation before retrying.'
    }
    $uninstaller = Join-Path $packageRoot 'setup\uninstall.ps1'
    if (-not (Test-Path -LiteralPath $uninstaller -PathType Leaf)) { throw 'Package uninstaller is missing.' }
    $candidate
  }

  $activation = Get-MailyOutlookActivation -Manifest $incomingManifest
  Write-Output (Get-MailyActivationSignal -Status $activation.status) | Out-Host
  if ($activation.status -eq 'needs-support') {
    $outlookActivationRequired = $true
    $outlookSupportRequired = $true
    Add-Step -Name 'outlook-activation' -Status warning -Detail $activation.detail
    exit 23
  }
  $recovery = Get-MailyOutlookActivationGuidance -Status $activation.status -ManifestPath $incomingManifest.Path
  if ($activation.status -eq 'repair-required') {
    $outlookActivationRequired = $true
    Add-Step -Name 'outlook-activation' -Status warning -Detail ($activation.detail + ' ' + $recovery)
    throw 'Outlook registration repair is required before replacing the application.'
  }
  if ($VerifyOutlookActivation) {
    $installedPath = Join-Path $installRootPath 'addin/manifest.xml'
    if ($activation.status -ne 'verified' -or -not (Test-Path -LiteralPath $installedPath) -or (Get-MailyManifestSignature $installedPath) -cne (Get-MailyManifestSignature $incomingManifest.Path)) {
      $outlookActivationRequired = $true
      Add-Step -Name 'outlook-activation' -Status warning -Detail ($activation.detail + ' Install this package, then ' + $recovery)
      exit 20
    }
    if (-not (Test-OutlookRunning)) { throw 'Open Outlook after restarting it, then choose Verify Outlook again.' }
    if (-not $isHostedPackage) {
      if (-not (Wait-HostReady -TimeoutSeconds 10)) { throw 'The installed Maily service is not healthy. Retry installation.' }
      if (-not [bool](Invoke-JsonGet $readinessUrl).ready) { throw 'The installed Maily service is not ready. Retry installation.' }
    }
    Add-Step -Name 'outlook-activation' -Status passed -Detail $activation.detail
    exit 0
  }

  $outlookExe = Resolve-OutlookExe
  if ([string]::IsNullOrWhiteSpace($outlookExe)) {
    Add-Step -Name "outlook-installed" -Status "failed" -Detail "Microsoft Outlook desktop was not found. Install Outlook desktop before using the add-in."
    throw "Microsoft Outlook desktop was not found."
  }
  Add-Step -Name "outlook-installed" -Status "passed" -Detail "Outlook desktop found." -Data @{ path = $outlookExe }

  $resolvedCodexHome = ""
  if ($isHostedPackage) {
    Add-Step -Name "codex-setup" -Status "skipped" -Detail "Hosted package does not require local Codex setup."
  }
  else {
    $resolvedCodexHome = Resolve-CodexHome
    $codexCheck = Resolve-RunnableCodexCli -ResolvedCodexHome $resolvedCodexHome
    if (-not $codexCheck.ok) {
      Add-Step -Name "codex-setup" -Status "failed" -Detail $codexCheck.message -Data $codexCheck
      throw $codexCheck.message
    }
    $resolvedCodexExecutablePath = if (-not [string]::IsNullOrWhiteSpace([string]$codexCheck.codexCommandPath)) {
      [string]$codexCheck.codexCommandPath
    }
    else {
      [string]$codexCheck.codexExecutablePath
    }
    Add-Step -Name "codex-setup" -Status "passed" -Detail $codexCheck.message -Data $codexCheck
    Add-Step -Name "codex-path" -Status "skipped" -Detail "Maily uses an explicit runtime path; the user's PATH is unchanged."
  }

  if ($CloseOutlook) {
    Invoke-InstallerStep -Name "outlook-close" -ScriptBlock {
      Close-OutlookForSetup
    } | Out-Null
  }
  elseif (Test-OutlookRunning) {
    Add-Step -Name "outlook-close" -Status "warning" -Detail "Outlook is running. Setup can complete, but Outlook must restart before it sees refreshed add-in metadata."
  }
  else {
    Add-Step -Name "outlook-close" -Status "passed" -Detail "Outlook is not running."
  }

  if ($isHostedPackage) {
    Add-Step -Name "stop-existing-maily" -Status "skipped" -Detail "Hosted package does not manage local Maily processes."
  }
  else {
    Invoke-InstallerStep -Name "stop-existing-maily" -ScriptBlock {
      Stop-MailyProcesses
    } | Out-Null
  }

  $upgradeSnapshot = Invoke-InstallerStep -Name "backup-existing-installation" -ScriptBlock { Backup-MailyInstallation }
  if ($upgradeSnapshot.exists) {
    Add-Step -Name 'upgrade-detected' -Status passed -Detail ("Replacing Maily {0} with {1}; settings and sign-in are preserved." -f $upgradeSnapshot.previousVersion, $incomingManifest.Version)
  }
  $applicationMutationStarted = $true
  Invoke-InstallerStep -Name 'remove-previous-version' -ScriptBlock {
    foreach ($name in Get-MailyManagedNames) {
      $target = Join-Path $installRootPath $name
      Assert-ChildPath -Parent $installRootPath -Child $target
    Assert-MailyUnredirectedPath -Path $target
      if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }
    }
    @{ previousVersion = $upgradeSnapshot.previousVersion; preserved = @('data', 'logs', 'certs', 'ChatGPT authentication') }
  } | Out-Null

  Invoke-InstallerStep -Name "copy-application" -ScriptBlock {
    New-Item -ItemType Directory -Force -Path $installRootPath | Out-Null
    $payloadDirectories = if ($isHostedPackage) {
      @("addin", "scripts")
    }
    else {
      @("app", "addin", "scripts", "smoke_data")
    }
    if ((Test-Path -LiteralPath (Join-Path $payloadRoot 'runtime/atk/node.exe') -PathType Leaf) -or
        (-not $isHostedPackage -and (Test-Path -LiteralPath (Join-Path $payloadRoot "runtime\codex\codex.exe") -PathType Leaf))) {
      $payloadDirectories += "runtime"
    }

    foreach ($name in $payloadDirectories) {
      Copy-PayloadDirectory -Name $name
    }

    if (-not $isHostedPackage -and (Test-Path -LiteralPath (Join-Path $installRootPath 'app/runtime-file-map.json'))) {
      . (Join-Path $PSScriptRoot 'MailyPackageRuntime.ps1')
      Restore-MailyPackageRuntime -AppRoot (Join-Path $installRootPath 'app')
    }

    $uninstallSource = Join-Path $packageRoot "setup\uninstall.ps1"
    if (-not (Test-Path -LiteralPath $uninstallSource)) {
      $uninstallSource = Join-Path $packageRoot "uninstall.ps1"
    }
    if (Test-Path -LiteralPath $uninstallSource) {
      $installedSetupDirectory = Join-Path $installRootPath "setup"
      New-Item -ItemType Directory -Force -Path $installedSetupDirectory | Out-Null
      Copy-Item -LiteralPath $uninstallSource -Destination (Join-Path $installedSetupDirectory "uninstall.ps1") -Force
      Copy-Item -LiteralPath (Join-Path (Split-Path -Parent $uninstallSource) "OutlookShutdown.ps1") -Destination (Join-Path $installedSetupDirectory "OutlookShutdown.ps1") -Force
    }
    else {
      throw "Package uninstaller script is missing at $uninstallSource."
    }

    $setupSource = Join-Path $packageRoot $setupExeName
    if (Test-Path -LiteralPath $setupSource) {
      Copy-Item -LiteralPath $setupSource -Destination (Join-Path $installRootPath $setupExeName) -Force
    }
    else {
      throw "Package GUI setup executable is missing at $setupSource."
    }

    $packageModeSource = Join-Path $packageRoot $packageModeFileName
    if (Test-Path -LiteralPath $packageModeSource) {
      Copy-Item -LiteralPath $packageModeSource -Destination (Join-Path $installRootPath $packageModeFileName) -Force
    }
    else {
      Set-Content -LiteralPath (Join-Path $installRootPath $packageModeFileName) -Value $resolvedPackageMode -Encoding ASCII
    }

    @{ installRoot = $installRootPath; copiedPayloadDirectories = $payloadDirectories }
  } | Out-Null

  if (-not $isHostedPackage -and (Test-Path -LiteralPath (Join-Path $payloadRoot "runtime\codex\codex.exe") -PathType Leaf)) {
    $resolvedCodexExecutablePath = Join-Path $installRootPath "runtime\codex\codex.exe"
    if (-not (Test-Path -LiteralPath $resolvedCodexExecutablePath -PathType Leaf)) {
      throw "The bundled AI runtime was not copied successfully."
    }
  }

  Invoke-InstallerStep -Name 'migrate-legacy-settings' -ScriptBlock {
    Update-LegacyMailyProfiles -SettingsPath (Join-Path $installRootPath 'data\addin\addin_profile_settings.json')
  } | Out-Null

  $manifest = Invoke-InstallerStep -Name "manifest" -ScriptBlock {
    $manifestPath = Join-Path $installRootPath "addin\manifest.xml"
    if ($isHostedPackage) {
      $manifestText = Get-Content -LiteralPath $manifestPath -Raw
      if ($manifestText -match "https?://localhost(?::\d+)?") {
        throw "Hosted package manifest still contains localhost URLs."
      }
    }

    Resolve-MailyManifest -ManifestPath $manifestPath
  }

  $certificate = $null
  if ($isHostedPackage) {
    Add-Step -Name "localhost-certificate" -Status "skipped" -Detail "Hosted package does not create or trust localhost certificates."
    Add-Step -Name "write-runtime-config" -Status "skipped" -Detail "Hosted package does not install a local web host."
  }
  else {
    try {
      $certificate = Ensure-LocalhostCertificate -CertificateDirectory (Join-Path $installRootPath "certs")
      Add-Step -Name "localhost-certificate" -Status "passed" -Detail "Completed." -Data ([pscustomobject]@{
        thumbprint = $certificate.thumbprint
        pfxPath = $certificate.pfxPath
        certificatePath = $certificate.certificatePath
        passwordConfigured = -not [string]::IsNullOrWhiteSpace([string]$certificate.password)
        notAfter = $certificate.notAfter
        reusedTrustedCertificate = $certificate.reusedTrustedCertificate
      })
    }
    catch {
      Add-Step -Name "localhost-certificate" -Status "failed" -Detail $_.Exception.Message
      throw
    }

    Invoke-InstallerStep -Name "write-runtime-config" -ScriptBlock {
      Write-InstalledConfiguration -ResolvedCodexHome $resolvedCodexHome -Certificate $certificate
    } | Out-Null
  }

  Invoke-InstallerStep -Name "install-state" -ScriptBlock {
    Write-InstallState -Manifest $manifest -Certificate $certificate -ResolvedCodexHome $resolvedCodexHome
  } | Out-Null
  $script:installationPayloadReady = $true


  if ($isHostedPackage) {
    Add-Step -Name "startup-registration" -Status "skipped" -Detail "Hosted package does not install supervisor startup, scheduled task, Run key, or Windows service."
  }
  elseif ($SkipStartupRegistration) {
    Add-Step -Name "supervisor-launch" -Status "skipped" -Detail "Skipped by request."
  }
  else {
    Invoke-InstallerStep -Name "startup-registration" -ScriptBlock {
      Register-MailyStartup
    } | Out-Null
  }

  if ($SkipOutlookRegistration) {
    Add-Step -Name "outlook-registration" -Status "skipped" -Detail "Skipped by request."
  }
  else {
    Invoke-InstallerStep -Name "outlook-registration" -ScriptBlock {
      Register-MailyOutlookAddin -Manifest $manifest
    } | Out-Null
  }

  Invoke-InstallerStep -Name "windows-uninstall-entry" -ScriptBlock {
    Register-UninstallEntry -Manifest $manifest
  } | Out-Null

  if ($isHostedPackage) {
    Add-Step -Name "backend-verification" -Status "skipped" -Detail "Hosted package does not install or verify a local backend."
    Add-Step -Name "style-refresh" -Status "skipped" -Detail "Hosted package does not use the local Outlook bridge for install-time style refresh."
  }
  elseif ($SkipVerification) {
    Add-Step -Name "backend-verification" -Status "skipped" -Detail "Skipped by request."
  }
  else {
    $script:startedHostForVerification = $false
    try {
      $verificationData = Invoke-InstallerStep -Name "backend-verification" -ScriptBlock {
        $script:startedHostForVerification = Start-HostForVerification
        if (-not (Wait-HostReady -TimeoutSeconds 45)) {
          throw "Backend health did not return 200 within 45 seconds."
        }

        $readiness = Invoke-JsonGet $readinessUrl
        if (-not [bool]$readiness.ready) {
          throw "Backend readiness returned degraded. See readiness data in install-summary.json."
        }

        $readiness
      }
      $null = $verificationData

      if ($SkipStyleRefresh -or [string]::IsNullOrWhiteSpace($MailboxAddress)) {
        Add-Step -Name "style-refresh" -Status "skipped" -Detail "Skipped by request or no mailbox address supplied."
      }
      else {
        Invoke-InstallerStep -Name "style-refresh" -ScriptBlock {
          $refreshUri = "https://localhost:7097/api/outlook/style-profile/refresh?mailboxAddress={0}&profileId={1}" -f [System.Uri]::EscapeDataString($MailboxAddress), [System.Uri]::EscapeDataString($ProfileId)
          Invoke-JsonPost $refreshUri
        } | Out-Null
      }
    }
    finally {
      Stop-HostIfInstallerStartedIt -StartedHost $script:startedHostForVerification
    }
  }

  if (Test-OutlookRunning) {
    Add-Step -Name "outlook-refresh" -Status "warning" -Detail "Outlook is running. If the ribbon still shows an older Maily command surface, close and reopen Outlook once."
  }
  else {
    Add-Step -Name "outlook-refresh" -Status "passed" -Detail "Outlook is closed. Mailbox registration is checked separately; opening Outlook alone does not install a custom mailbox add-in."
  }
  Add-Step -Name 'upgrade-complete' -Status passed -Detail ("Maily {0} application files are installed. Previous application backup: {1}" -f $manifest.Version, $upgradeSnapshot.backupRoot)
  # Cache refresh invalidates pre-install evidence. Only VerifyOutlookActivation after Outlook
  # repopulates its cache can complete registration verification.
  $outlookActivationRequired = $true
  $outlookRestartVerificationRequired = $activation.status -eq 'verified' -and -not $SkipOutlookRegistration -and -not $SkipVerification
  if ($outlookRestartVerificationRequired) {
    Add-Step -Name 'outlook-activation' -Status warning -Detail 'Maily cache was refreshed. Open or restart Outlook, then choose Verify Outlook. The mailbox manifest matched before refresh; no removal or reinstallation in My add-ins is needed unless verification reports a conflict.'
  }
  else { Add-Step -Name 'outlook-activation' -Status warning -Detail ($activation.detail + ' ' + $recovery) }
}
catch {
  if (-not ($steps | Where-Object { $_.status -eq "failed" })) {
    Add-Step -Name "installer" -Status "failed" -Detail $_.Exception.Message
  }
  if ($applicationMutationStarted -and $null -ne $upgradeSnapshot) {
    try {
      Write-Output 'MAILY_STEP|rollback|running|Restoring the previous Maily installation...' | Out-Host
      $rollbackStatus = Restore-MailyFailedInstallation -Snapshot $upgradeSnapshot
      if ($rollbackStatus -eq 'retained-for-retry') {
        Add-Step -Name 'rollback' -Status warning -Detail 'Installation failed after the complete application was copied. Application, registration targets and uninstall files were retained. Retry setup or use the package uninstaller; the failure above still requires resolution.'
      }
      else {
        Add-Step -Name 'rollback' -Status passed -Detail ("Installation files restored to their pre-attempt state (previous version: {0}). Partial application files are retained in the transaction backup. See the error above." -f $upgradeSnapshot.previousVersion)
      }
    }
    catch {
      $rollbackStatus = 'failed'
      Add-Step -Name 'rollback' -Status failed -Detail ("Automatic restore failed: {0}. Preserved backup: {1}" -f $_.Exception.Message, $upgradeSnapshot.backupRoot)
    }
  }
}
finally {
  try {
    $failedSteps = @($steps | Where-Object { $_.status -eq "failed" })
    $warningSteps = @($steps | Where-Object { $_.status -eq "warning" })
    $summaryDir = Join-Path $installRootPath "data"
    Assert-MailyUnredirectedPath -Path $summaryDir
    New-Item -ItemType Directory -Force -Path $summaryDir | Out-Null
    $summaryPath = Join-Path $summaryDir "install-summary.json"
    $summary = [pscustomobject]@{
      runId = $runId
      generatedAtUtc = [DateTime]::UtcNow.ToString("o")
      packageRoot = $packageRoot
      packageMode = $resolvedPackageMode
      installRoot = $installRootPath
      previousVersion = if ($null -ne $upgradeSnapshot) { $upgradeSnapshot.previousVersion } else { '' }
      backupPath = if ($null -ne $upgradeSnapshot) { $upgradeSnapshot.backupRoot } else { '' }
      rollbackStatus = $rollbackStatus
      activationState = if ($null -ne $activation) { $activation.status } else { 'unknown' }
      status = if ($outlookSupportRequired) { "needs-support" } elseif ($outlookActivationRequired) { "activation-required" } elseif ($failedSteps.Count -gt 0) { "failed" } elseif ($warningSteps.Count -gt 0) { "warning" } else { "passed" }
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
    Write-InstallerLog ("failed`tinstall-summary`t{0}" -f $_.Exception.Message)
  }
}

if ($outlookSupportRequired) { exit 23 }
if ($outlookRestartVerificationRequired) { exit 21 }
if ($outlookActivationRequired) { exit 20 }
if ($steps | Where-Object { $_.status -eq "failed" }) {
  exit 1
}

exit 0
