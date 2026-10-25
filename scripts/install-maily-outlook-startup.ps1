param(
  [switch]$BuildIfMissing,
  [switch]$ForcePublish,
  [ValidateSet("Debug", "Release")]
  [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$entryName = "MailyOutlookAddinHostStartup"
$taskName = "MailyOutlookAddinHostWatcher"
$supervisorProjectPath = Join-Path $repoRoot "services\Maily.OutlookAddin.Supervisor\Maily.OutlookAddin.Supervisor.csproj"
$projectPath = Join-Path $repoRoot "services\Maily.OutlookAddin.Web\Maily.OutlookAddin.Web.csproj"
$installRoot = Join-Path $repoRoot "output\startup"
$supervisorPublishDir = Join-Path $installRoot "Maily.OutlookAddin.Supervisor"
$hostPublishDir = Join-Path $installRoot "Maily.OutlookAddin.Web"
$supervisorExe = Join-Path $supervisorPublishDir "Maily.OutlookAddin.Supervisor.exe"
$hostExe = Join-Path $hostPublishDir "Maily.OutlookAddin.Web.exe"
$runKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$supervisorCommand = "`"$supervisorExe`""
$codexHome = if (-not [string]::IsNullOrWhiteSpace($env:CODEX_HOME)) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE ".codex" }

if (-not (Test-Path $supervisorProjectPath)) {
  throw "Maily Outlook supervisor project not found at $supervisorProjectPath."
}

Write-Host "Ensuring the localhost HTTPS development certificate is trusted..."
dotnet dev-certs https --check --trust | Out-Host

if ($ForcePublish) {
  Get-Process -Name Maily.OutlookAddin.Supervisor -ErrorAction SilentlyContinue | ForEach-Object {
    try {
      Stop-Process -Id $_.Id -Force -ErrorAction Stop
    }
    catch {
    }
  }

  Get-Process -Name Maily.OutlookAddin.Web -ErrorAction SilentlyContinue | ForEach-Object {
    try {
      Stop-Process -Id $_.Id -Force -ErrorAction Stop
    }
    catch {
    }
  }
}

if ($ForcePublish -or $BuildIfMissing -or -not (Test-Path $supervisorExe)) {
  Write-Host "Publishing the no-window Outlook supervisor..."
  New-Item -ItemType Directory -Force -Path $supervisorPublishDir | Out-Null
  & dotnet publish $supervisorProjectPath -c $Configuration -o $supervisorPublishDir /nodeReuse:false | Out-Host
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to publish the Outlook add-in supervisor."
  }
}

if ($ForcePublish -or $BuildIfMissing -or -not (Test-Path $hostExe)) {
  Write-Host "Publishing the Outlook add-in host..."
  New-Item -ItemType Directory -Force -Path $hostPublishDir | Out-Null
  & dotnet publish $projectPath -c $Configuration -o $hostPublishDir /nodeReuse:false | Out-Host
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to publish the Outlook add-in host."
  }
}

$supervisorConfig = @{
  MailySupervisor = @{
    ServiceName = "MailyOutlookAddinSupervisor"
    HostServiceName = "MailyOutlookAddinHost"
    HealthUrl = "https://localhost:7097/api/health"
    PollSeconds = 5
    ActivePollSeconds = 5
    IdlePollSeconds = 3
    HealthCheckSeconds = 10
    StopAfterOutlookExitSeconds = 45
    StartupTimeoutSeconds = 20
    LogRetentionDays = 31
    PreferWindowsService = $false
    StopHostWhenOutlookExits = $true
    RepoRoot = $repoRoot
    HostExecutablePath = $hostExe
    HostWorkingDirectory = $hostPublishDir
    LogDirectory = (Join-Path $repoRoot "output\logs")
  }
} | ConvertTo-Json -Depth 8
Set-Content -Path (Join-Path $supervisorPublishDir "appsettings.json") -Value $supervisorConfig -Encoding UTF8

$hostProductionConfig = @{
  MailyHosting = @{
    ServiceName = "MailyOutlookAddinHost"
    HttpsUrl = "https://localhost:7097"
    HttpUrl = "http://localhost:5289"
    LogDirectory = (Join-Path $repoRoot "output\logs")
    CertificatePath = ""
    CertificatePassword = ""
    RequireOutlookBridge = $true
    EnableHttp = $false
    EnableRequestLog = $false
    EnableEventLog = $false
    ReadinessCacheSeconds = 30
  }
  MailyGeneration = @{
    RepoRoot = $repoRoot
    PromptsRoot = (Join-Path $repoRoot "smoke_data\prompts")
    SettingsFilePath = (Join-Path $repoRoot "output\addin\addin_profile_settings.json")
    CodexHome = $codexHome
    CodexExecutablePath = "codex"
    Model = ""
    PromptVersion = "v1"
    BackendApiKey = ""
  }
} | ConvertTo-Json -Depth 8
Set-Content -Path (Join-Path $hostPublishDir "appsettings.Production.json") -Value $hostProductionConfig -Encoding UTF8

New-Item -Path $runKeyPath -Force | Out-Null
Set-ItemProperty -Path $runKeyPath -Name $entryName -Value $supervisorCommand

$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$action = New-ScheduledTaskAction -Execute $supervisorExe
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $currentUser
$settings = New-ScheduledTaskSettingsSet `
  -AllowStartIfOnBatteries `
  -DontStopIfGoingOnBatteries `
  -StartWhenAvailable `
  -MultipleInstances IgnoreNew `
  -RestartCount 999 `
  -RestartInterval (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId $currentUser -RunLevel Limited -LogonType Interactive

Register-ScheduledTask `
  -TaskName $taskName `
  -Action $action `
  -Trigger $trigger `
  -Settings $settings `
  -Principal $principal `
  -Force | Out-Null

Get-Process -Name Maily.OutlookAddin.Supervisor -ErrorAction SilentlyContinue | ForEach-Object {
  try {
    Stop-Process -Id $_.Id -Force -ErrorAction Stop
  }
  catch {
  }
}

Get-Process -Name Maily.OutlookAddin.Web -ErrorAction SilentlyContinue | ForEach-Object {
  try {
    Stop-Process -Id $_.Id -Force -ErrorAction Stop
  }
  catch {
  }
}

Start-Process -FilePath $supervisorExe -WorkingDirectory $supervisorPublishDir -WindowStyle Hidden

Write-Host "Installed current-user auto-start entry:"
Write-Host "  $entryName"
Write-Host "Installed current-user scheduled task:"
Write-Host "  $taskName"
Write-Host "Published install root:"
Write-Host "  $installRoot"
Write-Host "The no-window $Configuration supervisor was also started for the current session."
Write-Host "It watches for Outlook, starts the Maily add-in host hidden, silently recovers it if needed, and stops it after Outlook exits."
