param(
  [string]$MailboxAddress = "",
  [string]$ProfileId = "default",
  [switch]$SkipOutlookRegistration,
  [switch]$SkipStyleRefresh
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$runId = Get-Date -Format "yyyyMMdd-HHmmss"
$outputDir = Join-Path $repoRoot "output\production-install\$runId"
$summaryPath = Join-Path $outputDir "summary.json"
$manifestPath = Join-Path $repoRoot "addin\manifest.xml"
$healthUrl = "https://localhost:7097/api/health"
$readinessUrl = "https://localhost:7097/api/outlook/readiness"

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

function Add-Step {
  param(
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][string]$Status,
    [string]$Detail = "",
    [object]$Data = $null
  )

  [pscustomobject]@{
    name = $Name
    status = $Status
    detail = $Detail
    data = $Data
  }
}

function Invoke-StepCommand {
  param(
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][scriptblock]$Command
  )

  $logPath = Join-Path $outputDir ("{0}.log" -f ($Name -replace "[^A-Za-z0-9_-]", "-"))
  try {
    $global:LASTEXITCODE = 0
    $output = & $Command 2>&1
    $output | Set-Content -Path $logPath -Encoding UTF8
    if ($LASTEXITCODE -ne 0) {
      return Add-Step -Name $Name -Status "failed" -Detail "Exit code $LASTEXITCODE. See $logPath." -Data @{ logPath = $logPath }
    }

    return Add-Step -Name $Name -Status "passed" -Detail "See $logPath." -Data @{ logPath = $logPath }
  }
  catch {
    $_.Exception.Message | Set-Content -Path $logPath -Encoding UTF8
    return Add-Step -Name $Name -Status "failed" -Detail $_.Exception.Message -Data @{ logPath = $logPath }
  }
}

function Wait-HostReady {
  $deadline = (Get-Date).AddSeconds(45)
  while ((Get-Date) -lt $deadline) {
    try {
      $response = Invoke-WebRequest -UseBasicParsing -Uri $healthUrl -TimeoutSec 5
      if ($response.StatusCode -eq 200) {
        return $true
      }
    }
    catch {
    }

    Start-Sleep -Seconds 1
  }

  return $false
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

function Register-OutlookAddinDirect {
  param([Parameter(Mandatory = $true)][string]$ManifestPath)

  $xml = [xml](Get-Content -LiteralPath $ManifestPath -Raw)
  $addinId = [string]$xml.OfficeApp.Id
  if ([string]::IsNullOrWhiteSpace($addinId)) {
    throw "Could not read add-in id from $ManifestPath."
  }

  $developerKeyPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
  New-Item -Path $developerKeyPath -Force | Out-Null
  New-ItemProperty -Path $developerKeyPath -Name $addinId -PropertyType String -Value $ManifestPath -Force | Out-Null
  New-ItemProperty -Path $developerKeyPath -Name "RefreshAddins" -PropertyType DWord -Value 1 -Force | Out-Null

  [pscustomobject]@{
    addinId = $addinId
    manifestPath = $ManifestPath
    developerKeyPath = $developerKeyPath
  }
}

$steps = @()

if (-not (Test-Path $manifestPath)) {
  throw "Outlook manifest was not found at $manifestPath."
}

$steps += Invoke-StepCommand -Name "publish-and-install-startup" -Command {
  powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repoRoot "scripts\install-maily-outlook-startup.ps1") -BuildIfMissing -ForcePublish
}

if ($SkipOutlookRegistration) {
  $steps += Add-Step -Name "outlook-registration" -Status "skipped" -Detail "Skipped by request."
}
else {
  try {
    $registration = Register-OutlookAddinDirect -ManifestPath $manifestPath
    $steps += Add-Step -Name "outlook-registration" -Status "passed" -Detail "Registered Outlook manifest through current-user Office WEF registry without Node/npm." -Data $registration
  }
  catch {
    $steps += Add-Step -Name "outlook-registration" -Status "failed" -Detail $_.Exception.Message
  }
}

if (-not (Wait-HostReady)) {
  $steps += Add-Step -Name "host-ready" -Status "failed" -Detail "Host did not become healthy within 45 seconds. Make sure Outlook is running so the supervisor wakes the host."
}
else {
  $steps += Add-Step -Name "host-ready" -Status "passed" -Detail "Host health returned 200."
}

try {
  $readiness = Invoke-JsonGet $readinessUrl
  $steps += Add-Step -Name "backend-readiness" -Status ($(if ($readiness.ready) { "passed" } else { "failed" })) -Detail "Backend readiness endpoint completed." -Data $readiness
}
catch {
  $steps += Add-Step -Name "backend-readiness" -Status "failed" -Detail $_.Exception.Message
}

if ($SkipStyleRefresh -or [string]::IsNullOrWhiteSpace($MailboxAddress)) {
  $steps += Add-Step -Name "style-refresh" -Status "skipped" -Detail "Skipped by request or missing mailbox address."
}
else {
  try {
    $refreshUri = "https://localhost:7097/api/outlook/style-profile/refresh?mailboxAddress={0}&profileId={1}" -f [System.Uri]::EscapeDataString($MailboxAddress), [System.Uri]::EscapeDataString($ProfileId)
    $style = Invoke-JsonPost $refreshUri
    $styleState = if ($style.PSObject.Properties.Name -contains "style") { $style.style } else { $style }
    $styleSummary = if ($styleState.PSObject.Properties.Name -contains "summary") { [string]$styleState.summary } else { [string]$styleState.styleSummary }
    $styleOk = $styleState.sampleCount -gt 0 -and -not [string]::IsNullOrWhiteSpace($styleSummary)
    $steps += Add-Step -Name "style-refresh" -Status ($(if ($styleOk) { "passed" } else { "failed" })) -Detail "Refreshed personal style from Sent Items." -Data $styleState
  }
  catch {
    $steps += Add-Step -Name "style-refresh" -Status "failed" -Detail $_.Exception.Message
  }
}

$readinessArgs = @(
  "-NoProfile",
  "-ExecutionPolicy", "Bypass",
  "-File", (Join-Path $repoRoot "scripts\test-maily-production-readiness.ps1"),
  "-RequireOutlookRegistration"
)
if (-not [string]::IsNullOrWhiteSpace($MailboxAddress)) {
  $readinessArgs += "-MailboxAddress"
  $readinessArgs += $MailboxAddress
  $readinessArgs += "-ProfileId"
  $readinessArgs += $ProfileId
}

$steps += Invoke-StepCommand -Name "production-readiness-check" -Command {
  powershell @readinessArgs
}

$summary = [pscustomobject]@{
  runId = $runId
  generatedAtUtc = [DateTime]::UtcNow.ToString("o")
  mailboxAddress = $MailboxAddress
  profileId = $ProfileId
  outputDir = $outputDir
  summaryPath = $summaryPath
  steps = $steps
}

$summary | ConvertTo-Json -Depth 12 | Set-Content -Path $summaryPath -Encoding UTF8
$summary | ConvertTo-Json -Depth 12

$failed = $steps | Where-Object { $_.status -eq "failed" }
if ($failed) {
  exit 1
}
