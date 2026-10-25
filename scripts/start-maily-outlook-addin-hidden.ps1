param(
  [switch]$BuildIfMissing,
  [ValidateSet("Debug", "Release")]
  [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$serviceName = "MailyOutlookAddinHost"
$projectPath = Join-Path $repoRoot "services\Maily.OutlookAddin.Web\Maily.OutlookAddin.Web.csproj"
$publishedAppDir = Join-Path $repoRoot "output\startup\Maily.OutlookAddin.Web"
$releaseAppDir = Join-Path $repoRoot "services\Maily.OutlookAddin.Web"
$debugAppDir = Join-Path $repoRoot "services\Maily.OutlookAddin.Web"
$publishedAppExe = Join-Path $publishedAppDir "Maily.OutlookAddin.Web.exe"
$releaseAppExe = Join-Path $releaseAppDir "bin\Release\net10.0\Maily.OutlookAddin.Web.exe"
$debugAppExe = Join-Path $debugAppDir "bin\Debug\net10.0\Maily.OutlookAddin.Web.exe"
$healthUrl = "https://localhost:7097/api/health"
$logDir = Join-Path $repoRoot "output\logs"
$logPath = Join-Path $logDir "addin-startup.log"
$hostProcessName = "Maily.OutlookAddin.Web"

New-Item -ItemType Directory -Force -Path $logDir | Out-Null

function Write-StartupLog {
  param([string]$Message)

  $timestamp = [DateTimeOffset]::UtcNow.ToString("o")
  Add-Content -Path $logPath -Value "$timestamp`t$Message"
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

function Get-HostProcesses {
  return @(Get-Process -Name $hostProcessName -ErrorAction SilentlyContinue)
}

function Stop-StaleHostProcesses {
  $processes = Get-HostProcesses
  if (-not $processes -or $processes.Count -eq 0) {
    return
  }

  foreach ($process in $processes) {
    try {
      Write-StartupLog "Stopping stale host process id=$($process.Id)."
      Stop-Process -Id $process.Id -Force -ErrorAction Stop
    }
    catch {
      Write-StartupLog "Failed to stop stale host process id=$($process.Id): $($_.Exception.Message)"
    }
  }

  Start-Sleep -Milliseconds 750
}

function Get-HostExecutable {
  if (Test-Path $publishedAppExe) {
    return [pscustomobject]@{ Exe = $publishedAppExe; Directory = $publishedAppDir; Source = "published" }
  }

  if ($Configuration -eq "Release" -and (Test-Path $releaseAppExe)) {
    return [pscustomobject]@{ Exe = $releaseAppExe; Directory = $releaseAppDir; Source = "release-build" }
  }

  if (Test-Path $debugAppExe) {
    return [pscustomobject]@{ Exe = $debugAppExe; Directory = $debugAppDir; Source = "debug-build" }
  }

  return $null
}

try {
  Write-StartupLog "Ensuring localhost HTTPS development certificate is trusted."
  $certOutput = & dotnet dev-certs https --check --trust 2>&1
  $certOutput | ForEach-Object { Write-StartupLog $_ }
  if ($LASTEXITCODE -ne 0) {
    Write-StartupLog "HTTPS certificate validation failed with exit code $LASTEXITCODE."
    exit $LASTEXITCODE
  }

  if (Test-HostReady) {
    Write-StartupLog "Host already running."
    exit 0
  }

  $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
  if ($service) {
    if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Running) {
      Write-StartupLog "Windows service $serviceName is running but health check failed. Restarting service."
      Restart-Service -Name $serviceName -Force
    }
    elseif ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::StartPending) {
      Write-StartupLog "Windows service $serviceName is already starting."
    }
    else {
      Write-StartupLog "Starting Windows service $serviceName."
      Start-Service -Name $serviceName
    }
  }
  else {
    Stop-StaleHostProcesses

    $hostExecutable = Get-HostExecutable
    if ($null -eq $hostExecutable) {
      if (-not $BuildIfMissing) {
        Write-StartupLog "Host executable missing and BuildIfMissing was not specified."
        exit 1
      }

      Write-StartupLog "Host executable missing. Publishing project."
      $publishDir = Join-Path $repoRoot "output\startup\Maily.OutlookAddin.Web"
      New-Item -ItemType Directory -Force -Path $publishDir | Out-Null
      $buildOutput = & dotnet publish $projectPath -c $Configuration -o $publishDir /nodeReuse:false 2>&1
      $buildOutput | ForEach-Object { Write-StartupLog $_ }
      if ($LASTEXITCODE -ne 0) {
        Write-StartupLog "Publish failed with exit code $LASTEXITCODE."
        exit $LASTEXITCODE
      }

      $hostExecutable = Get-HostExecutable
      if ($null -eq $hostExecutable) {
        Write-StartupLog "Publish completed but host executable was not found."
        exit 1
      }
    }

    Write-StartupLog "Starting hidden native host executable from $($hostExecutable.Source)."
    Start-Process -FilePath $hostExecutable.Exe -WorkingDirectory $hostExecutable.Directory -WindowStyle Hidden
  }

  for ($i = 0; $i -lt 20; $i++) {
    Start-Sleep -Seconds 1
    if (Test-HostReady) {
      Write-StartupLog "Host became ready."
      exit 0
    }
  }

  Write-StartupLog "Host failed to become ready within timeout."
  exit 1
}
catch {
  Write-StartupLog "Startup failed: $($_.Exception.Message)"
  exit 1
}
