param(
  [switch]$BuildIfMissing,
  [switch]$RunOnce,
  [int]$PollSeconds = 5
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$startupScript = Join-Path $repoRoot "scripts\start-maily-outlook-addin-hidden.ps1"
$logDir = Join-Path $repoRoot "output\logs"
$logPath = Join-Path $logDir "addin-outlook-watcher.log"
$mutexName = "Local\MailyOutlookAddinWatcher"

New-Item -ItemType Directory -Force -Path $logDir | Out-Null

function Write-WatcherLog {
  param([string]$Message)

  $timestamp = [DateTimeOffset]::UtcNow.ToString("o")
  Add-Content -Path $logPath -Value "$timestamp`t$Message"
}

Write-WatcherLog "Watcher launch requested. RunOnce=$RunOnce BuildIfMissing=$BuildIfMissing PollSeconds=$PollSeconds"

$createdNew = $false
$mutex = New-Object System.Threading.Mutex($true, $mutexName, [ref]$createdNew)
if (-not $createdNew) {
  Write-WatcherLog "Another watcher instance is already active. Exiting."
  exit 0
}

function Test-OutlookRunning {
  return $null -ne (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
}

function Test-HostReady {
  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri "https://localhost:7097/api/health" -TimeoutSec 5
    return $response.StatusCode -eq 200
  }
  catch {
    return $false
  }
}

function Invoke-HostStartup {
  if ($BuildIfMissing) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $startupScript -BuildIfMissing
  }
  else {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $startupScript
  }

  return $LASTEXITCODE
}

try {
  do {
    try {
      if (Test-OutlookRunning) {
        Write-WatcherLog "Outlook detected. Entering host health watch loop."

        if ($RunOnce) {
          $exitCode = Invoke-HostStartup
          if ($exitCode -eq 0) {
            Write-WatcherLog "RunOnce host check/start succeeded."
          }
          else {
            Write-WatcherLog "RunOnce host check/start failed with exit code $exitCode."
          }
          break
        }

        while (Test-OutlookRunning) {
          if (-not (Test-HostReady)) {
            Write-WatcherLog "Host health check failed while Outlook is running. Attempting silent recovery."
            $exitCode = Invoke-HostStartup
            if ($exitCode -eq 0) {
              Write-WatcherLog "Host recovery succeeded."
            }
            else {
              Write-WatcherLog "Host recovery failed with exit code $exitCode."
            }
          }

          Start-Sleep -Seconds $PollSeconds
        }

        Write-WatcherLog "Outlook exited. Waiting for the next launch."
      }
    }
    catch {
      Write-WatcherLog "Watcher loop error: $($_.Exception.Message)"
      Start-Sleep -Seconds ([Math]::Max($PollSeconds, 5))
    }

    if ($RunOnce) {
      break
    }

    Start-Sleep -Seconds $PollSeconds
  } while ($true)
}
catch {
  Write-WatcherLog "Watcher fatal error: $($_.Exception.Message)"
  throw
}
finally {
  $mutex.ReleaseMutex() | Out-Null
  $mutex.Dispose()
  Write-WatcherLog "Watcher stopped."
}
