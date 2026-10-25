$ErrorActionPreference = "Stop"

$entryName = "MailyOutlookAddinHostStartup"
$taskName = "MailyOutlookAddinHostWatcher"
$runKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"

if (Get-ItemProperty -Path $runKeyPath -Name $entryName -ErrorAction SilentlyContinue) {
  Remove-ItemProperty -Path $runKeyPath -Name $entryName
  Write-Host "Removed current-user auto-start entry $entryName."
}
else {
  Write-Host "Current-user auto-start entry $entryName is not installed."
}

if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
  Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
  Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
  Write-Host "Removed current-user scheduled task $taskName."
}
else {
  Write-Host "Current-user scheduled task $taskName is not installed."
}

Get-Process -Name Maily.OutlookAddin.Supervisor -ErrorAction SilentlyContinue | ForEach-Object {
  try {
    Stop-Process -Id $_.Id -Force -ErrorAction Stop
    Write-Host "Stopped Maily Outlook supervisor process $($_.Id)."
  }
  catch {
    Write-Host "Could not stop supervisor process $($_.Id): $($_.Exception.Message)"
  }
}

Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe' OR Name = 'pwsh.exe'" |
  Where-Object { ($_.CommandLine -like "*watch-outlook-and-start-maily-addin.ps1*") } |
  ForEach-Object {
    try {
      Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop
      Write-Host "Stopped legacy PowerShell watcher process $($_.ProcessId)."
    }
    catch {
      Write-Host "Could not stop legacy watcher process $($_.ProcessId): $($_.Exception.Message)"
    }
  }
