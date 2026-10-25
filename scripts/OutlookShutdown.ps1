param([switch]$CloseOutlook, [switch]$CloseWindowsOnly)

# This helper never opens Outlook, reads mail, saves drafts or terminates Outlook.
# A separate STA process contains potentially blocked normal window-close calls.
function Release-MailyOutlookComObject {
  param($Value)
  if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
  }
}
function Invoke-MailyOutlookWindowClose {
  param([Parameter(Mandatory=$true)]$Outlook)
  # Application.Quit discards unsaved changes. Never call it, even for a
  # background-only instance: normal window closure or a safe failure is required.
  $inspectors = $null
  $explorers = $null
  try {
    # Inspect presence only, never read or alter the inline item's contents.
    # MailItem.Close is unsupported for this object, so the user must finish it.
    $explorers = $Outlook.Explorers
    for ($index = 1; $index -le $explorers.Count; $index++) {
      $window = $null; $inline = $null
      try {
        $window = $explorers.Item($index)
        $inline = $window.ActiveInlineResponse
        if ($null -ne $inline) { throw 'Finish/save and close inline reply first.' }
      } finally {
        Release-MailyOutlookComObject $inline
        Release-MailyOutlookComObject $window
      }
    }
    $inspectors = $Outlook.Inspectors
    $count = $inspectors.Count
    while ($count -gt 0) {
      $window = $null
      try {
        $window = $inspectors.Item($count)
        $window.Close(2) # olPromptForSave; never olSave or olDiscard.
      } finally { Release-MailyOutlookComObject $window }
      $after = $inspectors.Count
      if ($after -ge $count) { throw 'Outlook window close was cancelled or is still pending.' }
      $count = $after
    }
    $count = $explorers.Count
    while ($count -gt 0) {
      $window = $null
      try {
        $window = $explorers.Item($count)
        $inline = $null
        try {
          $inline = $window.ActiveInlineResponse
          if ($null -ne $inline) { throw 'Finish/save and close inline reply first.' }
        } finally { Release-MailyOutlookComObject $inline }
        $window.Close() # Normal Explorer close only when no inline reply is active.
      } finally { Release-MailyOutlookComObject $window }
      $after = $explorers.Count
      if ($after -ge $count) { throw 'Outlook window close was cancelled or is still pending.' }
      $count = $after
    }
  } finally {
    Release-MailyOutlookComObject $inspectors
    Release-MailyOutlookComObject $explorers
  }
}
function Get-MailyOutlookProcesses {
  $session = [Diagnostics.Process]::GetCurrentProcess().SessionId
  @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $session })
}
function Start-MailyOutlookCloseRequest {
  param([string]$HelperPath)
  $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
  Start-Process -FilePath $powershell -ArgumentList ('-NoProfile -STA -ExecutionPolicy Bypass -File "' + $HelperPath + '" -CloseWindowsOnly') -WindowStyle Hidden -PassThru
}
function Invoke-MailyOutlookShutdown {
  param([Parameter(Mandatory=$true)][string]$HelperPath, [int]$TimeoutSeconds = 30)
  $initial = @(Get-MailyOutlookProcesses)
  if ($initial.Count -eq 0) { return [pscustomobject]@{ found=0; closed=$true } }
  $quit = $null
  try {
    $quit = Start-MailyOutlookCloseRequest -HelperPath $HelperPath
    for ($attempt = 0; $attempt -lt ($TimeoutSeconds * 2); $attempt++) {
      $remaining = @(Get-MailyOutlookProcesses)
      if ($remaining.Count -eq 0) { return [pscustomobject]@{ found=$initial.Count; closed=$true } }
      # No window-message fallback: without COM we cannot rule out an inline reply.
      if ($quit.HasExited -and $quit.ExitCode -ne 0) {
        if ($quit.ExitCode -eq 4) { throw 'Finish/save and close inline reply first.' }
        throw 'Outlook close was cancelled or unavailable. Close Outlook normally, then retry.'
      }
      Start-Sleep -Milliseconds 500
    }
    throw 'Outlook is still open. Save your drafts and finish or cancel any Outlook prompts, then retry. No Outlook process was forcibly stopped.'
  }
  finally {
    if ($null -ne $quit) {
      # This is only the helper process created above, never OUTLOOK.EXE.
      try { if (-not $quit.HasExited) { $quit.Kill() } } catch { }
      $quit.Dispose()
    }
  }
}

if ($CloseWindowsOnly) {
  $outlook = $null
  try {
    try { $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { exit 2 }
    Invoke-MailyOutlookWindowClose -Outlook $outlook
    exit 0
  } catch { if ($_.Exception.Message -eq 'Finish/save and close inline reply first.') { exit 4 }; exit 3 }
  finally { if ($null -ne $outlook) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($outlook) } }
}
if ($CloseOutlook) {
  try { Invoke-MailyOutlookShutdown -HelperPath $PSCommandPath | Out-Null; exit 0 }
  catch { [Console]::Error.WriteLine('Outlook is still open. Finish/save and close inline replies and any Outlook prompts, then retry.'); exit 1 }
}
