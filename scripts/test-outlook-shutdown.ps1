$ErrorActionPreference = 'Stop'
$helper = Join-Path $PSScriptRoot 'OutlookShutdown.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($helper,[ref]$null,[ref]$null)
$functions = $ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$false)
& (New-Module {}) {
 param($functions,$helper)
 foreach($function in $functions){Invoke-Expression $function.Extent.Text}
 function Assert([bool]$ok,[string]$message){if(-not $ok){throw $message}}
 function Get-Process { throw 'Live process enumeration forbidden in shutdown fixtures' }
 function Start-Process { throw 'Live process launch forbidden in shutdown fixtures' }
 function Stop-Process { throw 'Process termination forbidden in shutdown fixtures' }
 function Start-Sleep { }
 function Get-MailyOutlookProcesses { if(-not $script:closed){$script:outlook} }
 function Start-MailyOutlookCloseRequest {
  param($HelperPath)
  $script:started++
  if($script:mode -eq 'quit'){$script:closed=$true}
  $script:worker
 }
 foreach($mode in @('closed','quit','fallback','prompt','cancel')){
  $script:mode=$mode; $script:closed=($mode -eq 'closed');$script:started=0;$script:windowCloses=0;$script:workerKilled=$false;$script:disposed=$false
  $script:outlook=[pscustomobject]@{HasExited=$false;MainWindowHandle=[IntPtr]::Zero}
  if($mode -eq 'fallback'){$script:outlook.MainWindowHandle=[IntPtr]1}
  $script:outlook | Add-Member ScriptMethod CloseMainWindow { $script:windowCloses++;$script:closed=$true; return $true }
  $script:outlook | Add-Member ScriptMethod Kill { throw 'Outlook must never be killed' }
  $script:worker=[pscustomobject]@{HasExited=($mode -ne 'prompt');ExitCode=$(if($mode -eq 'fallback'){2}elseif($mode -eq 'cancel'){3}else{0})}
  $script:worker | Add-Member ScriptMethod Kill { $script:workerKilled=$true }
  $script:worker | Add-Member ScriptMethod Dispose { $script:disposed=$true }
  $failed=$false;try{$result=Invoke-MailyOutlookShutdown -HelperPath $helper -TimeoutSeconds 1}catch{$failed=$true}
  Assert ($failed -eq ($mode -in @('prompt','cancel','fallback'))) "Wrong completion for $mode"
  if($mode -eq 'closed'){Assert ($started -eq 0) 'Already closed launched COM'}
  if($mode -eq 'quit'){Assert ($started -eq 1 -and $windowCloses -eq 0 -and $result.closed) 'Hidden/multiple-window application Quit failed'}
  if($mode -eq 'fallback'){Assert ($windowCloses -eq 0 -and $failed) 'Unavailable COM must fail closed without window fallback'}
  if($mode -eq 'prompt'){Assert ($workerKilled -and $windowCloses -eq 0) 'Prompt must time out without another close or force kill'}
  if($mode -eq 'cancel'){Assert ($windowCloses -eq 0 -and -not $workerKilled) 'Cancelled quit must not issue another close'}
  if($mode -ne 'closed'){Assert $disposed 'Worker handle leaked'}
  Write-Output "PASS shared shutdown: $mode"
 }
} $functions $helper
# Both installer wrappers must call exactly the shared helper, not a divergent COM path.
foreach($name in @('Install-MailyOutlookAddinPackage.ps1','Uninstall-MailyOutlookAddinPackage.ps1')){
 $source=[IO.File]::ReadAllText((Join-Path $PSScriptRoot $name))
 $tree=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$null)
 $function=$tree.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Close-OutlookForSetup'},$true)
 if($function.Extent.Text -notmatch 'Invoke-MailyOutlookShutdown' -or $function.Extent.Text -match 'GetActiveObject|CloseMainWindow'){throw "$name bypasses shared shutdown"}
 Write-Output "PASS shared shutdown wiring: $name"
}
