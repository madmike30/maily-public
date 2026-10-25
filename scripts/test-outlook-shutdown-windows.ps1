$ErrorActionPreference='Stop'
$path=Join-Path $PSScriptRoot 'OutlookShutdown.ps1'
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$null,[ref]$null)
$functions=$ast.FindAll({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst]},$false)
& (New-Module {}) {
 param($functions)
 foreach($function in $functions){Invoke-Expression $function.Extent.Text}
 foreach($case in @('close','inspector-cancel','explorer-cancel','background','inline','late-inline')){
  $script:case=$case;$script:inspectors=2;$script:explorers=2;$script:modes=@();$script:explorerCalls=0
  if($case -eq 'background'){$script:inspectors=0;$script:explorers=0}
  $inspector=New-Object PSObject
  $inspector | Add-Member ScriptMethod Close {param($mode) $script:modes+= $mode;if($script:case -ne 'inspector-cancel'){$script:inspectors--}}
  $explorer=New-Object PSObject; $explorer | Add-Member ScriptProperty ActiveInlineResponse {if($script:case -eq 'inline' -or ($script:case -eq 'late-inline' -and $script:inspectors -eq 0)){[pscustomobject]@{}}else{$null}}
  $explorer | Add-Member ScriptMethod Close {$script:explorerCalls++;if($script:case -ne 'explorer-cancel'){$script:explorers--}}
  $ic=New-Object PSObject;$ic | Add-Member ScriptProperty Count {$script:inspectors};$ic | Add-Member NoteProperty Window $inspector;$ic | Add-Member ScriptMethod Item {param($index) $this.Window}
  $ec=New-Object PSObject;$ec | Add-Member ScriptProperty Count {$script:explorers};$ec | Add-Member NoteProperty Window $explorer;$ec | Add-Member ScriptMethod Item {param($index) $this.Window}
  $app=[pscustomobject]@{Inspectors=$ic;Explorers=$ec};$app | Add-Member ScriptMethod Quit {throw 'Application.Quit must never be used'}
  $failed=$false;try{Invoke-MailyOutlookWindowClose -Outlook $app}catch{$failed=$true}
  if($failed -ne ($case -in @('inspector-cancel','explorer-cancel','inline','late-inline'))){throw "Wrong result for $case"}
  if(@($script:modes | Where-Object {$_ -ne 2}).Count){throw 'Inspector close used save/discard mode instead of prompt'}
  if($case -eq 'inspector-cancel' -and ($script:modes.Count -ne 1 -or $script:explorerCalls -ne 0)){throw 'Cancelled inspector did not stop shutdown'}
  if($case -eq 'explorer-cancel' -and $script:explorerCalls -ne 1){throw 'Cancelled Explorer close was retried'}
  if($case -eq 'close' -and ($script:inspectors -ne 0 -or $script:explorers -ne 0)){throw 'Not all windows closed'}
  if($case -eq 'inline' -and ($script:modes.Count -ne 0 -or $script:explorerCalls -ne 0)){throw 'Inline response must block all window close mutations'}
  if($case -eq 'late-inline' -and $script:explorerCalls -ne 0){throw 'New inline response must block Explorer close'}
  Write-Output "PASS non-destructive Outlook window closure: $case"
 }
} $functions
$source=[IO.File]::ReadAllText($path)
if($source -match '\.Quit\s*\('){throw 'Production helper still invokes destructive Application.Quit'}
Write-Output 'PASS production helper never calls Application.Quit'
