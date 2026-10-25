param([string]$LogDirectory = (Join-Path $PSScriptRoot ('../output/bridge-tests/' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $LogDirectory | Out-Null
Start-Transcript -Path (Join-Path $LogDirectory 'synthetic.log') | Out-Null
$failures = 0
$tokens = $null; $errors = $null
$bridgePath = Join-Path $PSScriptRoot 'outlook-bridge.ps1'
$ast = [System.Management.Automation.Language.Parser]::ParseFile($bridgePath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$ast.FindAll({param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $false) | ForEach-Object { Invoke-Expression $_.Extent.Text }
function Check([string]$Name, [scriptblock]$Test) {
  try { & $Test; Write-Output "PASS $Name" } catch { $script:failures++; Write-Output "FAIL $Name : $($_ | Out-String) $($_.ScriptStackTrace)" }
}
function Assert($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Assert-Throws([scriptblock]$Code) { $caught=$false; try { & $Code | Out-Null } catch { $caught=$true }; Assert $caught 'Expected fail-closed exception' }
function Get-FilteredOutlookItems { @() }
function Collect-MailPayloads { $script:fixture }
function Collect-CalendarPayloads { $script:fixture }
function Collect-MailPayloadsFromFolders { $script:fixture }
function Get-MailboxItem { [pscustomobject]@{} }
function Get-AttachmentsPayload { $script:fixture }
function Emit-Json { param([string]$Json) $Json }
$session=[pscustomobject]@{}
$session | Add-Member ScriptMethod GetDefaultFolder {param($id) [pscustomobject]@{ Name='Fixture' } }
$mailboxStore=$session
$session=[pscustomobject]@{}
$session | Add-Member ScriptMethod GetDefaultFolder {throw 'Default mailbox access is forbidden'}
$FolderName='inbox'; $Limit=10; $DateProperty=''; $BodyMode='full'; $startDateBound=$null; $endDateBound=$null
$switch=$ast.FindAll({param($node) $node -is [System.Management.Automation.Language.SwitchStatementAst]},$false) | Where-Object { $_.Condition.Extent.Text -eq '$Action' }
foreach($actionName in @('list-emails','list-emails-preview','list-search-emails','list-sent-emails','list-calendar-events','list-attachments')) {
  foreach($count in @(0,1,2)) {
    Check "$actionName emits array for $count results" {
      $script:fixture=@(for($i=0;$i -lt $count;$i++) { @{id="synthetic-$i"} })
      $clause=$switch.Clauses | Where-Object {$_.Item1.Value -eq $actionName}
      $json=Invoke-Expression $clause.Item2.Extent.Text
      if($json -is [scriptblock]) { $json=& $json }
      Assert ($json.Trim().StartsWith('[')) "Expected JSON array, got $json"
      $parsed = $json | ConvertFrom-Json
      Assert (@($parsed).Count -eq $count) 'Wrong result count'
    }
  }
}

Stop-Transcript | Out-Null
Write-Output "RESULT failures=$failures"
if ($failures) { exit 1 }
