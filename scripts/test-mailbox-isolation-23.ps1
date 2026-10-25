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
Check 'Mailbox selection resolves explicit secondary account and rejects missing mailbox' {
  $primary=[pscustomobject]@{StoreID='primary'}; $secondary=[pscustomobject]@{StoreID='secondary'}
  $accounts=@([pscustomobject]@{SmtpAddress='first@example.invalid';DeliveryStore=$primary}, [pscustomobject]@{SmtpAddress='second@example.invalid';DeliveryStore=$secondary})
  $s=[pscustomobject]@{Accounts=$accounts; DefaultStore=$primary}
  $resolved=Resolve-MailboxStore -Session $s -Address 'SECOND@example.invalid'
  Assert ($resolved.StoreID -eq 'secondary') 'Wrong mailbox store'
  Assert-Throws { Resolve-MailboxStore -Session $s -Address 'absent@example.invalid' }
  Assert-Throws { Resolve-MailboxStore -Session $s -Address '' }
}
Check 'Hydration refuses a cross-store item' {
  Get-Command Get-MailboxItem -ErrorAction Stop | Out-Null
  $s=[pscustomobject]@{}
  $s | Add-Member ScriptMethod GetItemFromID {param($id,$store) [pscustomobject]@{ Parent=[pscustomobject]@{StoreID='other'} } }
  Assert-Throws { Get-MailboxItem -Session $s -Store ([pscustomobject]@{StoreID='selected'}) -Id 'synthetic' }
}
Check 'Hydration uses requested StoreID and returns matching item' {
  $s=[pscustomobject]@{ SeenStore='' }
  $s | Add-Member ScriptMethod GetItemFromID {param($id,$store) $this.SeenStore=$store; [pscustomobject]@{Parent=[pscustomobject]@{StoreID=$store}; EntryID=$id} }
  $item=Get-MailboxItem -Session $s -Store ([pscustomobject]@{StoreID='secondary'}) -Id 'synthetic'
  Assert ($s.SeenStore -eq 'secondary' -and $item.EntryID -eq 'synthetic') 'Hydration used wrong identity'
}
Check 'Inbox and Sent Items use the resolved secondary store' {
  $store=[pscustomobject]@{StoreID='secondary'}
  $store | Add-Member ScriptMethod GetDefaultFolder {param($id) [pscustomobject]@{Name="secondary-$id"} }
  Assert ((Get-SearchFolders -Store $store -Scope 'inbox').Name -eq 'secondary-6') 'Inbox scope wrong'
  Assert ((Get-SearchFolders -Store $store -Scope 'sent-items').Name -eq 'secondary-5') 'Sent scope wrong'
}
Check 'Backend carries mailbox through hydration and list calls' {
  $program=Get-Content (Join-Path $PSScriptRoot '../services/Maily.OutlookAddin.Web/Program.cs') -Raw
  Assert ($program -match '\["MailId"\] = email.Id,\s*\["MailboxAddress"\] = mailboxAddress') 'Hydration lacks mailbox identity'
  Assert ($program -match '\["FolderName"\] = scope,\s*\["MailboxAddress"\] = mailboxAddress') 'Search lacks mailbox identity'
}

Check 'Connection test verifies a supplied mailbox without reading messages' {
  $mainTry = $ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.TryStatementAst] -and $_.Body.Extent.Text.Contains('$session = Get-OutlookSession') } | Select-Object -First 1
  $setup = [scriptblock]::Create(($mainTry.Body.Statements | Select-Object -First 3 | ForEach-Object { $_.Extent.Text }) -join "`n")
  function Get-OutlookSession {
    [pscustomobject]@{Accounts=@([pscustomobject]@{SmtpAddress='fixture@example.test';DeliveryStore=[pscustomobject]@{StoreID='fixture-store'}})}
  }
  $Action='test'
  $MailboxAddress='absent@example.test'
  Assert-Throws { & $setup }
  $MailboxAddress='fixture@example.test'
  . $setup
  Assert ($mailboxStore.StoreID -eq 'fixture-store') 'Connection probe did not resolve configured mailbox'
  $MailboxAddress=''
  . $setup
  Assert ($null -eq $mailboxStore) 'Generic bridge availability check unexpectedly required an account'
}
Stop-Transcript | Out-Null
Write-Output "RESULT failures=$failures"
if ($failures) { exit 1 }
