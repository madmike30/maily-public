param([string]$UninstallerPath = (Join-Path $PSScriptRoot 'Uninstall-MailyOutlookAddinPackage.ps1'))
& (New-Module {}) {
  param($UninstallerPath)
  $ErrorActionPreference = 'Stop'
  $tokens=$null; $errors=$null
  $ast=[Management.Automation.Language.Parser]::ParseFile($UninstallerPath,[ref]$tokens,[ref]$errors)
  if ($errors.Count) { throw 'Uninstaller parsing failed.' }
  $function=$ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-MailyRemovalOutcome'},$false) | Select-Object -First 1
  if (-not $function) { throw 'Uninstaller lacks an honest local-versus-mailbox completion outcome.' }
  Invoke-Expression $function.Extent.Text
  $removed=@([pscustomobject]@{name='install-files';status='passed'})
  $result=Get-MailyRemovalOutcome -Steps $removed
  if ($result.exitCode -ne 22 -or $result.status -ne 'mailbox-removal-unverified' -or -not $result.localRemovalComplete -or $result.mailboxRemovalVerified) { throw 'Local file removal must not claim mailbox removal.' }
  $result=Get-MailyRemovalOutcome -Steps @($removed + [pscustomobject]@{name='outlook-registration';status='failed'})
  if ($result.exitCode -ne 1 -or $result.status -ne 'failed' -or $result.localRemovalComplete) { throw 'A failed cleanup step must override partial completion.' }
  $result=Get-MailyRemovalOutcome -Steps @([pscustomobject]@{name='install-files';status='skipped'})
  if ($result.exitCode -ne 1 -or $result.localRemovalComplete) { throw 'Skipped removal cannot establish completion.' }
  $result=Get-MailyRemovalOutcome -Steps @()
  if ($result.exitCode -ne 1 -or $result.mailboxRemovalVerified) { throw 'Missing evidence must fail closed.' }
  Write-Output 'PASS uninstall completion: 4 synthetic cases; no registry, filesystem deletion or mailbox access.'
} $UninstallerPath
