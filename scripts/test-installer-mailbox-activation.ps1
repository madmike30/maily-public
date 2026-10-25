param([string]$InstallerPath = (Join-Path $PSScriptRoot 'Install-MailyOutlookAddinPackage.ps1'))
& (New-Module {}) {
  param($InstallerPath)
  $ErrorActionPreference = 'Stop'
  $tokens=$null; $errors=$null
  $ast=[Management.Automation.Language.Parser]::ParseFile($InstallerPath,[ref]$tokens,[ref]$errors)
  if ($errors.Count) { throw ($errors | Out-String) }
  foreach ($fn in $ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst]},$false)) { Invoke-Expression $fn.Extent.Text }
  $readCache=(Get-Command Get-MailyMailboxInstalledAddinCache).ScriptBlock
  $root=Join-Path (Split-Path $PSScriptRoot -Parent) ('output/mailbox-activation/'+[guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Path $root -Force | Out-Null
  $script:assertionCount=0
  function Assert-Equal($expected,$actual,$message) { $script:assertionCount++; if ($expected -ne $actual) { throw "$message : expected $expected, got $actual" } }
  function Get-MailyMailboxInstalledAddinCache { param($AddinId) return $script:cache }
  function New-Fixture($name,$version,$url,$id='11111111-1111-1111-1111-111111111111') {
    $path=Join-Path $root $name
    [IO.File]::WriteAllText($path,"<OfficeApp><Id>$id</Id><Version>$version</Version><SourceLocation DefaultValue='$url'/></OfficeApp>")
    return $path
  }
  $target=New-Fixture 'target.xml' '1.3.1.0' 'https://localhost:7097/taskpane.html?v=131'
  $old=New-Fixture 'old.xml' '1.2.7.0' 'https://localhost:7097/taskpane.html?v=127'
  $manifest=Resolve-MailyManifest $target
  foreach ($status in @('verified','activation-required','repair-required','needs-support')) {
    Assert-Equal ('MAILY_ACTIVATION|' + $status) (Get-MailyActivationSignal -Status $status) 'GUI activation signal must be stable and independent of guidance wording'
  }
  $cacheRoot=Join-Path $root 'wef'
  New-Item -ItemType Directory -Path $cacheRoot | Out-Null
  $metadataPath=Join-Path $cacheRoot ($manifest.Id+'.metadata')
  [IO.File]::WriteAllText($metadataPath,'<x:Metadata xmlns:x="urn:synthetic"><x:InstalledBy>User</x:InstalledBy></x:Metadata>')
  Assert-Equal 1 @(& $readCache -AddinId $manifest.Id -CacheRoot $cacheRoot).Count 'Metadata namespace prefix must not affect discovery'
  [IO.File]::WriteAllText($metadataPath,'<broken')
  $badCache=@(& $readCache -AddinId $manifest.Id -CacheRoot $cacheRoot)
  Assert-Equal 1 $badCache.Count 'Malformed Maily metadata must remain visible as unresolved evidence'
  Assert-Equal '' $badCache[0].manifestPath 'Malformed metadata cannot supply trustworthy manifest evidence'
  $script:cache=$badCache
  Assert-Equal 'needs-support' (Get-MailyOutlookActivation -Manifest $manifest).status 'Malformed ownership metadata must not recommend user removal'
  foreach ($source in @('Admin','Organization','Unknown','')) {
    [IO.File]::WriteAllText($metadataPath,"<Metadata><InstalledBy>$source</InstalledBy></Metadata>")
    $script:cache=@(& $readCache -AddinId $manifest.Id -CacheRoot $cacheRoot)
    $decision=Get-MailyOutlookActivation -Manifest $manifest
    Assert-Equal 'needs-support' $decision.status 'A non-user registration must remain untouched for support'
    if ($decision.detail -match 'remove|reinstall|corrupt|outdated') { throw 'Support guidance must not suggest changing a managed or unknown registration.' }
  }
  $script:cache=@([pscustomobject]@{installedBy='User';manifestPath=$old})
  Assert-Equal 'repair-required' (Get-MailyOutlookActivation -Manifest $manifest).status 'Stale mailbox copy must block replacement'
  $script:cache=@([pscustomobject]@{installedBy='User';manifestPath=$target})
  Assert-Equal 'verified' (Get-MailyOutlookActivation -Manifest $manifest).status 'Matching package must be accepted'
  $remote=New-Fixture 'remote.xml' '1.3.1.0' 'https://example.invalid/taskpane.html?v=131'
  $script:cache=@([pscustomobject]@{installedBy='User';manifestPath=$remote})
  Assert-Equal 'repair-required' (Get-MailyOutlookActivation -Manifest $manifest).status 'Same version at a different endpoint is not this package'
  $wrongId=New-Fixture 'wrong-id.xml' '1.3.1.0' 'https://localhost:7097/taskpane.html?v=131' '22222222-2222-2222-2222-222222222222'
  $script:cache=@([pscustomobject]@{installedBy='User';manifestPath=$wrongId})
  Assert-Equal 'repair-required' (Get-MailyOutlookActivation -Manifest $manifest).status 'Cached identity must be verified'
  $script:cache=@()
  Assert-Equal 'activation-required' (Get-MailyOutlookActivation -Manifest $manifest).status 'Cache absence never proves mailbox removal or registration'
  $script:cache=@([pscustomobject]@{installedBy='User';manifestPath=(Join-Path $root 'missing.xml')})
  Assert-Equal 'repair-required' (Get-MailyOutlookActivation -Manifest $manifest).status 'Unreadable mailbox manifest must fail closed'
  $script:cache=@([pscustomobject]@{installedBy='User';manifestPath=$target},[pscustomobject]@{installedBy='User';manifestPath=$old})
  Assert-Equal 'repair-required' (Get-MailyOutlookActivation -Manifest $manifest).status 'One current cache cannot hide another stale mailbox entry'
  $unreadable=Join-Path $root 'invalid.xml'
  [IO.File]::WriteAllText($unreadable,'<!DOCTYPE OfficeApp [<!ENTITY bad SYSTEM "file:///never-read">]><OfficeApp>&bad;</OfficeApp>')
  $script:cache=@([pscustomobject]@{installedBy='User';manifestPath=$unreadable})
  Assert-Equal 'repair-required' (Get-MailyOutlookActivation -Manifest $manifest).status 'DTD and external entity manifests must be rejected'
  foreach ($reverse in @($false,$true)) {
    $script:cache=@([pscustomobject]@{installedBy='User';manifestPath=$old},[pscustomobject]@{installedBy='Unknown';manifestPath=''})
    if ($reverse) { [array]::Reverse($script:cache) }
    Assert-Equal 'needs-support' (Get-MailyOutlookActivation -Manifest $manifest).status 'Unknown ownership must take priority over user repair in either order'
  }
  $waiting=Get-MailyOutlookActivationGuidance -Status 'activation-required' -ManifestPath $target
  if ($waiting -match 'remove|reinstall|retry setup') { throw 'Missing cache must not trigger destructive or repeating installation advice.' }
  if ($waiting -notmatch 'sync|24 hours') { throw 'Missing cache guidance must explain propagation before asking for connection.' }
  $support=Get-MailyOutlookActivationGuidance -Status 'needs-support' -ManifestPath $target
  if ($support -match 'remove|reinstall|Add from file') { throw 'Support recovery cannot tell users to change an unknown registration.' }
  # Integration contract: the mailbox gate must precede all application/process mutations.
  $entry=$ast.EndBlock.Statements | Where-Object { $_ -is [Management.Automation.Language.TryStatementAst] } | Select-Object -Last 1
  $body=$entry.Body.Extent.Text
  if ($body.IndexOf('$activation = Get-MailyOutlookActivation') -lt 0 -or $body.IndexOf('$activation = Get-MailyOutlookActivation') -gt $body.IndexOf('Stop-MailyProcesses')) { throw 'Mailbox preflight must precede process stop and replacement.' }
  if ($body -match 'Register-MailyOutlookAddin -Manifest \$manifest -SkipCacheCleanup') { throw 'An exact mailbox match must not bypass scoped Maily cache refresh.' }
  if ($body -notmatch '(?s)upgrade-complete.*\$outlookActivationRequired = \$true') { throw 'Post-refresh activation must await new Outlook evidence.' }
  if ($ast.Extent.Text -notmatch 'if \(\$outlookActivationRequired\) \{ exit 20 \}') { throw 'Incomplete activation must propagate a non-success exit code.' }
  $supportGate=$body.IndexOf("if (`$activation.status -eq 'needs-support')")
  if ($supportGate -lt 0 -or $supportGate -gt $body.IndexOf('Stop-MailyProcesses') -or $body -notmatch '(?s)needs-support.*?exit 23') { throw 'Support-required ownership must stop before application mutation with exit 23.' }
  Write-Output "PASS installer mailbox activation: $script:assertionCount value assertions, ownership/guidance checks and 5 integration contracts; no mailbox or host access."
} $InstallerPath
