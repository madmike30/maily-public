param([string]$InstallerPath = (Join-Path $PSScriptRoot 'Install-MailyOutlookAddinPackage.ps1'), [string]$OutputDirectory = '')
& (New-Module {}) { param($InstallerPath, $OutputDirectory)
$ErrorActionPreference = 'Stop'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($InstallerPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
# Load functions only: these fixtures must never execute the Windows installer entry point.
$functions = $ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)
foreach ($function in $functions) { Invoke-Expression $function.Extent.Text }
function Assert-True([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$testRoot = if ($OutputDirectory) { [IO.Path]::GetFullPath($OutputDirectory) } else { Join-Path (Split-Path (Split-Path $InstallerPath -Parent) -Parent) ('output/installer-lifecycle/' + [Guid]::NewGuid().ToString('N')) }
if ((Test-Path -LiteralPath $testRoot) -and (Get-ChildItem -LiteralPath $testRoot -Force | Select-Object -First 1)) { throw "Lifecycle output must be empty; fixtures are retained." }
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
# Simulate deletion only inside this module: retain each removed fixture as evidence.
function Remove-Item {
  [CmdletBinding()] param([Parameter(Mandatory=$true)][string]$LiteralPath, [switch]$Recurse, [switch]$Force)
  $target = [IO.Path]::GetFullPath($LiteralPath)
  $root = [IO.Path]::GetFullPath($testRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
  if (-not $target.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture removal escaped isolated root.' }
  $retained = Join-Path $root ('retained-removals/' + [Guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Path $retained -Force | Out-Null
  Move-Item -LiteralPath $target -Destination (Join-Path $retained ([IO.Path]::GetFileName($target)))
}
# Fail closed at external host boundaries before invoking any installer function.
foreach ($blocked in @('Get-ItemProperty','Set-ItemProperty','Remove-ItemProperty','New-ItemProperty','Get-Process','Start-Process','Stop-Process','New-SelfSignedCertificate','Import-Certificate','Export-PfxCertificate','Export-Certificate','New-Object')) {
  if ($blocked -eq 'New-Object') { continue } # Fixture collections use this; no COM functions are invoked.
  Set-Item -Path ("function:" + $blocked) -Value { throw 'Host registry/process/certificate access is forbidden in lifecycle fixtures.' }
}
function Test-Path {
  [CmdletBinding()] param([string]$LiteralPath, [string]$Path, [string]$PathType)
  $candidate = if ($LiteralPath) { $LiteralPath } else { $Path }
  $full = [IO.Path]::GetFullPath($candidate)
  if (-not $full.StartsWith([IO.Path]::GetFullPath($testRoot), [StringComparison]::OrdinalIgnoreCase)) { throw 'Filesystem access escaped lifecycle fixtures.' }
  $arguments = @{ LiteralPath = $full }
  if ($PathType) { $arguments.PathType = $PathType }
  Microsoft.PowerShell.Management\Test-Path @arguments
}
function Get-MailyPathAttributes {
  param([string]$Path)
  if (-not $Path.StartsWith($testRoot, [StringComparison]::OrdinalIgnoreCase)) { return [IO.FileAttributes]::Directory }
  if ([IO.Directory]::Exists($Path) -or [IO.File]::Exists($Path)) { return [IO.File]::GetAttributes($Path) }
  return $null
}
try {
  $candidates = New-Object 'System.Collections.Generic.List[object]'
  Add-CodexCandidate -Candidates $candidates -Path 'C:\fixture\codex.cmd' -Source 'fixture'
  Add-CodexCandidate -Candidates $candidates -Path 'C:\fixture\codex.bat' -Source 'fixture'
  Assert-True ($candidates.Count -eq 0) 'Bundled runtime selection must require a native executable.'
  Add-CodexCandidate -Candidates $candidates -Path 'C:\fixture\codex.exe' -Source 'fixture'
  Assert-True ($candidates.Count -eq 1) 'Native CLI candidate was discarded.'
  Write-Host 'PASS bundled runtime candidates require native executables'
  # Character escapes keep this BOM-free test readable by Windows PowerShell 5.1.
  $installRootPath = Join-Path $testRoot ('installed-' + [char]0x05D0 + [char]0x00E9)
  $packageRoot = Join-Path $testRoot 'package'
  $payloadRoot = Join-Path $packageRoot 'payload'
  $runId = 'fixture'
  $resolvedCodexExecutablePath = Join-Path $installRootPath 'runtime/codex/codex.exe'
  $healthUrl = 'https://localhost:7097/api/health'
  New-Item -ItemType Directory -Path (Join-Path $installRootPath 'app/Maily.OutlookAddin.Web'), (Join-Path $installRootPath 'app/Maily.OutlookAddin.Supervisor'), (Join-Path $installRootPath 'addin'), (Join-Path $installRootPath 'data/addin'), (Join-Path $payloadRoot 'app') -Force | Out-Null
  $manifest = '<OfficeApp><Id>fixture-maily</Id><Version>1.2.9.0</Version></OfficeApp>'
  Set-Content -LiteralPath (Join-Path $installRootPath 'addin/manifest.xml') -Value $manifest
  Set-Content -LiteralPath (Join-Path $installRootPath 'app/old.txt') -Value 'previous-version'
  $profilePath = Join-Path $installRootPath 'data/addin/addin_profile_settings.json'
  $original = '{"items":{"legacy":{"settings":{"model":"gpt-5.4-mini","language":"Hebrew","customInstructions":"Keep my exact preference","hidden":{"keep":true}}},"chosen":{"settings":{"model":"gpt-5.6-sol","tone":"friendly"}},"default":{"settings":{"model":""}}},"unknownRootField":17}'
  $legacyText = 'Preference ' + [char]0x05D0 + [char]0x00E9
  $customText = 'Tone ' + [char]0x05D1 + [char]0x00F1
  $original = $original.Replace('Keep my exact preference', $legacyText).Replace('friendly', $customText)
  [IO.File]::WriteAllText($profilePath, $original)
  $originalHash = (Get-FileHash -LiteralPath $profilePath).Hash

  # Emulate Windows PowerShell's legacy default decoder without changing the host code page.
  function Get-Content {
    [CmdletBinding()] param([string]$LiteralPath, [switch]$Raw, [string]$Encoding)
    if ([IO.Path]::GetFileName($LiteralPath) -in @('addin_profile_settings.json', 'upgrade-snapshot.json') -and -not $Encoding) {
      return [Text.Encoding]::GetEncoding(1252).GetString([IO.File]::ReadAllBytes($LiteralPath))
    }
    $arguments = @{ LiteralPath = $LiteralPath; Raw = $Raw }
    if ($Encoding) { $arguments.Encoding = $Encoding }
    Microsoft.PowerShell.Management\Get-Content @arguments
  }

  $snapshot = Backup-MailyInstallation
  Assert-True ($snapshot.exists -and $snapshot.previousVersion -eq '1.2.9.0') 'Existing version was not detected/backed up.'
  $migration = Update-LegacyMailyProfiles -SettingsPath $profilePath
  $migrated = [IO.File]::ReadAllText($profilePath) | ConvertFrom-Json
  Assert-True ($migrated.items.legacy.settings.customInstructions -ceq $legacyText) 'Migration corrupted Unicode instructions in the legacy profile.'
  Assert-True ($migrated.items.chosen.settings.tone -ceq $customText) 'Migration corrupted Unicode in an unrelated custom-model profile.'
  Assert-True ($migration.changedProfiles -eq 1 -and $migrated.items.legacy.settings.model -eq '') 'Legacy profile migration failed.'
  Assert-True ($migrated.items.chosen.settings.model -eq 'gpt-5.6-sol' -and $migrated.items.default.settings.model -eq '') 'Custom/default model was overwritten.'
  Assert-True ($migrated.items.legacy.settings.hidden.keep -and $migrated.unknownRootField -eq 17 -and $migrated.items.legacy.settings.language -eq 'Hebrew') 'User settings or unknown fields were lost.'
  Assert-True ((Get-FileHash -LiteralPath $migration.backupPath).Hash -eq $originalHash) 'Migration backup must match original bytes.'
  $migratedHash = (Get-FileHash -LiteralPath $profilePath).Hash
  $second = Update-LegacyMailyProfiles -SettingsPath $profilePath
  Assert-True ($second.changedProfiles -eq 0 -and (Get-FileHash -LiteralPath $profilePath).Hash -eq $migratedHash) 'Migration must be idempotent.'
  Write-Host 'PASS legacy-model migration preserves literal Unicode in all profiles under ANSI decoding, custom models, unknown fields, and exact backup bytes'

  $cert = [pscustomobject]@{ pfxPath = (Join-Path $installRootPath 'certs/localhost.pfx'); password = 'fixture-password' }
  Write-InstalledConfiguration -ResolvedCodexHome '' -Certificate $cert | Out-Null
  $production = Get-Content -LiteralPath (Join-Path $installRootPath 'app/Maily.OutlookAddin.Web/appsettings.Production.json') -Raw | ConvertFrom-Json
  Assert-True ($production.MailyGeneration.Model -eq '' -and $production.MailyGeneration.PromptVersion -eq 'v1') 'Installed runtime defaults are incorrect.'
  Assert-True ($production.MailyGeneration.CodexHome -eq '' -and $production.MailyGeneration.CodexExecutablePath -eq $resolvedCodexExecutablePath) 'Installed auth/runtime path changed incorrectly.'
  $configBytes = [IO.File]::ReadAllBytes((Join-Path $installRootPath 'app/Maily.OutlookAddin.Web/appsettings.Production.json'))
  Assert-True (-not ($configBytes[0] -eq 239 -and $configBytes[1] -eq 187 -and $configBytes[2] -eq 191)) 'Installed configuration contains a UTF-8 BOM.'
  Write-Host 'PASS installed Production configuration uses bundled runtime, existing auth, Codex recommended default model and v1 without BOM'

  # Inject failure after replacing application files and migrating settings; restore the captured snapshot.
  Set-Content -LiteralPath (Join-Path $payloadRoot 'app/new.txt') -Value 'new-version'
  Copy-PayloadDirectory -Name 'app'
  Assert-True (-not (Test-Path -LiteralPath (Join-Path $installRootPath 'app/old.txt'))) 'Replacement retained stale application files.'
  $markerPath = Join-Path $snapshot.backupRoot 'upgrade-snapshot.json'
  $markerBytes = [IO.File]::ReadAllBytes($markerPath)
  Assert-True (-not ($markerBytes[0] -eq 239 -and $markerBytes[1] -eq 187 -and $markerBytes[2] -eq 191)) 'Transaction marker must be BOM-free UTF-8.'
  $legacyMarker = [Text.Encoding]::GetEncoding(1252).GetString($markerBytes) | ConvertFrom-Json
  Assert-True ($legacyMarker.installRoot -ne $installRootPath) 'Unicode fixture must expose legacy ANSI decoding.'
  Restore-MailyInstallationFiles -Snapshot $snapshot
  Assert-True ((Get-Content -LiteralPath (Join-Path $installRootPath 'app/old.txt')).Trim() -eq 'previous-version') 'Rollback did not restore old application.'
  Assert-True (-not (Test-Path -LiteralPath (Join-Path $installRootPath 'app/new.txt'))) 'Rollback retained partial new application files.'
  Assert-True ((Get-FileHash -LiteralPath $profilePath).Hash -eq $originalHash) 'Rollback did not restore original user profile bytes.'
  Write-Host 'PASS injected replacement failure restores Unicode paths and original profile bytes under legacy ANSI decoding'

  $file = Join-Path $payloadRoot 'app/new.txt'
  $secondPackageFile = Join-Path $packageRoot 'helper.ps1'
  Set-Content -LiteralPath $secondPackageFile -Value 'synthetic-helper'
  $inventoryEntries = @(
    [pscustomobject]@{ path = 'payload/app/new.txt'; sha256 = (Get-FileHash -LiteralPath $file).Hash },
    [pscustomobject]@{ path = 'helper.ps1'; sha256 = (Get-FileHash -LiteralPath $secondPackageFile).Hash }
  )
  ConvertTo-Json -InputObject $inventoryEntries | Set-Content -LiteralPath (Join-Path $packageRoot 'SHA256SUMS.json')
  $inventoryResult = Test-MailyPackageIntegrity
  Assert-True ($inventoryResult.verifiedFiles -eq 2) 'Multi-entry JSON inventory was nested instead of enumerated (Windows PowerShell 5.1 regression).'
  Remove-Item -LiteralPath $secondPackageFile -Force
  ConvertTo-Json -InputObject @($inventoryEntries[0]) | Set-Content -LiteralPath (Join-Path $packageRoot 'SHA256SUMS.json')
  Test-MailyPackageIntegrity | Out-Null
  Set-Content -LiteralPath $file -Value 'tampered'
  $rejected = $false
  try { Test-MailyPackageIntegrity | Out-Null } catch { $rejected = $true }
  Assert-True $rejected 'Tampered payload was accepted.'
  ConvertTo-Json -InputObject @([pscustomobject]@{ path = '../outside.txt'; sha256 = 'bad' }) | Set-Content -LiteralPath (Join-Path $packageRoot 'SHA256SUMS.json')
  $rejected = $false
  try { Test-MailyPackageIntegrity | Out-Null } catch { $rejected = $true }
  Assert-True $rejected 'Package inventory traversal was accepted.'
  Write-Host 'PASS package checksum validation rejects tampering and path traversal'
  $hashEntry = [pscustomobject]@{ path = 'payload/app/new.txt'; sha256 = (Get-FileHash -LiteralPath $file).Hash }
  @($hashEntry, $hashEntry) | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $packageRoot 'SHA256SUMS.json')
  $rejected = $false
  try { Test-MailyPackageIntegrity | Out-Null } catch { $rejected = $true }
  Assert-True $rejected 'Duplicate inventory entries were accepted.'
  ConvertTo-Json -InputObject @($hashEntry) | Set-Content -LiteralPath (Join-Path $packageRoot 'SHA256SUMS.json')
  Set-Content -LiteralPath (Join-Path $packageRoot 'unlisted.exe') -Value 'unlisted'
  $rejected = $false
  try { Test-MailyPackageIntegrity | Out-Null } catch { $rejected = $true }
  Assert-True $rejected 'An unlisted executable was accepted.'
  Write-Host 'PASS duplicate and omitted inventory files are rejected'

  $cacheRoot = Join-Path $testRoot 'cache'
  $mailyId = '194a156f-26dd-420f-8eba-e78caa663953'
  $otherId = '204a156f-26dd-420f-8eba-e78caa663953'
  New-Item -ItemType Directory -Path (Join-Path $cacheRoot ('profile/' + $mailyId)), (Join-Path $cacheRoot ('profile/' + $otherId)) -Force | Out-Null
  $otherFile = Join-Path $cacheRoot ('profile/' + $otherId + '/important.txt')
  Set-Content -LiteralPath $otherFile -Value 'other-addin-data'
  Set-Content -LiteralPath (Join-Path $cacheRoot ('profile/' + $mailyId + '/metadata')) -Value 'old-maily'
  Set-Content -LiteralPath (Join-Path $cacheRoot ('Stream_' + ($mailyId -replace '-', '') + '.dat')) -Value 'old-maily'
  $removed = Remove-MailyCacheFiles -CacheRoot $cacheRoot -AddinId $mailyId
  Assert-True ($removed -eq 2 -and (Test-Path -LiteralPath $otherFile)) 'Cache cleanup removed another add-in or failed to clear Maily.'
  Assert-True (-not (Test-Path -LiteralPath (Join-Path $cacheRoot ('profile/' + $mailyId)))) 'Old Maily WEF cache remained.'
  $rejected = $false
  try { Remove-MailyCacheFiles -CacheRoot $cacheRoot -AddinId '*' | Out-Null } catch { $rejected = $true }
  Assert-True $rejected 'Wildcard add-in ID was accepted.'
  Write-Host 'PASS scoped cache cleanup clears Maily GUID/compact-GUID entries and preserves another add-in'
  if ($env:OS -ne 'Windows_NT') {
    $link = Join-Path $cacheRoot ('redirect-' + $mailyId)
    New-Item -ItemType SymbolicLink -Path $link -Target (Split-Path -Parent $otherFile) | Out-Null
    $rejected = $false
    try { Remove-MailyCacheFiles -CacheRoot $cacheRoot -AddinId $mailyId | Out-Null } catch { $rejected = $true }
    Assert-True ($rejected -and (Test-Path -LiteralPath $otherFile)) 'Cache cleanup followed a targeted reparse point.'
    Remove-Item -LiteralPath $link -Force
    Write-Host 'PASS redirected cache entry is rejected without touching its target'
  }
  $uninstallTokens = $null; $uninstallErrors = $null
  $uninstallAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path (Split-Path -Parent $InstallerPath) 'Uninstall-MailyOutlookAddinPackage.ps1'), [ref]$uninstallTokens, [ref]$uninstallErrors)
  Assert-True ($uninstallErrors.Count -eq 0) 'Uninstaller has syntax errors.'
  foreach ($functionName in @('Remove-MailyCacheFiles', 'Remove-MailyOutlookCache')) {
    $left = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $false)
    $right = $uninstallAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $false)
    Assert-True (($left.Extent.Text -replace "`r`n", "`n") -ceq ($right.Extent.Text -replace "`r`n", "`n")) 'Installer/uninstaller cache cleanup has diverged.'
  }
  Write-Host 'PASS installer and uninstaller share the verified cache-cleanup implementation'
  $ownershipFunction = $uninstallAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Test-MailyCertificateOwnership' }, $false)
  Invoke-Expression $ownershipFunction.Extent.Text
  $owned = [pscustomobject]@{ Thumbprint = 'ABC123'; FriendlyName = 'Maily Outlook Add-in Localhost'; Subject = 'CN=localhost'; HasPrivateKey = $true }
  Assert-True (Test-MailyCertificateOwnership -Certificate $owned -ExpectedThumbprint 'abc123') 'Owned Maily certificate was not recognized.'
  Assert-True (-not (Test-MailyCertificateOwnership -Certificate $owned -ExpectedThumbprint 'other')) 'A different thumbprint was accepted as owned.'
  $shared = [pscustomobject]@{ Thumbprint = 'ABC123'; FriendlyName = 'ASP.NET Core HTTPS development certificate'; Subject = 'CN=localhost'; HasPrivateKey = $true }
  Assert-True (-not (Test-MailyCertificateOwnership -Certificate $shared -ExpectedThumbprint 'ABC123')) 'A shared localhost certificate was classified as Maily-owned.'
  Assert-True (-not (Test-MailyCertificateOwnership -Certificate $null -ExpectedThumbprint 'ABC123')) 'Missing personal-store ownership should retain trusted certificates.'
  $owned.HasPrivateKey = $false
  Assert-True (-not (Test-MailyCertificateOwnership -Certificate $owned -ExpectedThumbprint 'ABC123')) 'A public-only certificate was classified as owned personal key material.'
  Write-Host 'PASS certificate ownership retains shared, mismatched, missing, and public-only certificates'
  $savedInstallRoot = $installRootPath
  $installRootPath = Join-Path $testRoot 'redirected-install'
  $redirectTarget = Join-Path $testRoot 'redirect-target'
  New-Item -ItemType Directory -Path $installRootPath, $redirectTarget -Force | Out-Null
  New-Item -ItemType Junction -Path (Join-Path $installRootPath 'data') -Target $redirectTarget | Out-Null
  $rejected = $false
  try { Backup-MailyInstallation | Out-Null } catch { $rejected = $_.Exception.Message -match 'redirected' }
  Assert-True $rejected 'Backup followed a redirected data ancestor.'
  Assert-True (-not (Test-Path -LiteralPath (Join-Path $redirectTarget 'upgrade-backups'))) 'Backup wrote through the junction before rejecting it.'
  $rejected = $false
  try { Restore-MailyInstallationFiles -Snapshot ([pscustomobject]@{ backupRoot = (Join-Path $installRootPath 'data/upgrade-backups/fixture'); names = @() }) } catch { $rejected = $_.Exception.Message -match 'redirected' }
  Assert-True $rejected 'Restore followed a redirected backup ancestor.'
  $installRootPath = $savedInstallRoot
  Write-Host 'PASS backup and restore reject redirected ancestors before writing'
  $savedInstallRoot = $installRootPath
  $installRootPath = Join-Path $testRoot 'fresh-install'
  New-Item -ItemType Directory -Path (Join-Path $installRootPath 'data') -Force | Out-Null
  $userFile = Join-Path $installRootPath 'data/keep-user-file.txt'
  [IO.File]::WriteAllText($userFile, 'retained user file')
  $freshSnapshot = Backup-MailyInstallation
  New-Item -ItemType Directory -Path (Join-Path $installRootPath 'app') -Force | Out-Null
  [IO.File]::WriteAllText((Join-Path $installRootPath 'app/partial.txt'), 'partial copied application')
  function Stop-MailyProcesses { }
  $tamperedSnapshot = $freshSnapshot | ConvertTo-Json -Depth 5 | ConvertFrom-Json
  $tamperedSnapshot.transactionId = 'wrong-transaction'
  $rejected = $false
  try { Restore-MailyInstallationFiles -Snapshot $tamperedSnapshot } catch { $rejected = $_.Exception.Message -match 'transaction marker' }
  Assert-True ($rejected -and (Test-Path -LiteralPath (Join-Path $installRootPath 'app/partial.txt'))) 'Rollback accepted an unowned transaction.'
  Restore-MailyFailedInstallation -Snapshot $freshSnapshot
  Assert-True (-not (Test-Path -LiteralPath (Join-Path $installRootPath 'app'))) 'Failed fresh installation blocks retry with partial managed files.'
  Assert-True (Test-Path -LiteralPath $userFile) 'Fresh install rollback lost unrelated user data.'
  Assert-True (@(Get-ChildItem -LiteralPath $freshSnapshot.backupRoot -Recurse -File | Where-Object Name -eq 'partial.txt').Count -eq 1) 'Rollback deleted partial application evidence instead of quarantining it.'
  Write-Host 'PASS failed fresh installation permits retry and retains user data and partial evidence'
  # Once the complete payload is installed, late registration/readiness failures
  # must leave usable retry/uninstall targets instead of dangling registrations.
  New-Item -ItemType Directory -Path (Join-Path $installRootPath 'app'), (Join-Path $installRootPath 'addin'), (Join-Path $installRootPath 'setup') -Force | Out-Null
  [IO.File]::WriteAllText((Join-Path $installRootPath 'app/complete.txt'), 'complete payload')
  [IO.File]::WriteAllText((Join-Path $installRootPath 'addin/manifest.xml'), '<OfficeApp><Id>fixture-maily</Id><Version>1.3.1.0</Version></OfficeApp>')
  [IO.File]::WriteAllText((Join-Path $installRootPath 'setup/uninstall.ps1'), 'fixture uninstall target')
  $script:installationPayloadReady = $true
  $lateResult = Restore-MailyFailedInstallation -Snapshot $freshSnapshot
  Assert-True ($lateResult -eq 'retained-for-retry') 'Late fresh-install failure did not report retained payload.'
  Assert-True ((Test-Path -LiteralPath (Join-Path $installRootPath 'app/complete.txt')) -and (Test-Path -LiteralPath (Join-Path $installRootPath 'setup/uninstall.ps1'))) 'Late failure broke registration or uninstall targets.'
  $script:installationPayloadReady = $false
  Write-Host 'PASS late fresh-install failure preserves complete retry/uninstall targets without registry mutation'
  $installRootPath = $savedInstallRoot
  function Invoke-CodexProcess { return @{ started = $true; exitCode = 0; stdout = 'Logged in with ChatGPT synthetic-private-marker'; stderr = ''; error = '' } }
  $fakeRuntime = Join-Path $testRoot 'codex.exe'
  [IO.File]::WriteAllText($fakeRuntime, 'fixture only')
  $status = Test-CodexSetup -ExecutablePath $fakeRuntime
  Assert-True (-not $status.ok -and ($status | ConvertTo-Json -Depth 5) -notmatch 'synthetic-private-marker') 'Ambiguous authentication must be rejected without exposing raw diagnostics.'
  function Invoke-CodexProcess { return @{ started = $true; exitCode = 0; stdout = 'Logged in using ChatGPT'; stderr = ''; error = '' } }
  $status = Test-CodexSetup -ExecutablePath $fakeRuntime
  Assert-True $status.ok 'Exact supported ChatGPT status should authenticate.'
  Write-Host 'PASS authentication readiness retains no raw CLI output'
  function Test-OutlookRunning { return $true }
  function Test-Path { throw 'Unexpected external filesystem/registry access during closed-Outlook guard test.' }
  $rejected = $false
  try { Remove-MailyOutlookCache -AddinId '194a156f-26dd-420f-8eba-e78caa663953' | Out-Null } catch { $rejected = $_.Exception.Message -match 'Close Outlook' }
  Assert-True $rejected 'Cache mutation must reject a running Outlook before accessing registry or files.'
  Write-Host 'PASS cache mutation fails closed while Outlook is running'
  Write-Host 'PASS installer lifecycle fixtures'
}
finally { Write-Host "Retained lifecycle fixtures: $testRoot" }
} $InstallerPath $OutputDirectory
