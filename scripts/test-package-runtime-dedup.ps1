param([string]$OutputRoot = '')
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
if(!$OutputRoot){$OutputRoot=Join-Path $repo ('output/runtime-dedup/'+[guid]::NewGuid().ToString('N'))}
New-Item -ItemType Directory -Path $OutputRoot | Out-Null
$fixture=Join-Path $OutputRoot 'app'
foreach($name in @('Maily.OutlookAddin.Web','Maily.OutlookAddin.Supervisor')){New-Item -ItemType Directory (Join-Path $fixture $name)|Out-Null; [IO.File]::WriteAllText((Join-Path $fixture "$name/shared.dll"),'synthetic same runtime');[IO.File]::WriteAllText((Join-Path $fixture "$name/unique.dll"),$name)}
. (Join-Path $PSScriptRoot 'MailyPackageRuntime.ps1')
$evidence=Join-Path $OutputRoot 'retained'
$before=Get-FileHash (Join-Path $fixture 'Maily.OutlookAddin.Supervisor/shared.dll')
$result=Compress-MailyPackageRuntime -AppRoot $fixture -EvidenceRoot $evidence
if($result.files -ne 1 -or (Test-Path (Join-Path $fixture 'Maily.OutlookAddin.Supervisor/shared.dll'))){throw 'Duplicate not compacted'}
Test-MailyPackageRuntime -AppRoot $fixture | Out-Null
Restore-MailyPackageRuntime -AppRoot $fixture
if((Get-FileHash (Join-Path $fixture 'Maily.OutlookAddin.Supervisor/shared.dll')).Hash -ne $before.Hash){throw 'Restored runtime differs'}
if(!(Test-Path (Join-Path $evidence 'shared.dll'))){throw 'Original evidence lost'}
'PASS compact and restore identical bytes'
function New-MapFixture([string]$Name) {
  $root=Join-Path $OutputRoot $Name
  foreach($part in @('Maily.OutlookAddin.Web','Maily.OutlookAddin.Supervisor')){New-Item -ItemType Directory (Join-Path $root $part) -Force|Out-Null}
  $source=Join-Path $root 'Maily.OutlookAddin.Web/shared.dll'
  [IO.File]::WriteAllText($source,'synthetic same runtime')
  $script:entry=@{path='shared.dll';bytes=(Get-Item $source).Length;sha256=(Get-FileHash $source).Hash}
  return $root
}
function Write-Map([string]$Root,$Entries) {[IO.File]::WriteAllText((Join-Path $Root 'runtime-file-map.json'),(@{version=1;files=@($Entries)}|ConvertTo-Json -Depth 6))}
function Reject([string]$Name,[scriptblock]$Setup) {
  $root=New-MapFixture $Name
  & $Setup $root
  $rejected=$false
  try {Restore-MailyPackageRuntime -AppRoot $root} catch {$rejected=$true}
  if(!$rejected){throw "Accepted unsafe map: $Name"}
  "PASS rejects $Name"
}
$old=New-MapFixture 'old-package'
Restore-MailyPackageRuntime -AppRoot $old
'PASS legacy package without map'
Reject 'traversal' {param($root) $entry.path='../shared.dll'; Write-Map $root @($entry)}
Reject 'rooted' {param($root) $entry.path='C:/shared.dll'; Write-Map $root @($entry)}
Reject 'ads' {param($root) $entry.path='shared.dll:alternate.dll'; Write-Map $root @($entry)}
Reject 'hash-mismatch' {param($root) $entry.sha256='0'*64; Write-Map $root @($entry)}
Reject 'size-mismatch' {param($root) $entry.bytes++; Write-Map $root @($entry)}
Reject 'duplicate' {param($root) $second=$entry.Clone();$second.path='SHARED.dll';Write-Map $root @($entry,$second)}
Reject 'missing-source' {param($root) $entry.path='missing.dll';Write-Map $root @($entry)}
Reject 'collision' {param($root) [IO.File]::WriteAllText((Join-Path $root 'Maily.OutlookAddin.Supervisor/shared.dll'),'preserve');Write-Map $root @($entry)}
Reject 'prevalidate-all' {param($root) $second=$entry.Clone();$second.path='../bad.dll';Write-Map $root @($entry,$second)}
if(Test-Path (Join-Path $OutputRoot 'prevalidate-all/Maily.OutlookAddin.Supervisor/shared.dll')){throw 'Wrote before validating all entries'}
Reject 'source-junction' {param($root) $outside=Join-Path $OutputRoot 'junction-content';New-Item -ItemType Directory $outside|Out-Null;[IO.File]::WriteAllText((Join-Path $outside 'shared.dll'),'synthetic same runtime'); New-Item -ItemType Junction -Path (Join-Path $root 'Maily.OutlookAddin.Web/nested') -Target $outside|Out-Null;$entry.path='nested/shared.dll';Write-Map $root @($entry)}
Reject 'target-junction' {param($root) $outside=Join-Path $OutputRoot 'target-content';New-Item -ItemType Directory $outside|Out-Null;New-Item -ItemType Directory (Join-Path $root 'Maily.OutlookAddin.Web/nested')|Out-Null;Copy-Item (Join-Path $root 'Maily.OutlookAddin.Web/shared.dll') (Join-Path $root 'Maily.OutlookAddin.Web/nested/shared.dll');New-Item -ItemType Junction -Path (Join-Path $root 'Maily.OutlookAddin.Supervisor/nested') -Target $outside|Out-Null;$entry.path='nested/shared.dll';Write-Map $root @($entry)}
'RESULT: runtime package sharing passed'
