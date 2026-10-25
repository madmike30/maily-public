param([string]$OutputDirectory = '')
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$root = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $repo ('output/package-retention/' + [Guid]::NewGuid().ToString('N')) }
New-Item -ItemType Directory -Path $root | Out-Null
$version = ([xml](Get-Content (Join-Path $repo 'addin/manifest.xml') -Raw)).OfficeApp.Version
$existing = Join-Path $root "MailyOutlookAddin-$version"
New-Item -ItemType Directory -Path $existing | Out-Null
$evidence = Join-Path $existing 'evidence.txt'
[IO.File]::WriteAllText($evidence, 'retained evidence')
function dotnet { throw 'Unexpected publish: existing package should be rejected first.' }
$caught = ''
try { & (Join-Path $PSScriptRoot 'build-maily-outlook-installer.ps1') -OutputRoot $root -NoZip } catch { $caught = $_.Exception.Message }
if (-not (Test-Path -LiteralPath $evidence)) { throw 'Build deleted an existing package artifact.' }
if ($caught -notmatch 'already exists') { throw "Existing package was not rejected: $caught" }
Write-Host 'PASS build refuses existing package and retains evidence'
$destination = Join-Path $root 'runtime'
function Invoke-WebRequest { param($Uri, $OutFile) [IO.File]::WriteAllText($OutFile, 'synthetic corrupt download') }
$caught = ''
try { & (Join-Path $PSScriptRoot 'Bundle-MailyCodexRuntime.ps1') -Destination $destination } catch { $caught = $_.Exception.Message }
if ($caught -notmatch 'SHA256 mismatch') { throw "Expected integrity rejection: $caught" }
if (Test-Path (Join-Path $destination 'codex.exe')) { throw 'Corrupt executable remained in runtime.' }
if (-not (Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object Name -Like '*.rejected-*')) { throw 'Corrupt download evidence was deleted.' }
Write-Host 'PASS corrupt download is quarantined outside the package'
Write-Host "Retained package fixtures: $root"
