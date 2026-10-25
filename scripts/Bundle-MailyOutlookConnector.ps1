param([Parameter(Mandatory=$true)][string]$Destination, [string]$Runtime='win-x64', [string]$CacheDirectory)
$ErrorActionPreference='Stop'
if ($Runtime -ne 'win-x64') { throw 'The Outlook connector is pinned for win-x64.' }
if (-not $CacheDirectory) { $CacheDirectory=Join-Path (Split-Path $PSScriptRoot -Parent) 'output/outlook-connect-downloads' }
& node (Join-Path $PSScriptRoot 'bundle-outlook-connect.cjs') $Destination $CacheDirectory
if ($LASTEXITCODE -ne 0) { throw 'Pinned Outlook connector bundling failed.' }
