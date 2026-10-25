param([string]$LogDirectory = (Join-Path $PSScriptRoot ('../output/bridge-logging/' + [guid]::NewGuid().ToString('N'))))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $LogDirectory | Out-Null
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'outlook-bridge.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Bridge parse failed' }
$logger = $ast.Find({param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Write-BridgeLog'}, $false)
Invoke-Expression $logger.Extent.Text
$bridgeLogPath = Join-Path $LogDirectory 'bridge.jsonl'
$bridgeCorrelationId = [guid]::NewGuid().ToString('N')
$Action = 'rewrite-active-compose'
$canary = 'SYNTHETIC_PRIVATE_CANARY'
Write-BridgeLog -EventName 'bridge_action_failed' -Status 'error' -Detail $canary -Properties @{subject=$canary; mailId=$canary; folderName=$canary; nested=@{body=$canary}; limit=50; collected=3; itemClass=26}
$Action = $canary
Write-BridgeLog -EventName $canary -Status $canary -Detail $canary -Properties @{limit=$canary; collected=-1; itemClass=$canary}
$raw = [IO.File]::ReadAllText($bridgeLogPath)
if ($raw.Contains($canary)) { throw 'Private synthetic content was persisted' }
$rows = @(Get-Content $bridgeLogPath | ForEach-Object { $_ | ConvertFrom-Json })
if ($rows.Count -ne 2 -or $rows[0].action -ne 'rewrite-active-compose' -or $rows[0].properties.limit -ne 50 -or $rows[0].properties.collected -ne 3 -or $rows[0].correlationId -ne $bridgeCorrelationId) { throw 'Safe diagnostic metadata was lost' }
if ($rows[1].action -ne 'unknown' -or $rows[1].event -ne 'unknown' -or $rows[1].status -ne 'info' -or @($rows[1].properties.PSObject.Properties).Count -ne 0) { throw 'Unknown labels or invalid counts were retained' }
$bytes = [IO.File]::ReadAllBytes($bridgeLogPath)
if ($bytes.Length -ge 3 -and $bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) { throw 'Log has UTF-8 BOM' }
'PASS bridge logging privacy, allowlisted metadata, and UTF-8 without BOM'
