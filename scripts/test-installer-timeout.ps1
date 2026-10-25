param([string]$InstallerPath = (Join-Path $PSScriptRoot 'Install-MailyOutlookAddinPackage.ps1'))
$ErrorActionPreference = 'Stop'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($InstallerPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
foreach ($name in @('Quote-CodexCommandArgument', 'Invoke-CodexProcess')) {
    $function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $false)
    Invoke-Expression $function.Extent.Text
}
$root = Join-Path (Split-Path $PSScriptRoot -Parent) ('output/installer-timeout/' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
$pidPath = Join-Path $root 'owned-process.pid'
$fixture = Join-Path $root 'sleep.ps1'
Set-Content -LiteralPath $fixture -Value ('Set-Content -LiteralPath ''' + $pidPath.Replace("'", "''") + ''' -Value $PID; Start-Sleep -Seconds 60')
$owned = $null
try {
    $result = Invoke-CodexProcess -ExecutablePath (Join-Path $PSHOME 'powershell.exe') -Arguments @('-NoProfile', '-File', $fixture)
    if (-not (Test-Path -LiteralPath $pidPath)) { throw 'Synthetic process did not start.' }
    $owned = Get-Process -Id ([int](Get-Content -LiteralPath $pidPath)) -ErrorAction SilentlyContinue
    $result | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'result.json')
    if ($null -ne $owned) { throw 'Timed-out owned process is still alive.' }
    if ($result.error -ne 'Codex CLI timed out.') { throw 'Timeout result was not preserved.' }
    Write-Host 'PASS Windows PowerShell installer timeout stops the owned synthetic process'
}
finally {
    if ($null -ne $owned -and -not $owned.HasExited) { $owned.Kill(); $owned.WaitForExit() }
    Write-Host "Retained timeout fixture: $root"
}
