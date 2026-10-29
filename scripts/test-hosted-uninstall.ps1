param([string]$UninstallerPath = (Join-Path $PSScriptRoot 'Uninstall-MailyOutlookAddinPackage.ps1'))
& (New-Module {}) {
  param($UninstallerPath)
  $ErrorActionPreference = 'Stop'
  $repo = Split-Path (Split-Path $UninstallerPath -Parent) -Parent
  $testRoot = Join-Path $repo ('output/hosted-uninstall/' + [guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Path $testRoot | Out-Null
  $tokens=$null; $errors=$null
  $ast=[Management.Automation.Language.Parser]::ParseFile($UninstallerPath,[ref]$tokens,[ref]$errors)
  if ($errors.Count) { throw 'Uninstaller parsing failed.' }
  foreach ($name in @('Assert-SafeInstallRoot','Assert-ChildPath','Remove-InstallFiles')) {
    $definition=$ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$false) | Where-Object Name -EQ $name | Select-Object -First 1
    if (-not $definition) { throw "Missing uninstaller function $name" }
    Invoke-Expression $definition.Extent.Text
  }
  # Exercise real file-removal decisions while preserving every synthetic file.
  function Remove-Item {
    [CmdletBinding()]param([Parameter(Mandatory=$true)][string]$LiteralPath,[switch]$Recurse,[switch]$Force)
    $target=[IO.Path]::GetFullPath($LiteralPath)
    $root=[IO.Path]::GetFullPath($testRoot).TrimEnd([char[]]'\/')
    if (-not $target.StartsWith($root + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Removal escaped fixture root.' }
    $retained=Join-Path $testRoot ('retained/' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $retained -Force | Out-Null
    Move-Item -LiteralPath $target -Destination (Join-Path $retained ([IO.Path]::GetFileName($target)))
  }
  $installRootPath=Join-Path $testRoot 'installed'
  foreach ($name in @('runtime/atk','scripts','addin','setup','data','logs','app','unrelated')) {
    $dir=Join-Path $installRootPath $name
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $dir 'sentinel.txt'),'synthetic evidence',[Text.UTF8Encoding]::new($false))
  }
  foreach ($name in @('MailyOutlookAddinSetup.exe','package-mode.txt','ai-runtime-mode.txt','package-summary.json')) { [IO.File]::WriteAllText((Join-Path $installRootPath $name),'synthetic') }
  $isHostedPackage=$true; $KeepUserData=$true; $KeepLogs=$true; $Force=$true; $packageModeFileName='package-mode.txt'
  $result=Remove-InstallFiles
  foreach ($name in @('runtime','scripts','addin','setup','MailyOutlookAddinSetup.exe','package-mode.txt','ai-runtime-mode.txt','package-summary.json')) {
    if (Test-Path -LiteralPath (Join-Path $installRootPath $name)) { throw "Hosted uninstall left managed payload: $name" }
    if ($result.removedNames -notcontains $name) { throw "Removal outcome omitted $name" }
  }
  foreach ($name in @('data','logs','app','unrelated')) {
    if (([IO.File]::ReadAllText((Join-Path $installRootPath "$name/sentinel.txt"))) -ne 'synthetic evidence') { throw "Hosted uninstall changed retained $name" }
  }
  Write-Output 'PASS hosted uninstall removes connector payload, preserving settings/logs and non-hosted files'
  Write-Output "Artifacts: $testRoot"
} $UninstallerPath