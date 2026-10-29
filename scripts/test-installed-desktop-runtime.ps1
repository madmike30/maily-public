param([string]$InstallerPath=(Join-Path $PSScriptRoot 'Install-MailyOutlookAddinPackage.ps1'))
$ErrorActionPreference='Stop'
$root=Join-Path (Split-Path $PSScriptRoot -Parent) ('output/desktop-runtime-tests/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $root|Out-Null
$ast=[Management.Automation.Language.Parser]::ParseFile($InstallerPath,[ref]$null,[ref]$null)
foreach($f in $ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst]},$false)){Invoke-Expression $f.Extent.Text}
function Assert([bool]$ok,[string]$why){if(!$ok){throw $why}}
function Expect-Failure([scriptblock]$Action){$failed=$false;try{&$Action|Out-Null}catch{$failed=$true};Assert $failed 'Expected fail-closed result'}
# Exercise the installer process boundary with synthetic values only. Never print environment values.
$apiKeyNames=@('OPENAI_API_KEY','CODEX_API_KEY','AZURE_OPENAI_API_KEY')
$previousApiValues=@{}
try {
  foreach($key in $apiKeyNames){$previousApiValues[$key]=[Environment]::GetEnvironmentVariable($key);[Environment]::SetEnvironmentVariable($key,'synthetic-not-a-credential')}
  $probe='if ($env:OPENAI_API_KEY -or $env:CODEX_API_KEY -or $env:AZURE_OPENAI_API_KEY) { exit 77 }; if ($env:CODEX_HOME -ne ''synthetic-chatgpt-home'') { exit 78 }; Write-Output ''subscription-environment-ok'''
  $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($probe))
  $ps=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
  $result=Invoke-CodexProcess -ExecutablePath $ps -Arguments @('-NoProfile','-EncodedCommand',$encoded) -ResolvedCodexHome 'synthetic-chatgpt-home'
  Assert ($result.started -and $result.exitCode -eq 0 -and $result.stdout.Trim() -eq 'subscription-environment-ok') 'Installer child must exclude API keys and preserve configured ChatGPT home'
  foreach($key in $apiKeyNames){Assert ([Environment]::GetEnvironmentVariable($key) -eq 'synthetic-not-a-credential') 'Child isolation changed parent environment'}
  Write-Output 'PASS installer child subscription environment excludes API keys and preserves home and parent'
}finally{foreach($key in $apiKeyNames){[Environment]::SetEnvironmentVariable($key,$previousApiValues[$key])}}
$packageRoot=Join-Path $root 'package';$payloadRoot=Join-Path $packageRoot 'payload';New-Item -ItemType Directory $payloadRoot -Force|Out-Null
Assert ((Resolve-MailyAiRuntimeMode) -eq 'Bundled') 'Legacy missing marker must remain bundled'
[IO.File]::WriteAllText((Join-Path $packageRoot 'ai-runtime-mode.txt'),'Invalid')
Expect-Failure {Resolve-MailyAiRuntimeMode}
[IO.File]::WriteAllText((Join-Path $packageRoot 'ai-runtime-mode.txt'),'InstalledApp')
$resolvedAiRuntimeMode=Resolve-MailyAiRuntimeMode
Assert ($resolvedAiRuntimeMode -eq 'InstalledApp') 'Installed app mode not read'
Write-Output 'PASS explicit mode and legacy marker handling'
$oldLocal=$env:LOCALAPPDATA
try {
$env:LOCALAPPDATA=Join-Path $root 'local';New-Item -ItemType Directory $env:LOCALAPPDATA|Out-Null
$CodexExecutablePath=Join-Path $root 'standalone.exe';[IO.File]::WriteAllText($CodexExecutablePath,'synthetic')
New-Item -ItemType Directory (Join-Path $payloadRoot 'runtime/codex') -Force|Out-Null
[IO.File]::WriteAllText((Join-Path $payloadRoot 'runtime/codex/codex.exe'),'synthetic')
function Get-Command {throw 'Installed-app discovery must not inspect PATH'}
Assert (@(Get-CodexCliCandidates).Count -eq 0) 'Standalone/bundled CLI must not qualify as desktop app'
$bin=Join-Path $env:LOCALAPPDATA 'OpenAI/Codex/bin'
New-Item -ItemType Directory (Join-Path $bin 'old'),(Join-Path $bin 'new') -Force|Out-Null
$older=Join-Path $bin 'old/codex.exe';$newer=Join-Path $bin 'new/codex.exe'
[IO.File]::WriteAllText($older,'synthetic-old');[IO.File]::WriteAllText($newer,'synthetic-new')
[IO.File]::SetLastWriteTimeUtc($older,[DateTime]::UtcNow.AddDays(-1))
$candidates=@(Get-CodexCliCandidates);Assert ($candidates[0].path -eq $newer) 'Newest installed app not preferred'
$outside=Join-Path $root 'redirected-runtime';New-Item -ItemType Directory $outside|Out-Null
[IO.File]::WriteAllText((Join-Path $outside 'codex.exe'),'synthetic-redirected')
New-Item -ItemType Junction -Path (Join-Path $bin 'redirected') -Target $outside|Out-Null
Assert (@(Get-CodexCliCandidates | Where-Object {$_.path -like '*redirected*'}).Count -eq 0) 'Redirected app version must be rejected'
$retired=Join-Path $root 'retired-runtime.exe';Move-Item -LiteralPath $newer -Destination $retired
Assert (@(Get-CodexCliCandidates)[0].path -eq $older) 'Removed app version remained selected'
Move-Item -LiteralPath $retired -Destination $newer
$script:attempts=@();$script:statusText='Logged in using an API key';$script:rejectNewest=$false;$script:versionText='codex-cli 0.164.0-alpha.11.1'
function Invoke-CodexProcess {param($ExecutablePath,$Arguments,$ResolvedCodexHome);$script:attempts+=($Arguments -join ' ');return @{started=$true;exitCode=0;stdout=if(($Arguments -join ' ') -eq '--version'){$script:versionText}elseif($script:rejectNewest -and $ExecutablePath -eq $newer){'Logged in using an API key'}else{$script:statusText};stderr='';error=''}}
$check=Resolve-RunnableCodexCli;Assert (!$check.ok -and $check.message -match 'desktop') 'App presence/API key is not ChatGPT sign-in'
foreach($mixed in @("Logged in using an API key`nChatGPT is available", "Logged in using ChatGPT`nWarning: diagnostic output", 'Not logged in using ChatGPT')) {
  $script:statusText=$mixed;$check=Resolve-RunnableCodexCli;Assert (!$check.ok) 'Mixed/negated authentication diagnostics must fail closed'
}
$script:statusText="  logged IN using chatgpt `r`n";$check=Resolve-RunnableCodexCli;Assert ($check.ok -and $check.codexCommandPath -eq $newer) 'Trimmed case-insensitive exact ChatGPT status not accepted'
$script:rejectNewest=$true;$script:attempts=@();$check=Resolve-RunnableCodexCli
Assert (!$check.ok -and @($script:attempts|Where-Object {$_ -eq 'login status'}).Count -eq 1) 'Newest app failure must not fall back to an older signed-in runtime'
$script:rejectNewest=$false
foreach($unsupported in @('codex-cli 0.151.0-alpha.1','codex-cli 0.150.9','codex-cli unknown','codex-cli 0.164.0 extra diagnostics')) {
  $script:versionText=$unsupported;$script:attempts=@();$check=Resolve-RunnableCodexCli
  Assert (!$check.ok -and $script:attempts.Count -eq 1 -and $script:attempts[0] -eq '--version') 'Unsupported/malformed runtime must stop before sign-in probe'
}
$script:versionText='codex-cli 0.151.0';$check=Resolve-RunnableCodexCli;Assert $check.ok 'Minimum supported runtime rejected'
$script:versionText='codex-cli 0.164.0-alpha.11.1';$check=Resolve-RunnableCodexCli;Assert $check.ok 'Supported prerelease runtime rejected'
Assert (@($script:attempts|Where-Object {$_ -notin @('login status','--version')}).Count -eq 0) 'Only read-only status/version probes may run'
Write-Output 'PASS installed-app-only discovery, newest version and sign-in checks'
$isHostedPackage=$false;$required=@(Get-MailyRequiredPayloadFiles);Assert ($required -notcontains 'runtime\codex\codex.exe') 'App mode must not require bundled CLI'
$resolvedAiRuntimeMode='Bundled';Assert (@(Get-MailyRequiredPayloadFiles) -contains 'runtime\codex\codex.exe') 'Legacy/bundled CLI requirement lost'
$isHostedPackage=$true;Assert (@(Get-MailyRequiredPayloadFiles).Count -eq 1) 'Hosted payload contract changed'
Write-Output 'PASS per-mode payload requirements'
$resolvedAiRuntimeMode='InstalledApp';$resolvedCodexExecutablePath='installed-app';$installRootPath=Join-Path $root 'installed';$healthUrl='https://localhost:7097/api/health'
New-Item -ItemType Directory (Join-Path $installRootPath 'app/Maily.OutlookAddin.Web'),(Join-Path $installRootPath 'app/Maily.OutlookAddin.Supervisor') -Force|Out-Null
Write-InstalledConfiguration -ResolvedCodexHome 'synthetic-home' -Certificate @{pfxPath='synthetic';password='synthetic'}|Out-Null
$config=Get-Content (Join-Path $installRootPath 'app/Maily.OutlookAddin.Web/appsettings.Production.json') -Raw|ConvertFrom-Json
Assert ($config.MailyGeneration.CodexExecutablePath -eq 'installed-app') 'Runtime resolver sentinel missing'
Assert ((Get-MailyManagedNames) -contains 'ai-runtime-mode.txt') 'Mode marker must participate in rollback'
Write-Output 'PASS dynamic runtime configuration and rollback ownership'
}finally{$env:LOCALAPPDATA=$oldLocal}
Write-Output 'RESULT installed desktop runtime checks passed; no live host access'
