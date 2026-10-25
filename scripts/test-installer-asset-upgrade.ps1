$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$run = Join-Path $repo ('output/asset-upgrade-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $run | Out-Null
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Install-MailyOutlookAddinPackage.ps1'), [ref]$tokens, [ref]$errors)
$definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-MailyManifestSignature' }, $true)
Invoke-Expression $definition.Extent.Text
function Signature([string]$url, [string]$version = '1.3.1.0') {
    $path = Join-Path $run ([guid]::NewGuid().ToString('N') + '.xml')
    [IO.File]::WriteAllText($path, "<OfficeApp><Id>194a156f-26dd-420f-8eba-e78caa663953</Id><Version>$version</Version><SourceLocation DefaultValue='$url'/></OfficeApp>", [Text.UTF8Encoding]::new($false))
    Get-MailyManifestSignature -Path $path
}
$before = Signature 'https://localhost:7097/addin/taskpane.html?v=20261023release3'
$after = Signature 'https://localhost:7097/addin/taskpane.html?v=20261025public1'
if ($before -cne $after) { throw 'A frontend-only local cache revision must not require mailbox re-registration.' }
Write-Output 'PASS local asset-only upgrade preserves mailbox compatibility'
foreach ($url in @('https://localhost:7097/addin/other.html?v=20261025public1', 'https://localhost:7098/addin/taskpane.html?v=20261025public1', 'https://localhost:7097/addin/taskpane.html?v=20261025public1&amp;surface=read')) {
    if ($before -ceq (Signature $url)) { throw 'A meaningful URL change must still require registration repair.' }
}
if ($before -ceq (Signature 'https://localhost:7097/addin/taskpane.html?v=20261023release3' '1.3.2.0')) { throw 'Version changes must remain distinct.' }
if ((Signature 'https://example.invalid/addin/taskpane.html?v=old') -ceq (Signature 'https://example.invalid/addin/taskpane.html?v=new')) { throw 'Hosted cache policy cannot be assumed.' }
$program = [IO.File]::ReadAllText((Join-Path $repo 'services/Maily.OutlookAddin.Web/Program.cs'))
if ($program -notmatch 'no-store, no-cache, must-revalidate, max-age=0') { throw 'Local compatibility requires uncached pane HTML.' }
Write-Output 'PASS meaningful URL/version and hosted changes remain distinct; HTML remains uncached'
Write-Output "Artifacts: $run"
