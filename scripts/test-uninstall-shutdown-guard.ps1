$ErrorActionPreference='Stop'
$path=Join-Path $PSScriptRoot 'Uninstall-MailyOutlookAddinPackage.ps1'
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$null,[ref]$null)
$step=$ast.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-UninstallerStep'},$true)
& (New-Module {}) {
 param($step)
 Invoke-Expression $step.Extent.Text
 function Add-Step {param($Name,$Status,$Detail,$Data)}
 $Force=$true;$mutated=$false;$failed=$false
 try {Invoke-UninstallerStep -Name outlook-close -ScriptBlock {throw 'Synthetic cancelled close'}; $mutated=$true} catch {$failed=$true}
 if(-not $failed -or $mutated){throw 'Force bypassed shutdown failure and allowed subsequent mutation'}
 Write-Output 'PASS forced uninstall still stops after shutdown failure'
 $failed=$false;try {Invoke-UninstallerStep -Name optional-step -ScriptBlock {throw 'Synthetic optional failure'}}catch{$failed=$true}
 if($failed){throw 'Force compatibility changed for unrelated steps'}
 Write-Output 'PASS unrelated Force step compatibility preserved'
} $step
