param([string]$LogDirectory = (Join-Path $PSScriptRoot ('../output/bridge-tests/' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $LogDirectory | Out-Null
Start-Transcript -Path (Join-Path $LogDirectory 'synthetic.log') | Out-Null
$failures = 0
$tokens = $null; $errors = $null
$bridgePath = Join-Path $PSScriptRoot 'outlook-bridge.ps1'
$ast = [System.Management.Automation.Language.Parser]::ParseFile($bridgePath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
$ast.FindAll({param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $false) | ForEach-Object { Invoke-Expression $_.Extent.Text }
function Check([string]$Name, [scriptblock]$Test) {
  try { & $Test; Write-Output "PASS $Name" } catch { $script:failures++; Write-Output "FAIL $Name : $($_ | Out-String) $($_.ScriptStackTrace)" }
}
function Assert($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Assert-Throws([scriptblock]$Code) { $caught=$false; try { & $Code | Out-Null } catch { $caught=$true }; Assert $caught 'Expected fail-closed exception' }
function New-Inspector([string]$Subject, [string]$Text) {
  $editor = [pscustomobject]@{ Content = [pscustomobject]@{ Text=$Text } }
  [pscustomobject]@{ CurrentItem=[pscustomobject]@{ Class=43; Subject=$Subject }; WordEditor=$editor }
}
function New-Outlook($Active, $Others=@()) {
  $inspectors = [pscustomobject]@{ Count=$Others.Count; Values=@($Others) }
  $inspectors | Add-Member ScriptMethod Item {param($index) $this.Values[$index-1]}
  $outlook = [pscustomobject]@{ Active=$Active; Inspectors=$inspectors }
  $outlook | Add-Member ScriptMethod ActiveInspector { $this.Active }
  $outlook | Add-Member ScriptMethod ActiveExplorer { $null }
  $outlook
}
Check 'Wrong subject active inspector is rejected' {
  $outlook=New-Outlook (New-Inspector 'Unrelated' 'unrelated draft')
  Assert-Throws { Find-ComposeTarget -Outlook $outlook -ExpectedSubject 'Expected' -ExpectedSeedText 'expected draft' }
}
Check 'Wrong seed active inspector is rejected' {
  $outlook=New-Outlook (New-Inspector 'Expected' 'unrelated draft')
  Assert-Throws { Find-ComposeTarget -Outlook $outlook -ExpectedSubject 'Expected' -ExpectedSeedText 'expected draft' }
}
Check 'Two matching distinct drafts are rejected' {
  $one=New-Inspector 'Expected' 'expected draft'; $two=New-Inspector 'Expected' 'expected draft'
  $outlook=New-Outlook $one @($one,$two)
  Assert-Throws { Find-ComposeTarget -Outlook $outlook -ExpectedSubject 'Expected' -ExpectedSeedText 'expected draft' }
}
Check 'Same inspector enumerated twice remains usable' {
  $one=New-Inspector 'Expected' 'expected draft'
  $target=Find-ComposeTarget -Outlook (New-Outlook $one @($one)) -ExpectedSubject 'Expected' -ExpectedSeedText 'expected draft'
  Assert ([object]::ReferenceEquals($target.Item,$one.CurrentItem)) 'Wrong target'
}
Check 'Ambiguous same-subject drafts without a seed are rejected' {
  $one=New-Inspector 'Expected' 'first'; $two=New-Inspector 'Expected' 'second'
  Assert-Throws { Find-ComposeTarget -Outlook (New-Outlook $one @($one,$two)) -ExpectedSubject 'Expected' -ExpectedSeedText '' }
}
Check 'Sent items cannot be rewrite targets' {
  $one=New-Inspector 'Expected' 'expected draft'
  $one.CurrentItem | Add-Member NoteProperty Sent $true
  Assert-Throws { Find-ComposeTarget -Outlook (New-Outlook $one) -ExpectedSubject 'Expected' -ExpectedSeedText 'expected draft' }
}
Check 'Compose target must belong to the requested store' {
  $one=New-Inspector 'Expected' 'expected draft'
  $one.CurrentItem | Add-Member NoteProperty Parent ([pscustomobject]@{StoreID='secondary'})
  Assert-Throws { Find-ComposeTarget -Outlook (New-Outlook $one) -ExpectedSubject 'Expected' -ExpectedSeedText 'expected draft' -ExpectedStoreId 'primary' }
  $target=Find-ComposeTarget -Outlook (New-Outlook $one) -ExpectedSubject 'Expected' -ExpectedSeedText 'expected draft' -ExpectedStoreId 'secondary'
  Assert ([object]::ReferenceEquals($target.Item,$one.CurrentItem)) 'Secondary compose not found'
}
Check 'Long nonmatching seed never permits a target' {
  Assert-Throws { Find-ComposeTarget -Outlook (New-Outlook (New-Inspector 'Expected' 'different draft')) -ExpectedSubject 'Expected' -ExpectedSeedText ('x' * 300) }
}
Check 'Matching 80-character prefix cannot authorize conflicting draft text' {
  $prefix='x' * 80
  $one=New-Inspector 'Expected' ($prefix + ' different continuation')
  Assert-Throws { Find-ComposeTarget -Outlook (New-Outlook $one) -ExpectedSubject 'Expected' -ExpectedSeedText ($prefix + ' expected continuation') }
}
Check 'Exact normalized multiline seed inside body and signature stays eligible' {
  $one=New-Inspector 'Expected' "Expected  multiline`r`n  draft text`r`nRegards`r`nSynthetic User"
  $target=Find-ComposeTarget -Outlook (New-Outlook $one) -ExpectedSubject 'Expected' -ExpectedSeedText "Expected multiline`n draft   text"
  Assert ([object]::ReferenceEquals($target.Item,$one.CurrentItem)) 'Full normalized seed was rejected'
}
function Invoke-SyntheticRewrite([string]$Text, [string]$Seed) {
  $script:mutationWrites=0
  $content=[pscustomobject]@{Text=$Text;Start=0}
  $paragraphs=@(); $offset=0
  foreach($line in ($Text -split "`r`n|`r|`n")) {
    $paragraphs += [pscustomobject]@{Range=[pscustomobject]@{Text=($line+"`r");Start=$offset}}
    $offset += $line.Length+1
  }
  $collection=[pscustomobject]@{Count=$paragraphs.Count;Values=$paragraphs}
  $collection | Add-Member ScriptMethod Item {param($index) $this.Values[$index-1]}
  $editor=[pscustomobject]@{Content=$content;Paragraphs=$collection;Application=[pscustomobject]@{UndoRecord=$null}}
  $editor | Add-Member ScriptMethod Range {
    param($start,$end)
    $range=[pscustomobject]@{}
    $range | Add-Member ScriptProperty Text {''} {$script:mutationWrites++}
    return $range
  }
  $inspector=[pscustomobject]@{CurrentItem=[pscustomobject]@{Class=43;Subject='Expected';Sent=$false;Parent=[pscustomobject]@{StoreID='fixture-store'}};WordEditor=$editor}
  $outlook=New-Outlook $inspector
  $outlook | Add-Member NoteProperty FixtureSession ([pscustomobject]@{Accounts=@([pscustomobject]@{SmtpAddress='test@example.invalid';DeliveryStore=[pscustomobject]@{StoreID='fixture-store'}})})
  $outlook | Add-Member ScriptMethod GetNamespace {param($name) $this.FixtureSession}
  function New-Object {param([string]$ComObject) if($ComObject -ne 'Outlook.Application') {throw 'Unexpected COM type'}; return $outlook}
  Rewrite-ActiveCompose -ExpectedSubject 'Expected' -HtmlSourcePath '' -ReplacementText 'replacement' -SeedPreview $Seed -DisplayName 'Synthetic User' -EmailAddress 'test@example.invalid' -Mode 'reply'
}
Check 'Mutation refuses a seed found only in quoted history' {
  Assert-Throws { Invoke-SyntheticRewrite "Other draft`rFrom: sender@example.invalid`rExpected draft" 'Expected draft' }
  Assert ($script:mutationWrites -eq 0) 'Quoted seed authorized a write'
}
Check 'Mutation refuses edits appended to the seed paragraph' {
  Assert-Throws { Invoke-SyntheticRewrite 'Expected draft newly typed words' 'Expected draft' }
  Assert ($script:mutationWrites -eq 0) 'Same-paragraph edits were overwritten'
}
Check 'Mutation refuses edits appended as another paragraph' {
  Assert-Throws { Invoke-SyntheticRewrite "Expected draft`rNewly typed paragraph" 'Expected draft' }
  Assert ($script:mutationWrites -eq 0) 'Extra-paragraph edits were overwritten'
}
Check 'Mutation supports full drafts beyond 1024 characters' {
  $seed=('Long synthetic draft words ' * 60).Trim()
  Invoke-SyntheticRewrite ($seed+"`rRegards`rSynthetic User") $seed | Out-Null
  Assert ($script:mutationWrites -eq 1) 'Long unchanged draft was not rewritten'
}
Check 'Mutation supports normalized multiline seed and preserves signature' {
  Invoke-SyntheticRewrite "Expected  multiline`rdraft text`rRegards`rSynthetic User" "Expected multiline`n draft   text" | Out-Null
  Assert ($script:mutationWrites -eq 1) 'Valid multiline draft was not rewritten'
}
Check 'Mutation refuses new draft content when the expected seed was empty' {
  Assert-Throws { Invoke-SyntheticRewrite 'Newly typed draft' '' }
  Assert ($script:mutationWrites -eq 0) 'New draft text was overwritten'
}
Check 'Mutation permits an unchanged empty draft' {
  Invoke-SyntheticRewrite '' '' | Out-Null
  Assert ($script:mutationWrites -eq 1) 'Empty draft insertion failed'
}
Check 'Mutation validates drafts beyond twelve paragraphs' {
  $seed=(1..15 | ForEach-Object {"Synthetic paragraph $_"}) -join "`r"
  Invoke-SyntheticRewrite ($seed+"`rRegards`rSynthetic User") $seed | Out-Null
  Assert ($script:mutationWrites -eq 1) 'Complete multi-paragraph draft was rejected'
}

Write-Output "RESULT failures=$failures"
Stop-Transcript | Out-Null
if ($failures) { exit 1 }
