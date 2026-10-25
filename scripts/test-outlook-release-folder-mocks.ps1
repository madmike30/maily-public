$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path $PSScriptRoot 'test-outlook-release-folder.ps1'
$result = & $scriptPath -SessionFactory { throw 'Factory must not run in dry-run.' }
if ($result.Mode -ne 'DryRun' -or $result.Passed -ne 0) { throw 'Expected a side-effect-free dry-run.' }
Write-Output 'PASS: dry-run does not instantiate Outlook.'
class ReleaseItemMock {
    [string]$EntryID = 'synthetic-item'
    [string]$Subject
    [string]$Body
    [int]$BodyFormat = 2
    [bool]$Sent = $false
    [object]$Recipients = @{ Count = 0 }
    [object]$Parent
    [bool]$Saved = $false
    [object]$SaveParent
    [scriptblock]$AfterSave
    [scriptblock]$MoveTransform
    [int]$MoveCount = 0
    [object]$MoveDestination
    [object]$MovedItem
    [void] Save() {
        # Outlook's default HTML mode can rewrite a plain Body assignment.
        # The Unicode fixture must opt into plain text before saving it.
        if ($this.BodyFormat -ne 1) { $this.Body += "`r`n" }
        $this.Saved = $true
        if ($null -ne $this.SaveParent) { $this.Parent = $this.SaveParent }
        if ($this.AfterSave) { & $this.AfterSave $this | Out-Null }
    }
    [object] Move([object]$destination) {
        $this.MoveCount++
        $this.MoveDestination = $destination
        if (-not $this.Saved) { throw 'Move before saving.' }
        $moved = [ReleaseItemMock]::new()
        $moved.EntryID = 'moved-synthetic-item'
        $moved.Subject = $this.Subject; $moved.Body = $this.Body
        $moved.BodyFormat = $this.BodyFormat
        $moved.Sent = $this.Sent; $moved.Recipients = @{ Count = $this.Recipients.Count }
        $moved.Parent = $destination; $moved.Saved = $true
        $result = if ($this.MoveTransform) { & $this.MoveTransform $moved } else { $moved }
        $this.MovedItem = $result
        return $result
    }
}
class ReleaseItemsMock {
    [ReleaseItemMock]$Item
    [object] Add([string]$kind) {
        if ($kind -ne 'IPM.Note') { throw 'Wrong item class.' }
        return $this.Item
    }
}
class ReleaseFoldersMock {
    [object]$Folder
    [string]$CreatedName
    [object] Add([string]$name) { $this.CreatedName = $name; return $this.Folder }
}
class ReleaseStoreMock {
    [string]$StoreID = 'synthetic-store'
    [object]$Root
    [object] GetRootFolder() { return $this.Root }
}
class ReleaseSessionMock {
    [object[]]$Accounts
    [ReleaseItemMock]$Item
    [bool]$WrongTarget = $false
    [scriptblock]$ReopenTransform
    [object]$Reopened
    [bool]$AlternateIds = $false
    [int]$ComparisonCount = 0
    [int]$ThrowComparison = 0
    [int]$RejectComparison = 0
    [string]$LastReopenId
    [bool] CompareEntryIDs([string]$first, [string]$second) {
        $this.ComparisonCount++
        if ($this.ThrowComparison -eq $this.ComparisonCount) { throw 'sensitive-marker' }
        if ($this.RejectComparison -eq $this.ComparisonCount) { return $false }
        if ($first -ceq $second) { return $true }
        return ($first -ceq 'alternate-item' -and $second -ceq 'synthetic-item') -or ($first -ceq 'alternate-folder' -and $second -ceq 'synthetic-folder')
    }
    [object] GetItemFromID([string]$id, [string]$store) {
        $source = if ($null -ne $this.Item.MovedItem) { $this.Item.MovedItem } else { $this.Item }
        if ($id -cne $source.EntryID -or $store -ne 'synthetic-store' -or -not $source.Saved) { throw 'Unscoped, stale or premature reopen.' }
        $this.LastReopenId = $id
        $parent = [pscustomobject]@{ EntryID = $source.Parent.EntryID; StoreID = $source.Parent.StoreID }
        $copy = [pscustomobject]@{ EntryID = $source.EntryID; Subject = $source.Subject; Parent = $parent; Body = $source.Body; Sent = $source.Sent; Recipients = @{ Count = $source.Recipients.Count } }
        if ($this.WrongTarget) { $copy.EntryID = 'other-item' }
        if ($this.AlternateIds) { $copy.EntryID = 'alternate-item'; $copy.Parent.EntryID = 'alternate-folder' }
        if ($this.ReopenTransform) { $copy = & $this.ReopenTransform $copy }
        $this.Reopened = $copy
        return $copy
    }
}
function New-ReleaseMock {
    $item = [ReleaseItemMock]::new()
    $items = [ReleaseItemsMock]::new(); $items.Item = $item
    $folder = [pscustomobject]@{ EntryID = 'synthetic-folder'; StoreID = 'synthetic-store'; Items = $items }
    $item.Parent = $folder
    $folders = [ReleaseFoldersMock]::new(); $folders.Folder = $folder
    $store = [ReleaseStoreMock]::new(); $store.Root = @{ Folders = $folders }
    $session = [ReleaseSessionMock]::new(); $session.Item = $item
    $session.Accounts = @([pscustomobject]@{ SmtpAddress = 'fixture@example.test'; DeliveryStore = $store })
    return $session
}
$mock = New-ReleaseMock
$result = & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock }
if ($result.Passed -ne 4 -or -not $result.FolderRetained -or $result.Mode -ne 'Live') { throw 'Expected four roundtrip checks.' }
if ($mock.Accounts[0].DeliveryStore.Root.Folders.CreatedName -notmatch '^Maily-Release-Test-[a-f0-9]{32}$') { throw 'Expected unique dedicated folder.' }
if ($mock.Item.Body -notmatch 'שלום' -or -not $mock.Item.Saved) { throw 'Unicode fixture was not saved.' }
if ($mock.Item.BodyFormat -ne 1) { throw 'Unicode plain-text fixture must explicitly select plain-text format.' }
if ([object]::ReferenceEquals($mock.Item, $mock.Reopened) -or [object]::ReferenceEquals($mock.Item.Parent, $mock.Reopened.Parent)) { throw 'Reopen must use independent item and parent objects.' }
Write-Output 'PASS: synthetic unsent Unicode item roundtrip is scoped to its created folder and store.'
$mock = New-ReleaseMock
$mock.Item.SaveParent = [pscustomobject]@{ EntryID = 'synthetic-default-drafts'; StoreID = 'synthetic-store' }
$result = & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock }
if ($result.Passed -ne 4 -or $mock.Item.MoveCount -ne 1 -or -not [object]::ReferenceEquals($mock.Item.MoveDestination, $mock.Accounts[0].DeliveryStore.Root.Folders.Folder) -or $mock.LastReopenId -cne 'moved-synthetic-item') { throw 'Default-saved owned draft must move once to the created folder and reopen using returned identity.' }
Write-Output 'PASS: a newly created synthetic draft saved in the default folder moves once and reopens by its returned ID.'
$mock = New-ReleaseMock; $mock.AlternateIds = $true
$result = & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock }
if ($result.Passed -ne 4 -or $mock.ComparisonCount -ne 3) { throw 'Equivalent entry IDs must be validated by Outlook identity comparison.' }
Write-Output 'PASS: alternate item and folder IDs are accepted only when Outlook confirms semantic identity.'
function Assert-Rejected([scriptblock]$Action) {
    $failed = $false
    try { & $Action | Out-Null } catch {
        $failed = $true
        if ($_.Exception.Message -match 'fixture@example|synthetic-store|synthetic-item|sensitive-marker') { throw 'Error exposed private metadata.' }
    }
    if (-not $failed) { throw 'Unsafe fixture request was accepted.' }
}
Assert-Rejected { & $scriptPath -Live -SessionFactory { throw 'Factory must not run without explicit scope.' } }
Write-Output 'PASS: explicit domain is required before Outlook access.'
foreach ($count in @(0, 2)) {
    $mock = New-ReleaseMock
    $account = $mock.Accounts[0]
    $mock.Accounts = if ($count -eq 0) { @() } else { @($account, $account) }
    Assert-Rejected { & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock } }
    if ($account.DeliveryStore.Root.Folders.CreatedName) { throw 'Ambiguous account caused folder creation.' }
}
Write-Output 'PASS: absent and ambiguous accounts fail before folder creation.'
$mock = New-ReleaseMock
Assert-Rejected { & $scriptPath -Live -AccountDomain other.example.test -SessionFactory { $mock } }
if ($mock.Accounts[0].DeliveryStore.Root.Folders.CreatedName) { throw 'Domain suffix was incorrectly accepted.' }
Write-Output 'PASS: account domain matching is exact.'
$mock = New-ReleaseMock; $mock.WrongTarget = $true
Assert-Rejected { & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock } }
Write-Output 'PASS: reopened item identity mismatch fails closed.'
$mock = New-ReleaseMock; $mock.Item.Recipients.Count = 1
Assert-Rejected { & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock } }
if ($mock.Item.Saved) { throw 'Recipient-bearing item was saved.' }
Write-Output 'PASS: recipient-bearing draft is rejected before saving.'
Assert-Rejected { & $scriptPath -Live -AccountDomain example.test -SessionFactory { throw 'sensitive-marker' } }
Write-Output 'PASS: COM errors are anonymous.'
function Assert-FailureStage([scriptblock]$Action, [string]$ExpectedStage) {
    $message = ''
    try { & $Action | Out-Null } catch { $message = $_.Exception.Message }
    $expected = "Outlook release fixture check failed at stage '$ExpectedStage'. Any created synthetic folder and item were retained. No existing mail was inspected."
    if ($message -cne $expected) { if ($message -match "stage '([a-z-]+)'") { throw ("Expected fixed stage {0}; actual fixed stage {1}." -f $ExpectedStage, $Matches[1]) }; throw 'Expected anonymous fixed failure stage was not reported.' }
}
Assert-FailureStage { & $scriptPath -Live -AccountDomain example.test -SessionFactory { throw 'sensitive-marker' } } 'session'
$mock = New-ReleaseMock; $mock.Accounts = @()
Assert-FailureStage { & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock } } 'account-selection'
$mock = New-ReleaseMock; $mock.WrongTarget = $true
Assert-FailureStage { & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock } } 'verify-item-id'
Write-Output 'PASS: failures identify only fixed stages without private exception data.'
$identityCases = @(
    @{ Stage = 'verify-parent-id'; Transform = { param($item) $item.Parent.EntryID = 'sensitive-marker'; $item } },
    @{ Stage = 'verify-parent-store'; Transform = { param($item) $item.Parent.StoreID = 'sensitive-marker'; $item } },
    @{ Stage = 'verify-parent'; Transform = { param($item) $item.Parent = $null; $item } },
    @{ Stage = 'verify-item-id'; Transform = { param($item) $item | Add-Member -MemberType ScriptProperty -Name EntryID -Value { throw 'sensitive-marker' } -Force; $item } },
    @{ Stage = 'verify-parent'; Transform = { param($item) $item | Add-Member -MemberType ScriptProperty -Name Parent -Value { throw 'sensitive-marker' } -Force; $item } },
    @{ Stage = 'verify-parent-id'; Transform = { param($item) $item.Parent | Add-Member -MemberType ScriptProperty -Name EntryID -Value { throw 'sensitive-marker' } -Force; $item } },
    @{ Stage = 'verify-parent-store'; Transform = { param($item) $item.Parent | Add-Member -MemberType ScriptProperty -Name StoreID -Value { throw 'sensitive-marker' } -Force; $item } }
)
# PowerShell script-property getter exceptions surface as null property values;
# the following verify stages must still reject them without leaking the exception.
foreach ($case in $identityCases) {
    $mock = New-ReleaseMock
    $mock.ReopenTransform = $case.Transform
    Assert-FailureStage { & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock } } $case.Stage
    if (-not $mock.Item.Saved -or $mock.Item.EntryID -cne 'synthetic-item' -or $mock.Item.Parent.EntryID -cne 'synthetic-folder' -or $mock.Item.Parent.StoreID -cne 'synthetic-store') { throw 'Reopen failure altered the original fixture identity.' }
}
Write-Output 'PASS: independent reopened parent/store mismatches, null parent and throwing identity getters fail closed with anonymous stages.'
foreach ($comparison in @(2, 3)) {
    $expectedStage = if ($comparison -eq 2) { 'verify-item-id' } else { 'verify-parent-id' }
    $mock = New-ReleaseMock; $mock.ThrowComparison = $comparison
    Assert-FailureStage { & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock } } $expectedStage
    $mock = New-ReleaseMock; $mock.RejectComparison = $comparison
    Assert-FailureStage { & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock } } $expectedStage
}
Write-Output 'PASS: Outlook comparison false results and exceptions fail closed for both item and folder identity.'
$placementCases = @(
    @{ Stage = 'verify-saved-store'; Moves = 0; Configure = { param($session) $session.Item.SaveParent.StoreID = 'other-store' } },
    @{ Stage = 'verify-saved-store'; Moves = 0; Configure = { param($session) $session.Item.AfterSave = { param($item) $item.Parent = $null } } },
    @{ Stage = 'verify-saved-parent'; Moves = 0; Configure = { param($session) $session.ThrowComparison = 1 } },
    @{ Stage = 'verify-before-move-recipients'; Moves = 0; Configure = { param($session) $session.Item.AfterSave = { param($item) $item.Recipients.Count = 1 } } },
    @{ Stage = 'verify-before-move-sent'; Moves = 0; Configure = { param($session) $session.Item.AfterSave = { param($item) $item.Sent = $true } } },
    @{ Stage = 'verify-before-move-body'; Moves = 0; Configure = { param($session) $session.Item.AfterSave = { param($item) $item.Body = 'sensitive-marker' } } },
    @{ Stage = 'verify-before-move-subject'; Moves = 0; Configure = { param($session) $session.Item.AfterSave = { param($item) $item.Subject = 'sensitive-marker' } } },
    @{ Stage = 'move-created-draft'; Moves = 1; Configure = { param($session) $session.Item.MoveTransform = { param($item) throw 'sensitive-marker' } } },
    @{ Stage = 'verify-moved-draft'; Moves = 1; Configure = { param($session) $session.Item.MoveTransform = { param($item) $null } } },
    @{ Stage = 'verify-moved-draft'; Moves = 1; Configure = { param($session) $session.Item.MoveTransform = { param($item) $item.Recipients.Count = 1; $item } } },
    @{ Stage = 'verify-moved-draft'; Moves = 1; Configure = { param($session) $session.Item.MoveTransform = { param($item) $item.Sent = $true; $item } } },
    @{ Stage = 'verify-moved-draft'; Moves = 1; Configure = { param($session) $session.Item.MoveTransform = { param($item) $item.Body = 'sensitive-marker'; $item } } },
    @{ Stage = 'verify-moved-draft'; Moves = 1; Configure = { param($session) $session.Item.MoveTransform = { param($item) $item.Subject = 'sensitive-marker'; $item } } },
    @{ Stage = 'verify-moved-store'; Moves = 1; Configure = { param($session) $session.Item.MoveTransform = { param($item) $item.Parent = [pscustomobject]@{ EntryID = 'synthetic-folder'; StoreID = 'other-store' }; $item } } },
    @{ Stage = 'verify-moved-parent'; Moves = 1; Configure = { param($session) $session.Item.MoveTransform = { param($item) $item.Parent = [pscustomobject]@{ EntryID = 'other-folder'; StoreID = 'synthetic-store' }; $item } } },
    @{ Stage = 'read-saved-item-id'; Moves = 1; Configure = { param($session) $session.Item.MoveTransform = { param($item) $item.EntryID = ''; $item } } }
)
foreach ($case in $placementCases) {
    $mock = New-ReleaseMock
    $mock.Item.SaveParent = [pscustomobject]@{ EntryID = 'synthetic-default-drafts'; StoreID = 'synthetic-store' }
    & $case.Configure $mock
    Assert-FailureStage { & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock } } $case.Stage
    if ($mock.Item.MoveCount -ne $case.Moves -or $mock.LastReopenId) { throw 'Unsafe placement continued or retried after rejection.' }
    if ($case.Moves -eq 1 -and -not [object]::ReferenceEquals($mock.Item.MoveDestination, $mock.Accounts[0].DeliveryStore.Root.Folders.Folder)) { throw 'Move destination escaped the newly created fixture folder.' }
}
Write-Output 'PASS: 16 guarded placement failures stop before reopen; cross-store/content drift cannot move and failed moves never retry.'
$bodyDiagnosticCases = @(
    @{ Stage = 'verify-before-move-body-trailing-newline'; Trailing = $true; LineEndings = $false; Delta = 2; Transform = { param($item) $item.Body += "`r`n" } },
    @{ Stage = 'verify-before-move-body-line-endings'; Trailing = $false; LineEndings = $true; Delta = -1; Transform = { param($item) $item.Body = $item.Body.Replace("`r`n", "`n") } },
    @{ Stage = 'verify-before-move-body'; Trailing = $false; LineEndings = $false; Delta = 1; Transform = { param($item) $item.Body += 'X' } }
)
foreach ($case in $bodyDiagnosticCases) {
    $mock = New-ReleaseMock
    $mock.Item.SaveParent = [pscustomobject]@{ EntryID = 'synthetic-default-drafts'; StoreID = 'synthetic-store' }
    $mock.Item.AfterSave = $case.Transform
    $failure = $null
    try { & $scriptPath -Live -AccountDomain example.test -SessionFactory { $mock } | Out-Null } catch { $failure = $_.Exception }
    if ($null -eq $failure -or $failure.Message -notmatch ("stage '" + $case.Stage + "'")) { throw 'Expected anonymous body-comparison failure stage.' }
    if ($mock.Item.MoveCount -ne 0 -or $mock.LastReopenId) { throw 'Diagnostic normalization must never permit Move or reopen.' }
    if ($failure.Data.Count -ne 1 -or -not $failure.Data.Contains('BodyComparison')) { throw 'Expected only safe body-comparison metadata.' }
    $diagnostic = $failure.Data['BodyComparison'] | ConvertFrom-Json
    $keys = @($diagnostic.PSObject.Properties.Name | Sort-Object) -join ','
    if ($keys -cne 'actualChars,equalAfterLineEndingNormalization,exact,expectedChars,trailingNewlinesOnly') { throw 'Unexpected body diagnostic fields.' }
    if ($diagnostic.exact -ne $false -or $diagnostic.equalAfterLineEndingNormalization -ne $case.LineEndings -or $diagnostic.trailingNewlinesOnly -ne $case.Trailing -or ($diagnostic.actualChars - $diagnostic.expectedChars) -ne $case.Delta) { throw 'Body-comparison flags/counts mismatch.' }
    if ($failure.Data['BodyComparison'] -match 'sensitive-marker|synthetic|fixture|store|item|folder|Maily') { throw 'Body-comparison diagnostics disclosed content or identity.' }
}
Write-Output 'PASS: exact-rejection diagnostics distinguish terminal newlines, line endings and content changes using only flags/counts.'
if ($IsWindows) {
    $probe = @'
$ErrorActionPreference = 'Stop'
try {
    $dryRun = & './scripts/test-outlook-release-folder.ps1' -SessionFactory { throw 'FACTORY_INVOKED' }
    if ($dryRun.Mode -ne 'DryRun') { exit 5 }
    & './scripts/test-outlook-release-folder.ps1' -Live -AccountDomain example.test -SessionFactory { throw 'FACTORY_INVOKED' }
    exit 3
} catch {
    if ($_.Exception.Message -eq 'Live validation requires PowerShell 7 or later. Run this script with pwsh.') { exit 0 }
    Write-Output 'Legacy PowerShell did not reject with the expected runtime guard.'
    exit 4
}
'@
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($probe))
    Push-Location (Split-Path $PSScriptRoot -Parent)
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -OutputFormat Text -EncodedCommand $encoded
        if ($LASTEXITCODE -ne 0) { throw 'PowerShell 5 live runtime guard failed.' }
    } finally { Pop-Location }
    Write-Output 'PASS: PowerShell 5 rejects live mode before the mocked COM factory.'
}
Write-Output ('RESULT: {0} mock groups passed; no Outlook access.' -f $(if ($IsWindows) { 16 } else { 15 }))
