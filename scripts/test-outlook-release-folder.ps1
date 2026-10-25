<#
Optional, explicitly approved COM fixture check. Dry-run never instantiates Outlook.
Live mode retains one newly created folder and one synthetic, recipient-free draft.
This does not validate Office.js, generation, ribbon loading or delivery.
SessionFactory is a test seam; production callers should omit it.
#>
[CmdletBinding()]
param(
    [switch]$Live,
    [string]$AccountDomain,
    [scriptblock]$SessionFactory = { (New-Object -ComObject Outlook.Application).GetNamespace('MAPI') }
)
$ErrorActionPreference = 'Stop'
if (-not $Live) {
    return [pscustomobject]@{ Mode = 'DryRun'; Passed = 0; FolderRetained = $false }
}
if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw 'Live validation requires PowerShell 7 or later. Run this script with pwsh.'
}
if ($AccountDomain -notmatch '^[a-zA-Z0-9](?:[a-zA-Z0-9.-]*[a-zA-Z0-9])?$' -or -not $AccountDomain.Contains('.')) {
    throw 'Live validation requires an explicit account domain.'
}
$stage = 'session'
$bodyComparison = $null
try {
    $session = & $SessionFactory
    $stage = 'account-selection'
    $matches = @($session.Accounts | Where-Object {
        $address = [string]$_.SmtpAddress
        $parts = $address.Split('@')
        $parts.Count -eq 2 -and $parts[0].Length -gt 0 -and $parts[1].Equals($AccountDomain, [StringComparison]::OrdinalIgnoreCase)
    })
    if ($matches.Count -ne 1) { throw 'Account selection must resolve exactly one account.' }
    $stage = 'store-selection'
    $store = $matches[0].DeliveryStore
    $storeId = [string]$store.StoreID
    if ([string]::IsNullOrWhiteSpace($storeId)) { throw 'Missing store identity.' }
    $stage = 'create-folder'
    $folder = $store.GetRootFolder().Folders.Add('Maily-Release-Test-' + [guid]::NewGuid().ToString('N'))
    $folderId = [string]$folder.EntryID
    if ([string]::IsNullOrWhiteSpace($folderId) -or [string]$folder.StoreID -cne $storeId) { throw 'Invalid created folder identity.' }
    $stage = 'create-draft'
    $item = $folder.Items.Add('IPM.Note')
    if ($item.Recipients.Count -ne 0 -or $item.Sent) { throw 'New item is not an empty draft.' }
    # ASCII JSON escapes keep this BOM-free script parseable by PowerShell 5,
    # allowing its runtime guard to run before any COM access.
    $unicode = ConvertFrom-Json '"\u05e9\u05dc\u05d5\u05dd \u2014 caf\u00e9 \u2014 \u65e5\u672c\u8a9e \ud83d\ude00"'
    $body = "Maily synthetic release fixture. No real mailbox content.`r`n" + $unicode
    $subject = 'Maily synthetic release fixture'
    $item.BodyFormat = 1 # olFormatPlain: preserve this exact Unicode plain-text fixture.
    $item.Subject = $subject
    $item.Body = $body
    $stage = 'save'
    $item.Save()
    # Outlook can save a new item in its default folder. Only this newly created
    # synthetic, recipient-free draft may be placed into our newly created folder.
    $stage = 'read-saved-parent'
    $savedParent = $item.Parent
    $stage = 'verify-saved-store'
    if ($null -eq $savedParent -or [string]$savedParent.StoreID -cne $storeId) { throw 'Saved draft is outside the selected store.' }
    $stage = 'verify-saved-parent'
    $savedParentId = [string]$savedParent.EntryID
    if ([string]::IsNullOrWhiteSpace($savedParentId)) { throw 'Missing saved parent identity.' }
    if (-not $session.CompareEntryIDs($savedParentId, $folderId)) {
        $stage = 'verify-before-move-sent'
        if ($item.Sent) { throw 'Created draft is sent.' }
        $stage = 'verify-before-move-recipients'
        if ($item.Recipients.Count -ne 0) { throw 'Created draft has recipients.' }
        $stage = 'verify-before-move-subject'
        if ([string]$item.Subject -cne $subject) { throw 'Created draft subject changed.' }
        $stage = 'verify-before-move-body'
        $actualBody = [string]$item.Body
        if ($actualBody -cne $body) {
            # Diagnostic comparisons never authorize a move. Only fixed flags and
            # character counts survive the privacy boundary; no text is retained.
            $bodyComparison = [ordered]@{
                exact = $false
                equalAfterLineEndingNormalization = ($actualBody.Replace("`r`n", "`n").Replace("`r", "`n") -ceq $body.Replace("`r`n", "`n").Replace("`r", "`n"))
                trailingNewlinesOnly = ($actualBody.TrimEnd([char[]]"`r`n") -ceq $body.TrimEnd([char[]]"`r`n"))
                expectedChars = $body.Length
                actualChars = $actualBody.Length
            }
            if ($bodyComparison.trailingNewlinesOnly) { $stage = 'verify-before-move-body-trailing-newline' }
            elseif ($bodyComparison.equalAfterLineEndingNormalization) { $stage = 'verify-before-move-body-line-endings' }
            throw 'Created draft body changed.'
        }
        $stage = 'move-created-draft'
        $movedItem = $item.Move($folder)
        $stage = 'verify-moved-draft'
        if ($null -eq $movedItem -or $movedItem.Sent -or $movedItem.Recipients.Count -ne 0 -or [string]$movedItem.Subject -cne $subject -or [string]$movedItem.Body -cne $body) { throw 'Moved draft does not match the synthetic fixture.' }
        $stage = 'read-moved-parent'
        $movedParent = $movedItem.Parent
        $stage = 'verify-moved-store'
        if ($null -eq $movedParent -or [string]$movedParent.StoreID -cne $storeId) { throw 'Moved draft is outside the selected store.' }
        $stage = 'verify-moved-parent'
        $movedParentId = [string]$movedParent.EntryID
        if ([string]::IsNullOrWhiteSpace($movedParentId) -or -not $session.CompareEntryIDs($movedParentId, $folderId)) { throw 'Moved draft is outside the created folder.' }
        $item = $movedItem
    }
    $stage = 'read-saved-item-id'
    $itemId = [string]$item.EntryID
    if ([string]::IsNullOrWhiteSpace($itemId)) { throw 'Missing saved item identity.' }
    $stage = 'reopen'
    $reopened = $session.GetItemFromID($itemId, $storeId)
    $stage = 'read-item-id'
    $reopenedId = [string]$reopened.EntryID
    $stage = 'verify-item-id'
    if ([string]::IsNullOrWhiteSpace($reopenedId) -or -not $session.CompareEntryIDs($reopenedId, $itemId)) { throw 'Saved item identity mismatch.' }
    $stage = 'read-parent'
    $reopenedParent = $reopened.Parent
    $stage = 'verify-parent'
    if ($null -eq $reopenedParent) { throw 'Saved item parent is missing.' }
    $stage = 'read-parent-id'
    $reopenedParentId = [string]$reopenedParent.EntryID
    $stage = 'verify-parent-id'
    if ([string]::IsNullOrWhiteSpace($reopenedParentId) -or -not $session.CompareEntryIDs($reopenedParentId, $folderId)) { throw 'Saved folder identity mismatch.' }
    $stage = 'read-parent-store'
    $reopenedStoreId = [string]$reopenedParent.StoreID
    $stage = 'verify-parent-store'
    if ($reopenedStoreId -cne $storeId) { throw 'Saved folder store mismatch.' }
    $stage = 'verify-body'
    if ([string]$reopened.Subject -cne $subject -or [string]$reopened.Body -cne $body) { throw 'Synthetic content roundtrip mismatch.' }
    $stage = 'verify-unsent'
    if ($reopened.Sent -or $reopened.Recipients.Count -ne 0) { throw 'Saved item is not a recipient-free draft.' }
    [pscustomobject]@{ Mode = 'Live'; Passed = 4; FolderRetained = $true }
} catch {
    # COM errors can contain account names, paths or message content. Keep output anonymous.
    $failure = [InvalidOperationException]::new("Outlook release fixture check failed at stage '$stage'. Any created synthetic folder and item were retained. No existing mail was inspected.")
    if ($null -ne $bodyComparison) { $failure.Data['BodyComparison'] = $bodyComparison | ConvertTo-Json -Compress }
    throw $failure
}
