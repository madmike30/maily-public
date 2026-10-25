param(
  [Parameter(Mandatory = $true)]
  [string]$Action,
  [string]$MailId = "",
  [string]$FolderName = "",
  [string]$HtmlPath = "",
  [string]$DraftText = "",
  [string]$Subject = "",
  [string]$SeedText = "",
  [string]$UserName = "",
  [string]$MailboxAddress = "",
  [string]$ComposeType = "",
  [string]$DestPath = "",
  [string]$DateProperty = "",
  [string]$OutputPath = "",
  [string]$StartDate = "",
  [string]$EndDate = "",
  [string]$BodyMode = "",
  [int]$AttachmentIndex = 0,
  [int]$Limit = 50
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$logDir = Join-Path $repoRoot "output\logs"
$bridgeLogPath = Join-Path $logDir ("outlook-bridge-{0}.jsonl" -f (Get-Date -Format "yyyyMMdd"))
$bridgeCorrelationId = [guid]::NewGuid().ToString('N')

New-Item -ItemType Directory -Force -Path $logDir | Out-Null

function Remove-ExpiredLogFiles {
  param(
    [Parameter(Mandatory = $true)]
    [string]$DirectoryPath,
    [int]$RetentionDays = 31
  )

  if (-not (Test-Path $DirectoryPath)) {
    return
  }

  $cutoff = [DateTime]::UtcNow.AddDays(-1 * [Math]::Abs($RetentionDays))
  Get-ChildItem -Path $DirectoryPath -File -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTimeUtc -lt $cutoff } |
    ForEach-Object {
      try {
        Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
      } catch {
      }
    }
}

function Write-BridgeLog {
  param(
    [Parameter(Mandatory = $true)]
    [string]$EventName,
    [string]$Status = "info",
    [string]$Detail = "",
    [hashtable]$Properties = @{}
  )

  try {
    # Only fixed categories and numeric counters belong in persisted diagnostics.
    # Detail and arbitrary properties may contain mailbox content or exception text.
    $safeEvents = @('bridge_item_skipped', 'bridge_item_failed', 'bridge_calendar_collect_completed', 'bridge_action_start', 'bridge_action_complete', 'bridge_action_failed')
    $safeActions = @('test', 'list-emails', 'list-emails-preview', 'list-search-emails', 'get-email', 'list-attachments', 'download-attachment', 'create-reply', 'create-reply-all', 'rewrite-active-compose', 'list-sent-emails', 'list-calendar-events', 'set-read', 'move-mail')
    $safeProperties = [ordered]@{}
    foreach ($key in @('limit', 'collected', 'itemClass')) {
      $value = $Properties[$key]
      if (($value -is [int] -or $value -is [long]) -and $value -ge 0) {
        $safeProperties[$key] = $value
      }
    }
    $payload = [ordered]@{
      timestamp = [DateTimeOffset]::UtcNow.ToString("o")
      correlationId = $bridgeCorrelationId
      event = $(if ($EventName -cin $safeEvents) { $EventName } else { 'unknown' })
      status = $(if ($Status -cin @('info', 'warning', 'error', 'ok')) { $Status } else { 'info' })
      action = $(if ($Action -cin $safeActions) { $Action } else { 'unknown' })
      properties = $safeProperties
    }

    [IO.File]::AppendAllText($bridgeLogPath, (($payload | ConvertTo-Json -Compress -Depth 4) + [Environment]::NewLine), (New-Object Text.UTF8Encoding $false))
  } catch {
  }
}

Remove-ExpiredLogFiles -DirectoryPath $logDir -RetentionDays 31

function Invoke-WithRetry {
  param(
    [Parameter(Mandatory = $true)]
    [scriptblock]$Block,
    [int]$Retries = 12,
    [int]$DelayMs = 500
  )

  for ($i = 0; $i -lt $Retries; $i++) {
    try {
      return & $Block
    } catch {
      if ($_.Exception.Message -like "*rejected by callee*") {
        Start-Sleep -Milliseconds $DelayMs
      } else {
        throw
      }
    }
  }

  throw "Outlook COM call failed after retries."
}

function Get-OutlookSession {
  $outlook = Invoke-WithRetry { New-Object -ComObject Outlook.Application }
  Invoke-WithRetry { $outlook.GetNamespace("MAPI") }
}

function Emit-Json {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Json
  )

  if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    Write-Output $Json
  } else {
    Set-Content -Path $OutputPath -Value $Json -Encoding utf8
  }
}

function Get-NonNullString {
  param([AllowNull()]$Value)

  if ($null -eq $Value) {
    return ""
  }

  return [string]$Value
}

function Format-JsonDateTime {
  param([AllowNull()]$Value)

  if ($null -eq $Value) {
    return ""
  }

  try {
    return ([datetimeoffset][datetime]$Value).ToString("o")
  } catch {
    return ""
  }
}

function Normalize-ComparableText {
  param([string]$Value)

  $normalized = (Get-NonNullString $Value)
  if ([string]::IsNullOrWhiteSpace($normalized)) {
    return ""
  }

  $normalized = [regex]::Replace($normalized, "\s+", " ")
  return $normalized.Trim().ToLowerInvariant()
}

function Convert-Importance {
  param([int]$Importance)

  switch ($Importance) {
    0 { "low" }
    2 { "high" }
    default { "normal" }
  }
}

function Get-RecipientList {
  param([string]$Value)

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return @()
  }

  $parts = $Value -split ";"
  $result = @()

  foreach ($part in $parts) {
    $trimmed = $part.Trim()
    if ($trimmed) {
      $result += @{
        emailAddress = @{
          name = $trimmed
          address = $trimmed
        }
      }
    }
  }

  return $result
}

function Get-MailPayload {
  param(
    $Mail,
    [bool]$IncludeBody = $true
  )

  $body = ""
  $htmlBody = ""
  $sentDateTime = $null
  $receivedDateTime = $null
  $fromName = ""
  $fromAddress = ""
  $subject = ""
  $conversationId = ""
  $categories = @()
  $isRead = $true

  # Search and preview flows still need plain-text evidence. Only the HTML body stays behind the full-body gate.
  try { $body = [string]$Mail.Body } catch {}
  if ($IncludeBody) {
    try { $htmlBody = [string]$Mail.HTMLBody } catch {}
  }
  try { $sentDateTime = Format-JsonDateTime -Value $Mail.SentOn } catch {}
  try { $receivedDateTime = Format-JsonDateTime -Value $Mail.ReceivedTime } catch {}
  try { $fromName = [string]$Mail.SenderName } catch {}
  try { $fromAddress = [string]$Mail.SenderEmailAddress } catch {}
  try { $subject = [string]$Mail.Subject } catch {}
  try { $conversationId = [string]$Mail.ConversationID } catch {}
  try { $isRead = -not [bool]$Mail.UnRead } catch {}
  $folderName = ""
  try { $folderName = [string]$Mail.Parent.Name } catch {}
  try {
    if (-not [string]::IsNullOrWhiteSpace([string]$Mail.Categories)) {
      $categories = ([string]$Mail.Categories -split ",") | ForEach-Object { $_.Trim() }
    }
  } catch {}

  $attachmentCount = 0

  try {
    $attachmentCount = [int]$Mail.Attachments.Count
  } catch {}

  @{
    "id" = [string]$Mail.EntryID
    "internetMessageId" = $null
    "conversationId" = $conversationId
    "folderName" = $folderName
    "sentDateTime" = $sentDateTime
    "receivedDateTime" = $receivedDateTime
    "subject" = $subject
    "from" = @{
        emailAddress = @{
          name = $fromName
          address = $fromAddress
        }
      }
    "toRecipients" = Get-RecipientList -Value ([string]$Mail.To)
    "ccRecipients" = Get-RecipientList -Value ([string]$Mail.CC)
    "bccRecipients" = Get-RecipientList -Value ([string]$Mail.BCC)
    "replyTo" = @()
    "importance" = Convert-Importance -Importance ([int]$Mail.Importance)
    "categories" = $categories
    "isRead" = $isRead
    "hasAttachments" = ($attachmentCount -gt 0)
    "attachmentNames" = @()
    "bodyPreview" = if ($body.Length -gt 255) { $body.Substring(0, 255) } else { $body }
    "body" = @{
      "contentType" = if (-not $IncludeBody -or [string]::IsNullOrWhiteSpace($htmlBody)) { "text" } else { "html" }
      "content" = if ($IncludeBody) { if ([string]::IsNullOrWhiteSpace($htmlBody)) { $body } else { $htmlBody } } else { "" }
      "text" = if ($IncludeBody) { $body } else { "" }
    }
  }
}

function Get-AttachmentsPayload {
  param($Mail)

  $payload = @()
  $count = 0

  try { $count = [int]$Mail.Attachments.Count } catch {}

  for ($i = 1; $i -le $count; $i++) {
    try {
      $att = Invoke-WithRetry { $Mail.Attachments.Item($i) }
      if ($null -eq $att) {
        continue
      }

      $payload += @{
        "id" = "{0}__{1}" -f [string]$Mail.EntryID, $i
        "index" = $i
        "name" = [string]$att.FileName
        "size" = [int64]$att.Size
        "isInline" = $false
        "contentType" = ""
      }
    } catch {
      continue
    }
  }

  $payload
}

function Find-FolderByName {
  param(
    $ParentFolder,
    [string]$TargetName
  )

  if ([string]::Equals([string]$ParentFolder.Name, $TargetName, [System.StringComparison]::OrdinalIgnoreCase)) {
    return $ParentFolder
  }

  foreach ($child in (Get-OutlookChildFolders -ParentFolder $ParentFolder)) {
    $found = Find-FolderByName -ParentFolder $child -TargetName $TargetName
    if ($null -ne $found) {
      return $found
    }
  }

  return $null
}

function Get-OutlookChildFolders {
  param($ParentFolder)

  $result = @()
  if ($null -eq $ParentFolder) {
    return $result
  }

  $folders = Invoke-WithRetry { $ParentFolder.Folders }
  if ($null -eq $folders) {
    return $result
  }

  $count = 0
  try { $count = [int]$folders.Count } catch { $count = 0 }
  for ($index = 1; $index -le $count; $index++) {
    try {
      $child = Invoke-WithRetry { $folders.Item($index) }
      if ($null -ne $child) {
        $result += $child
      }
    } catch {
      continue
    }
  }

  return @($result)
}

function Get-MailFoldersRecursive {
  param(
    $ParentFolder,
    [ref]$Result
  )

  foreach ($child in (Get-OutlookChildFolders -ParentFolder $ParentFolder)) {
    if ($null -eq $child) {
      continue
    }

    $defaultItemType = 0
    try { $defaultItemType = [int]$child.DefaultItemType } catch {}
    if ($defaultItemType -eq 0) {
      $Result.Value += $child
    }

    Get-MailFoldersRecursive -ParentFolder $child -Result $Result
  }
}

function Get-MailboxItem {
  param($Session, $Store, [string]$Id)
  if ($null -eq $Store -or [string]::IsNullOrWhiteSpace([string]$Store.StoreID) -or [string]::IsNullOrWhiteSpace($Id)) {
    throw 'Mailbox store and item identity are required.'
  }
  $item = Invoke-WithRetry { $Session.GetItemFromID($Id, [string]$Store.StoreID) }
  if ($null -eq $item -or -not [string]::Equals([string]$item.Parent.StoreID, [string]$Store.StoreID, [System.StringComparison]::Ordinal)) {
    throw 'Outlook item does not belong to the requested mailbox.'
  }
  return $item
}

function Get-SearchFolders {
  param(
    $Store,
    [string]$Scope
  )

  $normalizedScope = (Get-NonNullString $Scope).Trim().ToLowerInvariant()
  switch ($normalizedScope) {
    "inbox" {
      return @(Invoke-WithRetry { $Store.GetDefaultFolder(6) })
    }
    "sent-items" {
      return @(Invoke-WithRetry { $Store.GetDefaultFolder(5) })
    }
    default {
      $store = Invoke-WithRetry { $Store }
      $root = Invoke-WithRetry { $store.GetRootFolder() }
      $folders = @()
      Get-MailFoldersRecursive -ParentFolder $root -Result ([ref]$folders)
      $excluded = @("Deleted Items", "Junk Email", "Drafts", "Outbox", "Conversation History", "Sync Issues")
      return @($folders | Where-Object {
        $name = Get-NonNullString $_.Name
        -not ($excluded | Where-Object { [string]::Equals($_, $name, [System.StringComparison]::OrdinalIgnoreCase) })
      })
    }
  }
}

function Find-FirstRegexIndex {
  param(
    [string]$Value,
    [string[]]$Patterns
  )

  $cutIndex = if ($null -eq $Value) { 0 } else { $Value.Length }
  foreach ($pattern in ($Patterns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
    $match = [regex]::Match($Value, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline)
    if ($match.Success) {
      $cutIndex = [Math]::Min($cutIndex, $match.Index)
    }
  }

  if ($cutIndex -ge $Value.Length) {
    return -1
  }

  return $cutIndex
}

function Get-TextLinesWithOffsets {
  param([string]$Text)

  $result = @()
  if ($null -eq $Text) {
    return $result
  }

  $index = 0
  while ($index -lt $Text.Length) {
    $lineStart = $index
    while ($index -lt $Text.Length -and $Text[$index] -ne "`r" -and $Text[$index] -ne "`n") {
      $index += 1
    }

    $lineText = $Text.Substring($lineStart, $index - $lineStart)

    if ($index -lt $Text.Length -and $Text[$index] -eq "`r") {
      $index += 1
      if ($index -lt $Text.Length -and $Text[$index] -eq "`n") {
        $index += 1
      }
    } elseif ($index -lt $Text.Length -and $Text[$index] -eq "`n") {
      $index += 1
    }

    $result += [pscustomobject]@{
      Text = $lineText
      Start = $lineStart
      End = $index
    }
  }

  if ($Text.Length -eq 0) {
    $result += [pscustomobject]@{
      Text = ""
      Start = 0
      End = 0
    }
  }

  return $result
}

function Test-IsClosingLine {
  param([string]$Line)

  $normalized = ((Get-NonNullString $Line).Trim().ToLowerInvariant() -replace "[,:;.!-]+$", "").Trim()
  return @(
    "regards",
    "best regards",
    "kind regards",
    "thanks",
    "thank you",
    "cheers",
    "sincerely",
    "warm regards"
  ) -contains $normalized
}

function Test-IsSignatureLine {
  param(
    [string]$Line,
    [string]$NormalizedUserName,
    [string]$NormalizedMailboxAddress
  )

  $normalized = (Get-NonNullString $Line).Trim().ToLowerInvariant()
  if ([string]::IsNullOrWhiteSpace($normalized)) {
    return $false
  }

  if ($normalized -eq "--" -or (Test-IsClosingLine -Line $normalized)) {
    return $true
  }

  if (-not [string]::IsNullOrWhiteSpace($NormalizedMailboxAddress) -and $normalized.Contains($NormalizedMailboxAddress)) {
    return $true
  }

  if (-not [string]::IsNullOrWhiteSpace($NormalizedUserName) -and $normalized.Contains($NormalizedUserName)) {
    return $true
  }

  if ($normalized -match "@" -or $normalized -match "\+?\d[\d\s\-()]{6,}") {
    return $true
  }

  if ($normalized.Contains("|")) {
    return $true
  }

  return $false
}

function Find-QuotedThreadStartIndex {
  param([string]$Text)

  return Find-FirstRegexIndex -Value (Get-NonNullString $Text) -Patterns @(
    "^\s*From:\s",
    "^\s*Sent:\s",
    "^\s*Date:\s",
    "^\s*Subject:\s",
    "^\s*To:\s",
    "^\s*Cc:\s",
    "^\s*On .+wrote:\s*$",
    "^\s*-{2,}\s*Original Message\s*-{2,}\s*$"
  )
}

function Find-SignatureStartIndex {
  param(
    [string]$Text,
    [string]$UserName,
    [string]$MailboxAddress
  )

  $normalizedUserName = (Get-NonNullString $UserName).Trim().ToLowerInvariant()
  $normalizedMailboxAddress = (Get-NonNullString $MailboxAddress).Trim().ToLowerInvariant()
  $lines = Get-TextLinesWithOffsets -Text (Get-NonNullString $Text)

  if ($lines.Count -eq 0) {
    return -1
  }

  $lastMeaningfulIndex = -1
  for ($i = $lines.Count - 1; $i -ge 0; $i--) {
    if (-not [string]::IsNullOrWhiteSpace($lines[$i].Text)) {
      $lastMeaningfulIndex = $i
      break
    }
  }

  if ($lastMeaningfulIndex -lt 0) {
    return -1
  }

  for ($i = $lastMeaningfulIndex; $i -ge 0; $i--) {
    $lineText = $lines[$i].Text
    if ([string]::IsNullOrWhiteSpace($lineText)) {
      continue
    }

    if (Test-IsSignatureLine -Line $lineText -NormalizedUserName $normalizedUserName -NormalizedMailboxAddress $normalizedMailboxAddress) {
      $start = $i
      while ($start -gt 0) {
        $prior = $lines[$start - 1].Text
        if ([string]::IsNullOrWhiteSpace($prior)) {
          $start -= 1
          continue
        }

        if ((Test-IsClosingLine -Line $prior) -or (Test-IsSignatureLine -Line $prior -NormalizedUserName $normalizedUserName -NormalizedMailboxAddress $normalizedMailboxAddress)) {
          $start -= 1
          continue
        }

        break
      }

      return $lines[$start].Start
    }
  }

  return -1
}

function Get-ComposeReplacementPlan {
  param(
    [string]$Text,
    [string]$UserName,
    [string]$MailboxAddress
  )

  $value = Get-NonNullString $Text
  $quotedThreadStart = Find-QuotedThreadStartIndex -Text $value
  $leading = if ($quotedThreadStart -ge 0) { $value.Substring(0, $quotedThreadStart) } else { $value }
  $signatureStart = Find-SignatureStartIndex -Text $leading -UserName $UserName -MailboxAddress $MailboxAddress

  $replaceEnd = if ($signatureStart -ge 0) { $signatureStart } else { $leading.Length }
  while ($replaceEnd -gt 0 -and [char]::IsWhiteSpace($leading[$replaceEnd - 1])) {
    $replaceEnd -= 1
  }

  $detectedDraftText = if ($replaceEnd -gt 0) { $leading.Substring(0, $replaceEnd).Trim() } else { "" }

  return [pscustomobject]@{
    ReplaceStart = 0
    ReplaceEnd = [Math]::Max(0, $replaceEnd)
    SignatureStart = $signatureStart
    QuotedThreadStart = $quotedThreadStart
    DetectedDraftText = $detectedDraftText
    FullTextLength = $value.Length
    Strategy = "heuristic-boundary"
    SeedMatched = $false
  }
}

function Get-EditorPreviewText {
  param(
    $Editor,
    [int]$MaxChars = 1024
  )

  try {
    $content = Invoke-WithRetry { $Editor.Content }
    $text = [string](Invoke-WithRetry { $content.Text })
    if ($text.Length -gt $MaxChars) {
      return $text.Substring(0, $MaxChars)
    }

    return $text
  } catch {
    return ""
  }
}

function Find-SeedReplacementPlan {
  param(
    $Editor,
    $Content,
    [string]$SeedText,
    [string]$UserName,
    [string]$MailboxAddress
  )

  $normalizedSeed = Normalize-ComparableText -Value $SeedText
  if ([string]::IsNullOrWhiteSpace($normalizedSeed)) {
    return $null
  }

  $paragraphCount = 0
  try {
    $paragraphCount = [int](Invoke-WithRetry { $Editor.Paragraphs.Count })
  } catch {
    return $null
  }

  if ($paragraphCount -le 0) {
    return $null
  }

  $accumulated = ""
  $lastMatchedRange = $null
  $normalizedUserName = (Get-NonNullString $UserName).Trim().ToLowerInvariant()
  $normalizedMailboxAddress = (Get-NonNullString $MailboxAddress).Trim().ToLowerInvariant()

  for ($index = 1; $index -le $paragraphCount; $index++) {
    $paragraph = Invoke-WithRetry { $Editor.Paragraphs.Item($index) }
    if ($null -eq $paragraph) {
      continue
    }

    $range = Invoke-WithRetry { $paragraph.Range }
    $paragraphText = Get-NonNullString ([string](Invoke-WithRetry { $range.Text }))
    $paragraphText = $paragraphText -replace "[`r`n]+$", ""
    $normalizedParagraph = Normalize-ComparableText -Value $paragraphText

    if ([string]::IsNullOrWhiteSpace($normalizedParagraph)) {
      continue
    }

    if ((Test-IsClosingLine -Line $paragraphText) -or
      (Test-IsSignatureLine -Line $paragraphText -NormalizedUserName $normalizedUserName -NormalizedMailboxAddress $normalizedMailboxAddress) -or
      ((Find-QuotedThreadStartIndex -Text $paragraphText) -eq 0)) {
      break
    }

    if ($accumulated) {
      $accumulated = "$accumulated $normalizedParagraph"
    } else {
      $accumulated = $normalizedParagraph
    }

    if (-not $normalizedSeed.StartsWith($accumulated)) {
      break
    }

    $lastMatchedRange = $range
    if ($accumulated -eq $normalizedSeed) {
      $replaceEnd = ($lastMatchedRange.Start - $Content.Start) + $paragraphText.Length
      return [pscustomobject]@{
        ReplaceStart = 0
        ReplaceEnd = [Math]::Max(0, $replaceEnd)
        SignatureStart = -1
        QuotedThreadStart = -1
        DetectedDraftText = $SeedText
        FullTextLength = [int](Get-NonNullString ([string](Invoke-WithRetry { $Content.Text }))).Length
        Strategy = "seed-paragraph-match"
        SeedMatched = $true
      }
    }
  }

  return $null
}

function Resolve-MailboxStore {
  param($Session, [string]$Address)
  if ([string]::IsNullOrWhiteSpace($Address)) { throw 'An explicit mailbox address is required.' }
  $matches = @()
  foreach ($account in $Session.Accounts) {
    if ([string]::Equals(([string]$account.SmtpAddress).Trim(), $Address.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
      $store = $account.DeliveryStore
      if ($null -ne $store -and -not [string]::IsNullOrWhiteSpace([string]$store.StoreID)) { $matches += $store }
    }
  }
  $ids = @($matches | ForEach-Object { [string]$_.StoreID } | Select-Object -Unique)
  if ($ids.Count -ne 1) { throw 'Requested mailbox is unavailable or ambiguous in this Outlook profile.' }
  return $matches[0]
}

function Get-ComposeTargetScore {
  param(
    $Item,
    $Editor,
    [string]$ExpectedSubject,
    [string]$ExpectedSeedText,
    [string]$ExpectedStoreId = ''
  )

  if ($null -eq $Item -or $null -eq $Editor) {
    return -1
  }

  try {
    if ([int]$Item.Class -ne 43 -or $Item.Sent -eq $true) {
      return -1
    }

    if (-not [string]::IsNullOrWhiteSpace($ExpectedStoreId)) {
      $itemStoreId = [string]$Item.Parent.StoreID
      if (-not [string]::Equals($itemStoreId, $ExpectedStoreId, [System.StringComparison]::Ordinal)) { return -1 }
    }

    if (-not [string]::IsNullOrWhiteSpace($ExpectedSubject) -and
      -not [string]::Equals(([string]$Item.Subject).Trim(), $ExpectedSubject.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
      return -1
    }

    $score = 100
    if ([string]::IsNullOrWhiteSpace($ExpectedSubject)) {
      $score = 50
    }

    $normalizedSeed = Normalize-ComparableText -Value $ExpectedSeedText
    if ([string]::IsNullOrWhiteSpace($normalizedSeed)) {
      return $score + 10
    }

    $preview = Normalize-ComparableText -Value ([string](Invoke-WithRetry { $Editor.Content.Text }))
    if ([string]::IsNullOrWhiteSpace($preview)) {
      return -1
    }

    if ($preview.StartsWith($normalizedSeed)) {
      return $score + 1000
    }

    # A shared prefix does not prove this is the original draft. Require the
    # entire normalized seed at the start of the full current editor text.
    return -1
  } catch {
    return -1
  }
}

function New-ComposeTarget {
  param(
    $Item,
    $Editor,
    [string]$Source
  )

  try {
    return [pscustomobject]@{
      Item = $Item
      Editor = $Editor
      Source = $Source
      Subject = [string]$Item.Subject
      Preview = Get-EditorPreviewText -Editor $Editor
    }
  } catch {
    return $null
  }
}

function Find-ComposeTarget {
  param(
    $Outlook,
    [string]$ExpectedSubject,
    [string]$ExpectedSeedText,
    [string]$ExpectedStoreId = ''
  )

  $candidates = @()

  try {
    $activeInspector = Invoke-WithRetry { $Outlook.ActiveInspector() }
    if ($null -ne $activeInspector) {
      $item = Invoke-WithRetry { $activeInspector.CurrentItem }
      $editor = Invoke-WithRetry { $activeInspector.WordEditor }
      $target = New-ComposeTarget -Item $item -Editor $editor -Source "active-inspector"
      if ($null -ne $target) {
        $candidates += $target
      }
    }
  } catch {
  }

  try {
    $inspectors = Invoke-WithRetry { $Outlook.Inspectors }
    for ($index = 1; $index -le $inspectors.Count; $index++) {
      $candidate = Invoke-WithRetry { $inspectors.Item($index) }
      if ($null -eq $candidate) {
        continue
      }

      $item = Invoke-WithRetry { $candidate.CurrentItem }
      $editor = Invoke-WithRetry { $candidate.WordEditor }
      $target = New-ComposeTarget -Item $item -Editor $editor -Source ("inspector-{0}" -f $index)
      if ($null -ne $target) {
        $candidates += $target
      }
    }
  } catch {
  }

  try {
    $explorer = Invoke-WithRetry { $Outlook.ActiveExplorer() }
    if ($null -ne $explorer) {
      $inlineItem = Invoke-WithRetry { $explorer.ActiveInlineResponse }
      $inlineEditor = Invoke-WithRetry { $explorer.ActiveInlineResponseWordEditor }
      $target = New-ComposeTarget -Item $inlineItem -Editor $inlineEditor -Source "active-inline-response"
      if ($null -ne $target) {
        $candidates += $target
      }
    }
  } catch {
  }

  $eligible = @()
  foreach ($candidate in $candidates) {
    $score = Get-ComposeTargetScore -Item $candidate.Item -Editor $candidate.Editor -ExpectedSubject $ExpectedSubject -ExpectedSeedText $ExpectedSeedText -ExpectedStoreId $ExpectedStoreId
    if ($score -lt 0) { continue }
    # ActiveInspector also appears in Inspectors. Never choose between distinct
    # matching drafts using focus or a weak score bonus.
    $duplicate = $false
    foreach ($existing in $eligible) {
      if ([object]::ReferenceEquals($existing.Item, $candidate.Item)) { $duplicate = $true; break }
    }
    if (-not $duplicate) { $eligible += $candidate }
  }
  if ($eligible.Count -gt 1) { throw 'Multiple matching Outlook drafts are open. Close the other matching drafts and retry.' }
  if ($eligible.Count -eq 1) { return $eligible[0] }

  throw "Could not find the active Outlook compose target for subject '$ExpectedSubject'."
}

function Rewrite-ActiveCompose {
  param(
    [string]$ExpectedSubject,
    [string]$HtmlSourcePath,
    [string]$ReplacementText,
    [string]$SeedPreview,
    [string]$DisplayName,
    [string]$EmailAddress,
    [string]$Mode
  )

  $outlook = Invoke-WithRetry { New-Object -ComObject Outlook.Application }
  $session = Invoke-WithRetry { $outlook.GetNamespace('MAPI') }
  $store = Resolve-MailboxStore -Session $session -Address $EmailAddress
  $target = Find-ComposeTarget -Outlook $outlook -ExpectedSubject $ExpectedSubject -ExpectedSeedText $SeedPreview -ExpectedStoreId ([string]$store.StoreID)
  $mail = $target.Item
  $editor = $target.Editor
  $content = Invoke-WithRetry { $editor.Content }
  $fullText = [string](Invoke-WithRetry { $content.Text })
  $seedPlan = Find-SeedReplacementPlan -Editor $editor -Content $content -SeedText $SeedPreview -UserName $DisplayName -MailboxAddress $EmailAddress
  $heuristicPlan = Get-ComposeReplacementPlan -Text $fullText -UserName $DisplayName -MailboxAddress $EmailAddress
  $normalizedSeed = Normalize-ComparableText -Value $SeedPreview
  $normalizedCurrentDraft = Normalize-ComparableText -Value $heuristicPlan.DetectedDraftText
  $plans = @()
  if (-not [string]::IsNullOrWhiteSpace($normalizedSeed)) {
    if ($null -eq $seedPlan -or $normalizedCurrentDraft -ne $normalizedSeed) {
      throw 'The Outlook draft changed or could not be identified exactly. Generate a new reply from the current draft.'
    }
    $plans += $seedPlan
  } else {
    if (-not [string]::IsNullOrWhiteSpace($normalizedCurrentDraft)) {
      throw 'The Outlook draft is no longer empty. Generate a new reply from the current draft.'
    }
    $plans += $heuristicPlan
  }
  $wordReplacementText = ((Get-NonNullString $ReplacementText) -replace "`r`n", "`n" -replace "`r", "`n").Trim()
  $wordReplacementText = $wordReplacementText -replace "`n", "`r"
  $usedPlan = $null
  $attempts = @()

  $undoRecord = $null
  try {
    $undoRecord = $editor.Application.UndoRecord
    if ($null -ne $undoRecord) {
      $undoRecord.StartCustomRecord("Maily rewrite")
    }
  } catch {
    $undoRecord = $null
  }

  try {
    foreach ($plan in $plans) {
      if ($null -eq $plan) {
        continue
      }

      try {
        if (-not [string]::Equals([string](Invoke-WithRetry { $editor.Content.Text }), $fullText, [System.StringComparison]::Ordinal)) {
          throw 'The Outlook draft changed before replacement. Generate a new reply from the current draft.'
        }
        $targetRange = Invoke-WithRetry { $editor.Range($content.Start + $plan.ReplaceStart, $content.Start + $plan.ReplaceEnd) }
        $targetRange.Text = $wordReplacementText
        $usedPlan = $plan
        $attempts += @("{0}:{1}" -f (Get-NonNullString $plan.Strategy), $plan.ReplaceEnd)
        break
      } catch {
        $attempts += @("{0}:{1}:failed" -f (Get-NonNullString $plan.Strategy), $plan.ReplaceEnd)
        if ($plan -eq $plans[$plans.Count - 1]) {
          throw
        }
      }
    }
  }
  finally {
    if ($null -ne $undoRecord) {
      try {
        $undoRecord.EndCustomRecord()
      } catch {
      }
    }
  }

  $postText = [string](Invoke-WithRetry { $editor.Content.Text })
  return @{
    ok = $true
    subject = [string]$mail.Subject
    targetSource = Get-NonNullString $target.Source
    composeType = $Mode
    replaceStart = $usedPlan.ReplaceStart
    replaceEnd = $usedPlan.ReplaceEnd
    signatureStart = $usedPlan.SignatureStart
    quotedThreadStart = $usedPlan.QuotedThreadStart
    fullTextLength = $usedPlan.FullTextLength
    detectedDraftText = $usedPlan.DetectedDraftText
    strategy = Get-NonNullString $usedPlan.Strategy
    seedMatched = [bool]$usedPlan.SeedMatched
    rewriteAttempts = $attempts
    seedPreview = $SeedPreview
    replacementChars = $wordReplacementText.Length
    postTextPreview = if ($postText.Length -gt 200) { $postText.Substring(0, 200) } else { $postText }
  }
}

function Format-OutlookRestrictionDate {
  param([datetime]$Value)

  $Value.ToString("MM/dd/yyyy hh:mm tt", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-FilteredOutlookItems {
  param(
    $Folder,
    [string]$DateProperty,
    [Nullable[datetime]]$StartDateBound = $null,
    [Nullable[datetime]]$EndDateBound = $null,
    [bool]$Descending = $true
  )

  $items = Invoke-WithRetry { ,$Folder.Items }
  $filters = @()

  if ($null -ne $StartDateBound) {
    $filters += "[$DateProperty] >= '" + (Format-OutlookRestrictionDate -Value $StartDateBound) + "'"
  }

  if ($null -ne $EndDateBound) {
    $filters += "[$DateProperty] <= '" + (Format-OutlookRestrictionDate -Value $EndDateBound) + "'"
  }

  if ($filters.Count -gt 0) {
    $restriction = $filters -join " AND "
    $items = Invoke-WithRetry { ,$items.Restrict($restriction) }
  }

  try {
    Invoke-WithRetry { $items.Sort("[$DateProperty]", $Descending) } | Out-Null
  } catch {}

  if ([string]::Equals($DateProperty, "Start", [System.StringComparison]::OrdinalIgnoreCase)) {
    try {
      $items.IncludeRecurrences = $true
    } catch {}
  }

  return $items
}

function Get-ItemDateValue {
  param(
    $Mail,
    [string]$DateProperty
  )

  try {
    switch ($DateProperty) {
      "SentOn" { return [datetime]$Mail.SentOn }
      default { return [datetime]$Mail.ReceivedTime }
    }
  } catch {
    return $null
  }
}

function Collect-MailPayloads {
  param(
    $Items,
    [int]$Limit,
    [bool]$IncludeBody = $false,
    [string]$DateProperty = "ReceivedTime",
    [Nullable[datetime]]$StartDateBound = $null,
    [Nullable[datetime]]$EndDateBound = $null
  )

  $result = @()
  $seen = 0
  $index = 1

  while ($seen -lt $Limit -and $index -le $Items.Count) {
    $item = Invoke-WithRetry { $Items.Item($index) }
    $index += 1

    if ($null -eq $item) {
      continue
    }

    if ([int]$item.Class -ne 43) {
      continue
    }

    $itemDate = Get-ItemDateValue -Mail $item -DateProperty $DateProperty

    if ($null -ne $EndDateBound -and $null -ne $itemDate -and $itemDate -gt $EndDateBound) {
      continue
    }

    if ($null -ne $StartDateBound -and $null -ne $itemDate -and $itemDate -lt $StartDateBound) {
      break
    }

    $result += Get-MailPayload -Mail $item -IncludeBody:$IncludeBody
    $seen += 1
  }

  return $result
}

function Collect-MailPayloadsFromFolders {
  param(
    $Folders,
    [int]$Limit,
    [bool]$IncludeBody = $false,
    [string]$DateProperty = "ReceivedTime",
    [Nullable[datetime]]$StartDateBound = $null,
    [Nullable[datetime]]$EndDateBound = $null
  )

  $result = @()
  foreach ($folder in @($Folders)) {
    if ($result.Count -ge $Limit) {
      break
    }

    try {
      $items = Get-FilteredOutlookItems -Folder $folder -DateProperty $DateProperty -StartDateBound $StartDateBound -EndDateBound $EndDateBound
      $remaining = [Math]::Max(0, $Limit - $result.Count)
      $payloads = Collect-MailPayloads -Items $items -Limit $remaining -IncludeBody:$IncludeBody -DateProperty $DateProperty -StartDateBound $StartDateBound -EndDateBound $EndDateBound
      if ($payloads.Count -gt 0) {
        $result += $payloads
      }
    } catch {
      continue
    }
  }

  if ($result.Count -eq 0) {
    return @()
  }

  return @($result |
    Sort-Object {
      try { [datetime]$_.receivedDateTime } catch { [datetime]::MinValue }
    } -Descending |
    Select-Object -First $Limit)
}

function Get-CalendarPayload {
  param($Appointment)

  $subject = ""
  $location = ""
  $organizer = ""
  $body = ""
  $startDateTime = ""
  $endDateTime = ""

  try { $subject = [string]$Appointment.Subject } catch {}
  try { $location = [string]$Appointment.Location } catch {}
  try { $organizer = [string]$Appointment.Organizer } catch {}
  try { $body = [string]$Appointment.Body } catch {}
  try { $startDateTime = Format-JsonDateTime -Value $Appointment.Start } catch {}
  try { $endDateTime = Format-JsonDateTime -Value $Appointment.End } catch {}

  @{
    "subject" = $subject
    "location" = $location
    "organizer" = $organizer
    "startDateTime" = $startDateTime
    "endDateTime" = $endDateTime
    "bodyPreview" = if ($body.Length -gt 255) { $body.Substring(0, 255) } else { $body }
  }
}

function Collect-CalendarPayloads {
  param(
    $Items,
    [int]$Limit,
    [Nullable[datetime]]$StartDateBound = $null,
    [Nullable[datetime]]$EndDateBound = $null
  )

  $result = @()
  $seen = 0
  $item = $null
  try {
    $item = Invoke-WithRetry { $Items.GetFirst() }
  } catch {
    $item = $null
  }

  while ($null -ne $item -and $seen -lt $Limit) {
    $currentItem = $item
    try {
      $item = Invoke-WithRetry { $Items.GetNext() }
    } catch {
      $item = $null
    }

    if ($null -eq $currentItem) {
      continue
    }

    $itemClass = 0
    try {
      $itemClass = [int]$currentItem.Class
    } catch {
      Write-BridgeLog -EventName "bridge_item_skipped" -Status "warning" -Detail "Could not read Outlook calendar item class." -Properties @{
        action = $Action
      }
      continue
    }

    if ($itemClass -ne 26) {
      continue
    }

    $startValue = $null
    $endValue = $null
    try { $startValue = [datetime]$currentItem.Start } catch {}
    try { $endValue = [datetime]$currentItem.End } catch {}

    if ($null -ne $EndDateBound -and $null -ne $startValue -and $startValue -gt $EndDateBound) {
      continue
    }

    if ($null -ne $StartDateBound -and $null -ne $endValue -and $endValue -lt $StartDateBound) {
      continue
    }

    try {
      $result += Get-CalendarPayload -Appointment $currentItem
      $seen += 1
    } catch {
      $subject = ""
      try { $subject = [string]$currentItem.Subject } catch {}
      Write-BridgeLog -EventName "bridge_item_failed" -Status "warning" -Detail $_.Exception.Message -Properties @{
        action = $Action
        itemClass = $itemClass
        subject = $subject
      }
      continue
    }
  }

  Write-BridgeLog -EventName "bridge_calendar_collect_completed" -Status "info" -Detail "Collected calendar payloads from Outlook items." -Properties @{
    action = $Action
    collected = $result.Count
    limit = $Limit
  }

  return @($result |
    Sort-Object {
      try { [datetime]$_.startDateTime } catch { [datetime]::MaxValue }
    } |
    Select-Object -First $Limit)
}

Write-BridgeLog -EventName "bridge_action_start" -Status "info" -Detail "Outlook bridge action started." -Properties @{
  mailId = $MailId
  folderName = $FolderName
  subject = $Subject
  limit = $Limit
  startDate = $StartDate
  endDate = $EndDate
  composeType = $ComposeType
}

try {
  $session = Get-OutlookSession
  $mailboxStore = $null
  if (($Action -notin @('test', 'rewrite-active-compose')) -or ($Action -eq 'test' -and -not [string]::IsNullOrWhiteSpace($MailboxAddress))) {
    $mailboxStore = Resolve-MailboxStore -Session $session -Address $MailboxAddress
  }
  $startDateBound = if ([string]::IsNullOrWhiteSpace($StartDate)) { $null } else { [datetime]::Parse($StartDate) }
  $endDateBound = if ([string]::IsNullOrWhiteSpace($EndDate)) { $null } else { [datetime]::Parse($EndDate).Date.AddDays(1).AddSeconds(-1) }

  switch ($Action) {
    "test" {
      Emit-Json ((@{ ok = $true } | ConvertTo-Json -Compress -Depth 8))
    }
    "list-emails" {
      $folder = Invoke-WithRetry { $mailboxStore.GetDefaultFolder(6) }
      $items = Get-FilteredOutlookItems -Folder $folder -DateProperty "ReceivedTime" -StartDateBound $startDateBound -EndDateBound $endDateBound
      $result = Collect-MailPayloads -Items $items -Limit $Limit -IncludeBody:$false -DateProperty "ReceivedTime" -StartDateBound $startDateBound -EndDateBound $endDateBound

      if ($result.Count -eq 0) {
        Emit-Json "[]"
      } else {
        Emit-Json (ConvertTo-Json -InputObject @($result) -Compress -Depth 8)
      }
    }
    "list-emails-preview" {
      $folder = Invoke-WithRetry { $mailboxStore.GetDefaultFolder(6) }
      $items = Get-FilteredOutlookItems -Folder $folder -DateProperty "ReceivedTime" -StartDateBound $startDateBound -EndDateBound $endDateBound
      $result = Collect-MailPayloads -Items $items -Limit $Limit -IncludeBody:$false -DateProperty "ReceivedTime" -StartDateBound $startDateBound -EndDateBound $endDateBound

      if ($result.Count -eq 0) {
        Emit-Json "[]"
      } else {
        Emit-Json (ConvertTo-Json -InputObject @($result) -Compress -Depth 8)
      }
    }
    "list-search-emails" {
      $folders = Get-SearchFolders -Store $mailboxStore -Scope $FolderName
      $resolvedDateProperty = if ([string]::IsNullOrWhiteSpace($DateProperty)) {
        "ReceivedTime"
      } else {
        $DateProperty
      }

      $isSentItemsScope = [string]::Equals((Get-NonNullString $FolderName).Trim(), "sent-items", [System.StringComparison]::OrdinalIgnoreCase)
      if ($isSentItemsScope -and $Limit -gt 31 -and $null -ne $startDateBound -and $null -ne $endDateBound) {
        $collectedById = @{}
        $subwindowDays = 21
        $subwindowEnd = [datetime]$endDateBound

        while ($collectedById.Count -lt $Limit -and $subwindowEnd -ge [datetime]$startDateBound) {
          $subwindowStart = $subwindowEnd.Date.AddDays(1 - $subwindowDays)
          if ($subwindowStart -lt [datetime]$startDateBound) {
            $subwindowStart = [datetime]$startDateBound
          }

          $batch = Collect-MailPayloadsFromFolders -Folders $folders -Limit ([Math]::Min(31, $Limit - $collectedById.Count)) -IncludeBody:$false -DateProperty $resolvedDateProperty -StartDateBound $subwindowStart -EndDateBound $subwindowEnd
          foreach ($mail in @($batch)) {
            $mailId = Get-NonNullString $mail.id
            if (-not [string]::IsNullOrWhiteSpace($mailId)) {
              $collectedById[$mailId] = $mail
            }
          }

          if ($subwindowStart -le [datetime]$startDateBound) {
            break
          }

          $subwindowEnd = $subwindowStart.AddSeconds(-1)
        }

        $result = @($collectedById.Values |
          Sort-Object {
            try {
              if (-not [string]::IsNullOrWhiteSpace((Get-NonNullString $_.sentDateTime))) {
                return [datetime]$_.sentDateTime
              }

              if (-not [string]::IsNullOrWhiteSpace((Get-NonNullString $_.receivedDateTime))) {
                return [datetime]$_.receivedDateTime
              }

              return [datetime]::MinValue
            } catch {
              return [datetime]::MinValue
            }
          } -Descending |
          Select-Object -First $Limit)
      } else {
        $result = Collect-MailPayloadsFromFolders -Folders $folders -Limit $Limit -IncludeBody:$false -DateProperty $resolvedDateProperty -StartDateBound $startDateBound -EndDateBound $endDateBound
      }

      if ($result.Count -eq 0) {
        Emit-Json "[]"
      } else {
        Emit-Json (ConvertTo-Json -InputObject @($result) -Compress -Depth 8)
      }
    }
    "get-email" {
      $mail = Get-MailboxItem -Session $session -Store $mailboxStore -Id $MailId
      Emit-Json (((Get-MailPayload -Mail $mail) | ConvertTo-Json -Compress -Depth 8))
    }
    "list-attachments" {
      $mail = Get-MailboxItem -Session $session -Store $mailboxStore -Id $MailId
      $attachments = Get-AttachmentsPayload -Mail $mail
      if ($attachments.Count -eq 0) {
        Emit-Json "[]"
      } else {
        Emit-Json (ConvertTo-Json -InputObject @($attachments) -Compress -Depth 8)
      }
    }
    "download-attachment" {
      $mail = Get-MailboxItem -Session $session -Store $mailboxStore -Id $MailId
      $attachment = Invoke-WithRetry { $mail.Attachments.Item($AttachmentIndex) }
      Invoke-WithRetry { $attachment.SaveAsFile($DestPath) } | Out-Null
      Emit-Json ((@{ ok = $true } | ConvertTo-Json -Compress -Depth 8))
    }
    "create-reply" {
      $mail = Get-MailboxItem -Session $session -Store $mailboxStore -Id $MailId
      $reply = Invoke-WithRetry { $mail.Reply() }
      $html = Get-Content -Raw -Path $HtmlPath
      $reply.HTMLBody = $html + $reply.HTMLBody
      Invoke-WithRetry { $reply.Display() }
      Emit-Json ((@{ ok = $true } | ConvertTo-Json -Compress -Depth 8))
    }
    "create-reply-all" {
      $mail = Get-MailboxItem -Session $session -Store $mailboxStore -Id $MailId
      $reply = Invoke-WithRetry { $mail.ReplyAll() }
      $html = Get-Content -Raw -Path $HtmlPath
      $reply.HTMLBody = $html + $reply.HTMLBody
      Invoke-WithRetry { $reply.Display() }
      Emit-Json ((@{ ok = $true } | ConvertTo-Json -Compress -Depth 8))
    }
    "rewrite-active-compose" {
      $result = Rewrite-ActiveCompose -ExpectedSubject $Subject -HtmlSourcePath $HtmlPath -ReplacementText $DraftText -SeedPreview $SeedText -DisplayName $UserName -EmailAddress $MailboxAddress -Mode $ComposeType
      Emit-Json (($result | ConvertTo-Json -Compress -Depth 8))
    }
    "list-sent-emails" {
      $folder = Invoke-WithRetry { $mailboxStore.GetDefaultFolder(5) }  # 5 = Sent Items
      $items = Get-FilteredOutlookItems -Folder $folder -DateProperty "SentOn"
      $includeBody = [string]::Equals($BodyMode, "full", [System.StringComparison]::OrdinalIgnoreCase)
      $result = Collect-MailPayloads -Items $items -Limit $Limit -IncludeBody:$includeBody -DateProperty "SentOn"

      if ($result.Count -eq 0) {
        Emit-Json "[]"
      } else {
        Emit-Json (ConvertTo-Json -InputObject @($result) -Compress -Depth 8)
      }
    }
    "list-calendar-events" {
      $folder = Invoke-WithRetry { $mailboxStore.GetDefaultFolder(9) }  # 9 = Calendar
      $items = Get-FilteredOutlookItems -Folder $folder -DateProperty "Start" -StartDateBound $startDateBound -EndDateBound $endDateBound -Descending:$false
      $result = Collect-CalendarPayloads -Items $items -Limit $Limit -StartDateBound $startDateBound -EndDateBound $endDateBound

      if ($result.Count -eq 0) {
        Emit-Json "[]"
      } else {
        Emit-Json (ConvertTo-Json -InputObject @($result) -Compress -Depth 8)
      }
    }
    "set-read" {
      $mail = Get-MailboxItem -Session $session -Store $mailboxStore -Id $MailId
      $mail.UnRead = $false
      Invoke-WithRetry { $mail.Save() } | Out-Null
      Emit-Json ((@{ ok = $true } | ConvertTo-Json -Compress -Depth 8))
    }
    "move-mail" {
      $mail = Get-MailboxItem -Session $session -Store $mailboxStore -Id $MailId
      $store = $mail.Parent.Store
      $root = $store.GetRootFolder()
      $target = Find-FolderByName -ParentFolder $root -TargetName $FolderName
      if ($null -eq $target) {
        throw "Folder not found: $FolderName"
      }
      Invoke-WithRetry { $mail.Move($target) } | Out-Null
      Emit-Json ((@{ ok = $true } | ConvertTo-Json -Compress -Depth 8))
    }
    default {
      throw "Unknown action: $Action"
    }
  }

  Write-BridgeLog -EventName "bridge_action_complete" -Status "ok" -Detail "Outlook bridge action completed successfully." -Properties @{
    mailId = $MailId
    folderName = $FolderName
    limit = $Limit
  }
} catch {
  Write-BridgeLog -EventName "bridge_action_failed" -Status "error" -Detail $_.Exception.Message -Properties @{
    mailId = $MailId
    folderName = $FolderName
    subject = $Subject
    limit = $Limit
    startDate = $StartDate
    endDate = $EndDate
  }
  throw
}
