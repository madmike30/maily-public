param(
    [string]$MailboxAddress = "",
    [string]$ProfileId = "default",
    [switch]$FullWideSearch,
    [switch]$Live
)

$ErrorActionPreference = "Stop"

if (-not $Live) {
    [pscustomobject]@{ Mode = 'DryRun'; RequiresExplicitMailbox = $true; LiveAccess = $false }
    return
}
# Live mode reads multiple folders. Use only a separately approved synthetic test mailbox.
if ([string]::IsNullOrWhiteSpace($MailboxAddress)) {
    throw 'An explicit mailbox address is required for the search smoke suite.'
}
$MailboxAddress = $MailboxAddress.Trim()

Set-Location (Join-Path $PSScriptRoot "..")

$repoRoot = (Get-Location).Path
$runId = Get-Date -Format "yyyyMMdd-HHmmss"
$outputDir = Join-Path $repoRoot "output\live-search-smoke\$runId"
$summaryPath = Join-Path $outputDir "summary.json"
$hostHealthUrl = "https://localhost:7097/api/health"
$searchStartUrl = "https://localhost:7097/api/outlook/search/start"
$searchRefineUrl = "https://localhost:7097/api/outlook/search/refine"
$searchSessionUrl = "https://localhost:7097/api/outlook/search/session"

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

function Wait-Until {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Condition,
        [Parameter(Mandatory = $true)]
        [string]$Description,
        [int]$TimeoutSeconds = 30,
        [int]$PollMilliseconds = 500
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $result = & $Condition
        if ($result) {
            return $result
        }

        Start-Sleep -Milliseconds $PollMilliseconds
    }

    throw "Timed out while waiting for $Description."
}

function Ensure-HostReady {
    cmd /c run-maily-outlook-addin-latest.bat | Out-Host

    Wait-Until -Description "Maily search host health" -TimeoutSeconds 30 -Condition {
        try {
            $response = Invoke-WebRequest -UseBasicParsing -Uri $hostHealthUrl -TimeoutSec 5
            return $response.StatusCode -eq 200
        }
        catch {
            return $false
        }
    } | Out-Null
}

function Invoke-BridgeJson {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action,
        [hashtable]$Arguments = @{}
    )

    $argList = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", (Join-Path $repoRoot "scripts\outlook-bridge.ps1"),
        "-Action", $Action,
        "-MailboxAddress", $MailboxAddress
    )

    foreach ($entry in $Arguments.GetEnumerator()) {
        $argList += "-$($entry.Key)"
        $argList += [string]$entry.Value
    }

    $json = [string]::Join([Environment]::NewLine, @(& powershell @argList))
    if ($LASTEXITCODE -ne 0) {
        throw "Bridge action '$Action' failed."
    }

    if ([string]::IsNullOrWhiteSpace($json)) {
        return @()
    }

    return @($json | ConvertFrom-Json)
}

function Invoke-JsonPost {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $true)]
        [object]$Body
    )

    $payload = $Body | ConvertTo-Json -Depth 8
    $payloadBytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    $response = Invoke-WebRequest -UseBasicParsing -Uri $Uri -Method Post -Body $payloadBytes -ContentType "application/json; charset=utf-8"
    return $response.Content | ConvertFrom-Json
}

function Invoke-JsonGet {
    param([Parameter(Mandatory = $true)][string]$Uri)

    $response = Invoke-WebRequest -UseBasicParsing -Uri $Uri -Method Get
    return $response.Content | ConvertFrom-Json
}

function Save-JsonArtifact {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Value
    )

    $Value | ConvertTo-Json -Depth 12 | Set-Content -Path $Path -Encoding utf8
}

function Test-ContainsHebrew {
    param([string]$Value)

    return (($Value | ForEach-Object { if ($null -eq $_) { "" } else { [string]$_ } }) -match "[\u0590-\u05FF]")
}

function Normalize-Whitespace {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    return ([regex]::Replace($Value, "\s+", " ")).Trim()
}

function Get-SubjectQuery {
    param([string]$Subject)

    $subjectValue = if ($null -eq $Subject) { "" } else { [string]$Subject }
    $clean = Normalize-Whitespace ($subjectValue -replace "^(RE|FW|FWD)\s*:\s*", "")
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return "recent update"
    }

    if (Test-ContainsHebrew $clean) {
        return $clean
    }

    $tokens = [regex]::Matches($clean, "[A-Za-z0-9][A-Za-z0-9&\-/\.]{1,}") |
        ForEach-Object { $_.Value } |
        Where-Object { $_.Length -ge 3 } |
        Select-Object -First 5

    if ($tokens.Count -gt 0) {
        return ($tokens -join " ")
    }

    return $clean
}

function Get-SenderFilterValue {
    param($Mail)

    $from = $Mail.from.emailAddress.address
    if ([string]::IsNullOrWhiteSpace($from)) {
        return ""
    }

    return [string]$from
}

function Find-MailByBodySignal {
    param(
        [Parameter(Mandatory = $true)]$Mails,
        [Parameter(Mandatory = $true)][string[]]$Terms
    )

    foreach ($mail in $Mails) {
        $bodyPreview = Normalize-Whitespace $mail.bodyPreview
        $subject = Normalize-Whitespace $mail.subject
        $combined = "$subject`n$bodyPreview".ToLowerInvariant()
        foreach ($term in $Terms) {
            if ($combined.Contains($term.ToLowerInvariant())) {
                return $mail
            }
        }
    }

    return $null
}

function Find-HebrewMail {
    param([Parameter(Mandatory = $true)]$Mails)

    foreach ($mail in $Mails) {
        if (Test-ContainsHebrew (Normalize-Whitespace $mail.subject)) {
            return $mail
        }
    }

    return $null
}

function Get-MailScore {
    param(
        [Parameter(Mandatory = $true)]$Mail,
        [Parameter(Mandatory = $true)][string[]]$Terms
    )

    $combined = (((Normalize-Whitespace $Mail.subject) + " " + (Normalize-Whitespace $Mail.bodyPreview))).ToLowerInvariant()
    $score = 0
    foreach ($term in $Terms) {
        if ($combined.Contains($term.ToLowerInvariant())) {
            $score += 1
        }
    }

    return $score
}

function Find-StrongMailByTerms {
    param(
        [Parameter(Mandatory = $true)]$Mails,
        [Parameter(Mandatory = $true)][string[]]$RequiredTerms,
        [string[]]$PreferredTerms = @()
    )

    $ranked = foreach ($mail in $Mails) {
        $requiredScore = Get-MailScore -Mail $mail -Terms $RequiredTerms
        if ($requiredScore -lt $RequiredTerms.Count) {
            continue
        }

        $preferredScore = if ($PreferredTerms.Count -gt 0) { Get-MailScore -Mail $mail -Terms $PreferredTerms } else { 0 }
        [pscustomobject]@{
            mail = $mail
            requiredScore = $requiredScore
            preferredScore = $preferredScore
        }
    }

    return $ranked |
        Sort-Object `
            @{ Expression = "preferredScore"; Descending = $true }, `
            @{ Expression = { [datetimeoffset]$_.mail.receivedDateTime }; Descending = $true } |
        Select-Object -First 1 |
        ForEach-Object { $_.mail }
}

function Get-SubjectExpectationFragments {
    param([string]$Subject)

    $clean = Normalize-Whitespace ($Subject -replace "^(RE|FW|FWD)\s*:\s*", "")
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return @()
    }

    $fragments = @()
    $fragments += $clean

    $tokenFragments = [regex]::Matches($clean, "[\p{L}\p{Nd}][\p{L}\p{Nd}&/\-]{2,}") |
        ForEach-Object { $_.Value.Trim() } |
        Where-Object { $_.Length -ge 3 -and $_ -notin @("for", "the", "and") } |
        Select-Object -First 4

    $fragments += $tokenFragments
    return @($fragments | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Get-CombinedExpectationFragments {
    param([string[]]$Subjects)

    $all = foreach ($subject in $Subjects) {
        Get-SubjectExpectationFragments $subject
    }

    return @($all | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Assert-TopSubjectMatches {
    param(
        [Parameter(Mandatory = $true)]$Response,
        [Parameter(Mandatory = $true)][string[]]$ExpectedFragments,
        [Parameter(Mandatory = $true)][string]$CaseName
    )

    if ($Response.results.Count -lt 1) {
        throw "Case '$CaseName' returned no result rows."
    }

    $topSubject = Normalize-Whitespace $Response.results[0].subject
    $matched = $false
    foreach ($fragment in $ExpectedFragments) {
        if ($topSubject.ToLowerInvariant().Contains($fragment.ToLowerInvariant())) {
            $matched = $true
            break
        }
    }

    if (-not $matched) {
        throw "Case '$CaseName' top subject '$topSubject' did not match expected fragments: $($ExpectedFragments -join ', ')."
    }
}

function Get-ResultEvidenceText {
    param([Parameter(Mandatory = $true)]$Result)

    return (Normalize-Whitespace "$($Result.subject) $($Result.whyItMatches) $($Result.bodyPreview)").ToLowerInvariant()
}

function Test-ContainsAnyTerm {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string[]]$Terms
    )

    foreach ($term in $Terms) {
        if (-not [string]::IsNullOrWhiteSpace($term) -and $Text.Contains($term.ToLowerInvariant())) {
            return $true
        }
    }

    return $false
}

function Assert-ResultsContainGroupedEvidence {
    param(
        [Parameter(Mandatory = $true)]$Response,
        [Parameter(Mandatory = $true)][object[]]$RequiredGroups,
        [Parameter(Mandatory = $true)][string]$CaseName,
        [int]$MaxRank = 3
    )

    if ($Response.results.Count -lt 1) {
        throw "Case '$CaseName' returned no result rows."
    }

    $candidateResults = @($Response.results | Select-Object -First $MaxRank)
    foreach ($result in $candidateResults) {
        $text = Get-ResultEvidenceText -Result $result
        $matchedAllGroups = $true

        foreach ($group in $RequiredGroups) {
            if (-not (Test-ContainsAnyTerm -Text $text -Terms @($group))) {
                $matchedAllGroups = $false
                break
            }
        }

        if ($matchedAllGroups) {
            return
        }
    }

    $renderedGroups = @($RequiredGroups | ForEach-Object { "[" + ((@($_) -join ", ")) + "]" })
    throw "Case '$CaseName' did not return grouped evidence $($renderedGroups -join ' + ') within the top $MaxRank results."
}

function Assert-SearchResponse {
    param(
        [Parameter(Mandatory = $true)]$Response,
        [Parameter(Mandatory = $true)][string]$Stage
    )

    if ([string]::IsNullOrWhiteSpace($Response.meta.sessionId)) {
        throw "Missing search session id."
    }

    if ([string]::IsNullOrWhiteSpace($Response.answerSummary)) {
        throw "Answer summary was empty for stage '$Stage'."
    }

    if ($Stage -eq "refine" -and $Response.meta.evidenceMode -ne "hydrated-top-results") {
        throw "Expected hydrated evidence for refine stage."
    }

    if ($Response.meta.candidateCount -le 0 -and $Response.meta.resultCount -gt 0) {
        throw "Search stage '$Stage' reported results without candidates."
    }
}

function New-TestResult {
    param([Parameter(Mandatory = $true)][string]$Name)

    return [ordered]@{
        name = $Name
        status = "pending"
        errors = @()
        checks = [ordered]@{}
        artifacts = [ordered]@{}
    }
}

function Add-DateRangeExpansionResult {
    param(
        [Parameter(Mandatory = $true)]$Results,
        [Parameter(Mandatory = $true)][string]$OutputDir
    )

    $result = New-TestResult -Name "date-range-expansion-sent-autonomous"
    $caseDir = Join-Path $OutputDir "14-date-range-expansion-sent-autonomous"
    New-Item -ItemType Directory -Force -Path $caseDir | Out-Null

    try {
        $query = "autonomous"
        $start30 = Invoke-JsonPost -Uri $searchStartUrl -Body @{
            mailboxAddress = $MailboxAddress
            profileId = $ProfileId
            query = $query
            filters = @{ scope = "sent-items"; dateRange = "30d"; sender = ""; focus = "all" }
        }
        $start365 = Invoke-JsonPost -Uri $searchStartUrl -Body @{
            mailboxAddress = $MailboxAddress
            profileId = $ProfileId
            query = $query
            filters = @{ scope = "sent-items"; dateRange = "365d"; sender = ""; focus = "all" }
        }

        Assert-SearchResponse -Response $start30 -Stage "start"
        Assert-SearchResponse -Response $start365 -Stage "start"

        if ($start365.meta.candidateCount -le $start30.meta.candidateCount) {
            throw "Expected 365d candidate count to exceed 30d candidate count, but got $($start30.meta.candidateCount) vs $($start365.meta.candidateCount)."
        }

        $delta = $start365.meta.candidateCount - $start30.meta.candidateCount
        $materialIncrease = ($delta -ge 20) -or ($start365.meta.candidateCount -ge [math]::Ceiling($start30.meta.candidateCount * 1.2))
        if (-not $materialIncrease) {
            throw "Expected a material date-range expansion, but candidate count only moved from $($start30.meta.candidateCount) to $($start365.meta.candidateCount)."
        }

        $start30Path = Join-Path $caseDir "start-30d-response.json"
        $start365Path = Join-Path $caseDir "start-365d-response.json"
        Save-JsonArtifact -Path $start30Path -Value $start30
        Save-JsonArtifact -Path $start365Path -Value $start365

        $result.checks.query = $query
        $result.checks.candidate30d = $start30.meta.candidateCount
        $result.checks.candidate365d = $start365.meta.candidateCount
        $result.checks.delta = $delta
        $result.checks.topSubject30d = if ($start30.results.Count -gt 0) { $start30.results[0].subject } else { "" }
        $result.checks.topSubject365d = if ($start365.results.Count -gt 0) { $start365.results[0].subject } else { "" }
        $result.artifacts.start30dResponse = $start30Path
        $result.artifacts.start365dResponse = $start365Path
        $result.status = "passed"
    }
    catch {
        $result.status = "failed"
        $result.errors += $_.Exception.Message
    }

    $Results += [pscustomobject]$result
    return $Results
}

Ensure-HostReady

$inboxMails = Invoke-BridgeJson -Action "list-search-emails" -Arguments @{
    FolderName = "inbox"
    StartDate = (Get-Date).AddDays(-180).ToString("yyyy-MM-dd")
    Limit = 80
}
$sentMails = Invoke-BridgeJson -Action "list-search-emails" -Arguments @{
    FolderName = "sent-items"
    StartDate = (Get-Date).AddDays(-365).ToString("yyyy-MM-dd")
    Limit = 180
}
$allMail = Invoke-BridgeJson -Action "list-search-emails" -Arguments @{
    FolderName = "all-mail"
    StartDate = (Get-Date).AddDays(-365).ToString("yyyy-MM-dd")
    Limit = 220
}

if ($inboxMails.Count -lt 1 -or $sentMails.Count -lt 1 -or $allMail.Count -lt 5) {
    throw "Not enough live mailbox data was returned to run the search smoke suite."
}

$commitmentMail = Find-MailByBodySignal -Mails $allMail -Terms @("i will", "we will", "send", "share", "follow up", "follow-up", "next step")
$decisionMail = Find-MailByBodySignal -Mails $allMail -Terms @("agreed", "approved", "confirmed", "decision", "decided", "resolved")
$hebrewMail = Find-HebrewMail -Mails $allMail
$autonomousPricingMail = Find-StrongMailByTerms -Mails $sentMails -RequiredTerms @("autonomous", "estimate") -PreferredTerms @("investment proposal", "pricing", "quote", "cost", "adb")
$autonomousQuoteMail = Find-StrongMailByTerms -Mails $sentMails -RequiredTerms @("adb", "cost") -PreferredTerms @("rough cost estimate", "non binding", "quote", "autonomous")
$mediumAllMailRange = if ($FullWideSearch) { "180d" } else { "90d" }
$wideAllMailRange = if ($FullWideSearch) { "365d" } else { "180d" }

$broadMail = $allMail | Where-Object { -not [string]::IsNullOrWhiteSpace($_.subject) } | Select-Object -First 1
$inboxMail = $inboxMails[0]
$sentMail = $sentMails[0]
$secondInboxMail = if ($inboxMails.Count -gt 1) { $inboxMails[1] } else { $inboxMail }
$secondSentMail = if ($sentMails.Count -gt 1) { $sentMails[1] } else { $sentMail }

if ($null -eq $commitmentMail) { $commitmentMail = $allMail[0] }
if ($null -eq $decisionMail) { $decisionMail = $allMail[1] }
if ($null -eq $hebrewMail) { $hebrewMail = $inboxMails[0] }
if ($null -eq $autonomousPricingMail) { $autonomousPricingMail = $sentMail }
if ($null -eq $autonomousQuoteMail) { $autonomousQuoteMail = $secondSentMail }

$cases = @(
    @{
        Name = "inbox-exact-subject"
        Start = @{ query = Get-SubjectQuery $inboxMail.subject; filters = @{ scope = "inbox"; dateRange = "90d"; sender = ""; focus = "all" } }
    },
    @{
        Name = "sent-items-exact-subject"
        Start = @{ query = Get-SubjectQuery $sentMail.subject; filters = @{ scope = "sent-items"; dateRange = "365d"; sender = ""; focus = "all" } }
    },
    @{
        Name = "all-mail-broad-topic"
        Start = @{ query = Get-SubjectQuery $broadMail.subject; filters = @{ scope = "all-mail"; dateRange = $mediumAllMailRange; sender = ""; focus = "all" } }
    },
    @{
        Name = "sender-query"
        Start = @{ query = (Normalize-Whitespace $secondInboxMail.from.emailAddress.name); filters = @{ scope = "all-mail"; dateRange = $mediumAllMailRange; sender = ""; focus = "all" } }
    },
    @{
        Name = "hebrew-thread-query"
        Start = @{ query = Get-SubjectQuery $hebrewMail.subject; filters = @{ scope = "all-mail"; dateRange = $wideAllMailRange; sender = ""; focus = "all" } }
        ExpectedTopSubjectFragments = Get-SubjectExpectationFragments $hebrewMail.subject
    },
    @{
        Name = "commitments-focus"
        Start = @{ query = Get-SubjectQuery $commitmentMail.subject; filters = @{ scope = "all-mail"; dateRange = $wideAllMailRange; sender = ""; focus = "commitments" } }
    },
    @{
        Name = "decisions-focus"
        Start = @{ query = Get-SubjectQuery $decisionMail.subject; filters = @{ scope = "all-mail"; dateRange = $wideAllMailRange; sender = ""; focus = "decisions" } }
    },
    @{
        Name = "refine-by-sender"
        Start = @{ query = Get-SubjectQuery $inboxMail.subject; filters = @{ scope = "all-mail"; dateRange = $mediumAllMailRange; sender = ""; focus = "all" } }
        Refine = @{ sender = Get-SenderFilterValue $inboxMail; scope = "all-mail"; dateRange = $mediumAllMailRange; focus = "all" }
    },
    @{
        Name = "refine-by-scope"
        Start = @{ query = Get-SubjectQuery $secondSentMail.subject; filters = @{ scope = "all-mail"; dateRange = $wideAllMailRange; sender = ""; focus = "all" } }
        Refine = @{ sender = ""; scope = "sent-items"; dateRange = "365d"; focus = "all" }
    },
    @{
        Name = "refine-by-focus-and-time"
        Start = @{ query = Get-SubjectQuery $commitmentMail.subject; filters = @{ scope = "all-mail"; dateRange = $wideAllMailRange; sender = ""; focus = "all" } }
        Refine = @{ sender = ""; scope = "all-mail"; dateRange = "90d"; focus = "commitments" }
    },
    @{
        Name = "hebrew-autonomous-pricing-sent"
        Start = @{ query = "הצעות מחיר ששלחתי על אוטונומוס"; filters = @{ scope = "sent-items"; dateRange = "365d"; sender = ""; focus = "all" } }
        EvidenceRequiredGroups = @(
            @("autonomous", "adb", "אוטונומ", "lakehouse"),
            @("pricing", "commercial", "estimate", "quote", "cost", "מחיר", "תמחור", "הצעת מחיר")
        )
    },
    @{
        Name = "english-autonomous-pricing-sent"
        Start = @{ query = "emails I sent about autonomous database pricing"; filters = @{ scope = "sent-items"; dateRange = "365d"; sender = ""; focus = "all" } }
        EvidenceRequiredGroups = @(
            @("autonomous", "adb", "lakehouse"),
            @("pricing", "commercial", "estimate", "quote", "cost")
        )
    },
    @{
        Name = "rough-cost-estimate-adb"
        Start = @{ query = "non binding rough cost estimate adb"; filters = @{ scope = "sent-items"; dateRange = "365d"; sender = ""; focus = "all" } }
        EvidenceRequiredGroups = @(
            @("adb", "autonomous", "lakehouse"),
            @("estimate", "cost", "quote", "pricing", "commercial"),
            @("non binding", "rough cost", "investment proposal", "monthly total", "oci credits")
        )
    }
)

$results = @()

for ($index = 0; $index -lt $cases.Count; $index++) {
    $case = $cases[$index]
    $result = New-TestResult -Name $case.Name
    $caseDir = Join-Path $outputDir ("{0:D2}-{1}" -f ($index + 1), $case.Name)
    New-Item -ItemType Directory -Force -Path $caseDir | Out-Null

    try {
        $startBody = @{
            mailboxAddress = $MailboxAddress
            profileId = $ProfileId
            query = $case.Start.query
            filters = $case.Start.filters
        }

        $startResponse = Invoke-JsonPost -Uri $searchStartUrl -Body $startBody
        Assert-SearchResponse -Response $startResponse -Stage "start"

        $startArtifact = Join-Path $caseDir "start-response.json"
        Save-JsonArtifact -Path $startArtifact -Value $startResponse

        $result.checks.startQuery = $case.Start.query
        $result.checks.startScope = $case.Start.filters.scope
        $result.checks.startDateRange = $case.Start.filters.dateRange
        $result.checks.startFocus = $case.Start.filters.focus
        $result.checks.startCandidates = $startResponse.meta.candidateCount
        $result.checks.startResults = $startResponse.meta.resultCount
        $result.checks.startEvidence = $startResponse.meta.evidenceMode
        $result.checks.topSubject = if ($startResponse.results.Count -gt 0) { $startResponse.results[0].subject } else { "" }
        $result.artifacts.startResponse = $startArtifact

        if ($case.ContainsKey("ExpectedTopSubjectFragments")) {
            Assert-TopSubjectMatches -Response $startResponse -ExpectedFragments $case.ExpectedTopSubjectFragments -CaseName $case.Name
            $result.checks.expectedTopSubjectFragments = $case.ExpectedTopSubjectFragments
        }

        if ($case.ContainsKey("EvidenceRequiredGroups")) {
            Assert-ResultsContainGroupedEvidence -Response $startResponse -RequiredGroups $case.EvidenceRequiredGroups -CaseName $case.Name
            $result.checks.evidenceRequiredGroups = $case.EvidenceRequiredGroups
        }

        if ($case.ContainsKey("Refine")) {
            $refineBody = @{
                sessionId = $startResponse.meta.sessionId
                filters = $case.Refine
            }

            $refineResponse = Invoke-JsonPost -Uri $searchRefineUrl -Body $refineBody
            Assert-SearchResponse -Response $refineResponse -Stage "refine"

            if ($refineResponse.meta.candidateCount -gt $startResponse.meta.candidateCount) {
                throw "Refinement increased candidate count from $($startResponse.meta.candidateCount) to $($refineResponse.meta.candidateCount)."
            }

            $sessionResponse = Invoke-JsonGet -Uri "$searchSessionUrl/$($startResponse.meta.sessionId)"
            $refineArtifact = Join-Path $caseDir "refine-response.json"
            $sessionArtifact = Join-Path $caseDir "session-response.json"
            Save-JsonArtifact -Path $refineArtifact -Value $refineResponse
            Save-JsonArtifact -Path $sessionArtifact -Value $sessionResponse

            $result.checks.refineScope = $case.Refine.scope
            $result.checks.refineDateRange = $case.Refine.dateRange
            $result.checks.refineFocus = $case.Refine.focus
            $result.checks.refineSender = $case.Refine.sender
            $result.checks.refineCandidates = $refineResponse.meta.candidateCount
            $result.checks.refineResults = $refineResponse.meta.resultCount
            $result.checks.refineEvidence = $refineResponse.meta.evidenceMode
            $result.checks.refineTopSubject = if ($refineResponse.results.Count -gt 0) { $refineResponse.results[0].subject } else { "" }
            $result.artifacts.refineResponse = $refineArtifact
            $result.artifacts.sessionResponse = $sessionArtifact
        }

        $result.status = "passed"
    }
    catch {
        $result.status = "failed"
        $result.errors += $_.Exception.Message
    }

    $results += [pscustomobject]$result
}

$results = Add-DateRangeExpansionResult -Results $results -OutputDir $outputDir

$summary = [pscustomobject]@{
    runId = $runId
    generatedAtUtc = [DateTime]::UtcNow.ToString("o")
    mailboxAddress = $MailboxAddress
    fullWideSearch = [bool]$FullWideSearch
    outputDir = $outputDir
    results = $results
}

Save-JsonArtifact -Path $summaryPath -Value $summary
$summary | ConvertTo-Json -Depth 10

if ($results | Where-Object { $_.status -ne "passed" }) {
    exit 1
}
