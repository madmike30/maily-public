param(
    [switch]$KeepDraftsOpen
)

$ErrorActionPreference = "Stop"

Set-Location (Join-Path $PSScriptRoot "..")

$repoRoot = (Get-Location).Path
$runId = Get-Date -Format "yyyyMMdd-HHmmss"
$outputDir = Join-Path $repoRoot "output\live-outlook\$runId"
$installedRoot = Join-Path $env:LOCALAPPDATA "Maily\OutlookAddin"
$useInstalledHost = Test-Path -LiteralPath (Join-Path $installedRoot "app\Maily.OutlookAddin.Web\Maily.OutlookAddin.Web.exe")
$logRoot = if ($useInstalledHost) {
    Join-Path $installedRoot "logs"
} else {
    Join-Path $repoRoot "output\logs"
}
$requestLogPath = Join-Path $logRoot ("addin-web-requests-{0}.log" -f (Get-Date -Format "yyyyMMdd"))
$appLogPath = Join-Path $logRoot ("addin-web-{0}.jsonl" -f (Get-Date -Format "yyyyMMdd"))
$summaryPath = Join-Path $outputDir "summary.json"

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

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
    if ($useInstalledHost) {
        $supervisorExe = Join-Path $installedRoot "app\Maily.OutlookAddin.Supervisor\Maily.OutlookAddin.Supervisor.exe"
        if ((Test-Path -LiteralPath $supervisorExe) -and -not (Get-Process -Name Maily.OutlookAddin.Supervisor -ErrorAction SilentlyContinue)) {
            Start-Process -FilePath $supervisorExe -WorkingDirectory (Split-Path -Parent $supervisorExe) -WindowStyle Hidden
        }
    }
    else {
        cmd /c run-maily-outlook-addin-latest.bat | Out-Host
    }

    Wait-Until -Description "Maily add-in host health endpoint" -TimeoutSeconds 30 -Condition {
        try {
            $response = Invoke-WebRequest -UseBasicParsing -Uri "http://localhost:5289/api/health" -TimeoutSec 5
            return $response.StatusCode -eq 200
        }
        catch {
            return $false
        }
    } | Out-Null
}

function Test-IsTransientOutlookError {
    param(
        [AllowNull()]
        [System.Exception]$Exception
    )

    if (-not $Exception) {
        return $false
    }

    $message = [string]$Exception.Message
    return $message -match "RPC server is unavailable|has disconnected from its clients|application is busy|Call was rejected by callee"
}

function Invoke-OutlookOperationWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Operation,
        [Parameter(Mandatory = $true)]
        [string]$Description,
        [int]$Attempts = 3,
        [int]$DelayMilliseconds = 1500
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            return & $Operation
        }
        catch {
            if ($attempt -ge $Attempts -or -not (Test-IsTransientOutlookError -Exception $_.Exception)) {
                throw
            }

            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
}

function Get-OutlookApplication {
    return Invoke-OutlookOperationWithRetry -Description "Outlook application connection" -Operation {
        New-Object -ComObject Outlook.Application
    }
}

function Get-RootWindowByCaption {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CaptionPrefix
    )

    return Wait-Until -Description "Outlook window '$CaptionPrefix'" -TimeoutSeconds 20 -Condition {
        $root = [System.Windows.Automation.AutomationElement]::RootElement
        $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
        for ($i = 0; $i -lt $windows.Count; $i++) {
            $window = $windows.Item($i)
            if ($window.Current.ClassName -eq "rctrl_renwnd32" -and $window.Current.Name.StartsWith($CaptionPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $window
            }
        }

        return $null
    }
}

function Get-InspectorWindow {
    param(
        [Parameter(Mandatory = $true)]
        $Inspector,
        [Parameter(Mandatory = $true)]
        [string]$WindowCaption
    )

    try {
        $handle = [int]$Inspector.HWND
        if ($handle -gt 0) {
            $element = [System.Windows.Automation.AutomationElement]::FromHandle([System.IntPtr]::new($handle))
            if ($element) {
                return $element
            }
        }
    }
    catch {
    }

    return Get-RootWindowByCaption -CaptionPrefix $WindowCaption
}

function Get-WindowButtonNames {
    param(
        [Parameter(Mandatory = $true)]
        $Window,
        [int]$Limit = 30
    )

    $buttonCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button)
    $buttons = $Window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $buttonCondition)
    $names = New-Object System.Collections.Generic.List[string]

    for ($i = 0; $i -lt $buttons.Count; $i++) {
        $name = [string]$buttons.Item($i).Current.Name
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if (-not $names.Contains($name)) {
            $names.Add($name)
        }

        if ($names.Count -ge $Limit) {
            break
        }
    }

    return @($names)
}

function Get-MailySideloadState {
    $developerKeyPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer\683ea353-6d71-4315-afd4-147e968605a9"
    $manifestPath = Join-Path $repoRoot "addin\manifest.xml"

    return [pscustomobject]@{
        developerKeyPresent = Test-Path $developerKeyPath
        manifestPresent = Test-Path $manifestPath
        manifestVersion = if (Test-Path $manifestPath) {
            try {
                ([xml](Get-Content -LiteralPath $manifestPath -Raw)).OfficeApp.Version
            }
            catch {
                ""
            }
        } else {
            ""
        }
    }
}

function Expand-MailyRibbonMenu {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Window
    )

    $menuCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::MenuItem)
    $menus = $Window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $menuCondition)
    for ($i = 0; $i -lt $menus.Count; $i++) {
        $candidate = $menus.Item($i)
        $name = [string]$candidate.Current.Name
        if (-not [string]::Equals($name, "Maily", [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $expandPattern = $null
        if ($candidate.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$expandPattern)) {
            try {
                if ($expandPattern.Current.ExpandCollapseState -ne [System.Windows.Automation.ExpandCollapseState]::Expanded) {
                    $expandPattern.Expand()
                    Start-Sleep -Milliseconds 750
                }
                return $true
            }
            catch {
                return $false
            }
        }
    }

    return $false
}

function Find-GenerateReplyButton {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Window
    )

    $buttonCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button)
    $buttons = $Window.FindAll([System.Windows.Automation.TreeScope]::Descendants, $buttonCondition)
    for ($i = 0; $i -lt $buttons.Count; $i++) {
        $candidate = $buttons.Item($i)
        $name = [string]$candidate.Current.Name
        if ([string]::Equals($name, "Generate reply", [System.StringComparison]::OrdinalIgnoreCase) -or
            ($name.IndexOf("Generate", [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -and
             $name.IndexOf("reply", [System.StringComparison]::OrdinalIgnoreCase) -ge 0)) {
            return $candidate
        }
    }

    return $null
}

function Invoke-GenerateReplyRibbonButton {
    param(
        [Parameter(Mandatory = $true)]
        $Inspector,
        [Parameter(Mandatory = $true)]
        [string]$WindowCaption
    )

    $shell = New-Object -ComObject WScript.Shell
    $targetButton = $null
    $buttonNames = @()
    $deadline = (Get-Date).AddSeconds(25)
    while ((Get-Date) -lt $deadline) {
        try {
            $Inspector.Activate() | Out-Null
        }
        catch {
        }

        try {
            $shell.AppActivate($WindowCaption) | Out-Null
        }
        catch {
        }

        Start-Sleep -Milliseconds 250
        $window = Get-InspectorWindow -Inspector $Inspector -WindowCaption $WindowCaption
        if (-not $window) {
            Start-Sleep -Milliseconds 500
            continue
        }

        $buttonNames = Get-WindowButtonNames -Window $window
        $targetButton = Find-GenerateReplyButton -Window $window
        if (-not $targetButton -and (Expand-MailyRibbonMenu -Window $window)) {
            $targetButton = Find-GenerateReplyButton -Window $window
        }

        if ($targetButton) {
            break
        }

        Start-Sleep -Milliseconds 500
    }

    if (-not $targetButton) {
        $sideloadState = Get-MailySideloadState
        $availableButtons = if ($buttonNames.Count -gt 0) { $buttonNames -join "; " } else { "(none discovered)" }
        throw "Could not find the Outlook ribbon 'Generate reply' button. Visible button names: $availableButtons. SideloadKeyPresent: $($sideloadState.developerKeyPresent). ManifestPresent: $($sideloadState.manifestPresent). ManifestVersion: $($sideloadState.manifestVersion)."
    }

    $invokePattern = $null
    if (-not $targetButton.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$invokePattern)) {
        throw "The Outlook ribbon 'Generate reply' button does not expose an invoke pattern."
    }

    $invokePattern.Invoke()
}

function Invoke-NativeComposeUndo {
    param(
        [Parameter(Mandatory = $true)]
        $Inspector,
        [Parameter(Mandatory = $true)]
        [string]$WindowCaption
    )

    try {
        $Inspector.Activate() | Out-Null
    }
    catch {
    }

    try {
        $wordEditor = $Inspector.WordEditor
        $wordEditor.Windows.Item(1).Activate() | Out-Null
    }
    catch {
    }

    try {
        $wordBasic = $Inspector.WordEditor.Application.WordBasic
        $wordBasic.EditUndo() | Out-Null
        return "wordbasic-editundo"
    }
    catch {
    }

    $shell = New-Object -ComObject WScript.Shell
    if (-not $shell.AppActivate($WindowCaption)) {
        throw "Could not activate Outlook compose window '$WindowCaption' for undo."
    }

    Start-Sleep -Milliseconds 500
    $shell.SendKeys("^z")
    return "sendkeys"
}

function Get-JsonLogEntriesSince {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [datetimeoffset]$SinceUtc
    )

    if (-not (Test-Path $Path)) {
        return @()
    }

    $entries = @()
    foreach ($line in Get-Content -LiteralPath $Path) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $entry = $line | ConvertFrom-Json
            $timestamp = [datetimeoffset]::Parse($entry.timestamp)
            if ($timestamp -ge $SinceUtc) {
                $entries += $entry
            }
        }
        catch {
            continue
        }
    }

    return $entries
}

function Get-RequestLogEntriesSince {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [datetimeoffset]$SinceUtc
    )

    if (-not (Test-Path $Path)) {
        return @()
    }

    $entries = @()
    foreach ($line in Get-Content -LiteralPath $Path) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $parts = $line -split "`t"
        if ($parts.Count -lt 3) {
            continue
        }

        try {
            $timestamp = [datetimeoffset]::Parse($parts[0])
        }
        catch {
            continue
        }

        if ($timestamp -lt $SinceUtc) {
            continue
        }

        $entries += [pscustomobject]@{
            timestamp = $timestamp
            method = $parts[1]
            path = $parts[2]
            query = if ($parts.Count -ge 4) { $parts[3] } else { "" }
            statusCode = if ($parts.Count -ge 5) { $parts[4] } elseif ($parts.Count -ge 4 -and $parts[3] -match "^\d{3}$") { $parts[3] } else { "" }
        }
    }

    return $entries
}

function Get-LatestClientEvent {
    param(
        [AllowNull()]
        [object[]]$Entries,
        [Parameter(Mandatory = $true)]
        [string]$EventName
    )

    if (-not $Entries) {
        return $null
    }

    return $Entries |
        Where-Object { $_.category -eq "OutlookClient" -and $_.properties.EventName -eq $EventName } |
        Select-Object -Last 1
}

function Save-TextArtifact {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    Set-Content -LiteralPath $Path -Value $Value -Encoding UTF8
}

function Get-FirstInboxMail {
    param(
        [Parameter(Mandatory = $true)]
        $Outlook
    )

    return Invoke-OutlookOperationWithRetry -Description "Inbox mail discovery" -Operation {
        $namespace = $Outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)
        $items = $inbox.Items
        $items.Sort("[ReceivedTime]", $true)

        for ($index = 1; $index -le [Math]::Min($items.Count, 50); $index++) {
            try {
                $candidate = $items.Item($index)
            }
            catch {
                continue
            }

            if ($candidate -and $candidate.Class -eq 43 -and -not [string]::IsNullOrWhiteSpace([string]$candidate.Subject)) {
                return $candidate
            }
        }

        throw "Could not find a usable mail item in the Outlook Inbox for the reply-thread test."
    }
}

function New-TestCaseResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    return [ordered]@{
        name = $Name
        status = "pending"
        errors = @()
        artifacts = @{}
        checks = [ordered]@{}
    }
}

function Run-LiveComposeCase {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [scriptblock]$CreateDraft,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$SeedText,
        [Parameter(Mandatory = $true)]
        [scriptblock]$ValidateAfterBody,
        [scriptblock]$ValidateUndoBody = $null,
        [string]$ExpectedInsertMode = "replace-user-draft",
        [string]$ExpectedInsertStrategy = "bridge",
        [switch]$SkipRewriteRequest,
        [switch]$AllowMissingAfterBody
    )

    $caseResult = New-TestCaseResult -Name $Name
    $outlook = Get-OutlookApplication
    $draft = $null
    $inspector = $null

    try {
        $draft = Invoke-OutlookOperationWithRetry -Description "$Name draft creation" -Operation {
            & $CreateDraft $outlook
        }
        $inspector = Invoke-OutlookOperationWithRetry -Description "$Name inspector access" -Operation {
            $draft.GetInspector
        }
        $caption = [string]$inspector.Caption
        $window = Get-InspectorWindow -Inspector $inspector -WindowCaption $caption.Trim()

        $beforeBody = [string]$draft.HTMLBody
        $beforeText = [string]$inspector.WordEditor.Content.Text
        $beforePath = Join-Path $outputDir "$Name.before.html"
        Save-TextArtifact -Path $beforePath -Value $beforeBody

        $caseResult.artifacts.beforeHtml = $beforePath
        $caseResult.checks.caption = $caption.Trim()
        $caseResult.checks.seedText = $SeedText

        $startedAtUtc = [datetimeoffset]::UtcNow
        Invoke-GenerateReplyRibbonButton -Inspector $inspector -WindowCaption $caption.Trim()

        $completeEvent = Wait-Until -Description "$Name generate-complete log event" -TimeoutSeconds 45 -Condition {
            $entries = Get-JsonLogEntriesSince -Path $appLogPath -SinceUtc $startedAtUtc
            $event = Get-LatestClientEvent -Entries $entries -EventName "generate-complete"
            if ($event) {
                return $event
            }

            return $null
        }

        $rewriteRequest = $null
        if (-not $SkipRewriteRequest) {
            $rewriteRequest = Wait-Until -Description "$Name rewrite-active-compose request" -TimeoutSeconds 20 -Condition {
                $entries = Get-RequestLogEntriesSince -Path $requestLogPath -SinceUtc $startedAtUtc
                $match = $entries | Where-Object { $_.path -eq "/api/outlook/rewrite-active-compose" -and $_.statusCode -eq "200" } | Select-Object -Last 1
                if ($match) {
                    return $match
                }

                return $null
            }
        }

        $requestEntries = Get-RequestLogEntriesSince -Path $requestLogPath -SinceUtc $startedAtUtc
        $appEntries = Get-JsonLogEntriesSince -Path $appLogPath -SinceUtc $startedAtUtc
        $startEvent = Get-LatestClientEvent -Entries $appEntries -EventName "generate-start"
        $taskpaneEvent = Get-LatestClientEvent -Entries $appEntries -EventName "taskpane-open"

        $afterBody = ""
        try {
            Start-Sleep -Seconds 2
            $afterBody = Wait-Until -Description "$Name compose body rewrite" -TimeoutSeconds 30 -Condition {
                try {
                    $currentItem = $inspector.CurrentItem
                    $htmlBody = [string]$currentItem.HTMLBody
                    if ($htmlBody -and $htmlBody -ne $beforeBody) {
                        return $htmlBody
                    }
                }
                catch {
                }

                return $null
            }
        }
        catch {
            if (-not $AllowMissingAfterBody) {
                throw
            }
        }

        $afterPath = Join-Path $outputDir "$Name.after.html"
        $undoPath = Join-Path $outputDir "$Name.undo.txt"
        $appLogSlicePath = Join-Path $outputDir "$Name.app-log.json"
        $requestLogSlicePath = Join-Path $outputDir "$Name.request-log.json"
        $afterText = ""

        if ($afterBody) {
            Save-TextArtifact -Path $afterPath -Value $afterBody
            $caseResult.artifacts.afterHtml = $afterPath
            try {
                $afterText = [string]$inspector.WordEditor.Content.Text
            }
            catch {
                $afterText = ""
            }
        }

        $undoBody = ""
        $undoMethod = ""
        if ($ValidateUndoBody -and $afterBody -and $afterText) {
            $undoMethod = Invoke-NativeComposeUndo -Inspector $inspector -WindowCaption $caption.Trim()
            $undoBody = Wait-Until -Description "$Name undo" -TimeoutSeconds 20 -Condition {
                try {
                    $textBody = [string]$inspector.WordEditor.Content.Text
                    if ($textBody -and
                        $textBody -ne $afterText -and
                        $textBody.IndexOf($SeedText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                        return $textBody
                    }
                }
                catch {
                }

                return $null
            }
        }

        if ($undoBody) {
            Save-TextArtifact -Path $undoPath -Value $undoBody
            $caseResult.artifacts.undoHtml = $undoPath
        }
        ($appEntries | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $appLogSlicePath -Encoding UTF8
        ($requestEntries | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $requestLogSlicePath -Encoding UTF8

        $caseResult.artifacts.appLog = $appLogSlicePath
        $caseResult.artifacts.requestLog = $requestLogSlicePath
        $caseResult.checks.beforeBodyLength = $beforeBody.Length
        $caseResult.checks.afterBodyLength = $afterBody.Length
        $caseResult.checks.afterBodyReadback = if ($afterBody) { "captured" } else { "unavailable" }
        $caseResult.checks.undoBodyLength = $undoBody.Length
        $caseResult.checks.undoMethod = $undoMethod
        $caseResult.checks.undoVerified = [bool]$undoBody
        $caseResult.checks.requestPaths = @($requestEntries | ForEach-Object { "{0} {1}{2} {3}" -f $_.method, $_.path, $_.query, $_.statusCode })
        $caseResult.checks.generateStart = if ($startEvent) { $startEvent.properties.Detail } else { "" }
        $caseResult.checks.generateComplete = $completeEvent.properties.Detail
        $caseResult.checks.rewriteRequestStatus = if ($rewriteRequest) { $rewriteRequest.statusCode } else { "" }

        if (-not $taskpaneEvent) {
            $caseResult.errors += "No taskpane-open event was recorded after the Outlook ribbon click."
        }

        if (-not ($requestEntries | Where-Object { $_.path -like "/addin/taskpane.html*" })) {
            $caseResult.errors += "The request log did not show the task pane HTML request."
        }

        if (-not ($requestEntries | Where-Object { $_.path -eq "/api/outlook/generate-reply" -and $_.statusCode -eq "200" })) {
            $caseResult.errors += "The request log did not show a successful /api/outlook/generate-reply request."
        }

        if ($SkipRewriteRequest) {
            if ($requestEntries | Where-Object { $_.path -eq "/api/outlook/rewrite-active-compose" -and $_.statusCode -eq "200" }) {
                $caseResult.errors += "The request log unexpectedly showed /api/outlook/rewrite-active-compose for a no-seed prepend scenario."
            }
        }
        elseif (-not ($requestEntries | Where-Object { $_.path -eq "/api/outlook/rewrite-active-compose" -and $_.statusCode -eq "200" })) {
            $caseResult.errors += "The request log did not show a successful /api/outlook/rewrite-active-compose request."
        }

        if (-not $startEvent) {
            $caseResult.errors += "No generate-start client event was recorded."
        }
        elseif ($startEvent.properties.Detail -notmatch [regex]::Escape("SeedPreview: $SeedText")) {
            $caseResult.errors += "The generate-start log did not capture the expected seed text."
        }
        elseif ($startEvent.properties.Detail -notmatch "HostPlatform: outlook-classic") {
            $caseResult.errors += "The generate-start log did not identify the compose host as outlook-classic."
        }

        if ($completeEvent.properties.Detail -notmatch [regex]::Escape("InsertMode: $ExpectedInsertMode")) {
            $caseResult.errors += "The generate-complete log did not report the expected insert mode '$ExpectedInsertMode'."
        }
        if (-not [string]::IsNullOrWhiteSpace($ExpectedInsertStrategy) -and $completeEvent.properties.Detail -notmatch [regex]::Escape("InsertStrategy: $ExpectedInsertStrategy")) {
            $caseResult.errors += "The generate-complete log did not report the expected insert strategy '$ExpectedInsertStrategy'."
        }

        if (-not $AllowMissingAfterBody -and (-not $afterBody -or $afterBody -eq $beforeBody)) {
            $caseResult.errors += "The final compose body did not change after the bridge rewrite completed."
        }

        $validationErrors = & $ValidateAfterBody ([pscustomobject]@{
            beforeHtml = $beforeBody
            afterHtml = $afterBody
            beforeText = $beforeText
            afterText = $afterText
        })
        if ($validationErrors) {
            $caseResult.errors += @($validationErrors)
        }

        if ($ValidateUndoBody) {
            if (-not $undoBody) {
                $caseResult.errors += "Ctrl+Z did not restore the pre-generation draft."
            }
            else {
                $undoValidationErrors = & $ValidateUndoBody ([pscustomobject]@{
                    beforeHtml = $beforeBody
                    beforeText = $beforeText
                    afterHtml = $afterBody
                    afterText = $afterText
                    undoText = $undoBody
                })
                if ($undoValidationErrors) {
                    $caseResult.errors += @($undoValidationErrors)
                }
            }
        }

        $caseResult.status = if ($caseResult.errors.Count -eq 0) { "passed" } else { "failed" }
        return [pscustomobject]$caseResult
    }
    catch {
        $caseResult.errors += $_.Exception.Message
        $caseResult.status = "failed"
        return [pscustomobject]$caseResult
    }
    finally {
        if (-not $KeepDraftsOpen -and $draft) {
            try {
                $draft.Close(1) | Out-Null
            }
            catch {
            }
        }
    }
}

function Get-OutlookCurrentUserDisplayName {
    param($Outlook)

    try {
        $name = [string]$Outlook.Session.CurrentUser.Name
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            return $name
        }
    }
    catch {
    }

    return "Maily Tester"
}

function New-NewMailDraft {
    param($Outlook)

    $seedText = "need push to next wk asap"
    $displayName = Get-OutlookCurrentUserDisplayName -Outlook $Outlook
    $mail = $Outlook.CreateItem(0)
    $mail.Subject = "Maily live new mail $runId"
    $mail.BodyFormat = 2
    $mail.HTMLBody = "<div>$seedText</div><div><br></div><div>Regards,</div><div>$displayName</div>"
    $mail.Display()
    Start-Sleep -Seconds 2
    return $mail
}

function New-ReplyThreadDraft {
    param($Outlook)

    $seedText = "need push to next wk asap"
    $sourceMail = Get-FirstInboxMail -Outlook $Outlook
    $reply = $sourceMail.Reply()
    $reply.Display()
    Start-Sleep -Seconds 2
    $reply.HTMLBody = "<div>$seedText</div><div><br></div>" + [string]$reply.HTMLBody
    Start-Sleep -Seconds 1
    return $reply
}

function New-ReplyThreadDraftWithoutSeed {
    param($Outlook)

    $sourceMail = Get-FirstInboxMail -Outlook $Outlook
    $reply = $sourceMail.Reply()
    $reply.Display()
    Start-Sleep -Seconds 2
    return $reply
}

Ensure-HostReady

$results = @()

$results += Run-LiveComposeCase -Name "new-mail" -SeedText "need push to next wk asap" -CreateDraft {
    param($Outlook)
    New-NewMailDraft -Outlook $Outlook
} -ValidateAfterBody {
    param($State)

    $errors = @()
    if ($State.afterText.IndexOf("need push to next wk asap", [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        $errors += "The new-mail compose body still contains the original typed seed after rewrite."
    }
    if ($State.afterText.IndexOf("Regards,", [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
        $errors += "The new-mail compose body no longer contains the preserved signature."
    }

    return $errors
} -ValidateUndoBody {
    param($State)

    $errors = @()
    if ($State.undoText.IndexOf("need push to next wk asap", [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
        $errors += "Ctrl+Z did not restore the original typed seed."
    }
    if ($State.undoText.IndexOf("Regards,", [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
        $errors += "Ctrl+Z did not restore the preserved signature."
    }
    if ($State.beforeText.IndexOf("need push to next wk asap", [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
        return $errors
    }

    if ($State.undoText.Length -lt 8) {
        $errors += "Ctrl+Z restored too little text in the compose editor."
    }

    return $errors
}

$results += Run-LiveComposeCase -Name "reply-thread" -SeedText "need push to next wk asap" -CreateDraft {
    param($Outlook)
    New-ReplyThreadDraft -Outlook $Outlook
} -ValidateAfterBody {
    param($State)

    $errors = @()
    $threadMarkers = @(
        "divRplyFwdMsg",
        "Original Message",
        "From:",
        "wrote:"
    )

    $beforeMarker = $threadMarkers | Where-Object { $State.beforeHtml.IndexOf($_, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 } | Select-Object -First 1
    if (-not $beforeMarker) {
        $errors += "The reply-thread setup body did not contain a recognizable thread marker before generation."
        return $errors
    }

    if ($State.afterHtml -and $State.afterHtml.IndexOf($beforeMarker, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
        $errors += "The reply-thread compose body no longer contains the preserved thread marker '$beforeMarker'."
    }

    return $errors
} -AllowMissingAfterBody

$results += Run-LiveComposeCase -Name "reply-thread-no-seed" -SeedText "" -CreateDraft {
    param($Outlook)
    New-ReplyThreadDraftWithoutSeed -Outlook $Outlook
} -ExpectedInsertMode "prepend" -ExpectedInsertStrategy "" -SkipRewriteRequest -ValidateAfterBody {
    param($State)

    $errors = @()
    if ([string]::IsNullOrWhiteSpace($State.afterHtml)) {
        $errors += "The no-seed reply-thread compose body was not captured after generation."
        return $errors
    }

    if ($State.afterHtml -match [regex]::Escape("<!-- MAILY_REPLY_START -->") -and
        $State.afterHtml -eq $State.beforeHtml) {
        $errors += "The no-seed reply-thread body did not change after generation."
    }

    $threadMarkers = @(
        "divRplyFwdMsg",
        "Original Message",
        "From:",
        "wrote:"
    )

    $beforeMarker = $threadMarkers | Where-Object { $State.beforeHtml.IndexOf($_, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 } | Select-Object -First 1
    if ($beforeMarker -and $State.afterHtml.IndexOf($beforeMarker, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
        $errors += "The no-seed reply-thread body no longer contains the preserved thread marker '$beforeMarker'."
    }

    return $errors
}

$summary = [ordered]@{
    runId = $runId
    generatedAtUtc = [datetime]::UtcNow.ToString("o")
    appLogPath = $appLogPath
    requestLogPath = $requestLogPath
    outputDir = $outputDir
    results = $results
}

$summary | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
$summary | ConvertTo-Json -Depth 10

if ($results | Where-Object { $_.status -ne "passed" }) {
    exit 1
}
