param([string]$ScriptPath = (Join-Path $PSScriptRoot 'test-outlook-search-live-smoke.ps1'))
$ErrorActionPreference = 'Stop'
& (New-Module {}) {
    param($ScriptPath)
    $global:MailySearchSmokeCalls = [Collections.Generic.List[object]]::new()
    $global:MailySearchSmokeHostCalls = 0
    $global:MailySearchSmokeDirectoryCalls = 0
    # Intercept every external boundary used by the caller; never launch its host or bridge.
    function cmd { $global:MailySearchSmokeHostCalls++; return }
    function powershell {
        $global:MailySearchSmokeCalls.Add(@($args))
        $global:LASTEXITCODE = 0
        return '[]'
    }
    function Invoke-WebRequest { return @{ StatusCode = 200 } }
    function New-Item { $global:MailySearchSmokeDirectoryCalls++; return }
    function Set-Location { return }
    function Assert-True($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }

    $failedForEmptyFixture = $false
    try { & $ScriptPath -MailboxAddress 'synthetic@example.invalid' -Live | Out-Null }
    catch { $failedForEmptyFixture = $_.Exception.Message -eq 'Not enough live mailbox data was returned to run the search smoke suite.'; if (-not $failedForEmptyFixture) { Write-Output $_.Exception.Message } }
    Assert-True $failedForEmptyFixture 'Caller must stop after mocked empty fixture results.'
    Assert-True ($global:MailySearchSmokeCalls.Count -eq 3) 'Expected inbox, sent-items and all-mail bridge calls.'
    $expectedFolders = @('inbox', 'sent-items', 'all-mail')
    for ($index = 0; $index -lt 3; $index++) {
        $arguments = $global:MailySearchSmokeCalls[$index]
        $mailboxIndex = [Array]::IndexOf($arguments, '-MailboxAddress')
        Assert-True ($mailboxIndex -ge 0) 'Every search bridge call must include explicit mailbox scope.'
        Assert-True ($arguments[$mailboxIndex + 1] -eq 'synthetic@example.invalid') 'Bridge mailbox scope must match the selected account.'
        $folderIndex = [Array]::IndexOf($arguments, '-FolderName')
        Assert-True ($folderIndex -ge 0 -and $arguments[$folderIndex + 1] -eq $expectedFolders[$index]) 'Search folder selection must be retained.'
    }
    Write-Output 'PASS all three search calls forward the selected mailbox; all external access mocked'

    $global:MailySearchSmokeCalls.Clear()
    $global:MailySearchSmokeHostCalls = 0
    $global:MailySearchSmokeDirectoryCalls = 0
    $missingScopeRejected = $false
    try { & $ScriptPath -Live | Out-Null }
    catch { $missingScopeRejected = $_.Exception.Message -eq 'An explicit mailbox address is required for the search smoke suite.' }
    Assert-True $missingScopeRejected 'Live mode must require an explicit mailbox.'
    Assert-True ($global:MailySearchSmokeCalls.Count -eq 0 -and $global:MailySearchSmokeHostCalls -eq 0 -and $global:MailySearchSmokeDirectoryCalls -eq 0) 'Missing scope must fail before side effects.'
    Write-Output 'PASS missing live mailbox fails before host/process/directory access'

    $result = & $ScriptPath
    Assert-True ($result.Mode -eq 'DryRun') 'Default invocation must be dry-run.'
    Assert-True ($global:MailySearchSmokeCalls.Count -eq 0 -and $global:MailySearchSmokeHostCalls -eq 0 -and $global:MailySearchSmokeDirectoryCalls -eq 0) 'Dry-run must not access host/process/filesystem.'
    Write-Output 'PASS default dry-run has no host/process/filesystem access'
} $ScriptPath
