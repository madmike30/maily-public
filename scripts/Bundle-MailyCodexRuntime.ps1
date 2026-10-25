param(
    [Parameter(Mandatory = $true)][string] $Destination,
    [string] $Runtime = "win-x64"
)

$ErrorActionPreference = "Stop"
$lockPath = Join-Path $PSScriptRoot 'codex-runtime.lock.json'
$lock = Get-Content -LiteralPath $lockPath -Raw | ConvertFrom-Json
if ($Runtime -ne $lock.runtime) { throw "Bundled Codex is pinned for $($lock.runtime), not $Runtime." }

New-Item -ItemType Directory -Force -Path $Destination | Out-Null
foreach ($file in $lock.files) {
    if ([IO.Path]::GetFileName($file.name) -ne $file.name) { throw "Invalid locked filename." }
    $target = Join-Path $Destination $file.name
    $verified = (Test-Path -LiteralPath $target) -and ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -eq $file.sha256)
    if (-not $verified) {
        if (Test-Path -LiteralPath $target) {
            $prior = Join-Path (Split-Path $Destination -Parent) ($file.name + ".rejected-" + [Guid]::NewGuid().ToString("N"))
            Move-Item -LiteralPath $target -Destination $prior
        }
        Write-Host "Downloading pinned Codex $($lock.version): $($file.name)"
        Invoke-WebRequest -Uri $file.url -OutFile $target
    }
    $actual = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
    if ($actual -ne $file.sha256) {
        $quarantine = Join-Path (Split-Path $Destination -Parent) ($file.name + ".rejected-" + [Guid]::NewGuid().ToString("N"))
        Move-Item -LiteralPath $target -Destination $quarantine
        throw "SHA256 mismatch for $($file.name). Refusing to package an unverified runtime."
    }
}
Copy-Item -LiteralPath $lockPath -Destination (Join-Path $Destination 'codex-runtime.lock.json') -Force
# Never copy a user's CLI directory, CODEX_HOME, config, tokens or auth.json.
Write-Host "Pinned Codex $($lock.version) binary and upstream notices verified."
