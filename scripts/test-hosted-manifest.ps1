$ErrorActionPreference = "Stop"

Set-Location (Join-Path $PSScriptRoot "..")

$outputDir = Join-Path (Get-Location) ("output/hosted-manifest/" + [Guid]::NewGuid().ToString("N"))
$manifestPath = Join-Path $outputDir "manifest.xml"
$baseUrl = "https://maily.example.com/outlook"

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$result = & (Join-Path (Get-Location) "scripts\New-MailyHostedManifest.ps1") `
  -BaseUrl $baseUrl `
  -SourcePath (Join-Path (Get-Location) "addin\manifest.xml") `
  -DestinationPath $manifestPath

$manifestText = Get-Content -LiteralPath $manifestPath -Raw
if ($manifestText -match "https?://localhost(?::\d+)?") {
  throw "Hosted manifest still contains localhost URLs."
}

if ($manifestText -notmatch [regex]::Escape("$baseUrl/addin/taskpane.html")) {
  throw "Hosted manifest does not contain the expected hosted taskpane URL."
}

$xml = [xml]$manifestText
$appDomains = @($xml.GetElementsByTagName("AppDomain") | ForEach-Object { $_.InnerText.Trim() })
if (($appDomains | Where-Object { $_ -eq $baseUrl }).Count -ne 1) {
  throw "Hosted manifest should contain exactly one AppDomain for $baseUrl."
}

[pscustomobject]@{
  status = "passed"
  manifestPath = $manifestPath
  baseUrl = $baseUrl
  replacedLocalhostUrlCount = $result.replacedLocalhostUrlCount
  appDomains = $appDomains
} | ConvertTo-Json -Depth 5
