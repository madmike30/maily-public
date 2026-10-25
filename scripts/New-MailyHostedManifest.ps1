param(
  [Parameter(Mandatory = $true)]
  [string]$BaseUrl,
  [string]$SourcePath = "",
  [string]$DestinationPath = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($SourcePath)) {
  $SourcePath = Join-Path $repoRoot "addin\manifest.xml"
}
if ([string]::IsNullOrWhiteSpace($DestinationPath)) {
  $DestinationPath = Join-Path $repoRoot "dist\manifest.hosted.xml"
}

function Resolve-HttpsBaseUrl {
  param([Parameter(Mandatory = $true)][string]$Value)

  $trimmed = $Value.Trim().TrimEnd("/")
  if ([string]::IsNullOrWhiteSpace($trimmed)) {
    throw "Hosted base URL is required."
  }

  [System.Uri]$uri = $null
  if (-not [System.Uri]::TryCreate($trimmed, [System.UriKind]::Absolute, [ref]$uri)) {
    throw "Hosted base URL must be an absolute HTTPS URL."
  }

  if ($uri.Scheme -ne [System.Uri]::UriSchemeHttps) {
    throw "Hosted base URL must use HTTPS."
  }

  if (-not [string]::IsNullOrWhiteSpace($uri.Query) -or -not [string]::IsNullOrWhiteSpace($uri.Fragment)) {
    throw "Hosted base URL must not include a query string or fragment."
  }

  if (-not [string]::IsNullOrWhiteSpace($uri.UserInfo)) {
    throw "Hosted base URL must not include credentials."
  }

  $normalizedHost = $uri.Host.Trim().ToLowerInvariant()
  if ($normalizedHost -eq "localhost" -or $normalizedHost -eq "127.0.0.1" -or $normalizedHost -eq "::1") {
    throw "Hosted base URL must not point at localhost or a loopback address."
  }

  return $trimmed
}

function Remove-DuplicateAppDomains {
  param([Parameter(Mandatory = $true)][xml]$Manifest)

  $seen = New-Object "System.Collections.Generic.HashSet[string]" -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
  $nodes = @($Manifest.GetElementsByTagName("AppDomain"))
  foreach ($node in $nodes) {
    $value = ""
    if ($null -ne $node.InnerText) {
      $value = $node.InnerText.Trim()
    }
    if ([string]::IsNullOrWhiteSpace($value)) {
      continue
    }

    if (-not $seen.Add($value)) {
      [void]$node.ParentNode.RemoveChild($node)
    }
  }
}

$resolvedBaseUrl = Resolve-HttpsBaseUrl -Value $BaseUrl
$sourceFullPath = [System.IO.Path]::GetFullPath($SourcePath)
$destinationFullPath = [System.IO.Path]::GetFullPath($DestinationPath)

if (-not (Test-Path -LiteralPath $sourceFullPath)) {
  throw "Manifest source not found at $sourceFullPath."
}

$source = Get-Content -LiteralPath $sourceFullPath -Raw
$localhostUrlCount = ([regex]::Matches($source, "https?://localhost(?::\d+)?")).Count
$generated = $source.Replace("https://localhost:7097", $resolvedBaseUrl).Replace("http://localhost:5289", $resolvedBaseUrl)

if ($generated -match "https?://localhost(?::\d+)?") {
  throw "Hosted manifest generation left one or more localhost URLs in place."
}

$xml = New-Object xml
$xml.PreserveWhitespace = $true
$xml.LoadXml($generated)
Remove-DuplicateAppDomains -Manifest $xml

$destinationDirectory = Split-Path -Parent $destinationFullPath
New-Item -ItemType Directory -Force -Path $destinationDirectory | Out-Null

$settings = New-Object System.Xml.XmlWriterSettings
$settings.Encoding = New-Object System.Text.UTF8Encoding($false)
$settings.Indent = $true
$settings.OmitXmlDeclaration = $false
$writer = [System.Xml.XmlWriter]::Create($destinationFullPath, $settings)
try {
  $xml.Save($writer)
}
finally {
  $writer.Close()
}

$validationText = Get-Content -LiteralPath $destinationFullPath -Raw
if ($validationText -match "https?://localhost(?::\d+)?") {
  throw "Hosted manifest validation failed because localhost URLs are still present."
}

[pscustomobject]@{
  sourcePath = $sourceFullPath
  destinationPath = $destinationFullPath
  baseUrl = $resolvedBaseUrl
  replacedLocalhostUrlCount = $localhostUrlCount
}
