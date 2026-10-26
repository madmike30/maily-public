# Distribution-only sharing. Installed executables retain their original self-contained layouts.
function Assert-MailyRuntimePlainPath {
  param([string]$Path)
  $cursor = [IO.Path]::GetFullPath($Path)
  while ($cursor) {
    if (Test-Path -LiteralPath $cursor) {
      if ((Get-Item -LiteralPath $cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Runtime paths must not contain reparse points.' }
    }
    $parent = [IO.Path]::GetDirectoryName($cursor)
    if ($parent -eq $cursor) { break }
    $cursor = $parent
  }
}

function Get-MailyRuntimeMapEntries {
  param([string]$AppRoot)
  $root = [IO.Path]::GetFullPath($AppRoot)
  Assert-MailyRuntimePlainPath $root
  $mapPath = Join-Path $root 'runtime-file-map.json'
  Assert-MailyRuntimePlainPath $mapPath
  if (-not (Test-Path -LiteralPath $mapPath -PathType Leaf)) { return }
  $map = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($mapPath))
  if ($map.version -ne 1 -or $map.files -isnot [Array]) { throw 'Invalid runtime map.' }
  $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($entry in $map.files) {
    if ($entry.path -isnot [string] -or $entry.path -notmatch '^[A-Za-z0-9_.-]+(?:/[A-Za-z0-9_.-]+)*\.dll$' -or
        @($entry.path.Split('/') | Where-Object { $_ -eq '.' -or $_ -eq '..' }).Count -or
        $entry.sha256 -isnot [string] -or $entry.sha256 -notmatch '^[a-fA-F0-9]{64}$' -or
        ($entry.bytes -isnot [long] -and $entry.bytes -isnot [int]) -or $entry.bytes -lt 1 -or -not $seen.Add($entry.path)) { throw 'Invalid or duplicate runtime map entry.' }
    $source = Join-Path (Join-Path $root 'Maily.OutlookAddin.Web') $entry.path
    $target = Join-Path (Join-Path $root 'Maily.OutlookAddin.Supervisor') $entry.path
    Assert-MailyRuntimePlainPath $source
    Assert-MailyRuntimePlainPath $target
    if (-not (Test-Path -LiteralPath $source -PathType Leaf) -or (Test-Path -LiteralPath $target)) { throw 'Runtime source missing or destination collision.' }
    if ((Get-Item -LiteralPath $source).Length -ne $entry.bytes -or (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $entry.sha256) { throw 'Shared runtime integrity check failed.' }
    [pscustomobject]@{ source=$source; target=$target; bytes=$entry.bytes; sha256=$entry.sha256 }
  }
}

function Test-MailyPackageRuntime {
  param([string]$AppRoot)
  # Materialize and validate every entry before any reconstruction writes.
  $entries = @(Get-MailyRuntimeMapEntries -AppRoot $AppRoot)
  return @{ files=$entries.Count }
}

function Restore-MailyPackageRuntime {
  param([string]$AppRoot)
  $entries = @(Get-MailyRuntimeMapEntries -AppRoot $AppRoot)
  foreach ($entry in $entries) {
    Assert-MailyRuntimePlainPath $entry.source
    Assert-MailyRuntimePlainPath $entry.target
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($entry.target)) | Out-Null
    [IO.File]::Copy($entry.source, $entry.target, $false)
    if ((Get-Item -LiteralPath $entry.target).Length -ne $entry.bytes -or (Get-FileHash -LiteralPath $entry.target -Algorithm SHA256).Hash -ne $entry.sha256) { throw 'Restored runtime integrity check failed.' }
  }
}

function Compress-MailyPackageRuntime {
  param([string]$AppRoot, [string]$EvidenceRoot)
  $root = [IO.Path]::GetFullPath($AppRoot)
  $evidence = [IO.Path]::GetFullPath($EvidenceRoot)
  Assert-MailyRuntimePlainPath $root
  Assert-MailyRuntimePlainPath $evidence
  if ($evidence.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -or $evidence -eq $root -or (Test-Path -LiteralPath $evidence)) { throw 'Choose a new evidence path outside the app payload.' }
  $mapPath = Join-Path $root 'runtime-file-map.json'
  if (Test-Path -LiteralPath $mapPath) { throw 'Runtime payload is already compacted.' }
  $supervisor = Join-Path $root 'Maily.OutlookAddin.Supervisor'
  $entries = @()
  foreach ($file in Get-ChildItem -LiteralPath $supervisor -Filter '*.dll' -File -Recurse) {
    Assert-MailyRuntimePlainPath $file.FullName
    $relative = $file.FullName.Substring($supervisor.Length + 1).Replace('\','/')
    $source = Join-Path (Join-Path $root 'Maily.OutlookAddin.Web') $relative
    Assert-MailyRuntimePlainPath $source
    if ((Test-Path -LiteralPath $source -PathType Leaf) -and (Get-Item -LiteralPath $source).Length -eq $file.Length) {
      $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
      if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -eq $hash) {
        $entries += [pscustomobject]@{ path=$relative; bytes=$file.Length; sha256=$hash }
      }
    }
  }
  [IO.Directory]::CreateDirectory($evidence) | Out-Null
  foreach ($entry in $entries) {
    $original = Join-Path $supervisor $entry.path
    $retained = Join-Path $evidence $entry.path
    Assert-MailyRuntimePlainPath $original
    Assert-MailyRuntimePlainPath $retained
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($retained)) | Out-Null
    [IO.File]::Move($original, $retained)
  }
  [IO.File]::WriteAllText($mapPath, (@{version=1;files=@($entries)} | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))
  Test-MailyPackageRuntime -AppRoot $root | Out-Null
  return @{files=$entries.Count;bytes=($entries | Measure-Object bytes -Sum).Sum}
}
