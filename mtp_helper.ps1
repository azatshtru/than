param(
  [Parameter(Mandatory=$true)][string]$Action,
  [string]$Device = "",
  [string]$Storage = "",
  [string]$SubPath = ".",
  [string]$Dest = ""
)

$ErrorActionPreference = 'Stop'

function Get-ComputerNamespace {
  $shell = New-Object -ComObject Shell.Application
  $pc = $shell.NameSpace('shell:MyComputerFolder')
  if (-not $pc) { $pc = $shell.NameSpace(17) }
  return $pc
}

function Get-DeviceItem($deviceName) {
  $pc = Get-ComputerNamespace
  foreach ($i in @($pc.Items())) {
    if ($i.Name -eq $deviceName) { return $i }
  }
  return $null
}

function Get-StorageItem($deviceName, $storageName) {
  $dev = Get-DeviceItem $deviceName
  if (-not $dev) { throw "Device not found: $deviceName" }
  $df = $dev.GetFolder
  if (-not $df) { throw "Cannot access device: $deviceName" }
  foreach ($s in @($df.Items())) {
    if ($s.Name -eq $storageName) { return $s }
  }
  throw "Storage not found: $storageName"
}

function Normalize-SubPath($subPath) {
  if ($null -eq $subPath) { return "" }
  $p = $subPath.Trim()
  if ($p -eq "." -or $p -eq "/" -or $p -eq "\") { return "" }
  return $p
}

function Get-BestItemName($item) {
  $name = $null
  try { $name = [string]$item.Name } catch { $name = "" }
  $isFolder = $false
  try { $isFolder = [bool]$item.IsFolder } catch { $isFolder = $false }

  # For files, prefer the actual file name (System.FileName) over the display name.
  # MTP/Shell display names can hide extensions (e.g. "IMG-..." vs "IMG-....jpg")
  # and dotfiles may appear with an empty display name (e.g. ".nomedia").
  if (-not $isFolder) {
    try {
      $fn = $item.ExtendedProperty('System.FileName')
      if ($fn -is [string] -and -not [string]::IsNullOrWhiteSpace($fn)) { return $fn }
    } catch {}
  }

  if (-not [string]::IsNullOrWhiteSpace($name)) { return $name }

  try {
    $fn = $item.ExtendedProperty('System.FileName')
    if ($fn -is [string] -and -not [string]::IsNullOrWhiteSpace($fn)) { return $fn }
  } catch {}
  try {
    $dn = $item.ExtendedProperty('System.ItemNameDisplay')
    if ($dn -is [string] -and -not [string]::IsNullOrWhiteSpace($dn)) { return $dn }
  } catch {}
  return $name
}

function Get-ItemSize($item) {
  $size = 0
  try { $size = [int64]$item.Size } catch { $size = 0 }
  if ($size -le 0) {
    try {
      $s2 = $item.ExtendedProperty('System.Size')
      if ($null -ne $s2 -and "$s2" -ne "") { $size = [int64]$s2 }
    } catch {}
  }
  return $size
}

function Get-ItemBySubPath($deviceName, $storageName, $subPath) {
  $subPath = Normalize-SubPath $subPath
  $store = Get-StorageItem $deviceName $storageName
  $curFolder = $store.GetFolder
  if ([string]::IsNullOrWhiteSpace($subPath)) { return $store }
  $parts = @($subPath -split '[\\/]+' | Where-Object { $_ -ne '' })
  for ($i = 0; $i -lt $parts.Count; $i++) {
    $part = $parts[$i]
    $found = $null
    foreach ($it in @($curFolder.Items())) {
      $itName = Get-BestItemName $it
      $itDisplay = $null
      try { $itDisplay = [string]$it.Name } catch { $itDisplay = $null }
      if ($itName -eq $part -or $itDisplay -eq $part) { $found = $it; break }
    }
    if (-not $found) { throw "Path not found: $subPath (missing $part)" }
    if ($i -eq $parts.Count - 1) { return $found }
    if (-not $found.IsFolder) { throw "Not a folder: $part" }
    $curFolder = $found.GetFolder
    if (-not $curFolder) { throw "Cannot open folder: $part" }
  }
}

switch ($Action) {
  'list-devices' {
    $pc = Get-ComputerNamespace
    foreach ($i in @($pc.Items())) {
      try {
        $f = $i.GetFolder
        if ($f) {
          # Heuristic: portable devices usually have storages; include if has items or is non-filesystem path
          $has = $false
          foreach ($s in @($f.Items())) { $has = $true; break }
          $isMtp = ($i.Path -like '*usb#*') -or ($i.Path -like '::{*')
          if ($has -or $isMtp) { Write-Output $i.Name }
        }
      } catch {}
    }
  }
  'list-storages' {
    $dev = Get-DeviceItem $Device
    if (-not $dev) { throw "Device not found: $Device" }
    foreach ($s in @($dev.GetFolder.Items())) {
      Write-Output $s.Name
    }
  }
  'list-children' {
    $SubPath = Normalize-SubPath $SubPath
    if ([string]::IsNullOrWhiteSpace($SubPath)) {
      $store = Get-StorageItem $Device $Storage
      $folder = $store.GetFolder
    } else {
      $item = Get-ItemBySubPath $Device $Storage $SubPath
      if (-not $item.IsFolder) { throw "Not a folder: $SubPath" }
      $folder = $item.GetFolder
    }
    foreach ($it in @($folder.Items())) {
      $type = if ($it.IsFolder) { 'D' } else { 'F' }
      $name = Get-BestItemName $it
      $size = 0
      try { $size = [int64](Get-ItemSize $it) } catch { $size = 0 }
      Write-Output ($type + "`t" + $name + "`t" + $size)
    }
  }
  'index' {
    $SubPath = Normalize-SubPath $SubPath
    if ([string]::IsNullOrWhiteSpace($SubPath)) {
      $store = Get-StorageItem $Device $Storage
      $rootFolder = $store.GetFolder
    } else {
      $rootItem = Get-ItemBySubPath $Device $Storage $SubPath
      if (-not $rootItem.IsFolder) { throw "Not a folder: $SubPath" }
      $rootFolder = $rootItem.GetFolder
    }

    function Walk($folder, $prefix) {
      foreach ($it in @($folder.Items())) {
        $name = Get-BestItemName $it
        $rel = if ($prefix -eq '') { $name } else { $prefix + '\' + $name }
        if ($it.IsFolder) {
          Write-Output ("D`t" + $rel)
          try {
            $sub = $it.GetFolder
            if ($sub) { Walk $sub $rel }
          } catch {}
        } else {
          $size = 0
          try { $size = [int64](Get-ItemSize $it) } catch { $size = 0 }
          Write-Output ("F`t" + $rel + "`t" + $size)
        }
      }
    }
    Walk $rootFolder ''
  }
  'copy-file' {
    if ([string]::IsNullOrWhiteSpace($Dest)) { throw "Dest required" }
    $SubPath = Normalize-SubPath $SubPath
    $item = Get-ItemBySubPath $Device $Storage $SubPath
    if ($item.IsFolder) { throw "Is a folder: $SubPath" }
    $destPath = $Dest
    $destDir = Split-Path -Parent $destPath
    $destName = Split-Path -Leaf $destPath
    if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Force -Path $destDir | Out-Null }
    $shell = New-Object -ComObject Shell.Application
    $ns = $shell.NameSpace($destDir)
    if (-not $ns) { throw "Cannot access dest dir: $destDir" }
    # 4 = no progress dialog, 16 = yes to all
    $ns.CopyHere($item, 20)
    # CopyHere is async; wait for file to appear with expected name
    $target = Join-Path $destDir $destName
    $srcName = Get-BestItemName $item
    $srcTarget = Join-Path $destDir $srcName
    for ($i = 0; $i -lt 600; $i++) {
      if (Test-Path -LiteralPath $target) { Write-Output $target; exit 0 }
      if ($srcName -and $srcName -ne $destName -and (Test-Path -LiteralPath $srcTarget)) {
        Move-Item -LiteralPath $srcTarget -Destination $target -Force
        Write-Output $target; exit 0
      }
      Start-Sleep -Milliseconds 100
    }
    if ($srcName -and (Test-Path -LiteralPath $srcTarget)) { Write-Output $srcTarget; exit 0 }
    throw "Copy timed out for $SubPath"
  }
  default {
    throw "Unknown action: $Action"
  }
}
