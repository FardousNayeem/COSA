# COSA - Curated Open Source Apps
# Run: powershell -ExecutionPolicy Bypass -File .\cosa.ps1

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ----------------------------
# Paths + App Info
# ----------------------------
$CosaVersion = "0.3.4"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$CatalogPath = Join-Path $Root "catalog\apps.json"
$DataDir = Join-Path $Root "data"
$LogsDir = Join-Path $Root "logs"
$StatePath = Join-Path $DataDir "state.json"

New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath = Join-Path $LogsDir "COSA_$Timestamp.log"

function Write-Log {
  param(
    [Parameter(Mandatory=$true)][string]$Message,
    [ValidateSet("INFO","WARN","ERROR")][string]$Level = "INFO"
  )
  $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
  Add-Content -Path $LogPath -Value $line
  if ($Level -eq "ERROR") { Write-Host $line -ForegroundColor Red }
  elseif ($Level -eq "WARN") { Write-Host $line -ForegroundColor Yellow }
  else { Write-Host $line }
}

# ----------------------------
# Defensive helpers
# ----------------------------
function As-Array {
  param([Parameter(ValueFromPipeline=$true)]$Value)
  # IMPORTANT: the leading comma makes the array non-enumerating in the pipeline
  if ($null -eq $Value) { return ,@() }
  return ,@($Value)
}

function Safe-Count($Value) {
  $arr = As-Array $Value
  return $arr.Count
}

# ----------------------------
# JSON Helpers
# ----------------------------
function Read-JsonFile {
  param([Parameter(Mandatory=$true)][string]$Path)
  if (!(Test-Path $Path)) { return $null }
  $raw = Get-Content -Raw -Path $Path
  if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
  return $raw | ConvertFrom-Json
}

function Write-JsonFile {
  param(
    [Parameter(Mandatory=$true)][string]$Path,
    [Parameter(Mandatory=$true)]$Object
  )
  $Object | ConvertTo-Json -Depth 10 | Set-Content -Path $Path -Encoding UTF8
}

# ----------------------------
# Catalog + State (normalized)
# ----------------------------
function Load-Catalog {
  if (!(Test-Path $CatalogPath)) {
    throw "Catalog file not found: $CatalogPath"
  }
  $catalog = Read-JsonFile -Path $CatalogPath
  if ($null -eq $catalog) { throw "Catalog is empty or invalid JSON: $CatalogPath" }
  return (As-Array $catalog)
}

function Normalize-State($state) {
  if ($null -eq $state) { return $null }

  # Ensure managedApps is always an array
  if ($null -eq $state.managedApps) { $state.managedApps = @() }
  else { $state.managedApps = As-Array $state.managedApps }

  return $state
}

function Ensure-State {
  $state = Read-JsonFile -Path $StatePath
  if ($null -eq $state) {
    $state = [pscustomobject]@{
      cosaVersion  = $CosaVersion
      createdAt    = (Get-Date).ToString("o")
      lastRunAt    = $null
      managedApps  = @()   # array of objects: { wingetId, pinned, lastSeenVersion, lastStatus }
    }
    Write-JsonFile -Path $StatePath -Object $state
  }
  return (Normalize-State $state)
}

function Save-State($state) {
  $state.lastRunAt = (Get-Date).ToString("o")
  $state.cosaVersion = $CosaVersion
  $state = Normalize-State $state
  Write-JsonFile -Path $StatePath -Object $state
}

function Get-ManagedIndexById($state, [string]$wingetId) {
  $apps = As-Array $state.managedApps
  for ($i=0; $i -lt $apps.Count; $i++) {
    if ($apps[$i].wingetId -eq $wingetId) { return $i }
  }
  return -1
}

# ----------------------------
# Networking / Download Helpers
# ----------------------------
function Ensure-Tls12 {
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
}

function Download-File {
  param(
    [Parameter(Mandatory=$true)][string]$Url,
    [Parameter(Mandatory=$true)][string]$OutFile
  )
  Ensure-Tls12
  Write-Log "Downloading: $Url"
  Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
  if (!(Test-Path $OutFile)) { throw "Download failed: $Url" }
}

function Try-AddAppxPackage {
  param([Parameter(Mandatory=$true)][string]$Path)
  if (!(Test-Path $Path)) { throw "Package file missing: $Path" }

  try {
    Write-Log "Installing AppX/MSIX: $Path"
    Add-AppxPackage -Path $Path -ErrorAction Stop | Out-Null
    return $true
  } catch {
    Write-Log ("Add-AppxPackage failed for $Path : " + $_.Exception.Message) "WARN"
    return $false
  }
}

function Get-ArchTag {
  $arch = $env:PROCESSOR_ARCHITECTURE
  if ($arch -eq "ARM64") { return "arm64" }
  if ($arch -eq "AMD64") { return "x64" }
  return "x86"
}

# ----------------------------
# WinGet Bootstrap
# ----------------------------
function Test-Winget {
  $cmd = Get-Command winget -ErrorAction SilentlyContinue
  return ($null -ne $cmd)
}

function Install-UiXamlFramework {
  param(
    [Parameter(Mandatory=$true)][string]$TempDir,
    [Parameter(Mandatory=$true)][string]$ArchTag
  )

  $indexUrl = "https://api.nuget.org/v3-flatcontainer/microsoft.ui.xaml/index.json"
  Ensure-Tls12
  Write-Log "Fetching WinUI (Microsoft.UI.Xaml) versions from NuGet..."
  $idx = Invoke-RestMethod -Uri $indexUrl -UseBasicParsing

  if ($null -eq $idx -or $null -eq $idx.versions) {
    throw "Failed to retrieve Microsoft.UI.Xaml versions from NuGet."
  }

  $versions = As-Array ($idx.versions | Where-Object { $_ -match "^2\.8\." } | Sort-Object { [version]$_ } -Descending)
  if ($versions.Count -eq 0) { throw "No Microsoft.UI.Xaml 2.8.x versions found on NuGet." }

  $ver = $versions[0]
  Write-Log "Selected Microsoft.UI.Xaml version: $ver"

  $nupkgUrl = "https://api.nuget.org/v3-flatcontainer/microsoft.ui.xaml/$ver/microsoft.ui.xaml.$ver.nupkg"
  $nupkgPath = Join-Path $TempDir "microsoft.ui.xaml.$ver.nupkg"
  $zipPath = Join-Path $TempDir "microsoft.ui.xaml.$ver.zip"
  Download-File -Url $nupkgUrl -OutFile $nupkgPath

  Copy-Item -Path $nupkgPath -Destination $zipPath -Force
  $extractDir = Join-Path $TempDir "microsoft.ui.xaml.$ver"
  if (Test-Path $extractDir) { Remove-Item -Recurse -Force $extractDir }
  Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

  $appxDir = Join-Path $extractDir ("tools\AppX\{0}\Release" -f $ArchTag)
  if (!(Test-Path $appxDir)) { throw "Could not locate WinUI AppX directory: $appxDir" }

  $appx = Get-ChildItem -Path $appxDir -Filter "*.appx" | Select-Object -First 1
  if ($null -eq $appx) { throw "No .appx found in: $appxDir" }

  [void](Try-AddAppxPackage -Path $appx.FullName)
}

function Install-VCLibsFramework {
  param(
    [Parameter(Mandatory=$true)][string]$TempDir,
    [Parameter(Mandatory=$true)][string]$ArchTag
  )

  $url = $null
  if ($ArchTag -eq "x64") { $url = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx" }
  elseif ($ArchTag -eq "arm64") { $url = "https://aka.ms/Microsoft.VCLibs.arm64.14.00.Desktop.appx" }
  else { $url = "https://aka.ms/Microsoft.VCLibs.x86.14.00.Desktop.appx" }

  $out = Join-Path $TempDir ("Microsoft.VCLibs.{0}.14.00.Desktop.appx" -f $ArchTag)
  Download-File -Url $url -OutFile $out
  [void](Try-AddAppxPackage -Path $out)
}

function Install-WinGetAppInstaller {
  param([Parameter(Mandatory=$true)][string]$TempDir)

  $url = "https://aka.ms/getwinget"
  $out = Join-Path $TempDir "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
  Download-File -Url $url -OutFile $out
  [void](Try-AddAppxPackage -Path $out)
}

function Bootstrap-Winget {
  if (Test-Winget) {
    Write-Log "winget already present."
    return $true
  }

  Write-Log "winget not found. Attempting WinGet bootstrap (App Installer)..." "WARN"

  $arch = Get-ArchTag
  Write-Log "Detected architecture: $arch"

  $tempDir = Join-Path $env:TEMP ("COSA_bootstrap_{0}" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
  New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
  Write-Log "Bootstrap temp directory: $tempDir"

  try {
    Install-VCLibsFramework -TempDir $tempDir -ArchTag $arch
    Install-UiXamlFramework -TempDir $tempDir -ArchTag $arch
    Install-WinGetAppInstaller -TempDir $tempDir

    Start-Sleep -Seconds 2

    if (Test-Winget) {
      $ver = (& winget --version) 2>$null
      Write-Log "winget installed successfully. Version: $ver"
      return $true
    }

    Write-Log "Bootstrap completed but winget still not detected." "WARN"
    return $false
  }
  catch {
    Write-Log ("Bootstrap error: " + $_.Exception.Message) "ERROR"
    return $false
  }
  finally {
    Write-Log "Bootstrap temp files kept at: $tempDir"
  }
}

function Require-Winget {
  if (Test-Winget) { return }

  $ok = Bootstrap-Winget
  if ($ok -and (Test-Winget)) { return }

  Write-Log "winget is required but could not be installed automatically on this system." "ERROR"
  Write-Log "Try Windows PowerShell 5.1 as Administrator, or ensure AppX/Store installs are permitted." "ERROR"
  throw "winget is required but not installed."
}

# ----------------------------
# WinGet runner
# ----------------------------
function Invoke-Winget([string[]]$args) {
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = "winget"
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError  = $true
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true
  $psi.Arguments = ($args -join " ")

  $p = New-Object System.Diagnostics.Process
  $p.StartInfo = $psi
  [void]$p.Start()
  $stdout = $p.StandardOutput.ReadToEnd()
  $stderr = $p.StandardError.ReadToEnd()
  $p.WaitForExit()

  return [pscustomobject]@{
    ExitCode = $p.ExitCode
    StdOut   = $stdout
    StdErr   = $stderr
  }
}

# ----------------------------
# Selection Parsing: "1,3,7-10"
# ----------------------------
function Parse-Selection {
  param([Parameter(Mandatory=$true)][string]$InputText)

  $text = $InputText.Trim()
  if ([string]::IsNullOrWhiteSpace($text)) { return ,@() }

  $set = New-Object 'System.Collections.Generic.HashSet[int]'
  $parts = $text.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

  foreach ($part in $parts) {
    if ($part -match "^\d+$") {
      [void]$set.Add([int]$part)
      continue
    }
    if ($part -match "^(\d+)\s*-\s*(\d+)$") {
      $a = [int]$Matches[1]
      $b = [int]$Matches[2]
      if ($b -lt $a) { $tmp=$a; $a=$b; $b=$tmp }
      for ($i=$a; $i -le $b; $i++) { [void]$set.Add($i) }
      continue
    }
    throw "Invalid selection token: '$part'"
  }

  $arr = foreach ($v in $set) { [int]$v }
  return ,@($arr | Sort-Object)
}

# ----------------------------
# Display Helpers
# ----------------------------
function Confirm($prompt) {
  while ($true) {
    $ans = Read-Host "$prompt (Y/N, 0=Back)"
    $a = $ans.Trim().ToUpper()
    if ($a -eq "Y") { return $true }
    if ($a -eq "N") { return $false }
    if ($a -eq "0") { return $null }
    Write-Host "Please enter Y, N, or 0."
  }
}

function Show-Apps($apps) {
  $apps = As-Array $apps
  Write-Host ""
  Write-Host ("{0,-4} {1,-30} {2,-14} {3}" -f "ID","Name","Category","WingetId")
  Write-Host ("{0,-4} {1,-30} {2,-14} {3}" -f "--","----","--------","-------")
  foreach ($a in $apps) {
    Write-Host ("{0,-4} {1,-30} {2,-14} {3}" -f $a.id, $a.name, $a.category, $a.wingetId)
  }
  Write-Host ""
}

function Name-ForWingetId($catalog, [string]$wingetId) {
  $catalog = As-Array $catalog
  $match = $catalog | Where-Object { $_.wingetId -eq $wingetId } | Select-Object -First 1
  if ($null -ne $match) { return $match.name }
  return $wingetId
}

# ----------------------------
# State / Managed
# ----------------------------
function Ensure-Managed($state, [string]$wingetId) {
  $state = Normalize-State $state
  $idx = Get-ManagedIndexById $state $wingetId
  if ($idx -ge 0) { return }

  $state.managedApps = As-Array $state.managedApps
  $state.managedApps += [pscustomobject]@{
    wingetId        = $wingetId
    pinned          = $false
    lastSeenVersion = $null
    lastStatus      = "managed"
  }
}

# ----------------------------
# Install Actions
# ----------------------------
function Install-WingetId($state, [string]$wingetId, [string]$displayName) {
  Require-Winget

  Write-Log "Installing: $displayName ($wingetId)"
  $res = Invoke-Winget @(
    "install",
    "--id", "`"$wingetId`"",
    "-e",
    "--silent",
    "--accept-package-agreements",
    "--accept-source-agreements"
  )

  if ($res.ExitCode -eq 0) {
    Write-Log "SUCCESS: $wingetId"
    Ensure-Managed $state $wingetId
    $idx = Get-ManagedIndexById $state $wingetId
    if ($idx -ge 0) { (As-Array $state.managedApps)[$idx].lastStatus = "installed" }
    return $true
  }

  Write-Log "FAIL: $wingetId exit=$($res.ExitCode)" "ERROR"
  if ($res.StdErr) { Write-Log ("winget stderr: " + $res.StdErr.Trim()) "ERROR" }
  if ($res.StdOut) { Write-Log ("winget stdout: " + $res.StdOut.Trim()) "WARN" }
  return $false
}

function Install-App($state, $app) {
  return (Install-WingetId $state $app.wingetId $app.name)
}

# ----------------------------
# Bundles
# ----------------------------
$BundleBasics = @(
  "PeaZip.PeaZip",
  "VideoLAN.VLC",
  "Notepad++.Notepad++",
  "DuongDieuPhap.ImageGlass",
  "GIMP.GIMP",
  "TheDocumentFoundation.LibreOffice",
  "KDE.Okular",
  "qBittorrent.qBittorrent"
)

$BundleDev = @(
  "VSCodium.VSCodium",
  "GitHub.GitHubDesktop",
  "Python.Python.3",
  "EclipseAdoptium.Temurin.17.JDK",
  "OpenJS.NodeJS",
  "Git.Git",
  "MSYS2.MSYS2"
)

function Install-Bundle($catalog, $state, [string]$bundleName, [string[]]$wingetIds) {
  Write-Host ""
  Write-Host "Bundle: $bundleName"
  Write-Host "Apps to install:"
  foreach ($id in $wingetIds) {
    Write-Host (" - " + (Name-ForWingetId $catalog $id) + " [$id]")
  }
  Write-Host ""

  $confirm = Confirm "Install bundle '$bundleName' ($($wingetIds.Count) apps)?"
  if ($null -eq $confirm) { Write-Log "Back to menu."; return }
  if ($confirm -eq $false) { Write-Log "Cancelled by user."; return }

  foreach ($id in $wingetIds) {
    $name = Name-ForWingetId $catalog $id
    [void](Install-WingetId $state $id $name)
  }
}

# ----------------------------
# Browse All Apps Install
# ----------------------------
function Install-From-AllApps($catalog, $state) {
  $catalog = As-Array $catalog
  Show-Apps $catalog

  $selText = Read-Host "Enter app IDs to install (e.g. 1,3,7-10) or 0 to go back"
  if ($selText.Trim() -eq "0") { Write-Log "Back to menu."; return }

  $ids = As-Array (Parse-Selection $selText)
  if ($ids.Count -eq 0) { Write-Log "No selection."; return }

  $selected = @()
  foreach ($id in $ids) {
    $match = $catalog | Where-Object { $_.id -eq $id } | Select-Object -First 1
    if ($null -eq $match) { Write-Log "Unknown ID: $id" "WARN"; continue }
    $selected += $match
  }
  $selected = As-Array $selected

  if ($selected.Count -eq 0) { Write-Log "Nothing valid selected."; return }

  Write-Host ""
  Write-Host "Selected:"
  foreach ($a in $selected) { Write-Host " - $($a.id): $($a.name) [$($a.wingetId)]" }
  Write-Host ""

  $confirm = Confirm "Install $($selected.Count) app(s)?"
  if ($null -eq $confirm) { Write-Log "Back to menu."; return }
  if ($confirm -eq $false) { Write-Log "Cancelled by user."; return }

  foreach ($app in $selected) { [void](Install-App $state $app) }
}

# ----------------------------
# Managed-only Updates
# ----------------------------
function Get-ManagedUpdateCandidates($state) {
  Require-Winget
  $state = Normalize-State $state

  $managed = As-Array $state.managedApps
  if ($managed.Count -eq 0) { return ,@() }

  $needUpdate = @()

  foreach ($m in $managed) {
    if ($m.pinned -eq $true) { continue }

    $res = Invoke-Winget @("upgrade", "--id", "`"$($m.wingetId)`"", "-e")
    $out = ($res.StdOut + "`n" + $res.StdErr)

    if ($out -match "No applicable update found" -or $out -match "No installed package found") {
      continue
    }

    if (-not [string]::IsNullOrWhiteSpace($out)) {
      $needUpdate += $m.wingetId
    }
  }

  return (As-Array $needUpdate)
}

function Update-ManagedFlow($catalog, $state) {
  Require-Winget
  $state = Normalize-State $state
  $catalog = As-Array $catalog

  $managed = As-Array $state.managedApps
  if ($managed.Count -eq 0) {
    Write-Log "No managed apps yet."
    return
  }

  Write-Log "Checking updates for managed apps only..."
  $candidates = As-Array (Get-ManagedUpdateCandidates $state)

  if ($candidates.Count -eq 0) {
    Write-Log "No updates available for managed apps."
    return
  }

  Write-Host ""
  Write-Host "Updates available for:"
  foreach ($id in $candidates) {
    Write-Host (" - " + (Name-ForWingetId $catalog $id) + " [$id]")
  }
  Write-Host ""

  $confirm = Confirm "Update these $($candidates.Count) app(s) now?"
  if ($null -eq $confirm) { Write-Log "Back to menu."; return }
  if ($confirm -eq $false) { Write-Log "User chose not to update."; return }

  Write-Log "Updating managed apps (fetching + installing)..."
  foreach ($id in $candidates) {
    Write-Log "Upgrading: $id"
    $res = Invoke-Winget @(
      "upgrade",
      "--id", "`"$id`"",
      "-e",
      "--silent",
      "--accept-package-agreements",
      "--accept-source-agreements"
    )

    $idx = Get-ManagedIndexById $state $id

    if ($res.ExitCode -eq 0) {
      Write-Log "OK: $id"
      if ($idx -ge 0) { (As-Array $state.managedApps)[$idx].lastStatus = "upgraded_or_ok" }
    } else {
      Write-Log "FAIL upgrade: $id exit=$($res.ExitCode)" "ERROR"
      if ($res.StdErr) { Write-Log ("winget stderr: " + $res.StdErr.Trim()) "ERROR" }
      if ($idx -ge 0) { (As-Array $state.managedApps)[$idx].lastStatus = "upgrade_failed" }
    }
  }
}

function Show-Managed($state) {
  $state = Normalize-State $state
  $managed = As-Array $state.managedApps

  if ($managed.Count -eq 0) {
    Write-Log "No managed apps yet."
    return
  }

  Write-Host ""
  Write-Host ("{0,-38} {1,-8} {2}" -f "WingetId","Pinned","LastStatus")
  Write-Host ("{0,-38} {1,-8} {2}" -f "-------","------","----------")
  foreach ($m in $managed) {
    Write-Host ("{0,-38} {1,-8} {2}" -f $m.wingetId, $m.pinned, $m.lastStatus)
  }
  Write-Host ""
}

# ----------------------------
# Homepage
# ----------------------------
try {
  Write-Log "COSA v$CosaVersion starting..."
  Write-Log "Log file: $LogPath"

  $catalog = Load-Catalog
  $state = Ensure-State

  while ($true) {
    Write-Host ""
    Write-Host "=== COSA (Curated Open Source Apps) v$CosaVersion ==="
    Write-Host "0) Refresh / Home"
    Write-Host "1) Basics bundle (Recommended)"
    Write-Host "2) Development bundle"
    Write-Host "3) Browse Apps"
    Write-Host "4) Update Managed Apps"
    Write-Host "5) Show Managed Apps"
    Write-Host "6) Exit"
    Write-Host ""

    $choice = Read-Host "Select an option"

    switch ($choice.Trim()) {
      "0" { continue }
      "1" { Install-Bundle $catalog $state "Recommended Apps" $BundleBasics; Save-State $state }
      "2" { Install-Bundle $catalog $state "Development Apps" $BundleDev; Save-State $state }
      "3" { Install-From-AllApps $catalog $state; Save-State $state }
      "4" { Update-ManagedFlow $catalog $state; Save-State $state }
      "5" { Show-Managed $state }
      "6" { break }
      default { Write-Log "Invalid option: $choice" "WARN" }
    }
  }

  Save-State $state
  Write-Log "COSA exiting. Bye!"
}
catch {
  Write-Log ("Fatal error: " + $_.Exception.Message) "ERROR"

  if ($_.InvocationInfo) {
    Write-Log ("At: " + $_.InvocationInfo.ScriptName + ":" + $_.InvocationInfo.ScriptLineNumber) "ERROR"
    Write-Log ("Line: " + ($_.InvocationInfo.Line.Trim())) "ERROR"
  }

  if ($_.ScriptStackTrace) {
    Write-Log ("Stack: " + $_.ScriptStackTrace) "ERROR"
  }

  Write-Log "See log: $LogPath" "ERROR"
  exit 1
}
