$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Dist = Join-Path $Root "dist"
$Stage = Join-Path $Dist "stage"
$AppStage = Join-Path $Stage "app"
$PackageZip = Join-Path $Stage "notion-pdf-studio.zip"
$InstallerPs1 = Join-Path $Stage "install.ps1"
$SelfContainedInstaller = Join-Path $Dist "NotionPDFStudio-Setup.ps1"
$SetupExe = Join-Path $Dist "NotionPDFStudio-Setup.exe"
$SetupSed = Join-Path $Stage "setup.sed"

function Copy-CleanDirectory {
  param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Destination
  )
  if (Test-Path $Destination) {
    Remove-Item -LiteralPath $Destination -Recurse -Force
  }
  New-Item -ItemType Directory -Path $Destination -Force | Out-Null
  Copy-Item -Path (Join-Path $Source "*") -Destination $Destination -Recurse -Force
}

function Write-InstallerScript {
  param([string]$Path)
  @'
$ErrorActionPreference = "Stop"

$PackageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ZipPath = Join-Path $PackageRoot "notion-pdf-studio.zip"
$InstallRoot = Join-Path $env:LOCALAPPDATA "Programs\Notion PDF Studio"
$TempRoot = Join-Path $env:TEMP ("NotionPdfStudioInstall-" + [guid]::NewGuid().ToString("N"))
$ShortcutName = "Notion PDF Studio.lnk"

function Stop-AppServer {
  $connection = Get-NetTCPConnection -LocalPort 4173 -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $connection) { return }
  try {
    $process = Get-CimInstance Win32_Process -Filter "ProcessId = $($connection.OwningProcess)" -ErrorAction SilentlyContinue
    if ($process.CommandLine -match "Notion PDF Studio" -or $process.CommandLine -match "notion-pdf-studio" -or $process.CommandLine -match "server\.js") {
      Stop-Process -Id $connection.OwningProcess -Force -ErrorAction SilentlyContinue
    }
  } catch {}
}

function New-AppShortcut {
  param([string]$Path)
  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($Path)
  $shortcut.TargetPath = Join-Path $InstallRoot "Start-NotionPdfStudio.vbs"
  $shortcut.WorkingDirectory = $InstallRoot
  $shortcut.IconLocation = (Join-Path $InstallRoot "public\shortcut.ico") + ",0"
  $shortcut.Description = "Notion PDF Studio"
  $shortcut.Save()
}

function Write-Uninstaller {
  $path = Join-Path $InstallRoot "Uninstall-NotionPdfStudio.ps1"
  @"
`$ErrorActionPreference = "SilentlyContinue"
`$InstallRoot = "$InstallRoot"
Get-NetTCPConnection -LocalPort 4173 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique | ForEach-Object { Stop-Process -Id `$_ -Force -ErrorAction SilentlyContinue }
Remove-Item -LiteralPath (Join-Path ([Environment]::GetFolderPath("Desktop")) "$ShortcutName") -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path ([Environment]::GetFolderPath("Programs")) "$ShortcutName") -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Notion PDF Studio" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath `$InstallRoot -Recurse -Force -ErrorAction SilentlyContinue
"@ | Set-Content -LiteralPath $path -Encoding UTF8
}

if (-not (Test-Path $ZipPath)) {
  throw "No se encuentra el paquete de instalación: $ZipPath"
}

Stop-AppServer
New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null
New-Item -ItemType Directory -Path $InstallRoot -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $InstallRoot "data") -Force | Out-Null

Expand-Archive -LiteralPath $ZipPath -DestinationPath $TempRoot -Force
Copy-Item -Path (Join-Path $TempRoot "*") -Destination $InstallRoot -Recurse -Force
Remove-Item -LiteralPath $TempRoot -Recurse -Force -ErrorAction SilentlyContinue

Write-Uninstaller

$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) $ShortcutName
$startMenuShortcut = Join-Path ([Environment]::GetFolderPath("Programs")) $ShortcutName
New-AppShortcut -Path $desktopShortcut
New-AppShortcut -Path $startMenuShortcut

$uninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Notion PDF Studio"
New-Item -Path $uninstallKey -Force | Out-Null
Set-ItemProperty -Path $uninstallKey -Name DisplayName -Value "Notion PDF Studio"
Set-ItemProperty -Path $uninstallKey -Name DisplayIcon -Value ((Join-Path $InstallRoot "public\shortcut.ico") + ",0")
Set-ItemProperty -Path $uninstallKey -Name Publisher -Value "Notion PDF Studio"
Set-ItemProperty -Path $uninstallKey -Name InstallLocation -Value $InstallRoot
Set-ItemProperty -Path $uninstallKey -Name UninstallString -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$InstallRoot\Uninstall-NotionPdfStudio.ps1`""

Start-Process -FilePath (Join-Path $InstallRoot "Start-NotionPdfStudio.vbs")
'@ | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Write-SelfContainedInstaller {
  param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$ZipPath,
    [Parameter(Mandatory = $true)][string]$InstallScriptPath
  )
  $zipBytes = [Convert]::ToBase64String([IO.File]::ReadAllBytes($ZipPath))
  $chunks = for ($i = 0; $i -lt $zipBytes.Length; $i += 7600) {
    $length = [Math]::Min(7600, $zipBytes.Length - $i)
    '"' + $zipBytes.Substring($i, $length) + '"'
  }
  $installScript = Get-Content -LiteralPath $InstallScriptPath -Raw
  $installEncoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($installScript))
  @"
`$ErrorActionPreference = "Stop"
`$TempRoot = Join-Path `$env:TEMP ("NotionPdfStudioSetup-" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path `$TempRoot -Force | Out-Null
`$zipPath = Join-Path `$TempRoot "notion-pdf-studio.zip"
`$installPath = Join-Path `$TempRoot "install.ps1"
`$zipBase64 = @(
$($chunks -join "`r`n")
) -join ""
[IO.File]::WriteAllBytes(`$zipPath, [Convert]::FromBase64String(`$zipBase64))
`$installScript = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String("$installEncoded"))
Set-Content -LiteralPath `$installPath -Value `$installScript -Encoding UTF8
try {
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File `$installPath
} finally {
  Start-Sleep -Milliseconds 300
  Remove-Item -LiteralPath `$TempRoot -Recurse -Force -ErrorAction SilentlyContinue
}
"@ | Set-Content -LiteralPath $Path -Encoding UTF8
}

if (Test-Path $Dist) {
  Remove-Item -LiteralPath $Dist -Recurse -Force
}
New-Item -ItemType Directory -Path $AppStage -Force | Out-Null

$nodePath = (Get-Command node -ErrorAction Stop).Source
$runtimeDir = Join-Path $AppStage "runtime"
New-Item -ItemType Directory -Path $runtimeDir -Force | Out-Null
Copy-Item -LiteralPath $nodePath -Destination (Join-Path $runtimeDir "node.exe") -Force

Copy-Item -LiteralPath (Join-Path $Root "server.js") -Destination $AppStage -Force
Copy-Item -LiteralPath (Join-Path $Root "package.json") -Destination $AppStage -Force
Copy-Item -LiteralPath (Join-Path $Root "package-lock.json") -Destination $AppStage -Force
Copy-Item -LiteralPath (Join-Path $Root "Start-NotionPdfStudio.ps1") -Destination $AppStage -Force
Copy-Item -LiteralPath (Join-Path $Root "Start-NotionPdfStudio.vbs") -Destination $AppStage -Force
Copy-Item -LiteralPath (Join-Path $Root "README.md") -Destination $AppStage -Force
Copy-CleanDirectory -Source (Join-Path $Root "public") -Destination (Join-Path $AppStage "public")
Copy-CleanDirectory -Source (Join-Path $Root "node_modules") -Destination (Join-Path $AppStage "node_modules")
New-Item -ItemType Directory -Path (Join-Path $AppStage "data") -Force | Out-Null
Set-Content -LiteralPath (Join-Path $AppStage "data\.gitkeep") -Value "" -Encoding ASCII

Compress-Archive -Path (Join-Path $AppStage "*") -DestinationPath $PackageZip -CompressionLevel Optimal -Force
Write-InstallerScript -Path $InstallerPs1
Write-SelfContainedInstaller -Path $SelfContainedInstaller -ZipPath $PackageZip -InstallScriptPath $InstallerPs1

$iexpress = Get-Command iexpress.exe -ErrorAction SilentlyContinue
if ($iexpress) {
  @"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=0
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=
TargetName=$SetupExe
FriendlyName=Notion PDF Studio Setup
AppLaunched=powershell.exe -NoProfile -ExecutionPolicy Bypass -File install.ps1
PostInstallCmd=<None>
AdminQuietInstCmd=
UserQuietInstCmd=
SourceFiles=SourceFiles
[SourceFiles]
SourceFiles0=$Stage
[SourceFiles0]
install.ps1=
notion-pdf-studio.zip=
"@ | Set-Content -LiteralPath $SetupSed -Encoding ASCII
  & $iexpress.Source /N /Q $SetupSed | Out-Null
  $deadline = (Get-Date).AddSeconds(120)
  while ((Get-Date) -lt $deadline -and -not (Test-Path $SetupExe)) {
    Start-Sleep -Milliseconds 250
  }
}

Write-Host "Windows package ready:"
Write-Host "  Payload: $PackageZip"
Write-Host "  Installer script: $InstallerPs1"
Write-Host "  Self-contained installer: $SelfContainedInstaller"
if (Test-Path $SetupExe) {
  Write-Host "  Installer exe: $SetupExe"
} else {
  Write-Host "  Installer exe was not created. Run install.ps1 from the stage folder as fallback."
}
