$ErrorActionPreference = "Stop"

$AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Port = if ($env:PORT) { [int]$env:PORT } else { 4173 }
$Url = "http://localhost:$Port"
$ServerScript = Join-Path $AppRoot "server.js"
$OutLogPath = Join-Path $AppRoot "notion-pdf-studio.out.log"
$ErrLogPath = Join-Path $AppRoot "notion-pdf-studio.err.log"
$BrowserProfile = Join-Path $env:LOCALAPPDATA "NotionPdfStudio\BrowserProfile"
$BrowserIconStamp = Join-Path $env:LOCALAPPDATA "NotionPdfStudio\icon-version.txt"
$IconVersion = "20260512-sharp-icon"
$BundledNode = Join-Path $AppRoot "runtime\node.exe"
$NodeExe = if (Test-Path $BundledNode) { $BundledNode } else { (Get-Command node -ErrorAction SilentlyContinue).Source }

function Show-AppError {
  param([string]$Message)
  Add-Type -AssemblyName PresentationFramework
  [System.Windows.MessageBox]::Show($Message, "Notion PDF Studio", "OK", "Error") | Out-Null
}

function Get-PortProcess {
  param([int]$LocalPort)
  $connection = Get-NetTCPConnection -LocalPort $LocalPort -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $connection) {
    return $null
  }

  return Get-Process -Id $connection.OwningProcess -ErrorAction SilentlyContinue
}

function Test-ManagedServer {
  param([System.Diagnostics.Process]$Process)
  if (-not $Process) {
    return $false
  }

  $escapedRoot = [regex]::Escape($AppRoot)
  $serverProcess = Get-CimInstance Win32_Process -Filter "ProcessId = $($Process.Id)" -ErrorAction SilentlyContinue
  return $serverProcess.CommandLine -match "server\.js" -and $serverProcess.CommandLine -match $escapedRoot
}

function Test-AppServerOnPort {
  param([int]$LocalPort)
  try {
    $health = Invoke-RestMethod -Uri "http://localhost:$LocalPort/api/health" -TimeoutSec 2
    return $health.name -eq "notion-pdf-studio"
  } catch {
    return $false
  }
}

function Wait-PortClosed {
  param([int]$LocalPort)
  $deadline = (Get-Date).AddSeconds(5)
  while ((Get-Date) -lt $deadline -and (Get-PortProcess -LocalPort $LocalPort)) {
    Start-Sleep -Milliseconds 150
  }
}

function Find-Browser {
  $commands = @("chrome", "msedge")
  foreach ($command in $commands) {
    $found = Get-Command $command -ErrorAction SilentlyContinue
    if ($found) {
      return $found.Source
    }
  }

  $paths = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
  )

  foreach ($path in $paths) {
    if ($path -and (Test-Path $path)) {
      return $path
    }
  }

  return $null
}

function Test-BrowserOpen {
  $escapedProfile = [regex]::Escape($BrowserProfile)
  $processes = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -match $escapedProfile }

  return $null -ne ($processes | Select-Object -First 1)
}

function Clear-BrowserIconCacheIfNeeded {
  $current = if (Test-Path $BrowserIconStamp) { Get-Content -LiteralPath $BrowserIconStamp -Raw -ErrorAction SilentlyContinue } else { "" }
  if ($current -eq $IconVersion) {
    return
  }

  $targets = @(
    (Join-Path $BrowserProfile "Default\Favicons"),
    (Join-Path $BrowserProfile "Default\Favicons-journal"),
    (Join-Path $BrowserProfile "Default\Top Sites"),
    (Join-Path $BrowserProfile "Default\Top Sites-journal"),
    (Join-Path $BrowserProfile "Default\Web Data"),
    (Join-Path $BrowserProfile "Default\Web Data-journal")
  )
  foreach ($target in $targets) {
    Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
  }
  New-Item -ItemType Directory -Path (Split-Path -Parent $BrowserIconStamp) -Force | Out-Null
  Set-Content -LiteralPath $BrowserIconStamp -Value $IconVersion -Encoding ASCII
}

Set-Location $AppRoot

if (-not $NodeExe) {
  Show-AppError "No encuentro el runtime de Notion PDF Studio. Reinstala la aplicación y vuelve a abrirla."
  exit 1
}

$serverProcess = Get-PortProcess -LocalPort $Port

if ($serverProcess) {
  if ((Test-ManagedServer -Process $serverProcess) -or (Test-AppServerOnPort -LocalPort $Port)) {
    Stop-Process -Id $serverProcess.Id -Force -ErrorAction SilentlyContinue
    Wait-PortClosed -LocalPort $Port
  } else {
    Show-AppError "El puerto $Port ya lo está usando otra aplicación. Cierra esa aplicación o cambia el puerto de Notion PDF Studio."
    exit 1
  }
}

$serverProcess = Start-Process -FilePath $NodeExe `
  -ArgumentList "`"$ServerScript`"" `
  -WorkingDirectory $AppRoot `
  -WindowStyle Hidden `
  -RedirectStandardOutput $OutLogPath `
  -RedirectStandardError $ErrLogPath `
  -PassThru

$deadline = (Get-Date).AddSeconds(10)
while ((Get-Date) -lt $deadline -and -not (Get-PortProcess -LocalPort $Port)) {
  Start-Sleep -Milliseconds 250
}

if (-not (Get-PortProcess -LocalPort $Port)) {
  Show-AppError "No he podido iniciar Notion PDF Studio. Revisa los logs de la carpeta de la app."
  exit 1
}

$browser = Find-Browser
if (-not $browser) {
  Show-AppError "No encuentro Google Chrome ni Microsoft Edge. Instala Chrome para abrir Notion PDF Studio como aplicación."
  exit 1
}

New-Item -ItemType Directory -Path $BrowserProfile -Force | Out-Null
Clear-BrowserIconCacheIfNeeded

$browserProcess = Start-Process -FilePath $browser -ArgumentList @(
  "--app=$Url",
  "--user-data-dir=$BrowserProfile",
  "--no-first-run",
  "--disable-extensions"
) -PassThru

Start-Sleep -Seconds 2
while (($browserProcess -and -not $browserProcess.HasExited) -or (Test-BrowserOpen)) {
  Start-Sleep -Seconds 1
}

# Give the browser a short grace period to flush the final token save beacon.
Start-Sleep -Milliseconds 900

$activeServer = Get-PortProcess -LocalPort $Port
if ((Test-ManagedServer -Process $activeServer) -or (Test-AppServerOnPort -LocalPort $Port)) {
  Stop-Process -Id $activeServer.Id -Force -ErrorAction SilentlyContinue
}
