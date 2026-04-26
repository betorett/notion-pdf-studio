$ErrorActionPreference = "Stop"

$AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Port = if ($env:PORT) { [int]$env:PORT } else { 4173 }
$Url = "http://localhost:$Port"
$OutLogPath = Join-Path $AppRoot "notion-pdf-studio.out.log"
$ErrLogPath = Join-Path $AppRoot "notion-pdf-studio.err.log"
$BrowserProfile = Join-Path $env:LOCALAPPDATA "NotionPdfStudio\BrowserProfile"

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

function Find-Browser {
  $commands = @("msedge", "chrome")
  foreach ($command in $commands) {
    $found = Get-Command $command -ErrorAction SilentlyContinue
    if ($found) {
      return $found.Source
    }
  }

  $paths = @(
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
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

Set-Location $AppRoot

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  Show-AppError "No encuentro Node.js en este equipo. Instala Node.js 20 o superior y vuelve a abrir Notion PDF Studio."
  exit 1
}

$serverProcess = Get-PortProcess -LocalPort $Port
$startedServer = $false

if (-not $serverProcess) {
  $serverProcess = Start-Process -FilePath "node" `
    -ArgumentList "server.js" `
    -WorkingDirectory $AppRoot `
    -WindowStyle Hidden `
    -RedirectStandardOutput $OutLogPath `
    -RedirectStandardError $ErrLogPath `
    -PassThru
  $startedServer = $true
}

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
  Show-AppError "No encuentro Microsoft Edge ni Google Chrome. Instala uno de los dos para abrir Notion PDF Studio como aplicación."
  exit 1
}

New-Item -ItemType Directory -Path $BrowserProfile -Force | Out-Null

Start-Process -FilePath $browser -ArgumentList @(
  "--app=$Url",
  "--user-data-dir=$BrowserProfile",
  "--no-first-run",
  "--disable-extensions"
) | Out-Null

Start-Sleep -Seconds 2
while (Test-BrowserOpen) {
  Start-Sleep -Seconds 1
}

$activeServer = Get-PortProcess -LocalPort $Port
if ($startedServer -and $activeServer) {
  Stop-Process -Id $activeServer.Id -Force -ErrorAction SilentlyContinue
} elseif (Test-ManagedServer -Process $activeServer) {
  Stop-Process -Id $activeServer.Id -Force -ErrorAction SilentlyContinue
}
