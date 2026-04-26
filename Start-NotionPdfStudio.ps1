$ErrorActionPreference = "Stop"

$AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Port = if ($env:PORT) { [int]$env:PORT } else { 4173 }
$Url = "http://localhost:$Port"
$OutLogPath = Join-Path $AppRoot "notion-pdf-studio.out.log"
$ErrLogPath = Join-Path $AppRoot "notion-pdf-studio.err.log"

function Test-PortOpen {
  param([int]$LocalPort)
  $connection = Get-NetTCPConnection -LocalPort $LocalPort -ErrorAction SilentlyContinue | Select-Object -First 1
  return $null -ne $connection
}

Set-Location $AppRoot

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  Add-Type -AssemblyName PresentationFramework
  [System.Windows.MessageBox]::Show(
    "No encuentro Node.js en este equipo. Instala Node.js 20 o superior y vuelve a abrir Notion PDF Studio.",
    "Notion PDF Studio",
    "OK",
    "Error"
  ) | Out-Null
  exit 1
}

if (-not (Test-PortOpen -LocalPort $Port)) {
  Start-Process -FilePath "node" `
    -ArgumentList "server.js" `
    -WorkingDirectory $AppRoot `
    -WindowStyle Hidden `
    -RedirectStandardOutput $OutLogPath `
    -RedirectStandardError $ErrLogPath

  $deadline = (Get-Date).AddSeconds(8)
  while ((Get-Date) -lt $deadline -and -not (Test-PortOpen -LocalPort $Port)) {
    Start-Sleep -Milliseconds 250
  }
}

Start-Process $Url
