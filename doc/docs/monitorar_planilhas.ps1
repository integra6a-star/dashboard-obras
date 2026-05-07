$ErrorActionPreference = "Continue"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$docs = Join-Path $root "docs"
$atualizar = Join-Path $root "atualizar.bat"

if (!(Test-Path $docs)) {
  Write-Host "Nao encontrei a pasta docs em: $docs"
  exit 1
}
if (!(Test-Path $atualizar)) {
  Write-Host "Nao encontrei atualizar.bat em: $atualizar"
  exit 1
}

Write-Host "Monitorando planilhas em: $root e $docs"
Write-Host "Quando salvar qualquer .xlsx, vou executar atualizar.bat."
Write-Host "Deixe esta janela aberta. Para parar, pressione Ctrl+C."
Write-Host ""

$script:lastRun = Get-Date "2000-01-01"

function Run-Atualizacao {
  $now = Get-Date
  if (($now - $script:lastRun).TotalSeconds -lt 4) { return }
  $script:lastRun = $now

  Write-Host ""
  Write-Host "Alteracao detectada em planilha. Atualizando JSONs..." -ForegroundColor Cyan
  Push-Location $root
  try {
    & $atualizar nopause
  } finally {
    Pop-Location
  }
  Write-Host "Monitoramento retomado." -ForegroundColor Green
}

function Add-Watcher($path) {
  $watcher = New-Object System.IO.FileSystemWatcher
  $watcher.Path = $path
  $watcher.Filter = "*.xlsx"
  $watcher.IncludeSubdirectories = $false
  $watcher.EnableRaisingEvents = $true

  Register-ObjectEvent $watcher Changed -Action { Run-Atualizacao } | Out-Null
  Register-ObjectEvent $watcher Created -Action { Run-Atualizacao } | Out-Null
  Register-ObjectEvent $watcher Renamed -Action { Run-Atualizacao } | Out-Null
  return $watcher
}

$watchers = @()
$watchers += Add-Watcher $root
$watchers += Add-Watcher $docs

while ($true) {
  Start-Sleep -Seconds 1
}
