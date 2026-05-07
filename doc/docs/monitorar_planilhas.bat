@echo off
cd /d "%~dp0"
echo ==========================================
echo MONITOR DE PLANILHAS - DASHBOARD OBRAS
echo ==========================================
echo.

if not exist "%~dp0docs" (
  echo ERRO: nao encontrei a pasta docs.
  echo.
  echo Se voce abriu este arquivo dentro do ZIP, primeiro clique em "Extrair tudo".
  echo Depois rode este .bat dentro da pasta extraida.
  echo.
  pause
  exit /b 1
)

if not exist "%~dp0monitorar_planilhas.ps1" (
  echo ERRO: nao encontrei monitorar_planilhas.ps1.
  echo.
  pause
  exit /b 1
)

echo Vou abrir uma janela do PowerShell para monitorar as planilhas.
echo Deixe a nova janela aberta enquanto estiver editando as planilhas.
echo.
start "Monitorar Planilhas" powershell -NoExit -NoProfile -ExecutionPolicy Bypass -File "%~dp0monitorar_planilhas.ps1"
