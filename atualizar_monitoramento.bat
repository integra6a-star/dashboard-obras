@echo off

setlocal

cd /d %~dp0

set "NODE_PATH=C:\Users\micro\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules;C:\Users\micro\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules\.pnpm\node_modules;C:\Users\micro\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules\.pnpm\playwright@1.61.1\node_modules"
set "SOLCAD_OBRA=INTERCEPTOR ITI-15"

"C:\Users\micro\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" "scripts\solcadgis_monitoramento.js"

if errorlevel 1 (

  echo.

  echo Falha ao atualizar monitoramento topografico.

  pause

  exit /b 1

)

echo.

echo Monitoramento topografico atualizado.

pause
