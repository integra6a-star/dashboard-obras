@echo off
setlocal EnableExtensions

cd /d "%~dp0"

echo ==========================================
echo ABRINDO DASHBOARD (SERVIDOR LOCAL)
echo Pasta: %cd%
echo ==========================================
echo.

REM mata servidor antigo (se estiver rodando)
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :8000 ^| findstr LISTENING') do (
  taskkill /PID %%a /F >nul 2>&1
)

REM sobe servidor local
start "server" /min cmd /c "python -m http.server 8000 --bind 127.0.0.1"

REM espera 1 segundo
ping 127.0.0.1 -n 2 >nul

REM abre o dashboard no navegador
start "" "http://127.0.0.1:8000/docs/index.html"

echo.
echo ✅ Aberto em: http://127.0.0.1:8000/docs/index.html
echo (Para parar o servidor: feche a janela "server" ou finalize o python)
echo.
endlocal