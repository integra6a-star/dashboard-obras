@echo off
cd /d C:\dashboard-obras
python scripts\gerar_dados_json.py
echo.
echo OK - dados.json atualizado. Abra/atualize o dashboard.
pause
