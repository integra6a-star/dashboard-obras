@echo off
cd /d "%~dp0"

echo ==========================================
echo ATUALIZANDO JSONs (OBRAS + FUNCIONARIOS)
echo Pasta: %cd%
echo ==========================================
echo.

echo [1/2] Gerando dados.json (OBRAS)...
python scripts\gerar_dados_json.py
if errorlevel 1 (
  echo ERRO ao gerar dados.json
  pause
  exit /b 1
)

echo.
echo [2/2] Gerando funcionarios.json (FUNCIONARIOS)...
python scripts\gerar_funcionarios_json.py
if errorlevel 1 (
  echo ERRO ao gerar funcionarios.json
  pause
  exit /b 1
)

echo.
echo ==========================================
echo OK! Tudo atualizado com sucesso
echo Saidas:
echo - docs\dados.json
echo - docs\funcionarios.json
echo ==========================================
pause