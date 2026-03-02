@echo off
cd /d "%~dp0"

echo ==========================================
echo ATUALIZANDO JSONs (OBRAS + FUNCIONARIOS + HISTORICO + MEDICAO)
echo Pasta: %cd%
echo ==========================================
echo.

echo [1/4] Gerando dados.json (OBRAS)...
python scripts\gerar_dados_json.py
if errorlevel 1 (
  echo ERRO ao gerar dados.json
  pause
  exit /b 1
)

echo.
echo [2/4] Gerando funcionarios.json (FUNCIONARIOS)...
python scripts\gerar_funcionarios_json.py
if errorlevel 1 (
  echo ERRO ao gerar funcionarios.json
  pause
  exit /b 1
)

echo.
echo [3/4] Atualizando funcionarios_historico.json (HISTORICO MENSAL)...
python scripts\atualizar_historico_funcionarios.py
if errorlevel 1 (
  echo ERRO ao atualizar funcionarios_historico.json
  pause
  exit /b 1
)

echo.
echo [4/4] Gerando medicao.json (MEDICAO)...
python scripts\medicao_json.py
if errorlevel 1 (
  echo ERRO ao gerar medicao.json
  pause
  exit /b 1
)

echo.
echo ==========================================
echo OK! Tudo atualizado com sucesso
echo Saidas:
echo - docs\dados.json
echo - docs\funcionarios.json
echo - docs\funcionarios_historico.json
echo - docs\medicao.json
echo ==========================================
pause