@echo off
cd /d "%~dp0"

echo ==========================================
echo ATUALIZANDO DASHBOARD ANTIGO - OBRAS 6A
echo Pasta: %cd%
echo ==========================================
echo.

echo [1/5] Gerando dados.json + eap_producao.json...
python scripts\gerar_dados_json.py
if errorlevel 1 goto erro

echo.
echo [2/5] Gerando JSONs do almoxarifado...
python scripts\almoxarifado_json.py
if errorlevel 1 goto erro

echo.
echo [3/5] Gerando funcionarios.json...
python scripts\gerar_funcionarios_json.py
if errorlevel 1 echo AVISO: funcionarios.json nao foi atualizado.

echo.
echo [4/5] Atualizando historico de funcionarios...
python scripts\atualizar_historico_funcionarios.py
if errorlevel 1 echo AVISO: funcionarios_historico.json nao foi atualizado.

echo.
echo [5/5] Gerando medicao.json...
python scripts\medicao_json.py
if errorlevel 1 echo AVISO: medicao.json nao foi atualizado.

echo.
echo ==========================================
echo OK! JSONs atualizados.
echo Confira e depois suba/substitua a pasta docs no GitHub.
echo ==========================================
pause
exit /b 0

:erro
echo.
echo ==========================================
echo ERRO: A atualizacao foi interrompida.
echo Confira se as planilhas estao fechadas e se o Python esta instalado.
echo ==========================================
pause
exit /b 1
