@echo off
setlocal
cd /d "%~dp0"

set "MAINT_FILE=docs\manutencao.json"
set "MSG_ON=Dashboard em manutencao. Atualizando dados..."
set "MSG_ERR=Dashboard em manutencao. Atualizacao interrompida."

call :SET_MAINT_ON "%MSG_ON%"

echo ==========================================
echo ATUALIZANDO JSONs (OBRAS + FUNCIONARIOS + HISTORICO + MEDICAO + EAP + PDS)
echo Pasta: %cd%
echo ==========================================
echo.

echo [1/5] Gerando dados.json (OBRAS)...
python scripts\gerar_dados_json.py
if errorlevel 1 goto :FAIL_DADOS

echo.
echo [2/5] Gerando funcionarios.json (FUNCIONARIOS)...
python scripts\gerar_funcionarios_json.py
if errorlevel 1 goto :FAIL_FUNC

echo.
echo [3/5] Atualizando funcionarios_historico.json (HISTORICO MENSAL)...
python scripts\atualizar_historico_funcionarios.py
if errorlevel 1 goto :FAIL_HIST

echo.
echo [4/5] Gerando medicao.json (MEDICAO)...
python scripts\medicao_json.py
if errorlevel 1 goto :FAIL_MEDICAO

echo.
echo [5/6] Atualizando EAP / curva de producao...
python scripts\gerar_eap_json.py
if errorlevel 1 goto :FAIL_EAP

echo.
echo [6/6] Atualizando PDS...
python scripts\converter_pds_corrigido.py
if errorlevel 1 goto :FAIL_PDS

echo.
if exist "docs\eap_producao.json" (
    echo OK: docs\eap_producao.json atualizado.
) else (
    echo AVISO: docs\eap_producao.json nao encontrado.
)

call :SET_MAINT_OFF

echo.
echo ==========================================
echo OK! Tudo atualizado com sucesso
echo Saidas:
echo - docs\dados.json
echo - docs\funcionarios.json
echo - docs\funcionarios_historico.json
echo - docs\medicao.json
echo - docs\eap_producao.json
echo - docs\pds_data.json
echo - docs\manutencao.json
if exist "%MAINT_FILE%" echo - manutencao desligada automaticamente
echo ==========================================
pause
exit /b 0

:FAIL_DADOS
echo ERRO ao gerar dados.json
goto :FAIL

:FAIL_FUNC
echo ERRO ao gerar funcionarios.json
goto :FAIL

:FAIL_HIST
echo ERRO ao atualizar funcionarios_historico.json
goto :FAIL

:FAIL_MEDICAO
echo ERRO ao gerar medicao.json
goto :FAIL

:FAIL_EAP
echo ERRO ao atualizar EAP / curva de producao
goto :FAIL

:FAIL_PDS
echo ERRO ao atualizar pds_data.json
goto :FAIL

:FAIL
call :SET_MAINT_ON "%MSG_ERR%"
echo.
echo A manutencao permaneceu LIGADA para evitar acesso durante falha.
pause
exit /b 1

:SET_MAINT_ON
if not exist "docs" mkdir "docs" >nul 2>&1
>"%MAINT_FILE%" echo {
>>"%MAINT_FILE%" echo   "ativo": true,
>>"%MAINT_FILE%" echo   "mensagem": "%~1"
>>"%MAINT_FILE%" echo }
echo [MANUTENCAO] Ligada.
exit /b 0

:SET_MAINT_OFF
if not exist "docs" mkdir "docs" >nul 2>&1
>"%MAINT_FILE%" echo {
>>"%MAINT_FILE%" echo   "ativo": false,
>>"%MAINT_FILE%" echo   "mensagem": ""
>>"%MAINT_FILE%" echo }
echo [MANUTENCAO] Desligada.
exit /b 0
