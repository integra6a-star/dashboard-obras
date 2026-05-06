@echo off
cd /d "%~dp0"
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

set "PY=python"
py -3 --version >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if "%PY%"=="python" (
  python --version >nul 2>nul
  if errorlevel 1 (
    if exist "%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" (
      set "PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    )
  )
)

%PY% --version >nul 2>nul
if errorlevel 1 goto sempython

echo ==========================================
echo ATUALIZANDO DASHBOARD ANTIGO - OBRAS 6A
echo Pasta: %cd%
echo Python: %PY%
echo ==========================================
echo.

echo [1/11] Gerando dados.json + eap_producao.json...
%PY% scripts\gerar_dados_json.py
if errorlevel 1 goto erro

echo.
echo [2/11] Gerando dados_mapa.json da planilha_base_mapa.xlsx...
%PY% scripts\atualizar_mapa_json.py
if errorlevel 1 goto erro

echo.
echo [3/11] Importando PDS Word mais recente...
%PY% scripts\importar_pds_word.py
if errorlevel 1 echo AVISO: PDS Word nao foi importado automaticamente.

echo.
echo [4/11] Gerando pds_data.json...
%PY% scripts\converter_pds.py
if errorlevel 1 echo AVISO: pds_data.json nao foi atualizado.

echo.
echo [5/11] Gerando JSONs do almoxarifado...
%PY% scripts\almoxarifado_json.py
if errorlevel 1 goto erro

echo.
echo [6/11] Gerando funcionarios.json...
%PY% scripts\gerar_funcionarios_json.py
if errorlevel 1 echo AVISO: funcionarios.json nao foi atualizado.

echo.
echo [7/11] Atualizando historico de funcionarios...
%PY% scripts\atualizar_historico_funcionarios.py
if errorlevel 1 echo AVISO: funcionarios_historico.json nao foi atualizado.

echo.
echo [8/11] Gerando medicao.json...
%PY% scripts\medicao_json.py
if errorlevel 1 echo AVISO: medicao.json nao foi atualizado.

echo.
echo [9/11] Conferencia rapida da planilha de mapa...
%PY% -c "import json, pathlib; p=pathlib.Path('docs/dados_mapa.json'); j=json.loads(p.read_text(encoding='utf-8')); print('dados_mapa:', len(j.get('obras',[])), 'obras,', len(j.get('pontos',[])), 'pontos,', len(j.get('trechos',[])), 'trechos')"
if errorlevel 1 echo AVISO: nao consegui conferir dados_mapa.json.

echo.
echo [10/11] Espelhando JSONs da pasta docs para a raiz...
%PY% -c "from pathlib import Path; import shutil; root=Path('.'); docs=root/'docs'; [shutil.copy2(p, root/p.name) for p in docs.glob('*.json')]; print('JSONs espelhados:', len(list(docs.glob('*.json'))))"
if errorlevel 1 goto erro

echo.
echo [11/11] Gerando validacao_dashboard.json e relatorio_atualizacao.txt...
%PY% scripts\validar_dashboard.py
if errorlevel 1 echo AVISO: validacao_dashboard.json nao foi atualizado.

echo.
echo Espelhando JSONs finais da pasta docs para a raiz...
%PY% -c "from pathlib import Path; import shutil; root=Path('.'); docs=root/'docs'; [shutil.copy2(p, root/p.name) for p in docs.glob('*.json')]; print('JSONs espelhados:', len(list(docs.glob('*.json'))))"
if errorlevel 1 goto erro

echo.
echo ==========================================
echo OK! JSONs atualizados.
echo Todas as planilhas foram convertidas para JSON e espelhadas para o link.
echo Relatorio da atualizacao: %cd%\relatorio_atualizacao.txt
echo ==========================================
if /i "%~1"=="nopause" exit /b 0
pause
exit /b 0

:erro
echo.
echo ==========================================
echo ERRO: A atualizacao foi interrompida.
echo Confira se as planilhas estao fechadas e se o Python esta instalado.
echo ==========================================
if /i "%~1"=="nopause" exit /b 1
pause
exit /b 1

:sempython
echo.
echo ==========================================
echo ERRO: nao encontrei Python neste computador.
echo Instale Python ou rode pelo Codex para atualizar os JSONs.
echo ==========================================
if /i "%~1"=="nopause" exit /b 1
pause
exit /b 1
