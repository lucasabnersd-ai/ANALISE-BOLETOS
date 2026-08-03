@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "PYTHON=C:\Users\lucas\AppData\Local\Programs\Python\Python312\python.exe"
set "PYARGS="

if not exist "%PYTHON%" (
  where py >nul 2>nul
  if errorlevel 1 (
    set "PYTHON=python"
  ) else (
    set "PYTHON=py"
    set "PYARGS=-3"
  )
)

echo ============================================================
echo   PAINEL ANALISE DE BOLETOS - Associacoes Encontradas
echo ============================================================
echo.

"%PYTHON%" %PYARGS% "%~dp0gerar_painel.py" %*
set "CODIGO=%ERRORLEVEL%"

echo.
if "%CODIGO%"=="0" (
  echo CONCLUIDO. Abrindo o painel...
  start "" "%~dp0PUBLICAR\index.html"
) else (
  echo OCORREU UM ERRO. Confira a mensagem acima.
)
echo.

if not defined PAINEL_SEM_PAUSA pause
exit /b %CODIGO%
