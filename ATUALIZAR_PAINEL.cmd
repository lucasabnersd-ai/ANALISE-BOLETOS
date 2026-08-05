@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

rem Atualiza o painel a partir de
rem   ...\LUCAS ABNER ARAUJO\TRATAMENTO PYTHON BOLETOS.xlsx
rem que e a base padrao do gerar_painel.py. Para ler outra planilha:
rem   ATUALIZAR_PAINEL.cmd --base "C:\caminho\outra.xlsx"

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
echo   PAINEL ANALISE DE BOLETOS
echo ============================================================
echo.

echo [1/2] Lendo a planilha e montando o painel...
echo.
"%PYTHON%" %PYARGS% "%~dp0gerar_painel.py" %*
set "CODIGO=%ERRORLEVEL%"
if not "%CODIGO%"=="0" goto :erro

echo.
echo [2/2] Subindo a carteira para o painel publicado...
echo.
rem Sem esta etapa o painel no ar continuaria com os dados da carga anterior:
rem o gerar_painel.py so escreve os arquivos aqui na maquina.
if not exist "%~dp0.analise_boletos_token" (
  echo   AVISO: falta o arquivo .analise_boletos_token nesta pasta.
  echo   O painel LOCAL esta atualizado, mas o painel publicado continua
  echo   com os dados da carga anterior.
  goto :abrir
)

"%PYTHON%" %PYARGS% "%~dp0publicar_dados.py"
set "CODIGO=%ERRORLEVEL%"
if not "%CODIGO%"=="0" goto :erro

:abrir
echo.
echo CONCLUIDO. Abrindo o painel local...
rem dev.html e o painel COM os dados, so nesta maquina. O PUBLICAR\index.html
rem nasce vazio de proposito -- e o que vai para o GitHub Pages.
start "" "%~dp0dev.html"
echo.
if not defined PAINEL_SEM_PAUSA pause
exit /b 0

:erro
echo.
echo OCORREU UM ERRO. Confira a mensagem acima.
echo.
if not defined PAINEL_SEM_PAUSA pause
exit /b %CODIGO%
