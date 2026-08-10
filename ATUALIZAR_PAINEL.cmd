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

echo [1/3] Lendo a planilha e montando o painel...
echo.
"%PYTHON%" %PYARGS% "%~dp0gerar_painel.py" %*
set "CODIGO=%ERRORLEVEL%"
if not "%CODIGO%"=="0" goto :erro

echo.
echo [2/3] Subindo a carteira para o painel publicado...
echo.
rem Sem esta etapa o painel no ar continuaria com os dados da carga anterior:
rem o gerar_painel.py so escreve os arquivos aqui na maquina.
if not exist "%~dp0.analise_boletos_token" (
  echo   AVISO: falta o arquivo .analise_boletos_token nesta pasta.
  echo   O painel LOCAL esta atualizado, mas a carteira publicada continua
  echo   com os dados da carga anterior.
  goto :publicar_git
)

"%PYTHON%" %PYARGS% "%~dp0publicar_dados.py"
set "CODIGO=%ERRORLEVEL%"
if not "%CODIGO%"=="0" goto :erro

:publicar_git
echo.
echo [3/3] Publicando PUBLICAR\ no GitHub Pages (commit e push)...
echo.
rem Sem esta etapa o layout gerado (PUBLICAR\index.html) fica so nesta maquina:
rem o deploy.py faz git add/commit/push do que o .gitignore deixar passar.
"%PYTHON%" %PYARGS% "%~dp0deploy.py"
set "CODIGO=%ERRORLEVEL%"
if "%CODIGO%"=="2" (
  echo   Nada mudou no painel publicado desde o ultimo commit.
  goto :fim
)
if not "%CODIGO%"=="0" goto :erro

:fim
echo.
echo CONCLUIDO. Painel publicado em https://lucasabnersd-ai.github.io/ANALISE-BOLETOS/PUBLICAR/
echo.
if not defined PAINEL_SEM_PAUSA pause
exit /b 0

:erro
echo.
echo OCORREU UM ERRO. Confira a mensagem acima.
echo.
if not defined PAINEL_SEM_PAUSA pause
exit /b %CODIGO%
