@echo off
setlocal
chcp 65001 >nul

rem ATENCAO: `shift`, mais abaixo, desloca TAMBEM o %0 -- depois de ler os
rem argumentos o %~dp0 vira a pasta do argumento, e nao a deste .cmd. Por isso a
rem pasta e guardada AGORA, antes de qualquer coisa, e daqui para baixo so se
rem usa %AQUI%. Sem isso, passar --base "D:\x\y.xlsx" fazia o script procurar o
rem gerar_painel.py dentro de D:\x\.
set "AQUI=%~dp0"
cd /d "%AQUI%"

rem UM CLIQUE = TUDO. Antes eram dois programas em sequencia, na mao: primeiro o
rem RODAR_ASSOCIADOR_BOLETOS.CMD para refazer a base, depois este aqui. Quem
rem esquecesse o primeiro publicava o painel em cima da base do dia anterior sem
rem perceber -- o painel nao tem como saber que a planilha ficou para tras.
rem
rem   [1/4] gera a base   ...\LUCAS ABNER ARAUJO\TRATAMENTO PYTHON BOLETOS.xlsx
rem         (RODAR_ASSOCIADOR_BOLETOS.CMD, que no fim dele tambem sincroniza e
rem         publica o painel BOLETOS-PENDENTES -- os dois saem da mesma rodada
rem         de associacao, entao ou os dois andam, ou um fica mentindo)
rem   [2/4] le a base e monta o painel      (gerar_painel.py)
rem   [3/4] sobe a carteira para o Supabase (publicar_dados.py)
rem   [4/4] publica no GitHub Pages         (deploy.py)
rem
rem Para republicar depressa, sem refazer a associacao (a base ja esta boa):
rem   ATUALIZAR_PAINEL.cmd /sembase
rem Para ler outra planilha:
rem   ATUALIZAR_PAINEL.cmd --base "C:\caminho\outra.xlsx"

set "PYTHON=C:\Users\lucas\AppData\Local\Programs\Python\Python312\python.exe"
set "PYARGS="
set "ASSOCIADOR=%AQUI%..\..\RODAR_ASSOCIADOR_BOLETOS\RODAR_ASSOCIADOR_BOLETOS.CMD"

rem /sembase e nosso, o resto segue para o gerar_painel.py. Sem separar os dois,
rem o argparse do Python receberia /sembase e derrubaria a rodada inteira.
set "GERAR_BASE=1"
set "EXTRA="
:ler_argumentos
if "%~1"=="" goto :argumentos_prontos
if /i "%~1"=="/sembase"   goto :pular_base
if /i "%~1"=="--sem-base" goto :pular_base
set "EXTRA=%EXTRA% "%~1""
shift
goto :ler_argumentos
:pular_base
set "GERAR_BASE="
shift
goto :ler_argumentos
:argumentos_prontos

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

if not defined GERAR_BASE (
  echo [1/4] Base NAO regerada ^(/sembase^): usando a planilha como ela esta.
  echo.
  goto :montar_painel
)

echo [1/4] Gerando a base TRATAMENTO PYTHON BOLETOS.xlsx...
echo       ^(associacao das 6 bases; demora alguns minutos^)
echo.
if not exist "%ASSOCIADOR%" (
  echo   ERRO: nao encontrei o gerador da base:
  echo   "%ASSOCIADOR%"
  echo   Rode-o na mao e depois use: ATUALIZAR_PAINEL.cmd /sembase
  set "CODIGO=1"
  goto :erro
)
rem Sem isto o associador para num "pressione qualquer tecla" e a rodada trava.
set "ASSOCIADOR_SEM_PAUSA=1"
call "%ASSOCIADOR%"
set "CODIGO=%ERRORLEVEL%"
rem O associador troca a pasta corrente; o resto daqui conta com a nossa.
cd /d "%AQUI%"
if not "%CODIGO%"=="0" (
  echo.
  echo   A BASE NAO FOI GERADA -- o painel NAO foi atualizado.
  echo   Nada foi publicado com base velha de proposito: seria impossivel
  echo   perceber depois que a planilha ficou para tras.
  echo   Para publicar assim mesmo, com a base que ja esta na pasta:
  echo       ATUALIZAR_PAINEL.cmd /sembase
  goto :erro
)

:montar_painel
echo.
echo [2/4] Lendo a planilha e montando o painel...
echo.
"%PYTHON%" %PYARGS% "%AQUI%gerar_painel.py"%EXTRA%
set "CODIGO=%ERRORLEVEL%"
if not "%CODIGO%"=="0" goto :erro

echo.
echo [3/4] Subindo a carteira para o painel publicado...
echo.
rem Sem esta etapa o painel no ar continuaria com os dados da carga anterior:
rem o gerar_painel.py so escreve os arquivos aqui na maquina.
if not exist "%AQUI%.analise_boletos_token" (
  echo   AVISO: falta o arquivo .analise_boletos_token nesta pasta.
  echo   O painel LOCAL esta atualizado, mas a carteira publicada continua
  echo   com os dados da carga anterior.
  goto :publicar_git
)

"%PYTHON%" %PYARGS% "%AQUI%publicar_dados.py"
set "CODIGO=%ERRORLEVEL%"
if not "%CODIGO%"=="0" goto :erro

:publicar_git
echo.
echo [4/4] Publicando PUBLICAR\ no GitHub Pages (commit e push)...
echo.
rem Sem esta etapa o layout gerado (PUBLICAR\index.html) fica so nesta maquina:
rem o deploy.py faz git add/commit/push do que o .gitignore deixar passar.
"%PYTHON%" %PYARGS% "%AQUI%deploy.py"
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

