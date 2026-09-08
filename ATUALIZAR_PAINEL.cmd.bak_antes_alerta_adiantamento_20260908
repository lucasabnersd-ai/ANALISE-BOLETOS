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

rem ===========================================================================
rem DUAS MAQUINAS, UM SCRIPT SO.
rem Esta pasta e a MESMA biblioteca do OneDrive montada em dois PCs, com nomes
rem diferentes em cada um:
rem   la     C:\Users\lucas\OneDrive - Grupo S&D\Arquivos de Gabriella Karla
rem          Oliveira Milas - FINANCEIRO COMPARTILHADO\LUCAS ABNER ARAUJO
rem   aqui   C:\Users\gabriella.milas\OneDrive - S & D Florestal\0 GABRIELLA
rem          \FINANCEIRO COMPARTILHADO\LUCAS ABNER ARAUJO
rem Como o OneDrive sincroniza a edicao de volta, trocar um caminho fixo pelo
rem outro so mudaria de lugar a maquina quebrada. Por isso daqui para baixo nao
rem existe "C:\Users\<fulano>": o interpretador sai de %LOCALAPPDATA% (que ja e
rem o do usuario logado) e as planilhas saem da pasta DESTE .cmd, que mora
rem dentro da propria biblioteca compartilhada.
rem ===========================================================================

rem --- Python -----------------------------------------------------------
rem Antes era o caminho fixo do Python do usuario "lucas". Fora daquela maquina
rem ele nao existe, e a rodada caia no que estivesse no PATH -- inclusive no
rem atalho da Microsoft Store, que abre a loja e nao roda nada.
set "PYTHON=%PAINEL_PYTHON%"
if not defined PYTHON set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
set "PYARGS="
if not exist "%PYTHON%" call :achar_python
if not defined PYTHON (
  echo   ERRO: nao encontrei o Python nesta maquina.
  echo   Instale o Python 3.12 ou aponte o seu antes de rodar:
  echo       set PAINEL_PYTHON=C:\caminho\para\python.exe
  set "CODIGO=1"
  goto :erro
)

set "ASSOCIADOR=%AQUI%..\..\RODAR_ASSOCIADOR_BOLETOS\RODAR_ASSOCIADOR_BOLETOS.CMD"

rem /sembase e nosso, o resto segue para o gerar_painel.py. Sem separar os dois,
rem o argparse do Python receberia /sembase e derrubaria a rodada inteira.
set "GERAR_BASE=1"
set "EXTRA="
set "BASE_PROPRIA="
:ler_argumentos
if "%~1"=="" goto :argumentos_prontos
if /i "%~1"=="/sembase"   goto :pular_base
if /i "%~1"=="--sem-base" goto :pular_base
rem Quem passa a propria planilha manda mais do que a conferencia la de baixo:
rem sem esta marca, --base "D:\x\y.xlsx" morreria na checagem da base PADRAO,
rem que nem e a que seria lida.
if /i "%~1"=="--base" set "BASE_PROPRIA=1"
rem Sem as aspas do `set "VAR=..."`: o caminho pode ter & no meio
rem ("...\OneDrive - S & D Florestal\..." nesta maquina), e ali o & ficaria
rem FORA das aspas -- cmd o leria como separador de comando e engoliria o
rem resto do caminho. Assim o & entra protegido pelas aspas do proprio valor.
set EXTRA=%EXTRA% "%~1"
shift
goto :ler_argumentos
:pular_base
set "GERAR_BASE="
shift
goto :ler_argumentos
:argumentos_prontos

rem --- As planilhas de origem -------------------------------------------
rem Os defaults do gerar_painel.py sao caminhos fixos da OUTRA maquina. E base
rem que nao e achada nao derruba o painel: a aba dela sai vazia com um AVISO no
rem meio da rodada, que e justamente o que ninguem le. Entao os caminhos sao
rem derivados daqui e passados na linha de comando -- assim valem nas duas.
rem   %AQUI%  ...\LUCAS ABNER ARAUJO\<AUTOMACOES LUCAS>\ANALISES BOLETOS\PAINEL ANALISE BOLETOS\
rem   %RAIZ%  ...\LUCAS ABNER ARAUJO
for %%D in ("%AQUI%..")       do set "ANALISES=%%~fD"
for %%D in ("%AQUI%..\..\..") do set "RAIZ=%%~fD"
rem Onde o associador grava o log de cada rodada -- derivado, e nao escrito na
rem mao, porque o nome da pasta tem cedilha e til de verdade.
for %%D in ("%AQUI%..\..\logs") do set "LOGS=%%~fD"
set "GENERICOS=%RAIZ%\BASES GENERICOS"
set "BASE_PAINEL=%RAIZ%\TRATAMENTO PYTHON BOLETOS.xlsx"

set "BASES="
call :usar_base --base           "%BASE_PAINEL%"
call :usar_base --base-pendentes "%RAIZ%\COPIAR E COLAR BOLETOS PENDENTES.xlsx"
call :usar_base --base-sf1       "%GENERICOS%\SF1.xlsx"
call :usar_base --base-sefaz     "%ANALISES%\SEFAZ.xlsx"
rem 01/09/2026. A manifestacao do destinatario (o que a empresa respondeu sobre a
rem nota: ciencia, confirmacao, desconhecimento, nao realizada). SO PREENCHE a
rem coluna da aba da SEFAZ -- nao acrescenta nota nenhuma ao painel, e de
rem proposito: 643 das 1.380 linhas do arquivo sao notas de terceiros para a
rem BRACELL, em que a S&D so aparece como transportadora. Ver manifestacao.py.
call :usar_base --base-manifestacao "%ANALISES%\MESMA PREMISSA.xlsx"
rem 04/09/2026. PREMISSA 2: ao contrario da de cima, esta ACRESCENTA nota ao
rem painel -- as do grupo que a SEFAZ.xlsx nao tem entram com Origem
rem "PREMISSA 2" e passam pela SF1 e pelos pedidos.
rem NAO passa caminho de proposito: o premissa2.py procura sozinho, na ordem
rem CONSOLIDADA -> ORGANIZADA, e le tanto a aba NOTAS quanto o CSV cru. Fixar
rem um nome aqui foi o que derrubou a leitura em 04/09, quando o ORGANIZADA
rem voltou a ser CSV e as 1.466 notas passaram a morar no CONSOLIDADA.
rem As tres da aba NF x Pedido de Compra. Ate 31/08/2026 elas nao tinham
rem argumento: o caminho vivia dentro do sc7.py/sefaz.py e so dava para apontar
rem outra pasta editando o .py. Agora seguem a mesma regra das outras -- o que
rem nao existir cai no default do modulo, com AVISO nome por nome.
call :usar_base --base-sc7       "%GENERICOS%\SC7.xlsx"
call :usar_base --base-sc1       "%GENERICOS%\SC1.xlsx"
call :usar_base --base-empresas  "%GENERICOS%\LISTAGEM EMPRESAS BIOFLOR.xlsx"
rem "SE2 - POSICAO DIARIA.xlsx" tem cedilha e til no nome de verdade. Escrever o
rem nome aqui deixaria o acerto por conta da pagina de codigo com que o .cmd for
rem lido; o curinga pega o nome real do disco. As copias ("... - Copia.xlsx")
rem ficam de fora: o padrao termina em DIARIA.xlsx.
for %%A in ("%GENERICOS%\SE2 - POSI*O DIARIA.xlsx") do call :usar_base --base-se2 "%%~fA"

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
  echo   O motivo esta na mensagem acima e no log mais recente de:
  echo   "%LOGS%"
  echo.
  echo   Para publicar assim mesmo, com a base que ja esta na pasta:
  echo       ATUALIZAR_PAINEL.cmd /sembase
  goto :erro
)

:montar_painel
rem A base pode ter acabado de nascer na etapa [1/4], entao a conferencia e
rem AQUI. Sem ela, quem some com a planilha ve o gerar_painel.py reclamar do
rem caminho DA OUTRA MAQUINA (o default dele) e vai procurar uma pasta que nem
rem existe neste PC.
if not defined BASE_PROPRIA if not exist "%BASE_PAINEL%" (
  echo.
  echo   ERRO: nao encontrei a base do painel:
  echo   "%BASE_PAINEL%"
  echo   Ela e gerada pelo RODAR_ASSOCIADOR_BOLETOS.CMD ^(etapa 1/4^) e chega
  echo   pela sincronizacao do OneDrive. Sem ela nao ha o que publicar.
  set "CODIGO=1"
  goto :erro
)

echo.
echo [2/4] Lendo a planilha e montando o painel...
echo.
"%PYTHON%" %PYARGS% "%AQUI%gerar_painel.py"%BASES%%EXTRA%
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

rem --- rotinas ----------------------------------------------------------

:achar_python
rem O `py` vem primeiro de proposito: ele escolhe uma instalacao de verdade. O
rem `python` do PATH pode ser o atalho da Microsoft Store, que so abre a loja.
where py >nul 2>nul
if not errorlevel 1 (
  set "PYTHON=py"
  set "PYARGS=-3"
  goto :eof
)
where python >nul 2>nul
if not errorlevel 1 (
  set "PYTHON=python"
  goto :eof
)
set "PYTHON="
goto :eof

:usar_base
rem So passa adiante o caminho que EXISTE. O que faltar segue com o default do
rem gerar_painel.py -- que avisa, nome por nome, qual base ficou de fora.
if exist "%~2" (
  set BASES=%BASES% %1 "%~2"
) else (
  echo   AVISO: nao encontrei %~nx2 -- a aba que depende dela nao sera atualizada.
)
goto :eof
