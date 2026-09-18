@echo off
setlocal
chcp 65001 >nul
set "AQUI=%~dp0"
cd /d "%AQUI%"

rem ===========================================================================
rem OS QUATRO AVISOS DIARIOS, sem depender de ninguem clicar.
rem
rem   alerta_adiantamento.py  titulo AMARRAR ADIANTAMENTO que esta COM boleto
rem   alerta_dia_util.py      titulo em aberto vencendo em sabado/domingo/feriado
rem   alerta_boleto_sc.py     titulo BOLETO S/C COM linha digitavel e cod. barras
rem   alerta_bordero.py       titulo com bordero emitido e SEM data de baixa
rem
rem Os dois primeiros vao para a Graziela. O terceiro (18/09/2026) vai para o
rem Rafael e a Graziela, com a Gabriella em copia -- destinatarios em
rem .alerta_boleto_sc_para e .alerta_boleto_sc_copia, os dois FORA do repo.
rem O quarto (18/09/2026) vai para o Lucas e a Graziela (.alerta_bordero_para).
rem
rem REDE DE PROTECAO, NAO O PRINCIPAL. Quem manda primeiro e a rodada do painel
rem (ATUALIZAR_PAINEL.cmd, passos 3, 4, 5 e 6) ou a do SE2 (PAINEL_SE2.cmd, passo 5) --
rem la os dados sao do dia. Este .cmd roda no fim da tarde pelo Agendador e so
rem tem efeito no dia em que NENHUMA rodada aconteceu: a trava diaria de cada
rem script (DADOSlertas_enviados.json) impede o e-mail repetido.
rem
rem Por que nao agendar de manha: a trava gastaria o envio do dia lendo a SE2 de
rem ONTEM, e a rodada da tarde -- com o dado certo -- ficaria calada. O rodape do
rem e-mail sempre diz de quando e a base, para quem le saber o que esta olhando.
rem
rem ATENCAO: precisa da sessao do usuario ABERTA: o envio e pelo Outlook classico (COM),
rem que nao existe fora de uma sessao interativa. Por isso a tarefa e "executar
rem somente quando o usuario estiver conectado".
rem ===========================================================================

set "PYTHON=%PAINEL_PYTHON%"
if not defined PYTHON set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
set "PYARGS="
if not exist "%PYTHON%" (
  where py >nul 2>nul
  if not errorlevel 1 ( set "PYTHON=py" & set "PYARGS=-3" ) else ( set "PYTHON=python" )
)

echo ============================================================
echo   AVISOS DIARIOS - PAINEL DE BOLETOS
echo   %DATE% %TIME%
echo ============================================================
echo.

echo [1/4] AMARRAR ADIANTAMENTO com boleto...
"%PYTHON%" %PYARGS% "%AQUI%alerta_adiantamento.py"
echo.

echo [2/4] Vencimento em sabado, domingo ou feriado...
"%PYTHON%" %PYARGS% "%AQUI%alerta_dia_util.py"
echo.

echo [3/4] BOLETO S/C com linha digitavel e codigo de barras...
"%PYTHON%" %PYARGS% "%AQUI%alerta_boleto_sc.py"
echo.

echo [4/4] Bordero emitido e SEM data de baixa...
"%PYTHON%" %PYARGS% "%AQUI%alerta_bordero.py"
echo.

rem Sai sempre 0: falha de e-mail nao e' motivo para o Agendador marcar erro
rem vermelho e o alarme virar barulho de fundo.
if not defined ALERTAS_SEM_PAUSA if "%1"=="" pause
exit /b 0
