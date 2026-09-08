@echo off
setlocal
chcp 65001 >nul
set "AQUI=%~dp0"
cd /d "%AQUI%"

rem ===========================================================================
rem OS DOIS AVISOS DIARIOS DA GRAZIELA, sem depender de ninguem clicar.
rem
rem   alerta_adiantamento.py  titulo AMARRAR ADIANTAMENTO que esta COM boleto
rem   alerta_dia_util.py      titulo em aberto vencendo em sabado/domingo/feriado
rem
rem REDE DE PROTECAO, NAO O PRINCIPAL. Quem manda primeiro e a rodada do painel
rem (ATUALIZAR_PAINEL.cmd, passos 3 e 4) ou a do SE2 (PAINEL_SE2.cmd, passo 5) --
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

echo [1/2] AMARRAR ADIANTAMENTO com boleto...
"%PYTHON%" %PYARGS% "%AQUI%alerta_adiantamento.py"
echo.

echo [2/2] Vencimento em sabado, domingo ou feriado...
"%PYTHON%" %PYARGS% "%AQUI%alerta_dia_util.py"
echo.

rem Sai sempre 0: falha de e-mail nao e' motivo para o Agendador marcar erro
rem vermelho e o alarme virar barulho de fundo.
if not defined ALERTAS_SEM_PAUSA if "%1"=="" pause
exit /b 0
