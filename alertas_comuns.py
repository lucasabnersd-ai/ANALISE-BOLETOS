# -*- coding: utf-8 -*-
"""O que os alertas diarios por e-mail tem em comum.

Hoje sao dois, os dois disparados pela rodada do painel (que le a SE2):

  alerta_adiantamento.py  titulo classificado como AMARRAR ADIANTAMENTO que
                          esta COM boleto associado
  alerta_dia_util.py      titulo em aberto vencendo em sabado, domingo ou
                          feriado -- pedido de ajuste para o proximo dia util

Aqui mora o que seria copiado entre os dois: destinatario, envio pelo Outlook,
a TRAVA DIARIA e o calendario de feriados.

⚠ A TRAVA DIARIA existe porque a rodada do painel roda VARIAS VEZES por dia (e
agora tambem pela rodada do SE2). Sem ela a Graziela receberia o mesmo e-mail a
cada clique, e alerta que chega demais vira alerta que ninguem le. "Diariamente"
= uma vez por dia, no primeiro que rodar.
"""

from __future__ import annotations

import datetime as dt
import html
import json
import subprocess
import sys
from pathlib import Path

PASTA = Path(__file__).resolve().parent
DADOS = PASTA / "DADOS"
ENVIADOS = DADOS / "alertas_enviados.json"

# ⚠ Destinatario NUNCA escrito em .py: este repositorio e servido pelo GitHub
# Pages, que entrega o repo inteiro (ANALISE-BOLETOS/gerar_painel.py responde
# 200). E-mail de pessoa em arquivo versionado seria e-mail publicado na
# internet. Estes arquivos ficam no .gitignore e viajam pelo OneDrive.
#   .alerta_<nome>_para      especifico de um alerta
#   .alerta_adiantamento_para  o padrao, quando nao ha o especifico
PADRAO_DESTINO = ".alerta_adiantamento_para"

FONTE = "font-family:Calibri,Arial,sans-serif;"


# --------------------------------------------------------------------------
# destinatario
# --------------------------------------------------------------------------

def destinatarios(nome: str = "") -> str:
    """Para quem este alerta vai. "" quando nao ha ninguem configurado."""
    candidatos = []
    if nome:
        candidatos.append(PASTA / f".alerta_{nome}_para")
    candidatos.append(PASTA / PADRAO_DESTINO)
    for arquivo in candidatos:
        if not arquivo.exists():
            continue
        # utf-8-sig: o Bloco de Notas e o `Set-Content -Encoding utf8` do
        # PowerShell gravam BOM, e o BOM entrava GRUDADO no endereco.
        linhas = arquivo.read_text(encoding="utf-8-sig").splitlines()
        enderecos = [l.strip().lstrip("﻿") for l in linhas]
        enderecos = [e for e in enderecos if e and not e.startswith("#")]
        if enderecos:
            return "; ".join(enderecos)
    return ""


def copias(nome: str = "") -> str:
    """Quem entra em CC neste alerta. "" quando ninguem.

    Arquivo separado do destinatario de proposito: quem esta em copia nao e'
    quem tem de agir. Misturar os dois numa lista so apagaria essa diferenca --
    e e' ela que decide quem responde.

    Nao ha copia PADRAO: sem `.alerta_<nome>_copia` o alerta sai sem CC. Copia
    herdada por engano e' e-mail de gente que nao pediu para receber.
    """
    if not nome:
        return ""
    arquivo = PASTA / f".alerta_{nome}_copia"
    if not arquivo.exists():
        return ""
    linhas = arquivo.read_text(encoding="utf-8-sig").splitlines()
    enderecos = [l.strip().lstrip("﻿") for l in linhas]
    enderecos = [e for e in enderecos if e and not e.startswith("#")]
    return "; ".join(enderecos)


# --------------------------------------------------------------------------
# trava diaria
# --------------------------------------------------------------------------

def _registro() -> dict:
    if not ENVIADOS.exists():
        return {}
    try:
        return json.loads(ENVIADOS.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        # Arquivo corrompido nao pode calar o alerta: melhor mandar de novo do
        # que deixar de mandar.
        return {}


def ja_enviado_hoje(chave: str, hoje: dt.date | None = None) -> bool:
    hoje = hoje or dt.date.today()
    return _registro().get(chave, {}).get("data") == hoje.isoformat()


def marcar_enviado(chave: str, quantos: int, hoje: dt.date | None = None) -> None:
    hoje = hoje or dt.date.today()
    reg = _registro()
    reg[chave] = {"data": hoje.isoformat(), "quantos": quantos,
                  "hora": dt.datetime.now().strftime("%H:%M")}
    DADOS.mkdir(parents=True, exist_ok=True)
    ENVIADOS.write_text(json.dumps(reg, ensure_ascii=False, indent=2), encoding="utf-8")


def quando_enviou(chave: str) -> str:
    r = _registro().get(chave)
    if not r:
        return "nunca"
    d = r.get("data", "")
    try:
        d = dt.date.fromisoformat(d).strftime("%d/%m/%Y")
    except ValueError:
        pass
    return f"{d} as {r.get('hora','?')} ({r.get('quantos','?')} titulo(s))"


# --------------------------------------------------------------------------
# feriados e dia util
# --------------------------------------------------------------------------

def _pascoa(ano: int) -> dt.date:
    """Domingo de Pascoa (algoritmo gregoriano anonimo).

    Calculado, e nao tabelado, de proposito: Carnaval, Sexta-feira Santa e
    Corpus Christi mudam todo ano. Uma lista fixa de 2026 pareceria certa ate
    janeiro e depois passaria a mentir em silencio -- o pior defeito possivel
    num alerta que existe justamente para achar data errada.
    """
    a = ano % 19
    b, c = divmod(ano, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes, dia = divmod(h + l - 7 * m + 114, 31)
    return dt.date(ano, mes, dia + 1)


def feriados(ano: int) -> dict[dt.date, str]:
    """Feriados NACIONAIS em que o banco nao compensa, no ano pedido.

    Municipais e estaduais variam por filial e nao cabem numa lista fixa: quem
    precisar acrescenta em `.feriados_locais`, uma linha por feriado,
    "DD/MM/AAAA;Nome". Linha vazia e linha com # sao ignoradas.
    """
    p = _pascoa(ano)
    lista = {
        dt.date(ano, 1, 1): "Confraternizacao Universal",
        p - dt.timedelta(days=48): "Carnaval (segunda)",
        p - dt.timedelta(days=47): "Carnaval (terca)",
        p - dt.timedelta(days=2): "Sexta-feira Santa",
        dt.date(ano, 4, 21): "Tiradentes",
        dt.date(ano, 5, 1): "Dia do Trabalho",
        p + dt.timedelta(days=60): "Corpus Christi",
        dt.date(ano, 9, 7): "Independencia",
        dt.date(ano, 10, 12): "Nossa Senhora Aparecida",
        dt.date(ano, 11, 2): "Finados",
        dt.date(ano, 11, 15): "Proclamacao da Republica",
        dt.date(ano, 11, 20): "Consciencia Negra",
        dt.date(ano, 12, 25): "Natal",
    }
    locais = PASTA / ".feriados_locais"
    if locais.exists():
        for linha in locais.read_text(encoding="utf-8-sig").splitlines():
            linha = linha.strip()
            if not linha or linha.startswith("#"):
                continue
            data, _, nome = linha.partition(";")
            try:
                d = dt.datetime.strptime(data.strip(), "%d/%m/%Y").date()
            except ValueError:
                continue
            if d.year == ano:
                lista[d] = (nome.strip() or "Feriado local") + " (local)"
    return lista


def porque_nao_e_util(d: dt.date) -> str:
    """"" quando o dia e util; senao o motivo, pronto para a tabela."""
    if d.weekday() == 5:
        return "Sabado"
    if d.weekday() == 6:
        return "Domingo"
    return feriados(d.year).get(d, "")


def proximo_dia_util(d: dt.date) -> dt.date:
    seguinte = d + dt.timedelta(days=1)
    while porque_nao_e_util(seguinte):
        seguinte += dt.timedelta(days=1)
    return seguinte


# --------------------------------------------------------------------------
# e-mail
# --------------------------------------------------------------------------

class EnvioTravado(RuntimeError):
    """O Outlook nao aceitou o e-mail no prazo -- a mensagem diz por que."""


def _caixa_de_permissao_aberta() -> str:
    """"" quando o caminho esta livre; senao a frase do que esta segurando.

    O Outlook classico tem um guarda do modelo de objeto: script mandando
    e-mail faz aparecer a caixa "Programa tentando enviar email em seu nome",
    com os botoes Permitir e Negar. Enquanto ninguem clica, o .Send() NAO
    VOLTA. E numa instancia -Embedding (a que o COM abre sozinho, sem janela
    na barra de tarefas) ninguem ve a caixa para clicar.

    10/09/2026: foi exatamente isso. A rodada do SE2 ficou muda no [5/5] com
    QUATRO caixas dessas empilhadas -- uma por rodada do dia -- e um
    OUTLOOK.EXE queimando 12 minutos de CPU. Tentar enviar com uma aberta so
    empilha a quinta, entao aqui a gente olha ANTES e nem tenta.
    """
    try:
        import win32gui
    except ImportError:
        return ""  # sem pywin32 nao ha o que olhar; o prazo do envio protege

    def botoes(janela) -> list[str]:
        rotulos: list[str] = []

        def filho(h, _):
            if win32gui.GetClassName(h) == "Button":
                rotulos.append(win32gui.GetWindowText(h).replace("&", "").strip())
            return True

        try:
            win32gui.EnumChildWindows(janela, filho, None)
        except Exception:  # noqa: BLE001 - janela pode morrer no meio da varredura
            pass
        return rotulos

    achadas: list[int] = []

    def topo(h, _):
        # #32770 e' a classe de CAIXA DE DIALOGO do Windows.
        if win32gui.GetClassName(h) != "#32770":
            return True
        if "Outlook" not in win32gui.GetWindowText(h):
            return True
        if "Permitir" in botoes(h):
            achadas.append(h)
        return True

    try:
        win32gui.EnumWindows(topo, None)
    except Exception:  # noqa: BLE001 - inspecao de janela nunca derruba a rodada
        return ""
    if not achadas:
        return ""
    return (f"o Outlook esta pedindo autorizacao NA TELA ({len(achadas)} caixa(s) "
            '"Programa tentando enviar email em seu nome"); clique em Permitir, '
            "ou encerre o OUTLOOK.EXE -Embedding no Gerenciador de Tarefas")


def enviar(assunto: str, corpo_html: str, para: str, segundos: int = 45,
           copia: str = "") -> None:
    """Manda pelo Outlook CLASSICO desta maquina, COM PRAZO.

    `Outlook.Application` e' o COM registrado em Office16\\OUTLOOK.EXE -- o
    Outlook novo (olk.exe, da Microsoft Store) nao expoe esse conector e nao
    participa, mesmo estando aberto.

    10/09/2026 -- POR QUE O ENVIO VAI PARA UM PROCESSO FILHO: o .Send() pode
    parar para sempre esperando clique na caixa de permissao (ver
    _caixa_de_permissao_aberta). Thread presa dentro do COM nao morre;
    processo filho morre. Estourado o prazo o filho e' encerrado, a rodada
    segue e a trava diaria NAO e' marcada -- assim a proxima rodada tenta de
    novo em vez de dar o dia por avisado.
    """
    travado = _caixa_de_permissao_aberta()
    if travado:
        raise EnvioTravado(travado)

    # ensure_ascii padrao (True) de proposito: o que anda no cano e' ASCII
    # puro (acento vira \uXXXX), entao nenhum decodificador do outro lado tem
    # como estragar. Trava junto com o decode explicito do _enviar_outlook.
    pedido = json.dumps({"para": para, "copia": copia,
                         "assunto": assunto, "corpo": corpo_html})
    filho = subprocess.Popen(
        [sys.executable, str(PASTA / "_enviar_outlook.py")],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace")
    try:
        saida, _ = filho.communicate(pedido, timeout=segundos)
    except subprocess.TimeoutExpired:
        filho.kill()
        filho.communicate()
        raise EnvioTravado(
            f"o Outlook nao respondeu em {segundos}s -- quase sempre e' a caixa de "
            "permissao esperando clique. Abra o Outlook CLASSICO e clique em "
            "Permitir, ou encerre o OUTLOOK.EXE -Embedding no Gerenciador de Tarefas"
        ) from None
    for linha in (saida or "").splitlines():
        if linha.strip():
            print("   " + linha.strip())
    if filho.returncode != 0:
        raise EnvioTravado(f"o envio falhou (codigo {filho.returncode})")


# --------------------------------------------------------------------------
# tabela do corpo do e-mail
# --------------------------------------------------------------------------

def tabela(colunas: list[tuple[str, str]], linhas: list[dict],
           direita: tuple[str, ...] = (), destaque: tuple[str, ...] = (),
           links: dict | None = None, sublinha=None) -> str:
    """Tabela HTML com estilo em cada celula.

    O Outlook ignora <style> em muitos cenarios -- por isso nada de folha de
    estilo: tudo em atributo, celula por celula.

    `links` transforma a celula de um campo em link: {campo: endereco}, onde o
    endereco e' um texto fixo ou uma funcao que recebe a linha e devolve o
    endereco daquela linha. O TEXTO da celula continua escapado; so o endereco
    entra no href -- celula nenhuma recebe HTML cru vindo da base.

    `sublinha` e' uma funcao que recebe a linha e devolve um texto para uma
    FAIXA logo abaixo dela, ocupando a largura inteira ("" = linha sem faixa).
    Serve para o dado comprido -- linha digitavel sao 47 digitos numa palavra
    so -- que como COLUNA empurraria a tabela para fora da tela (foi o que
    aconteceu no alerta BOLETO S/C em 18/09/2026). Na faixa ele quebra sozinho
    (`word-break`) e a tabela continua estreita.
    """
    th = (FONTE + "font-size:11pt;background:#1F3864;color:#FFFFFF;"
          "padding:6px 9px;border:1px solid #1F3864;text-align:left;white-space:nowrap;")
    td = (FONTE + "font-size:11pt;padding:5px 9px;border:1px solid #BFBFBF;vertical-align:top;")
    # digito de largura fixa: e assim que valor bate com valor de olho
    td_num = td + "text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums;"

    cab = "".join(f'<th style="{th}">{html.escape(rotulo)}</th>' for rotulo, _ in colunas)
    corpo = []
    for i, linha in enumerate(linhas):
        fundo = "background:#F2F2F2;" if i % 2 else ""
        celulas = []
        for rotulo, campo in colunas:
            valor = str(linha.get(campo) or "").strip() or "—"
            estilo = td_num if campo in direita else td
            if campo in destaque:
                estilo += "font-weight:bold;white-space:nowrap;"
            conteudo = html.escape(valor)
            endereco = (links or {}).get(campo)
            if callable(endereco):
                endereco = endereco(linha)
            if endereco and valor != "—":
                conteudo = (f'<a href="{html.escape(str(endereco), quote=True)}" '
                            f'style="color:#1F3864;">{conteudo}</a>')
            celulas.append(f'<td style="{estilo}{fundo}">{conteudo}</td>')
        corpo.append("<tr>" + "".join(celulas) + "</tr>")
        extra = sublinha(linha) if sublinha else ""
        if extra:
            # colspan na largura inteira: a faixa acompanha a linha de cima,
            # inclusive no zebrado, para nao parecer registro solto
            td_faixa = (FONTE + "font-size:10pt;color:#404040;padding:4px 9px;"
                        "border:1px solid #BFBFBF;border-top:none;"
                        "word-break:break-all;")
            corpo.append(f'<tr><td colspan="{len(colunas)}" '
                         f'style="{td_faixa}{fundo}">{html.escape(str(extra))}</td></tr>')
    return (f'<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;{FONTE}">'
            f"<thead><tr>{cab}</tr></thead><tbody>{''.join(corpo)}</tbody></table>")


def rodape(origem: str, extra: str = "") -> str:
    return (f'<p style="margin:14px 0 0 0;font-size:10pt;color:#595959;{FONTE}">'
            f'Painel de Análise de Boletos — '
            f'<a href="https://analise-boletos.vercel.app/">abrir o painel</a><br>'
            f'Aviso automático diário. {html.escape(origem)}'
            f'{"<br>" + html.escape(extra) if extra else ""}</p>')
