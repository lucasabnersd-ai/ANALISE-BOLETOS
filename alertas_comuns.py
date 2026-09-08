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

def enviar(assunto: str, corpo_html: str, para: str) -> None:
    """Manda pelo Outlook CLASSICO desta maquina.

    `Outlook.Application` e' o COM registrado em Office16\\OUTLOOK.EXE -- o
    Outlook novo (olk.exe, da Microsoft Store) nao expoe esse conector e nao
    participa, mesmo estando aberto.

    ⚠ Com o Outlook fechado o COM sobe uma instancia e a mensagem pode ficar
    parada na Caixa de Saida ate alguem abrir o programa. Por isso o
    SendAndReceive logo depois do Send.
    """
    import win32com.client  # so aqui: em maquina sem Outlook o resto ainda roda

    outlook = win32com.client.Dispatch("Outlook.Application")
    email = outlook.CreateItem(0)  # 0 = olMailItem
    email.To = para
    email.Subject = assunto
    email.HTMLBody = corpo_html
    email.Send()
    try:
        outlook.GetNamespace("MAPI").SendAndReceive(False)
    except Exception as erro:  # noqa: BLE001 - avisar e seguir
        print(f"   AVISO: nao consegui forcar o envio ({erro}).")
        print("   Se o Outlook estiver fechado, o e-mail sai quando ele abrir.")


# --------------------------------------------------------------------------
# tabela do corpo do e-mail
# --------------------------------------------------------------------------

def tabela(colunas: list[tuple[str, str]], linhas: list[dict],
           direita: tuple[str, ...] = (), destaque: tuple[str, ...] = ()) -> str:
    """Tabela HTML com estilo em cada celula.

    O Outlook ignora <style> em muitos cenarios -- por isso nada de folha de
    estilo: tudo em atributo, celula por celula.
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
            celulas.append(f'<td style="{estilo}{fundo}">{html.escape(valor)}</td>')
        corpo.append("<tr>" + "".join(celulas) + "</tr>")
    return (f'<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;{FONTE}">'
            f"<thead><tr>{cab}</tr></thead><tbody>{''.join(corpo)}</tbody></table>")


def rodape(origem: str, extra: str = "") -> str:
    return (f'<p style="margin:14px 0 0 0;font-size:10pt;color:#595959;{FONTE}">'
            f'Painel de Análise de Boletos — '
            f'<a href="https://analise-boletos.vercel.app/">abrir o painel</a><br>'
            f'Aviso automático diário. {html.escape(origem)}'
            f'{"<br>" + html.escape(extra) if extra else ""}</p>')
