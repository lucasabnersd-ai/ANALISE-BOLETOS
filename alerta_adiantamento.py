# -*- coding: utf-8 -*-
"""Avisa a Graziela quando um titulo de AMARRAR ADIANTAMENTO aparece COM boleto.

O `tipo_pgto_sf1` vem da classificacao (SF1). "AMARRAR ADIANTAMENTO" quer dizer
que aquele titulo vai ser quitado contra um adiantamento ja pago -- ou seja,
NAO era para existir boleto nenhum ligado a ele. Quando existe, ou a
classificacao esta errada, ou o boleto foi associado ao titulo errado, ou o
fornecedor mandou cobranca de uma coisa ja adiantada. Nos tres casos alguem
paga duas vezes se ninguem olhar.

O painel ja mostra isso, mas so para quem entra e procura. Este script vira o
aviso do avesso: a rodada e que procura, e so manda e-mail quando ACHA.

Uso:
    python alerta_adiantamento.py                 # manda o e-mail se houver caso
    python alerta_adiantamento.py --teste         # NAO manda: grava o HTML e mostra
    python alerta_adiantamento.py --para fulano@x # manda para outro endereco

Codigo de saida: SEMPRE 0 quando o painel foi lido. Falha de e-mail nao pode
derrubar a rodada do painel -- mesma regra da publicacao do BOLETOS-PENDENTES
la no associador. O que da errado sai no texto, nao no errorlevel.
"""

from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import sys
from pathlib import Path

PASTA = Path(__file__).resolve().parent
DADOS = PASTA / "DADOS" / "analise_boletos.json"
PREVIA = PASTA / "DADOS" / "alerta_adiantamento_previa.html"

# ⚠ O destinatario NAO fica escrito aqui. Este repositorio e servido pelo GitHub
# Pages, e o Pages entrega o repo inteiro, nao so a pasta PUBLICAR:
# https://lucasabnersd-ai.github.io/ANALISE-BOLETOS/gerar_painel.py responde 200
# para qualquer um. Endereco de pessoa dentro de .py versionado seria e-mail de
# colega publicado na internet. Fica em .alerta_adiantamento_para, que o
# .gitignore barra (mesma regra do .analise_boletos_token) e que viaja entre as
# duas maquinas pelo OneDrive.
DESTINATARIOS = PASTA / ".alerta_adiantamento_para"


def para_quem() -> str:
    """Le o destinatario do arquivo de fora do repositorio.

    Uma linha por endereco; linha vazia e linha comecando com # sao ignoradas.
    Devolve "" quando nao ha para quem mandar -- quem chama avisa e segue.
    """
    if not DESTINATARIOS.exists():
        return ""
    enderecos = [l.strip() for l in DESTINATARIOS.read_text(encoding="utf-8").splitlines()]
    return "; ".join(e for e in enderecos if e and not e.startswith("#"))

# O que a classificacao escreve na coluna. Comparacao normalizada (maiuscula e
# sem espaco sobrando) porque isso e digitado na mao la na SF1.
TIPO_ALVO = "AMARRAR ADIANTAMENTO"

# Um campo destes preenchido = existe boleto ligado ao titulo. Nao dependemos da
# coluna `status`: ela e um resumo ("SEM BOLETO", "REVISAR", "OK") e ja mudou de
# vocabulario antes. Os campos abaixo so nascem quando houve associacao.
CAMPOS_DE_BOLETO = (
    "boletos", "linha_digitavel", "valor_boleto", "vencimento_boleto",
    "fonte_boleto", "fornecedor_boleto", "cnpj_boleto", "nf_doc_boleto",
)

# Coluna do e-mail -> campo da linha do painel. A ordem daqui e a ordem da
# tabela. O UUID vem PRIMEIRO de proposito: e por ele que ela acha o titulo no
# painel e no TOTVS, e foi o que o Lucas pediu explicitamente.
COLUNAS = (
    ("Código (UUID)", "campo_uuid"),
    ("Filial", "filial"),
    ("Nº título", "no_titulo"),
    ("Parc.", "parcela"),
    ("Fornecedor", "razao_social"),
    ("Vlr título", "vlr_titulo"),
    ("Vlr boleto", "valor_boleto"),
    ("Vencimento", "vencimento"),
    ("Venc. boleto", "vencimento_boleto"),
    ("NF do boleto", "nf_doc_boleto"),
    ("Status", "status"),
    ("Critério do match", "criterio_match"),
    ("Origem do boleto", "fonte_boleto"),
    ("Linha digitável", "linha_digitavel"),
)


def tem_boleto(c: dict) -> bool:
    """Existe boleto associado a este titulo?"""
    for campo in CAMPOS_DE_BOLETO:
        valor = str(c.get(campo) or "").strip()
        # `boletos` e contagem: "0" existe e quer dizer NENHUM.
        if campo == "boletos" and valor in ("", "0"):
            continue
        if valor:
            return True
    return False


def achar(carga: dict) -> list[dict]:
    """Os titulos AMARRAR ADIANTAMENTO que tem boleto, em todas as abas.

    Hoje o `tipo_pgto_sf1` so existe na aba Titulos Associados, mas varremos
    todas: se amanha a coluna entrar em outra aba, o aviso acompanha sozinho --
    o contrario (fixar a aba aqui) sumiria em silencio, que e o pior jeito de
    um alerta falhar.
    """
    achados = []
    for aba in carga.get("abas", []):
        for linha in aba.get("linhas", []):
            c = linha.get("c", {})
            tipo = str(c.get("tipo_pgto_sf1") or "").strip().upper()
            if tipo != TIPO_ALVO:
                continue
            if not tem_boleto(c):
                continue
            achados.append({"aba": aba["nome"], "uuid": linha["uuid"], "c": c})
    # Mesma ordem toda vez: sem isto, dois e-mails com os mesmos titulos
    # pareceriam diferentes so pela ordem em que o painel montou as abas.
    achados.sort(key=lambda a: (a["c"].get("filial") or "",
                                a["c"].get("no_titulo") or "",
                                a["c"].get("parcela") or ""))
    return achados


def montar_html(achados: list[dict], carga: dict) -> str:
    """A previa em tabela, dentro do corpo do e-mail.

    Nada de anexo: a planilha exigiria abrir arquivo de fora da empresa para ver
    tres linhas. Tudo em HTML inline -- o Outlook ignora <style> em muitos
    cenarios, entao o estilo vai atributo por atributo em cada celula.
    """
    fonte = "font-family:Calibri,Arial,sans-serif;"
    th = (fonte + "font-size:11pt;background:#1F3864;color:#FFFFFF;"
          "padding:6px 9px;border:1px solid #1F3864;text-align:left;"
          "white-space:nowrap;")
    td = (fonte + "font-size:11pt;padding:5px 9px;border:1px solid #BFBFBF;"
          "vertical-align:top;")
    # Coluna de numero alinhada a direita e com digito de largura fixa: e assim
    # que valor bate com valor na hora de comparar de olho.
    td_num = td + "text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums;"
    td_cod = td + "white-space:nowrap;"

    cabecalho = "".join(f'<th style="{th}">{html.escape(nome)}</th>'
                        for nome, _ in COLUNAS)

    linhas = []
    for i, a in enumerate(achados):
        # zebra: so para nao perder a linha no meio da tabela larga
        fundo = "background:#F2F2F2;" if i % 2 else ""
        celulas = []
        for nome, campo in COLUNAS:
            valor = str(a["c"].get(campo) or "").strip() or "—"
            estilo = td_num if nome.startswith("Vlr") else (
                td_cod if campo in ("campo_uuid", "linha_digitavel") else td)
            if campo == "campo_uuid":
                estilo += "font-weight:bold;"
            celulas.append(f'<td style="{estilo}{fundo}">{html.escape(valor)}</td>')
        linhas.append("<tr>" + "".join(celulas) + "</tr>")

    quantos = len(achados)
    titulo = ("1 título classificado como <b>AMARRAR ADIANTAMENTO</b> está com boleto associado"
              if quantos == 1 else
              f"{quantos} títulos classificados como <b>AMARRAR ADIANTAMENTO</b> estão com boleto associado")

    return f"""<div style="{fonte}font-size:11pt;color:#000000;">
  <p style="margin:0 0 12px 0;">Bom dia, Graziela,</p>

  <p style="margin:0 0 12px 0;">{titulo}.</p>

  <p style="margin:0 0 14px 0;">Título de <b>amarrar adiantamento</b> é quitado
  contra adiantamento já pago, então não era para haver boleto ligado a ele.
  Vale conferir se a classificação está certa, se o boleto foi associado ao
  título correto, ou se o fornecedor cobrou algo que já foi adiantado.</p>

  <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;{fonte}">
    <thead><tr>{cabecalho}</tr></thead>
    <tbody>{''.join(linhas)}</tbody>
  </table>

  <p style="margin:14px 0 0 0;font-size:10pt;color:#595959;">
    Painel de Análise de Boletos —
    <a href="https://lucasabnersd-ai.github.io/ANALISE-BOLETOS/PUBLICAR/">abrir o painel</a><br>
    Aviso automático da rodada de {html.escape(carga.get('atualizado_em', ''))}.
    Base da carteira salva em {html.escape(carga.get('salva_em', ''))}.
  </p>
</div>"""


def enviar(assunto: str, corpo_html: str, para: str) -> None:
    """Manda pelo Outlook desta maquina (a mesma conta que ja esta logada).

    ⚠ Com o Outlook FECHADO o COM sobe uma instancia e a mensagem pode ficar
    parada na Caixa de Saida ate alguem abrir o programa. Por isso o
    SendAndReceive logo depois do Send: sem ele, o script dizia "enviado" e o
    e-mail so saia horas depois.
    """
    import win32com.client  # so aqui: em maquina sem Outlook o resto ainda roda

    outlook = win32com.client.Dispatch("Outlook.Application")
    email = outlook.CreateItem(0)  # 0 = olMailItem
    email.To = para
    email.Subject = assunto
    email.HTMLBody = corpo_html
    email.Send()

    try:
        namespace = outlook.GetNamespace("MAPI")
        namespace.SendAndReceive(False)
    except Exception as erro:  # noqa: BLE001 - avisar e seguir
        print(f"   AVISO: nao consegui forcar o envio ({erro}).")
        print("   Se o Outlook estiver fechado, o e-mail sai quando ele abrir.")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--para", default="")
    parser.add_argument("--teste", action="store_true",
                        help="nao envia: grava a previa em DADOS e mostra o resumo")
    parser.add_argument("--dados", type=Path, default=DADOS)
    args = parser.parse_args()

    if not args.dados.exists():
        print(f"   AVISO: nao achei {args.dados} -- rode o gerar_painel.py antes.")
        return 0

    carga = json.loads(args.dados.read_text(encoding="utf-8"))
    achados = achar(carga)

    if not achados:
        print("   Nenhum título de AMARRAR ADIANTAMENTO com boleto. Nada a avisar.")
        return 0

    destino = args.para or para_quem()
    quantos = len(achados)
    hoje = dt.datetime.now().strftime("%d/%m/%Y")
    assunto = (f"[Painel Boletos] {quantos} título{'s' if quantos > 1 else ''} de "
               f"AMARRAR ADIANTAMENTO com boleto associado - {hoje}")
    corpo = montar_html(achados, carga)

    print(f"   {quantos} título(s) de AMARRAR ADIANTAMENTO com boleto:")
    for a in achados:
        c = a["c"]
        print(f"     {c.get('campo_uuid')} | {c.get('filial')} | "
              f"tit {c.get('no_titulo')}/{c.get('parcela')} | "
              f"{(c.get('razao_social') or '')[:34]} | "
              f"título {c.get('vlr_titulo')} x boleto {c.get('valor_boleto')}")

    if args.teste:
        PREVIA.parent.mkdir(parents=True, exist_ok=True)
        PREVIA.write_text(corpo, encoding="utf-8")
        print(f"\n   MODO TESTE: nada foi enviado.")
        print(f"   Assunto: {assunto}")
        print(f"   Para:    {destino or '(ninguém -- veja o aviso abaixo)'}")
        print(f"   Prévia:  {PREVIA}")

    if not destino:
        # Sem destinatario o alerta e inutil, mas a rodada segue: o painel ja
        # esta montado e vai ser publicado do mesmo jeito.
        print(f"   AVISO: nao achei para quem mandar -- falta {DESTINATARIOS.name}")
        print("   Crie o arquivo com um e-mail por linha (ele fica fora do repositorio).")
        return 0

    if args.teste:
        return 0

    try:
        enviar(assunto, corpo, destino)
        print(f"   E-mail enviado para {destino}.")
    except Exception as erro:  # noqa: BLE001
        # Nao derruba a rodada: o painel ja esta montado e vai ser publicado.
        print(f"   AVISO: o e-mail NAO foi enviado ({erro}).")
        print("   O painel segue normalmente. Para mandar na mão:")
        print(f'       python "{Path(__file__).name}"')
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
