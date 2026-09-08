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
    python alerta_adiantamento.py            # manda, no maximo uma vez por dia
    python alerta_adiantamento.py --teste    # NAO manda: grava o HTML e mostra
    python alerta_adiantamento.py --forcar   # manda de novo no mesmo dia

Codigo de saida: SEMPRE 0 quando o painel foi lido. Falha de e-mail nao pode
derrubar a rodada -- mesma regra da publicacao do BOLETOS-PENDENTES la no
associador. O que da errado sai no texto, nao no errorlevel.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path

import alertas_comuns as comuns

PASTA = Path(__file__).resolve().parent
DADOS = PASTA / "DADOS" / "analise_boletos.json"
PREVIA = PASTA / "DADOS" / "alerta_adiantamento_previa.html"
CHAVE = "adiantamento"

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

# Coluna do e-mail -> campo da linha do painel. O UUID vem PRIMEIRO de
# proposito: e por ele que ela acha o titulo no painel e no TOTVS.
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
            if str(c.get("tipo_pgto_sf1") or "").strip().upper() != TIPO_ALVO:
                continue
            if not tem_boleto(c):
                continue
            achados.append(dict(c, _aba=aba["nome"]))
    # Mesma ordem toda vez: sem isto, dois e-mails com os mesmos titulos
    # pareceriam diferentes so pela ordem em que o painel montou as abas.
    achados.sort(key=lambda a: (a.get("filial") or "", a.get("no_titulo") or "",
                                a.get("parcela") or ""))
    return achados


def montar_html(achados: list[dict], carga: dict) -> str:
    quantos = len(achados)
    frase = ("1 título classificado como <b>AMARRAR ADIANTAMENTO</b> está com boleto associado"
             if quantos == 1 else
             f"{quantos} títulos classificados como <b>AMARRAR ADIANTAMENTO</b> "
             f"estão com boleto associado")
    return f"""<div style="{comuns.FONTE}font-size:11pt;color:#000000;">
  <p style="margin:0 0 12px 0;">Bom dia, Graziela,</p>

  <p style="margin:0 0 12px 0;">{frase}.</p>

  <p style="margin:0 0 14px 0;">Título de <b>amarrar adiantamento</b> é quitado
  contra adiantamento já pago, então não era para haver boleto ligado a ele.
  Vale conferir se a classificação está certa, se o boleto foi associado ao
  título correto, ou se o fornecedor cobrou algo que já foi adiantado.</p>

  {comuns.tabela(list(COLUNAS), achados,
                 direita=("vlr_titulo", "valor_boleto"),
                 destaque=("campo_uuid",))}

  {comuns.rodape(
      f"Rodada de {carga.get('atualizado_em', '')}.",
      f"Base da carteira salva em {carga.get('salva_em', '')}.")}
</div>"""


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--para", default="")
    parser.add_argument("--teste", action="store_true",
                        help="nao envia: grava a previa e mostra o resumo")
    parser.add_argument("--forcar", action="store_true",
                        help="envia mesmo se ja mandou hoje")
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

    quantos = len(achados)
    print(f"   {quantos} título(s) de AMARRAR ADIANTAMENTO com boleto:")
    for a in achados:
        print(f"     {a.get('campo_uuid')} | {a.get('filial')} | "
              f"tit {a.get('no_titulo')}/{a.get('parcela') or '-'} | "
              f"{(a.get('razao_social') or '')[:34]} | "
              f"título {a.get('vlr_titulo')} x boleto {a.get('valor_boleto')}")

    corpo = montar_html(achados, carga)
    assunto = (f"[Painel Boletos] {quantos} título{'s' if quantos > 1 else ''} de "
               f"AMARRAR ADIANTAMENTO com boleto associado - "
               f"{dt.datetime.now():%d/%m/%Y}")

    if args.teste:
        PREVIA.parent.mkdir(parents=True, exist_ok=True)
        PREVIA.write_text(corpo, encoding="utf-8")
        print(f"\n   MODO TESTE: nada foi enviado.\n   Assunto: {assunto}")
        print(f"   Prévia:  {PREVIA}")
        return 0

    destino = args.para or comuns.destinatarios(CHAVE)
    if not destino:
        print(f"   AVISO: nao achei para quem mandar -- falta {comuns.PADRAO_DESTINO}")
        print("   Crie o arquivo com um e-mail por linha (ele fica fora do repositorio).")
        return 0

    if not args.forcar and comuns.ja_enviado_hoje(CHAVE):
        print(f"   Já enviado hoje ({comuns.quando_enviou(CHAVE)}) -- não mando de novo.")
        print("   Para mandar assim mesmo: --forcar")
        return 0

    try:
        comuns.enviar(assunto, corpo, destino)
        comuns.marcar_enviado(CHAVE, quantos)
        print(f"   E-mail enviado para {destino}.")
    except Exception as erro:  # noqa: BLE001
        # Nao derruba a rodada: o painel ja esta montado e vai ser publicado.
        print(f"   AVISO: o e-mail NAO foi enviado ({erro}).")
        print("   O painel segue normalmente. Para mandar na mão:")
        print(f'       python "{Path(__file__).name}" --forcar')
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
