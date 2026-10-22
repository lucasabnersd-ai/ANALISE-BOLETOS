# -*- coding: utf-8 -*-
"""Sobe a carteira do painel para o Supabase.

Os dados NUNCA vao no HTML publicado (o Pages serve publicamente mesmo com o
repositorio privado). Quem guarda e o Postgres, atras de RLS; o painel so
recebe depois do login.

A escrita passa pela Edge Function `analise-boletos-carga`, que usa a
service_role por dentro. Aqui na maquina fica so o token estreito, em
`.analise_boletos_token` (fora do repositorio), que nao serve para mais nada.

Uso:
    python publicar_dados.py
"""

from __future__ import annotations

import json
import sys
import urllib.error
import urllib.request
from pathlib import Path

PASTA = Path(__file__).resolve().parent
DADOS = PASTA / "DADOS" / "analise_boletos.json"
TOKEN = PASTA / ".analise_boletos_token"
FUNCAO = "https://pyrniqluywejmgzqkari.supabase.co/functions/v1/analise-boletos-carga"


# Uuid reservado do cabecalho de cada aba. Nao colide com os titulos: eles sao
# UUID hexadecimal ou "BOL:<digitos>".
META = "#meta:"

# As abas que aparecem no painel e nunca ficam vazias num dia normal. Se uma
# delas faltar ou vier sem linha, a carga NAO sobe (ver main()).
OBRIGATORIAS = ("associacoes", "correspondencias", "futuros", "pendentes_itau", "pedidos")


def cabecalho_da_aba(carga: dict, aba: dict, posicao: int) -> dict:
    """O cabecalho da aba (colunas, parcelas, contagens) numa LINHA PROPRIA.

    Antes ele pegava carona no primeiro titulo da aba. Nao da mais: agora as NFs
    nao associadas ficam no banco mesmo quando saem da base, entao uma linha
    velha continuaria servindo um cabecalho antigo -- e o painel montaria as
    colunas erradas. Com uuid fixo, cada carga sobrescreve o cabecalho.
    """
    return {
        "uuid": META + aba["id"],
        "dados": {
            "aba": aba["id"],
            "_meta": {
                "id": aba["id"],
                # o painel le os titulos ordenados por uuid, entao a ordem das
                # abas tem de vir explicita -- senao quem aparece primeiro e
                # quem tiver o menor uuid
                "ordem": posicao,
                "nome": aba["nome"],
                # ⚠ Aba DESLIGADA. Precisa viajar (e nao simplesmente sumir da
                # carga) porque a lista de abas do painel sai das linhas
                # `#meta:` que estao no Postgres: uma aba que para de ser
                # enviada continua la, com o cabecalho antigo, e voltaria a
                # aparecer. Esta e a marca que a apaga da tela.
                "oculta": aba.get("oculta", False),
                # nome da aba no .xlsx exportado
                "guia": aba["guia"],
                "cor": aba["cor"],
                # quando existe, o painel abre esta aba em duas (futura/passado)
                "particoes": aba.get("particoes"),
                # veredito da SE2: quem ja tem boleto lancado ou ja foi baixado.
                # Fica AQUI, no cabecalho, porque a carga nao regrava titulo ja
                # tratado (regra 3 desta funcao) -- num campo dentro da linha,
                # o selo nunca chegaria em quem mais precisa dele.
                "se2": aba.get("se2"),
                # rotulos dos contadores da linha de resumo (mudam por aba)
                "resumo": aba.get("resumo"),
                "gerado_em": carga["gerado_em"],
                "atualizado_em": carga["atualizado_em"],
                "salva_em": carga.get("salva_em", ""),
                "sem_codigo": carga["sem_codigo"],
                "com_alerta": aba["com_alerta"],
                "divergentes": aba["divergentes"],
                # Quantas linhas a aba tem. Existe por causa do carregamento SOB
                # DEMANDA (13/08/2026): o painel desenha os botoes das abas antes
                # de ter as linhas de qualquer uma delas, e sem este numero o
                # botao nasceria com "0" e so acertaria depois do clique.
                # ⚠ E o total ANTES dos cortes do navegador (SE2, SAIU DA BASE,
                # vencimento): assim que a aba carrega, o painel recalcula.
                "total": len(aba["linhas"]),
                "compacta": aba["compacta"],
                "parcelas": aba["parcelas"],
                "pills": aba["pills"],
                "ocr": aba["ocr"],
            },
        },
    }


def enxugar(carga: dict) -> list[dict]:
    """Uma linha por titulo, so com o que a visao Conferencia usa.

    O `busca` e o `o` sao remontados no navegador a partir do `c`; levar os
    tres triplicaria o texto sem ganhar nada.
    """
    titulos = []
    for posicao, aba in enumerate(carga["abas"]):
        usadas = {c["chave"] for c in aba["compacta"]} | {"alerta_fatura"}
        titulos.append(cabecalho_da_aba(carga, aba, posicao))
        for r in aba["linhas"]:
            titulos.append({
                "uuid": r["uuid"],
                "dados": {
                    "aba": aba["id"],
                    "c": {k: v for k, v in r["c"].items() if k in usadas and v},
                    # `o` e a chave de ORDENACAO: numero cru e data ISO. Sem ele
                    # o painel ordenava pelo texto formatado ("R$ 5.427,50",
                    # "01/08/2026") e a ordem saia errada.
                    "o": {k: v for k, v in r["o"].items()
                          if k in usadas and v not in (None, "") and v != r["c"].get(k)},
                    "selo": r["selo"],
                    "ocr": r["ocr"],
                    "copia": r["copia"],
                    "difere": r["difere"],
                    "delta": r["delta"],
                    "forte": r["forte"],
                    "feito": r["feito"],
                    "alerta": r["alerta"],
                    # Sem criterio nao ha quadro para abrir. Na aba da SEFAZ isso
                    # e a regra (o selo e um flag, nao um veredito), e mandar o
                    # dicionario vazio custava 95 KB para nao dizer nada.
                    **({"match": r["match"]} if r["match"].get("itens") else {}),
                    # texto longo (Info Comp da SEFAZ, Criterio do cruzamento):
                    # nao cabe na celula e abre no quadro. So viaja onde existe.
                    **({"texto": r["texto"]} if r.get("texto") else {}),
                    # 05/09/2026. O POR QUE dos PC provaveis (aba NF x Pedido).
                    # ⚠ Esta lista e uma WHITELIST: campo novo na linha que nao
                    # entrar aqui simplesmente nao existe no painel publicado --
                    # e nada avisa. Foi assim que o icone dos provaveis apareceu no
                    # dev.html (JSON embutido) e nunca no ar. Contraparte no
                    # supa.js (`montarLinha`), que remonta a linha campo a campo.
                    **({"pcprov": r["pcprov"]}
                       if (r.get("pcprov") or {}).get("itens") or (r.get("pcprov") or {}).get("aviso")
                       else {}),
                },
            })
    return titulos


def main() -> int:
    if not DADOS.exists():
        print(f"ERRO: rode gerar_painel.py antes -- falta {DADOS}", file=sys.stderr)
        return 1
    if not TOKEN.exists():
        print(f"ERRO: falta o token em {TOKEN}", file=sys.stderr)
        return 1

    carga = json.loads(DADOS.read_text(encoding="utf-8"))

    # TRAVA (22/10/2026). Naquele dia a SC1 mudou de layout, o gerar_painel.py so
    # AVISOU ("aba de pedidos de compra nao atualizada") e montou a carga SEM a
    # aba NF x Pedido. A Edge Function arquiva o que nao vem na carga -- as ~400
    # notas da aba teriam sumido do ar. Aba visivel (nao `oculta`) faltando ou
    # vazia = nao publica: o .cmd termina com ERRO e o gatilho tenta de novo,
    # e o painel no ar continua o anterior, inteiro.
    visiveis = {}
    for aba in carga["abas"]:
        meta = aba.get("_meta") or aba
        if not meta.get("oculta"):
            visiveis[meta.get("id")] = len(aba.get("linhas") or [])
    faltam = [a for a in OBRIGATORIAS if a not in visiveis]
    vazias = [a for a, n in visiveis.items() if a in OBRIGATORIAS and n == 0]
    if faltam or vazias:
        print("ERRO: carga NAO publicada -- o painel no ar continua o anterior.", file=sys.stderr)
        if faltam:
            print(f"  aba(s) FALTANDO na carga: {', '.join(faltam)}", file=sys.stderr)
        if vazias:
            print(f"  aba(s) VAZIA(S) na carga: {', '.join(vazias)}", file=sys.stderr)
        print("  Veja os AVISOS do gerar_painel.py acima (base com layout novo, "
              "planilha aberta, arquivo faltando).", file=sys.stderr)
        return 1

    titulos = enxugar(carga)
    corpo = json.dumps({"titulos": titulos}, ensure_ascii=False).encode("utf-8")

    pedido = urllib.request.Request(
        FUNCAO, data=corpo, method="POST",
        # header proprio: o Authorization e reservado ao JWT do Supabase
        headers={"Content-Type": "application/json",
                 "x-carga-token": TOKEN.read_text(encoding="utf-8").strip()},
    )
    try:
        with urllib.request.urlopen(pedido, timeout=60) as resposta:
            saida = json.loads(resposta.read().decode("utf-8"))
    except urllib.error.HTTPError as erro:
        print(f"ERRO {erro.code}: {erro.read().decode('utf-8', 'replace')}", file=sys.stderr)
        return 1

    print(f"Enviados {len(titulos)} titulos ({len(corpo)/1024:.1f} KB).")
    print(f"Resposta: {saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
