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


def enxugar(carga: dict) -> list[dict]:
    """Uma linha por titulo, so com o que a visao Conferencia usa.

    O `busca` e o `o` sao remontados no navegador a partir do `c`; levar os
    tres triplicaria o texto sem ganhar nada.
    """
    usadas = {c["chave"] for c in carga["compacta"]} | {"alerta_fatura"}
    titulos = []
    for r in carga["linhas"]:
        titulos.append({
            "uuid": r["uuid"],
            "dados": {
                "c": {k: v for k, v in r["c"].items() if k in usadas and v},
                "copia": r["copia"],
                "difere": r["difere"],
                "delta": r["delta"],
                "forte": r["forte"],
                "feito": r["feito"],
                "alerta": r["alerta"],
                # o cabecalho da carteira viaja junto do 1o titulo lido; e
                # pequeno e evita uma 2a tabela so para metadados
                "_meta": {
                    "gerado_em": carga["gerado_em"],
                    "atualizado_em": carga["atualizado_em"],
                    "sem_codigo": carga["sem_codigo"],
                    "associados": carga["associados"],
                    "com_alerta": carga["com_alerta"],
                    "divergentes": carga["divergentes"],
                    "compacta": carga["compacta"],
                    "parcelas": carga["parcelas"],
                } if not titulos else None,
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
