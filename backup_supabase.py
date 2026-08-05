# -*- coding: utf-8 -*-
r"""Backup completo do painel ANALISE BOLETOS.

Salva num .json TUDO o que esta no banco -- inclusive o que o painel NAO mostra:

  titulos      todos, inclusive os arquivados (ativo=false), com saiu_em
  marcacoes    check, tratativa, quem tratou e quando
  historico    TODA alteracao de marcacao ja feita (criou / alterou / apagou),
               que e a garantia de nao perder modificacao: mesmo desmarcar um
               check deixa linha aqui
  autorizados  quem tem acesso ao painel

O historico mora no schema `private`, que o PostgREST nao enxerga -- por isso o
dump sai pela mesma Edge Function da carga, autenticado pelo token estreito de
`.analise_boletos_token` (fora do repositorio).

Uso:
    python backup_supabase.py                 # grava em BACKUPS\...
    python backup_supabase.py --saida X.json  # grava onde voce quiser
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
import urllib.error
import urllib.request
from pathlib import Path

PASTA = Path(__file__).resolve().parent
TOKEN = PASTA / ".analise_boletos_token"
BACKUPS = PASTA / "BACKUPS"
FUNCAO = "https://pyrniqluywejmgzqkari.supabase.co/functions/v1/analise-boletos-carga"


def baixar() -> dict:
    corpo = json.dumps({"acao": "backup"}).encode("utf-8")
    pedido = urllib.request.Request(
        FUNCAO, data=corpo, method="POST",
        # header proprio: o Authorization e reservado ao JWT do Supabase
        headers={"Content-Type": "application/json",
                 "x-carga-token": TOKEN.read_text(encoding="utf-8").strip()},
    )
    with urllib.request.urlopen(pedido, timeout=180) as resposta:
        return json.loads(resposta.read().decode("utf-8"))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--saida", type=Path, default=None)
    args = parser.parse_args()

    if not TOKEN.exists():
        print(f"ERRO: falta o token em {TOKEN}", file=sys.stderr)
        return 1

    try:
        dump = baixar()
    except urllib.error.HTTPError as erro:
        print(f"ERRO {erro.code}: {erro.read().decode('utf-8', 'replace')}", file=sys.stderr)
        return 1
    if "erro" in dump:
        print(f"ERRO: {dump['erro']}", file=sys.stderr)
        return 1

    saida = args.saida or (
        BACKUPS / f"analise_boletos_backup_{dt.datetime.now():%Y-%m-%d_%H%M}.json")
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(json.dumps(dump, ensure_ascii=False, indent=1), encoding="utf-8")

    titulos = dump.get("titulos") or []
    arquivados = sum(1 for t in titulos if not t.get("ativo"))
    print(f"Backup: {saida}")
    print(f"  titulos     : {len(titulos)} ({arquivados} arquivados)")
    print(f"  marcacoes   : {len(dump.get('marcacoes') or [])}")
    print(f"  historico   : {len(dump.get('historico') or [])} alteracoes")
    print(f"  autorizados : {len(dump.get('autorizados') or [])}")
    print(f"  tamanho     : {saida.stat().st_size/1048576:.1f} MB")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
