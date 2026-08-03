# -*- coding: utf-8 -*-
"""Gera o painel HTML da aba 'Associacoes Encontradas'.

Le TRATAMENTO PYTHON BOLETOS.xlsx e escreve um index.html unico, sem
dependencias externas, com os indicadores, os criterios do match, os alertas
e a linha digitavel de cada titulo associado.

Uso:
    python gerar_painel.py [--base CAMINHO] [--saida CAMINHO]
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import shutil
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

BASE_PADRAO = Path(
    r"C:\Users\lucas\OneDrive - Grupo S&D\Arquivos de Gabriella Karla Oliveira Milas - FINANCEIRO COMPARTILHADO"
    r"\LUCAS ABNER ARAUJO\TRATAMENTO PYTHON BOLETOS.xlsx"
)
PASTA = Path(__file__).resolve().parent
SAIDA_PADRAO = PASTA / "PUBLICAR" / "index.html"
MODELO = PASTA / "painel_modelo.html"

ABA_ASSOCIACOES = "Associações Encontradas"
ABA_RESUMO = "Resumo"

# Coluna do Excel -> chave no JSON. O codigo de barras fica de fora de
# proposito: no painel so entra a linha digitavel.
CAMPOS = {
    "Filial": "filial",
    "Prefixo": "prefixo",
    "Tipo": "tipo",
    "No. Titulo": "titulo",
    "Parcela": "parcela",
    "Razão Social": "razao_social",
    "Fornecedor Boleto": "fornecedor_boleto",
    "CNPJ Boleto": "cnpj",
    "NF/Doc Boleto": "nf_boleto",
    "Vlr.Titulo": "valor_titulo",
    "Valor Boleto": "valor_boleto",
    "Vencimento": "vencimento",
    "Vencto Real": "vencimento_real",
    "Vencimento Boleto": "vencimento_boleto",
    "Status": "status",
    "Score Match": "score",
    "2º Score": "score2",
    "Margem": "margem",
    "Alerta Fatura": "alerta",
    "Critério Match": "criterio",
    "Fonte Boleto": "fonte",
    "Linha Digitável": "linha_digitavel",
    "Natureza": "natureza",
    "DT Emissao": "emissao",
    "Historico": "historico",
    "C. de Custo": "centro_custo",
}


def abrir_planilha(caminho: Path):
    """Abre a planilha; se estiver travada pelo Excel/OneDrive, usa uma copia."""
    try:
        return load_workbook(caminho, read_only=True, data_only=True), None
    except (PermissionError, OSError):
        pasta = tempfile.mkdtemp(prefix="painel_boletos_")
        copia = Path(pasta) / caminho.name
        shutil.copy2(caminho, copia)
        return load_workbook(copia, read_only=True, data_only=True), pasta


def texto(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, (dt.datetime, dt.date)):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return re.sub(r"\s+", " ", str(valor).replace("_x000D_", " ")).strip()


def numero(valor):
    if valor in (None, ""):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    limpo = re.sub(r"[^0-9,.\-]", "", str(valor))
    if "," in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")
    try:
        return float(limpo)
    except ValueError:
        return None


def data_iso(valor):
    if isinstance(valor, (dt.datetime, dt.date)):
        return valor.strftime("%Y-%m-%d")
    txt = texto(valor)
    for formato in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(txt, formato).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None


def formatar_linha_digitavel(valor) -> str:
    """Deixa a linha no formato lido pelo banco: 5.5 5.6 5.6 1 14."""
    digitos = re.sub(r"\D", "", texto(valor))
    if len(digitos) != 47:
        return texto(valor)
    return " ".join([
        f"{digitos[0:5]}.{digitos[5:10]}",
        f"{digitos[10:15]}.{digitos[15:21]}",
        f"{digitos[21:26]}.{digitos[26:32]}",
        digitos[32],
        digitos[33:47],
    ])


def classificar_criterio(trecho: str) -> str:
    """Verde quando o criterio bate exato, ambar quando tem divergencia."""
    baixo = trecho.lower()
    if "dif" in baixo or "aprox" in baixo or "similar" in baixo:
        return "atencao"
    return "ok"


def quebrar_criterios(valor) -> list[dict]:
    bruto = texto(valor)
    if not bruto:
        return []
    partes = [p.strip() for p in bruto.split("+") if p.strip()]
    return [{"texto": p, "tipo": classificar_criterio(p)} for p in partes]


def ler_associacoes(wb) -> list[dict]:
    ws = wb[ABA_ASSOCIACOES]
    linhas = list(ws.iter_rows(values_only=True))
    cabecalho = [texto(c) for c in linhas[0]]

    registros = []
    for bruta in linhas[1:]:
        if all(v in (None, "") for v in bruta):
            continue
        origem = dict(zip(cabecalho, bruta))
        item = {destino: texto(origem.get(coluna)) for coluna, destino in CAMPOS.items()}

        item["valor_titulo"] = numero(origem.get("Vlr.Titulo"))
        item["valor_boleto"] = numero(origem.get("Valor Boleto"))
        item["score"] = int(numero(origem.get("Score Match")) or 0)
        item["score2"] = int(numero(origem.get("2º Score")) or 0)
        item["margem"] = int(numero(origem.get("Margem")) or 0)
        item["vencimento_iso"] = data_iso(origem.get("Vencimento"))
        item["vencimento_real_iso"] = data_iso(origem.get("Vencto Real"))
        item["criterios"] = quebrar_criterios(origem.get("Critério Match"))
        item["linha_digitavel"] = formatar_linha_digitavel(origem.get("Linha Digitável"))
        item["linha_digitavel_crua"] = re.sub(r"\D", "", texto(origem.get("Linha Digitável")))

        titulo, boleto = item["valor_titulo"], item["valor_boleto"]
        item["diferenca"] = None if None in (titulo, boleto) else round(boleto - titulo, 2)

        registros.append(item)

    registros.sort(key=lambda r: (-r["score"], r["razao_social"]))
    return registros


def ler_resumo(wb) -> dict:
    resumo = {}
    for chave, valor in wb[ABA_RESUMO].iter_rows(values_only=True):
        rotulo = texto(chave)
        if rotulo and valor is not None:
            resumo[rotulo] = valor if isinstance(valor, (int, float)) else texto(valor)
    return resumo


def montar_indicadores(resumo: dict, registros: list[dict]) -> dict:
    def do_resumo(rotulo, padrao=0):
        valor = resumo.get(rotulo, padrao)
        return int(valor) if isinstance(valor, (int, float)) else padrao

    total = do_resumo("Total de titulos TOTVS (BOLETO SC)")
    com_codigo = do_resumo("Já possuem código de barras")
    sem_codigo = do_resumo("Títulos sem código de barras")
    return {
        "gerado_em": resumo.get("Gerado em") or dt.date.today().strftime("%d/%m/%Y"),
        "total_totvs": total,
        "com_codigo": com_codigo,
        "sem_codigo": sem_codigo,
        "associados": len(registros),
        "forte": sum(1 for r in registros if "FORTE" in r["status"].upper()),
        "provavel": sum(1 for r in registros if "PROV" in r["status"].upper()),
        "com_alerta": sum(1 for r in registros if r["alerta"]),
        "com_divergencia": sum(1 for r in registros if r["diferenca"]),
        "valor_associado": round(sum(r["valor_boleto"] or 0 for r in registros), 2),
        "sem_match": do_resumo("Sem match encontrado"),
        "nfs_multiplas": do_resumo("NFs com múltiplas parcelas"),
        "futuros": do_resumo("Boletos futuros com NF não associada"),
        "fontes": {
            nome: int(resumo[nome])
            for nome in ("BOLETOS_ITAU", "CENTRAL_GRAZIELA", "CENTRAL_BH", "CENTRAL_LUCAS")
            if isinstance(resumo.get(nome), (int, float))
        },
    }


def gerar(base: Path, saida: Path) -> dict:
    if not base.exists():
        raise FileNotFoundError(f"Base nao encontrada: {base}")
    if not MODELO.exists():
        raise FileNotFoundError(f"Modelo do painel nao encontrado: {MODELO}")

    wb, temporaria = abrir_planilha(base)
    try:
        registros = ler_associacoes(wb)
        indicadores = montar_indicadores(ler_resumo(wb), registros)
    finally:
        wb.close()
        if temporaria:
            shutil.rmtree(temporaria, ignore_errors=True)

    dados = {
        "indicadores": indicadores,
        "associacoes": registros,
        "atualizado_em": dt.datetime.now().strftime("%d/%m/%Y às %H:%M"),
    }

    html = MODELO.read_text(encoding="utf-8").replace(
        "/*__DADOS__*/null",
        json.dumps(dados, ensure_ascii=False, separators=(",", ":")),
    )
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(html, encoding="utf-8")
    return indicadores


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base", type=Path, default=BASE_PADRAO)
    parser.add_argument("--saida", type=Path, default=SAIDA_PADRAO)
    args = parser.parse_args()

    try:
        indicadores = gerar(args.base, args.saida)
    except Exception as exc:  # noqa: BLE001 - mensagem amigavel no .cmd
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1

    print(f"Associacoes no painel: {indicadores['associados']}"
          f" (fortes {indicadores['forte']} | provaveis {indicadores['provavel']}"
          f" | com alerta {indicadores['com_alerta']})")
    print(f"Painel gerado: {args.saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
