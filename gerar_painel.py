# -*- coding: utf-8 -*-
"""Gera o painel HTML da aba 'Associacoes Encontradas'.

Le TRATAMENTO PYTHON BOLETOS.xlsx e escreve um index.html unico, sem
dependencias externas: a mesma planilha, com as mesmas colunas e na mesma
ordem. O painel so acrescenta o que a planilha nao consegue mostrar -- o
destaque quando o boleto diverge do titulo e o botao de copiar UUID e linha
digitavel.

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
import unicodedata
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

# So a linha digitavel entra; o codigo de barras fica de fora (pedido do usuario).
IGNORAR = {"Código de Barras"}

# Coluna de marcacao: fica no painel mesmo vazia na planilha, virando caixa de
# selecao. O que ja estiver preenchido na base entra marcado.
COLUNA_CHECK = "CHECK (FEITO)"

MOEDA = {"Vlr.Titulo", "Valor Boleto", "Saldo", "Desconto", "Multa", "Juros",
         "Correcao", "Val Liq Baix"}
DATA = {"Vencimento", "Vencto Real", "Vencimento Boleto", "DT Emissao",
        "DT Baixa", "DT Contab."}
INTEIRO = {"Score Match", "2º Score", "Margem"}

# Colunas com botao de copiar (o usuario cola no SE2 / no banco).
COPIAVEIS = {"Campo UUID": "UUID", "Linha Digitável": "copiar"}

# Confrontos que viram destaque na celula do boleto.
#   coluna destacada -> (coluna do titulo, tipo)
CONFRONTOS = {
    "Valor Boleto": ("Vlr.Titulo", "moeda"),
    "Vencimento Boleto": ("Vencimento", "data"),
}

LARGURAS = {
    "Campo UUID": 296, "Razão Social": 250, "Fornecedor Boleto": 200,
    "Nome Fornece": 175, "Critério Match": 280, "Alerta Fatura": 280,
    "Linha Digitável": 290, "Historico": 300, "CNPJ Boleto": 145,
    "No. Titulo": 104, "Status": 122, "Fonte Boleto": 135,
}
LARGURA_PADRAO = 108


def chave_de(rotulo: str) -> str:
    texto = unicodedata.normalize("NFKD", rotulo)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"_+", "_", re.sub(r"[^0-9a-z]+", "_", texto.lower())).strip("_")


def abrir_planilha(caminho: Path):
    """Abre a planilha; se estiver travada pelo Excel/OneDrive, usa uma copia."""
    try:
        return load_workbook(caminho, read_only=True, data_only=True), None
    except (PermissionError, OSError):
        pasta = tempfile.mkdtemp(prefix="painel_boletos_")
        shutil.copy2(caminho, Path(pasta) / caminho.name)
        return load_workbook(Path(pasta) / caminho.name, read_only=True, data_only=True), pasta


def texto(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, (dt.datetime, dt.date)):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    limpo = re.sub(r"\s+", " ", str(valor).replace("_x000D_", " ")).strip()
    return "" if limpo in ("/ /", "//") else limpo


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
    for formato in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(texto(valor), formato).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None


def moeda_br(valor) -> str:
    if valor is None:
        return ""
    inteiro = f"{valor:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return f"R$ {inteiro}"


def formatar_linha_digitavel(valor) -> str:
    """Deixa a linha no formato lido pelo banco: 5.5 5.6 5.6 1 14."""
    digitos = re.sub(r"\D", "", texto(valor))
    if len(digitos) != 47:
        return texto(valor)
    return " ".join([
        f"{digitos[0:5]}.{digitos[5:10]}", f"{digitos[10:15]}.{digitos[15:21]}",
        f"{digitos[21:26]}.{digitos[26:32]}", digitos[32], digitos[33:47],
    ])


def tipo_da_coluna(rotulo: str) -> str:
    if rotulo in MOEDA:
        return "moeda"
    if rotulo in DATA:
        return "data"
    if rotulo in INTEIRO:
        return "inteiro"
    return "texto"


def montar_colunas(cabecalho: list[str], usadas: set[str]) -> list[dict]:
    return [
        {
            "chave": chave_de(rotulo),
            "rotulo": rotulo,
            "tipo": tipo_da_coluna(rotulo),
            "copiar": COPIAVEIS.get(rotulo, ""),
            "check": rotulo == COLUNA_CHECK,
            "largura": 58 if rotulo == COLUNA_CHECK else LARGURAS.get(rotulo, LARGURA_PADRAO),
        }
        for rotulo in cabecalho
        if rotulo and rotulo not in IGNORAR and (rotulo in usadas or rotulo == COLUNA_CHECK)
    ]


def diferenca(origem: dict, coluna_boleto: str) -> str:
    """Texto curto da divergencia entre o boleto e o titulo ('' se conferem)."""
    coluna_titulo, tipo = CONFRONTOS[coluna_boleto]
    if tipo == "moeda":
        valor_t, valor_b = numero(origem.get(coluna_titulo)), numero(origem.get(coluna_boleto))
        if None in (valor_t, valor_b):
            return ""
        delta = round(valor_b - valor_t, 2)
        return ("+" if delta > 0 else "−") + moeda_br(abs(delta)) if delta else ""

    data_t, data_b = data_iso(origem.get(coluna_titulo)), data_iso(origem.get(coluna_boleto))
    if not (data_t and data_b):
        return ""
    dias = (dt.date.fromisoformat(data_b) - dt.date.fromisoformat(data_t)).days
    return f"{dias:+d} {'dia' if abs(dias) == 1 else 'dias'}" if dias else ""


def ler_associacoes(wb):
    ws = wb[ABA_ASSOCIACOES]
    linhas = list(ws.iter_rows(values_only=True))
    cabecalho = [texto(c) for c in linhas[0]]
    brutas = [b for b in linhas[1:] if not all(v in (None, "") for v in b)]

    # Uma coluna 100% vazia nas linhas associadas nao ajuda ninguem a conferir.
    usadas = {
        rotulo for i, rotulo in enumerate(cabecalho)
        if rotulo and any(texto(b[i]) for b in brutas)
    }
    colunas = montar_colunas(cabecalho, usadas)

    registros = []
    for bruta in brutas:
        origem = dict(zip(cabecalho, bruta))
        celulas, ordem = {}, {}

        for coluna in colunas:
            rotulo, chave, tipo = coluna["rotulo"], coluna["chave"], coluna["tipo"]
            valor = origem.get(rotulo)
            if tipo == "moeda":
                bruto = numero(valor)
                celulas[chave], ordem[chave] = moeda_br(bruto), bruto
            elif tipo == "data":
                celulas[chave], ordem[chave] = texto(valor), data_iso(valor) or ""
            elif tipo == "inteiro":
                bruto = numero(valor)
                celulas[chave] = "" if bruto is None else str(int(bruto))
                ordem[chave] = bruto
            elif rotulo == "Linha Digitável":
                celulas[chave] = formatar_linha_digitavel(valor)
                ordem[chave] = re.sub(r"\D", "", texto(valor))
            else:
                celulas[chave] = ordem[chave] = texto(valor)

        registros.append({
            "c": celulas,
            "o": ordem,
            "uuid": texto(origem.get("Campo UUID")),
            "feito": bool(texto(origem.get(COLUNA_CHECK))),
            # valor que o botao copia (o exibido tem pontuacao de leitura)
            "copia": {chave_de("Linha Digitável"): re.sub(r"\D", "", texto(origem.get("Linha Digitável")))},
            "difere": {chave_de(col): diferenca(origem, col) for col in CONFRONTOS
                       if diferenca(origem, col)},
            "forte": "FORTE" in texto(origem.get("Status")).upper(),
            "busca": " ".join(texto(v) for v in origem.values() if texto(v)).lower(),
        })

    registros.sort(key=lambda r: -(numero(r["o"].get("score_match")) or 0))
    return colunas, registros


def ler_resumo(wb) -> dict:
    resumo = {}
    for chave, valor in wb[ABA_RESUMO].iter_rows(values_only=True):
        rotulo = texto(chave)
        if rotulo and valor is not None:
            resumo[rotulo] = valor if isinstance(valor, (int, float)) else texto(valor)
    return resumo


def gerar(base: Path, saida: Path) -> dict:
    if not base.exists():
        raise FileNotFoundError(f"Base nao encontrada: {base}")
    if not MODELO.exists():
        raise FileNotFoundError(f"Modelo do painel nao encontrado: {MODELO}")

    wb, temporaria = abrir_planilha(base)
    try:
        colunas, registros = ler_associacoes(wb)
        resumo = ler_resumo(wb)
        abas = [
            {"nome": n, "linhas": max(wb[n].max_row - 1, 0), "ativa": n == ABA_ASSOCIACOES}
            for n in wb.sheetnames if n != ABA_RESUMO
        ]
    finally:
        wb.close()
        if temporaria:
            shutil.rmtree(temporaria, ignore_errors=True)

    def do_resumo(rotulo):
        valor = resumo.get(rotulo, 0)
        return int(valor) if isinstance(valor, (int, float)) else 0

    dados = {
        "gerado_em": resumo.get("Gerado em") or dt.date.today().strftime("%d/%m/%Y"),
        "atualizado_em": dt.datetime.now().strftime("%d/%m/%Y às %H:%M"),
        "sem_codigo": do_resumo("Títulos sem código de barras"),
        "associados": len(registros),
        "com_alerta": sum(1 for r in registros if r["c"].get("alerta_fatura")),
        "divergentes": sum(1 for r in registros if r["difere"]),
        "colunas": colunas,
        "linhas": registros,
        "abas": abas,
    }

    html = MODELO.read_text(encoding="utf-8").replace(
        "/*__DADOS__*/null",
        json.dumps(dados, ensure_ascii=False, separators=(",", ":")),
    )
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(html, encoding="utf-8")
    return dados


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base", type=Path, default=BASE_PADRAO)
    parser.add_argument("--saida", type=Path, default=SAIDA_PADRAO)
    args = parser.parse_args()

    try:
        dados = gerar(args.base, args.saida)
    except Exception as exc:  # noqa: BLE001 - mensagem amigavel no .cmd
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1

    print(f"Linhas: {dados['associados']} | colunas: {len(dados['colunas'])}"
          f" | com alerta: {dados['com_alerta']} | boleto difere do titulo: {dados['divergentes']}")
    print(f"Painel gerado: {args.saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
