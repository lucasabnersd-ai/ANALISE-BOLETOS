# -*- coding: utf-8 -*-
"""Gera o painel HTML da aba 'Associacoes Encontradas'.

Le TRATAMENTO PYTHON BOLETOS.xlsx e escreve um index.html unico, sem
dependencias externas, no formato de planilha: as mesmas colunas da base, na
mesma ordem, com os indicadores, os criterios do match, os alertas e a linha
digitavel de cada titulo associado.

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

# Colunas visiveis quando o painel abre. O resto continua na planilha e aparece
# no modo "todas as colunas".
ESSENCIAIS = [
    "Campo UUID", "Filial", "Prefixo", "No. Titulo", "Parcela", "Razão Social",
    "CNPJ Boleto", "NF/Doc Boleto", "Vlr.Titulo", "Valor Boleto", "Vencimento",
    "Vencto Real", "Status", "Score Match", "Margem", "Alerta Fatura",
    "Critério Match", "Linha Digitável",
]

MOEDA = {"Vlr.Titulo", "Valor Boleto", "Saldo", "Desconto", "Multa", "Juros",
         "Correcao", "Val Liq Baix"}
DATA = {"Vencimento", "Vencto Real", "Vencimento Boleto", "DT Emissao",
        "DT Baixa", "DT Contab."}
INTEIRO = {"Score Match", "2º Score", "Margem"}

# Colunas com tratamento proprio na tela (pastilha, barra, chips, botao copiar).
ESPECIAIS = {"Campo UUID": "uuid", "Status": "status", "Score Match": "score",
             "Alerta Fatura": "alerta", "Critério Match": "criterio",
             "Linha Digitável": "linha"}

LARGURAS = {
    "Campo UUID": 300, "Razão Social": 260, "Fornecedor Boleto": 210,
    "Nome Fornece": 180, "Critério Match": 300, "Alerta Fatura": 300,
    "Linha Digitável": 300, "Historico": 320, "CNPJ Boleto": 150,
    "No. Titulo": 108, "Status": 130, "Natureza": 100, "C. de Custo": 100,
    "Fonte Boleto": 140,
}
LARGURA_PADRAO = 112

# Visao de conferencia: o que o usuario compara no SE2, lado a lado.
# (rotulo, coluna do titulo, coluna do boleto, tipo de comparacao, largura)
CONFRONTOS = [
    ("Documento", "No. Titulo", "NF/Doc Boleto", "documento", 190),
    ("Fornecedor", "Razão Social", "Fornecedor Boleto", "texto", 300),
    ("Valor", "Vlr.Titulo", "Valor Boleto", "moeda", 190),
    ("Vencimento", "Vencimento", "Vencimento Boleto", "data", 180),
    ("Venc. real x boleto", "Vencto Real", "Vencimento Boleto", "data", 180),
]


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


def classificar_criterio(trecho: str) -> str:
    """Verde quando o criterio bate exato, ambar quando tem divergencia."""
    baixo = trecho.lower()
    return "atencao" if ("dif" in baixo or "aprox" in baixo or "similar" in baixo) else "ok"


def quebrar_criterios(valor) -> list[dict]:
    partes = [p.strip() for p in texto(valor).split("+") if p.strip()]
    return [{"texto": p, "tipo": classificar_criterio(p)} for p in partes]


def so_digitos(valor) -> str:
    return re.sub(r"\D", "", texto(valor)).lstrip("0")


def so_letras(valor) -> str:
    limpo = unicodedata.normalize("NFKD", texto(valor).upper())
    return re.sub(r"[^A-Z0-9]", "", "".join(c for c in limpo if not unicodedata.combining(c)))


def confrontar(origem: dict) -> list[dict]:
    """Monta os pares titulo x boleto que o usuario confere no SE2."""
    resultado = []
    for rotulo, campo_titulo, campo_boleto, tipo, _ in CONFRONTOS:
        bruto_t, bruto_b = origem.get(campo_titulo), origem.get(campo_boleto)
        par = {
            "rotulo": rotulo,
            "titulo": texto(bruto_t),
            "boleto": texto(bruto_b),
            "situacao": "vazio",
            "delta": "",
            "ordem": 0,
        }

        if tipo == "moeda":
            valor_t, valor_b = numero(bruto_t), numero(bruto_b)
            par["titulo"], par["boleto"] = moeda_br(valor_t), moeda_br(valor_b)
            if None not in (valor_t, valor_b):
                diferenca = round(valor_b - valor_t, 2)
                par["ordem"] = abs(diferenca)
                if diferenca:
                    par["situacao"] = "difere"
                    par["delta"] = ("+" if diferenca > 0 else "−") + moeda_br(abs(diferenca))
                else:
                    par["situacao"] = "igual"
                    par["delta"] = "sem diferença"

        elif tipo == "data":
            data_t, data_b = data_iso(bruto_t), data_iso(bruto_b)
            if data_t and data_b:
                dias = (dt.date.fromisoformat(data_b) - dt.date.fromisoformat(data_t)).days
                par["ordem"] = abs(dias)
                if dias:
                    par["situacao"] = "difere"
                    par["delta"] = f"{dias:+d} {'dia' if abs(dias) == 1 else 'dias'}"
                else:
                    par["situacao"] = "igual"
                    par["delta"] = "mesma data"

        elif tipo == "documento":
            if par["titulo"] and par["boleto"]:
                iguais = so_digitos(bruto_t) == so_digitos(bruto_b)
                par["situacao"] = "igual" if iguais else "difere"
                par["delta"] = "mesmo número" if iguais else "número diferente"

        else:  # texto
            if par["titulo"] and par["boleto"]:
                letras_t, letras_b = so_letras(bruto_t), so_letras(bruto_b)
                menor = min(len(letras_t), len(letras_b))
                iguais = menor >= 6 and letras_t[:menor] == letras_b[:menor]
                par["situacao"] = "igual" if iguais else "difere"
                par["delta"] = "mesmo nome" if iguais else "nome diferente"

        resultado.append(par)
    return resultado


def montar_colunas(cabecalho: list[str], usadas: set[str]) -> list[dict]:
    colunas = []
    for rotulo in cabecalho:
        if not rotulo or rotulo in IGNORAR or rotulo not in usadas:
            continue
        if rotulo in MOEDA:
            tipo = "moeda"
        elif rotulo in DATA:
            tipo = "data"
        elif rotulo in INTEIRO:
            tipo = "inteiro"
        else:
            tipo = "texto"
        colunas.append({
            "chave": chave_de(rotulo),
            "rotulo": rotulo,
            "tipo": tipo,
            "especial": ESPECIAIS.get(rotulo, ""),
            "largura": LARGURAS.get(rotulo, LARGURA_PADRAO),
            "essencial": rotulo in ESSENCIAIS,
        })
    return colunas


def ler_associacoes(wb):
    ws = wb[ABA_ASSOCIACOES]
    linhas = list(ws.iter_rows(values_only=True))
    cabecalho = [texto(c) for c in linhas[0]]

    brutas = [b for b in linhas[1:] if not all(v in (None, "") for v in b)]

    # Uma coluna 100% vazia nas 19 linhas nao ajuda ninguem a conferir.
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
            rotulo, chave = coluna["rotulo"], coluna["chave"]
            valor = origem.get(rotulo)
            if coluna["tipo"] == "moeda":
                bruto = numero(valor)
                celulas[chave], ordem[chave] = moeda_br(bruto), bruto
            elif coluna["tipo"] == "data":
                celulas[chave], ordem[chave] = texto(valor), data_iso(valor) or ""
            elif coluna["tipo"] == "inteiro":
                bruto = numero(valor)
                celulas[chave] = "" if bruto is None else str(int(bruto))
                ordem[chave] = bruto
            elif rotulo == "Linha Digitável":
                celulas[chave] = formatar_linha_digitavel(valor)
                ordem[chave] = re.sub(r"\D", "", texto(valor))
            else:
                celulas[chave] = ordem[chave] = texto(valor)

        titulo, boleto = numero(origem.get("Vlr.Titulo")), numero(origem.get("Valor Boleto"))
        registros.append({
            "c": celulas,
            "o": ordem,
            "confrontos": confrontar(origem),
            "uuid": texto(origem.get("Campo UUID")),
            "forte": "FORTE" in texto(origem.get("Status")).upper(),
            "alerta": texto(origem.get("Alerta Fatura")),
            "criterios": quebrar_criterios(origem.get("Critério Match")),
            "linha_crua": re.sub(r"\D", "", texto(origem.get("Linha Digitável"))),
            "historico": texto(origem.get("Historico")),
            "score": int(numero(origem.get("Score Match")) or 0),
            "score2": int(numero(origem.get("2º Score")) or 0),
            "valor_boleto": boleto,
            "diferenca": None if None in (titulo, boleto) else round(boleto - titulo, 2),
            "busca": " ".join(texto(v) for v in origem.values() if texto(v)).lower(),
        })

    registros.sort(key=lambda r: -r["score"])
    return colunas, registros


def ler_resumo(wb) -> dict:
    resumo = {}
    for chave, valor in wb[ABA_RESUMO].iter_rows(values_only=True):
        rotulo = texto(chave)
        if rotulo and valor is not None:
            resumo[rotulo] = valor if isinstance(valor, (int, float)) else texto(valor)
    return resumo


def montar_indicadores(resumo: dict, registros: list[dict]) -> dict:
    def do_resumo(rotulo):
        valor = resumo.get(rotulo, 0)
        return int(valor) if isinstance(valor, (int, float)) else 0

    return {
        "gerado_em": resumo.get("Gerado em") or dt.date.today().strftime("%d/%m/%Y"),
        "total_totvs": do_resumo("Total de titulos TOTVS (BOLETO SC)"),
        "com_codigo": do_resumo("Já possuem código de barras"),
        "sem_codigo": do_resumo("Títulos sem código de barras"),
        "associados": len(registros),
        "forte": sum(1 for r in registros if r["forte"]),
        "provavel": sum(1 for r in registros if not r["forte"]),
        "com_alerta": sum(1 for r in registros if r["alerta"]),
        "com_divergencia": sum(1 for r in registros if r["diferenca"]),
        "valor_associado": round(sum(r["valor_boleto"] or 0 for r in registros), 2),
        "sem_match": do_resumo("Sem match encontrado"),
        "futuros": do_resumo("Boletos futuros com NF não associada"),
    }


def gerar(base: Path, saida: Path) -> dict:
    if not base.exists():
        raise FileNotFoundError(f"Base nao encontrada: {base}")
    if not MODELO.exists():
        raise FileNotFoundError(f"Modelo do painel nao encontrado: {MODELO}")

    wb, temporaria = abrir_planilha(base)
    try:
        colunas, registros = ler_associacoes(wb)
        indicadores = montar_indicadores(ler_resumo(wb), registros)
        abas = [
            {"nome": n, "linhas": max(wb[n].max_row - 1, 0), "ativa": n == ABA_ASSOCIACOES}
            for n in wb.sheetnames if n != ABA_RESUMO
        ]
    finally:
        wb.close()
        if temporaria:
            shutil.rmtree(temporaria, ignore_errors=True)

    dados = {
        "indicadores": indicadores,
        "colunas": colunas,
        "confrontos": [
            {"rotulo": rotulo, "titulo": campo_titulo, "boleto": campo_boleto,
             "tipo": tipo, "largura": largura}
            for rotulo, campo_titulo, campo_boleto, tipo, largura in CONFRONTOS
        ],
        "linhas": registros,
        "abas": abas,
        "atualizado_em": dt.datetime.now().strftime("%d/%m/%Y às %H:%M"),
    }

    html = MODELO.read_text(encoding="utf-8").replace(
        "/*__DADOS__*/null",
        json.dumps(dados, ensure_ascii=False, separators=(",", ":")),
    )
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(html, encoding="utf-8")
    return {"indicadores": indicadores, "colunas": len(colunas)}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base", type=Path, default=BASE_PADRAO)
    parser.add_argument("--saida", type=Path, default=SAIDA_PADRAO)
    args = parser.parse_args()

    try:
        info = gerar(args.base, args.saida)
    except Exception as exc:  # noqa: BLE001 - mensagem amigavel no .cmd
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1

    ind = info["indicadores"]
    print(f"Associacoes no painel: {ind['associados']} em {info['colunas']} colunas"
          f" (fortes {ind['forte']} | provaveis {ind['provavel']}"
          f" | com alerta {ind['com_alerta']})")
    print(f"Painel gerado: {args.saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
