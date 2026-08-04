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
# Fora do repositorio (.gitignore): a carteira para subir e o painel de uso local.
DADOS_JSON = PASTA / "DADOS" / "analise_boletos.json"
DEV = PASTA / "dev.html"

ABA_ASSOCIACOES = "Associações Encontradas"
ABA_RESUMO = "Resumo"
ABA_PARCELAS = "NFs Múltiplas Parcelas"

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

# Colunas com botao de copiar (o usuario cola no SE2 / no banco). O rotulo so
# aparece nas que sao SO_BOTAO; nas outras o proprio valor e o botao.
COPIAVEIS = {"Campo UUID": "UUID", "Linha Digitável": "LINHA",
             "Fornecedor": "", "No. Titulo": ""}

# Dessas o valor nem aparece: a coluna e so o botao de copiar (sao longas demais
# e o usuario nunca le, so cola).
SO_BOTAO = {"Campo UUID", "Linha Digitável"}

# Tipos de alerta que ganham selo proprio na coluna Alerta. A ordem importa:
# o texto de fatura tambem fala em "parcelas", entao FATURA e testado primeiro.
ALERTAS = ((("FATURA",), "FATURA"),
           (("MAIS DE UMA PARCELA", "OUTRAS PARCELAS", "PARCELA"), "PARCELA"))

# Confrontos que viram destaque na celula do boleto.
#   coluna destacada -> (coluna do titulo, tipo)
CONFRONTOS = {
    "Valor Boleto": ("Vlr.Titulo", "moeda"),
    "Vencimento Boleto": ("Vencimento", "data"),
}

# Visao Conferencia (unica visao do painel): subconjunto das colunas NA MESMA
# ORDEM DA BASE. Como a base ja poe titulo e boleto lado a lado, a unica coisa
# acrescentada e a coluna de diferenca (@delta_*), logo depois do par.
#   (rotulo, lado, grupo, largura relativa)
# Empresa pagadora e contraparte NAO entram aqui: elas sao colunas do drill
# (o usuario corrigiu isso em 04/08/2026). A tabela principal segue enxuta.
COMPACTA = [
    ("CHECK (FEITO)",     "",       "",           2.5),
    ("Campo UUID",        "",       "",           3.5),
    ("Filial",            "",       "",           3),
    ("Prefixo",           "",       "",           3),
    # A NF do boleto sai do lugar dela na base para ficar colada na do titulo
    # (pedido do usuario): e essa comparacao que ele faz o tempo todo.
    ("No. Titulo",        "titulo", "NF",         6.5),
    ("NF/Doc Boleto",     "boleto", "NF",         6),
    ("Parcela",           "titulo", "",           3),
    ("Fornecedor",        "titulo", "",           5),
    ("Razão Social",      "titulo", "",          13),
    ("Vlr.Titulo",        "titulo", "VALOR",      6.5),
    ("Valor Boleto",      "boleto", "VALOR",      6.5),
    ("@delta_valor",      "delta",  "VALOR",      5),
    ("Vencimento",        "titulo", "VENCIMENTO", 5.5),
    ("Vencto Real",       "titulo", "VENCIMENTO", 5.5),
    ("Vencimento Boleto", "boleto", "VENCIMENTO", 5.5),
    ("@delta_dias",       "delta",  "VENCIMENTO", 3.5),
    ("Status",            "",       "",           7),
    ("Alerta Fatura",     "",       "",           5),
    ("@tratativa",        "",       "",           3.5),
    ("Linha Digitável",   "",       "",           4),
]

# Cabecalho mais curto na Conferencia (o painel casa a coluna pela chave, nao
# pelo rotulo -- renomear aqui e seguro).
ROTULOS_COMPACTA = {"@delta_valor": "Δ valor", "@delta_dias": "Δ dias",
                    "Alerta Fatura": "Alerta", "Campo UUID": "UUID",
                    "Linha Digitável": "Linha dig.", "@tratativa": "Tratativa",
                    "No. Titulo": "Nº título", "NF/Doc Boleto": "NF boleto",
                    "Razão Social": "Razão social", "Fornecedor": "Cód.",
                    "Vlr.Titulo": "Vlr. título", "Vencto Real": "Vencto real",
                    "Vencimento Boleto": "Venc. boleto"}

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


def delta_bruto(origem: dict, coluna_boleto: str):
    """Diferenca numerica boleto - titulo (reais ou dias); None se nao da para comparar."""
    coluna_titulo, tipo = CONFRONTOS[coluna_boleto]
    if tipo == "moeda":
        valor_t, valor_b = numero(origem.get(coluna_titulo)), numero(origem.get(coluna_boleto))
        return None if None in (valor_t, valor_b) else round(valor_b - valor_t, 2)

    data_t, data_b = data_iso(origem.get(coluna_titulo)), data_iso(origem.get(coluna_boleto))
    if not (data_t and data_b):
        return None
    return (dt.date.fromisoformat(data_b) - dt.date.fromisoformat(data_t)).days


def diferenca(origem: dict, coluna_boleto: str) -> str:
    """Texto curto da divergencia entre o boleto e o titulo ('' se conferem)."""
    delta = delta_bruto(origem, coluna_boleto)
    if not delta:
        return ""
    if CONFRONTOS[coluna_boleto][1] == "moeda":
        return ("+" if delta > 0 else "−") + moeda_br(abs(delta))
    return f"{delta:+d} {'dia' if abs(delta) == 1 else 'dias'}"


def ler_associacoes(wb):
    ws = wb[ABA_ASSOCIACOES]
    linhas = list(ws.iter_rows(values_only=True))
    cabecalho = [texto(c) for c in linhas[0]]
    brutas = [b for b in linhas[1:] if not all(v in (None, "") for v in b)]
    compacta = montar_compacta(cabecalho)

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
            # deltas numericos: alimentam a visao Conferencia e a ordenacao
            "delta": {"valor": delta_bruto(origem, "Valor Boleto"),
                      "dias": delta_bruto(origem, "Vencimento Boleto")},
            "forte": "FORTE" in texto(origem.get("Status")).upper(),
            "alerta": montar_alerta(origem),
            "busca": " ".join(texto(v) for v in origem.values() if texto(v)).lower(),
        })

    registros.sort(key=lambda r: -(numero(r["o"].get("score_match")) or 0))
    return colunas, compacta, registros


def montar_compacta(cabecalho: list[str]) -> list[dict]:
    """Colunas da visao Conferencia, respeitando a ordem em que estao na base."""
    ordem_base = {rotulo: i for i, rotulo in enumerate(cabecalho) if rotulo}
    colunas = []
    for rotulo, lado, grupo, largura in COMPACTA:
        virtual = rotulo.startswith("@")
        if not virtual and rotulo not in ordem_base:
            continue  # a base mudou de nome: melhor faltar a coluna do que errar
        colunas.append({
            "chave": rotulo.lstrip("@") if virtual else chave_de(rotulo),
            "rotulo": ROTULOS_COMPACTA.get(rotulo, rotulo),
            "tipo": ("tratativa" if rotulo == "@tratativa" else "delta") if virtual
                    else tipo_da_coluna(rotulo),
            "lado": lado,
            "grupo": grupo,
            "largura": largura,
            "copiavel": rotulo in COPIAVEIS,
            "copiar": COPIAVEIS.get(rotulo, ""),
            "so_botao": rotulo in SO_BOTAO,
            "check": rotulo == COLUNA_CHECK,
        })
    return colunas


def tipo_alerta(valor) -> str:
    """PARCELA / FATURA / ALERTA -- rotulo curto do selo ('' quando nao ha alerta)."""
    texto_alerta = texto(valor).upper()
    if not texto_alerta:
        return ""
    for marcadores, rotulo in ALERTAS:
        if any(m in texto_alerta for m in marcadores):
            return rotulo
    return "ALERTA"


def montar_alerta(origem: dict) -> dict:
    """Tipo do alerta + a NF que o drill de parcelas abre.

    A NF vem do TEXTO do alerta ("... - NF 000013303: ...") e nao da coluna
    NF/Doc Boleto: quem escreveu o alerta sabia de qual NF estava falando, e a
    coluna as vezes traz dois numeros no mesmo campo.
    """
    bruto = texto(origem.get("Alerta Fatura"))
    tipo = tipo_alerta(bruto)
    if not tipo:
        return {"tipo": "", "nf": ""}
    achada = re.search(r"\bNF\s+([0-9]+)", bruto, re.IGNORECASE)
    nf = nf_chave(achada.group(1)) if achada else nf_chave(origem.get("NF/Doc Boleto"))
    # O quadro que aparece ao passar o mouse: o que a base tem para justificar
    # o alerta, sem obrigar o usuario a abrir o Excel.
    return {
        "tipo": tipo,
        "nf": nf,
        "texto": bruto,
        "criterio": texto(origem.get("Critério Match")),
        "fonte": texto(origem.get("Fonte Boleto")),
        "score": texto(origem.get("Score Match")),
        "score2": texto(origem.get("2º Score")),
        "margem": texto(origem.get("Margem")),
    }


def nf_chave(valor) -> str:
    """NF sem zeros a esquerda -- e o unico jeito de casar as duas abas."""
    return re.sub(r"\D", "", texto(valor)).lstrip("0")


def ler_parcelas(wb) -> dict[str, list[dict]]:
    """Boletos da aba 'NFs Multiplas Parcelas', agrupados por NF.

    A NF aparece em dois lugares e nem sempre no mesmo: em parte das linhas ela
    esta na coluna 'NF (9 Digitos)', em outras essa coluna vem VAZIA e a NF so
    existe dentro do 'Nosso Numero (Ref.)', no formato 0013303-01. Ignorar o
    segundo caminho deixaria a NF 13303 (3 boletos) fora do drill.
    """
    if ABA_PARCELAS not in wb.sheetnames:
        return {}

    linhas = list(wb[ABA_PARCELAS].iter_rows(values_only=True))
    if not linhas:
        return {}
    cabecalho = [texto(c) for c in linhas[0]]

    por_nf: dict[str, list[dict]] = {}
    for bruta in linhas[1:]:
        origem = dict(zip(cabecalho, bruta))
        nosso = texto(origem.get("Nosso Número (Ref.)"))
        chave = nf_chave(origem.get("NF (9 Dígitos)")) or nf_chave(nosso.split("-")[0])
        if not chave:
            continue
        digitos = re.sub(r"\D", "", texto(origem.get("Linha Digitável")))
        por_nf.setdefault(chave, []).append({
            "parcela": texto(origem.get("Parcela")),
            "venc": texto(origem.get("Data Vencimento")),
            "emissao": texto(origem.get("Data Emissão")),
            "valor": moeda_br(numero(origem.get("Valor (R$)"))),
            "valor_n": numero(origem.get("Valor (R$)")) or 0.0,
            "fornecedor": texto(origem.get("Favorecido/Fornecedor")),
            "empresa": texto(origem.get("Empresa Pagadora")),
            "nosso": nosso,
            "linha": formatar_linha_digitavel(origem.get("Linha Digitável")),
            "copia": digitos,
        })

    for boletos in por_nf.values():
        boletos.sort(key=lambda b: (b["parcela"], b["venc"]))
    return por_nf


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
        colunas, compacta, registros = ler_associacoes(wb)
        resumo = ler_resumo(wb)
        parcelas = ler_parcelas(wb)
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
        # `colunas` nao vai para o HTML: a visao Planilha saiu do painel e o
        # painel so desenha `compacta`. A lista continua servindo para montar
        # as celulas de cada registro.
        "compacta": compacta,
        "linhas": registros,
        # So as NFs que algum alerta de parcela realmente abre -- nao adianta
        # levar as 5 NFs da aba se o painel so tem 3 alertas.
        "parcelas": {nf: parcelas[nf] for nf in
                     {r["alerta"]["nf"] for r in registros
                      if r["alerta"]["tipo"] == "PARCELA" and r["alerta"]["nf"]}
                     if nf in parcelas},
    }

    carga = json.dumps(dados, ensure_ascii=False, separators=(",", ":"))
    modelo = MODELO.read_text(encoding="utf-8")

    # Publicacao: o HTML sai SEM a carteira (mesmo modelo dos outros paineis --
    # o Pages serve publicamente mesmo com o repo privado). A carteira vai para
    # um JSON a parte, que o supa.js entrega so depois do login.
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(modelo, encoding="utf-8")
    conferir_sem_dados(saida, dados)

    DADOS_JSON.parent.mkdir(parents=True, exist_ok=True)
    DADOS_JSON.write_text(carga, encoding="utf-8")

    # Uso local: dev.html com tudo embutido, fora do repositorio (.gitignore).
    DEV.write_text(modelo.replace("/*__DADOS__*/null", carga), encoding="utf-8")
    return dados


# Pedacos que so podem existir no dev.html. Se aparecerem no publicado, a
# geracao para -- e a mesma trava que o painel do Geraldo tem.
def conferir_sem_dados(arquivo: Path, dados: dict) -> None:
    html = arquivo.read_text(encoding="utf-8")
    suspeitos = []
    for registro in dados["linhas"]:
        for valor in (registro["uuid"], registro["copia"].get(chave_de("Linha Digitável"), "")):
            if valor and valor in html:
                suspeitos.append(valor)
    if "/*__DADOS__*/null" not in html:
        suspeitos.append("marcador /*__DADOS__*/null ausente")
    if suspeitos:
        arquivo.unlink(missing_ok=True)
        raise RuntimeError(
            f"{arquivo.name} sairia com dados dentro ({len(suspeitos)} ocorrencia(s), "
            f"ex.: {suspeitos[0][:40]}). Arquivo apagado -- nada foi publicado."
        )


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

    print(f"Linhas: {dados['associados']} | colunas: {len(dados['compacta'])}"
          f" | com alerta: {dados['com_alerta']} | boleto difere do titulo: {dados['divergentes']}")
    print(f"Publicavel (sem dados): {args.saida}")
    print(f"Carteira para subir:    {DADOS_JSON}")
    print(f"Painel local (com dados): {DEV}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
