# -*- coding: utf-8 -*-
"""Titulo com BORDERO ja emitido (hoje ou antes) e SEM data de baixa na SE2.

O bordero saiu -- o pagamento foi mandado para o banco -- mas a baixa nunca
voltou para o titulo. Ou o pagamento nao aconteceu, ou aconteceu e ninguem deu
baixa: nos dois casos o titulo fica figurando em aberto, entra de novo em
relatorio de pendencia e pode virar pagamento em dobro. O e-mail so mostra o
fato; quem olha decide qual dos dois lados esta errado.

  o recorte   `Dt. Bordero` preenchida e MENOR OU IGUAL a hoje, e `DT Baixa`
              vazia. Foi o pedido dele, ao pe da letra, em 18/09/2026: "data de
              bordero do dia atual para tras, e sem data de baixa".
  o corte     bordero a partir de 01/01/2026. Pedido dele em 18/09/2026,
              depois do primeiro disparo: "traga somente as datas de bordero a
              partir de 01/01/2026, descarte 2025". Sairam os 13 titulos de
              2025 que estavam na primeira lista.
  o detalhe   `Tipo Pgto`, `Vencimento` e `Vencto Real` em coluna, e o CODIGO
              (UUID) mais a LINHA DIGITAVEL numa faixa embaixo da linha
              (pedido dele: "coloque o codigo UUID" e "se for boleto mande
              linha digitavel").

Por que UUID e linha digitavel nao sao COLUNA: sao 36 e 47 caracteres numa
palavra so, que nao quebram. Cada um como coluna empurra a tabela para fora da
tela -- medido no alerta BOLETO S/C em 18/09/2026, e medido de novo aqui: a
tabela sem eles fica em 869px. Na faixa de largura inteira os dois quebram
sozinhos, continuam INTEIROS para copiar, e a tabela nao mexe.

Ela sai CRUA, sem mascara de pontos e espacos: o uso e' copiar e colar no banco,
e separador inventado aqui pode ser recusado do outro lado.

A MESMA lista vai ANEXADA em Excel (uma linha por titulo, valores em coluna),
com duas colunas a mais que o e-mail: Vlr. Titulo e Cod. Barras. A planilha
fica em DADOS/ -- que esta no .gitignore, e por isso nao vai para o repositorio
publico junto com fornecedor e valor.

Vai para o Lucas e a Graziela (`.alerta_bordero_para`, fora do repo).

Uso:
    python alerta_bordero.py                 # manda, no maximo uma vez por dia
    python alerta_bordero.py --teste         # nao manda: grava a previa
    python alerta_bordero.py --forcar        # manda de novo no mesmo dia
    python alerta_bordero.py --desde 01/01/2025  # muda o corte
    python alerta_bordero.py --base "C:\\caminho\\SE2.xlsx"

Codigo de saida: SEMPRE 0 quando conseguiu ler a base. Falha de e-mail nao
derruba a rodada do painel.
"""

from __future__ import annotations

import argparse
import datetime as dt
import itertools
import shutil
import sys
import tempfile
from pathlib import Path

import alertas_comuns as comuns
import caminhos

PASTA = Path(__file__).resolve().parent
PREVIA = PASTA / "DADOS" / "alerta_bordero_previa.html"
CHAVE = "bordero"

RAIZ = caminhos.raiz_lucas()
BASES = caminhos.bases_genericos()

ABA = "SE2"

# O painel da SE2 no ar: o numero do titulo na tabela abre ele. Endereco da
# RAIZ porque o painel nao le filtro pela URL (conferido em 18/09/2026).
PAINEL_SE2 = "https://se2-lucas.vercel.app/"

# O corte por baixo: bordero a partir daqui. Pedido dele em 18/09/2026,
# "descarte 2025". Data ESCRITA, e nao "ano corrente" calculado: virando
# 2027 um ano corrente esvaziaria o e-mail sozinho e calado, e o alerta
# passaria meses dizendo que nao ha nada.
DESDE_PADRAO = dt.date(2026, 1, 1)

# Conferidas pelo NOME no cabecalho, nunca por posicao: o TOTVS ja inseriu
# coluna no meio da exportacao antes, e indice fixo erraria tudo em silencio.
COLUNAS = {
    "filial": "Filial", "prefixo": "Prefixo", "tipo": "Tipo",
    "num": "No. Titulo", "parcela": "Parcela", "fornecedor": "Fornecedor",
    "nome": "Nome Fornece", "razao": "Razão Social",
    "vencimento": "Vencimento", "real": "Vencto Real",
    "valor": "Vlr.Titulo", "baixa": "DT Baixa", "saldo": "Saldo",
    "tipo_pgto": "Tipo Pgto", "linha_dig": "Linha Dig.",
    "cod_barras": "Cod.Barras", "uuid": "Campo UUID",
    "dt_bordero": "Dt. Bordero", "num_bordero": "Num Bordero",
}


def achar_base() -> Path | None:
    """A SE2 mais recente. O nome tem cedilha e til de verdade -- por isso o
    curinga, e nao o nome escrito aqui."""
    if not BASES.is_dir():
        return None
    achadas = [p for p in BASES.glob("SE2 - POSI*O DIARIA.xlsx")
               if "backup" not in p.name.lower() and "copia" not in p.name.lower()]
    return max(achadas, key=lambda p: p.stat().st_mtime) if achadas else None


def como_data(valor):
    if isinstance(valor, dt.datetime):
        return valor.date()
    if isinstance(valor, dt.date):
        return valor
    if isinstance(valor, str):
        for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
            try:
                return dt.datetime.strptime(valor.strip(), formato).date()
            except ValueError:
                pass
    return None


def como_numero(valor) -> float:
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        texto = valor.strip().replace(".", "").replace(",", ".")
        try:
            return float(texto)
        except ValueError:
            return 0.0
    return 0.0


def reais(valor: float) -> str:
    return "R$ " + f"{valor:,.2f}".replace(",", "~").replace(".", ",").replace("~", ".")


def texto(valor) -> str:
    """Texto util da celula. O TOTVS devolve o campo vazio de varios jeitos --
    None, a string "None", data em branco "  /  /    " -- e todos significam a
    mesma coisa: nao tem nada ali."""
    if valor is None:
        return ""
    limpo = str(valor).strip()
    if limpo.lower() in ("none", "nan", "null"):
        return ""
    if limpo and set(limpo) <= set(" /:-"):
        return ""
    return limpo


def ler_se2(caminho: Path, hoje: dt.date,
            desde: dt.date) -> tuple[list[dict], dict]:
    """Varre a SE2 e devolve os titulos com bordero emitido e sem baixa.

    ⚠ A SE2 costuma estar ABERTA no Excel na hora da rodada, e ai o openpyxl
    leva PermissionError. Copiar para a pasta temporaria e ler a copia resolve
    -- e' leitura, nao muda nada no original.
    """
    import openpyxl

    origem = caminho
    temporario = None
    try:
        arquivo = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    except PermissionError:
        temporario = Path(tempfile.gettempdir()) / f"se2_bordero_{dt.date.today():%Y%m%d}.xlsx"
        shutil.copy2(caminho, temporario)
        arquivo = openpyxl.load_workbook(temporario, read_only=True, data_only=True)
        print("   (a SE2 estava aberta; li uma cópia)")

    if ABA not in arquivo.sheetnames:
        arquivo.close()
        raise KeyError(f'a planilha nao tem a aba "{ABA}" (tem: {arquivo.sheetnames})')
    pagina = arquivo[ABA]
    linhas = pagina.iter_rows(values_only=True)
    cabecalho = [str(c).strip() if c is not None else "" for c in next(linhas)]

    onde = {}
    faltando = []
    for chave, rotulo in COLUNAS.items():
        if rotulo in cabecalho:
            onde[chave] = cabecalho.index(rotulo)
        else:
            faltando.append(rotulo)
    # Sem estas o alerta nao existe. Faltando uma, ERRA ALTO em vez de mandar
    # e-mail vazio: "nenhum titulo" e "nao consegui olhar" sao coisas
    # diferentes, e so a segunda precisa de conserto.
    for obrigatoria in ("dt_bordero", "baixa", "num", "tipo_pgto"):
        if obrigatoria not in onde:
            arquivo.close()
            raise KeyError("colunas ausentes na SE2: " + ", ".join(faltando))

    def campo(linha, chave):
        i = onde.get(chave)
        return linha[i] if i is not None and i < len(linha) else None

    achados = []
    resumo = {"linhas": 0, "com_bordero": 0, "ate_hoje": 0, "com_linha_dig": 0,
              "antes_do_corte": 0}
    for linha in linhas:
        if campo(linha, "num") in (None, ""):
            continue
        resumo["linhas"] += 1

        data_bordero = como_data(campo(linha, "dt_bordero"))
        if not data_bordero:
            continue
        resumo["com_bordero"] += 1
        if data_bordero > hoje:
            continue
        resumo["ate_hoje"] += 1
        if como_data(campo(linha, "baixa")):
            continue
        # O corte por baixo vem DEPOIS da baixa, de proposito: assim
        # "antes_do_corte" conta so o que ficou de fora por ser velho, e nao o
        # que ja estava baixado. E' esse numero que vai no rodape do e-mail --
        # ele precisa saber QUANTO esta sendo escondido pelo corte.
        if data_bordero < desde:
            resumo["antes_do_corte"] += 1
            continue

        # ⚠ Saldo NAO entra no filtro, de proposito: ele pediu "sem data de
        # baixa", e so isso. Titulo com saldo zerado e sem baixa e' exatamente
        # o tipo de bagunca que este alerta existe para mostrar. O saldo vai
        # como COLUNA, para quem le julgar.
        saldo = como_numero(campo(linha, "saldo"))
        linha_dig = texto(campo(linha, "linha_dig"))
        if linha_dig:
            resumo["com_linha_dig"] += 1

        vencimento = como_data(campo(linha, "vencimento"))
        real = como_data(campo(linha, "real"))
        achados.append({
            "uuid": texto(campo(linha, "uuid")),
            "filial": texto(campo(linha, "filial")),
            "prefixo": texto(campo(linha, "prefixo")),
            "tipo": texto(campo(linha, "tipo")),
            "titulo": texto(campo(linha, "num")),
            "parcela": texto(campo(linha, "parcela")),
            "titulo_parcela": (texto(campo(linha, "num"))
                               + ("/" + texto(campo(linha, "parcela"))
                                  if texto(campo(linha, "parcela")) else "")),
            "fornecedor": texto(campo(linha, "razao")) or texto(campo(linha, "nome")),
            "bordero": f"{data_bordero:%d/%m/%Y}",
            "num_bordero": texto(campo(linha, "num_bordero")),
            "tipo_pgto": texto(campo(linha, "tipo_pgto")),
            "vencimento": f"{vencimento:%d/%m/%Y}" if vencimento else "",
            "real": f"{real:%d/%m/%Y}" if real else "",
            "valor": reais(como_numero(campo(linha, "valor"))),
            "saldo": reais(saldo),
            "linha_dig": linha_dig,
            "cod_barras": texto(campo(linha, "cod_barras")),
            "_data": data_bordero, "_saldo": saldo,
            # crus para a planilha: no Excel data tem de ser data e valor tem
            # de ser numero, senao nao soma nem ordena
            "_valor": como_numero(campo(linha, "valor")),
            "_venc": vencimento, "_real": real,
        })
    arquivo.close()
    if temporario and temporario.exists():
        try:
            temporario.unlink()
        except OSError:
            pass
    resumo["base"] = origem
    resumo["salva_em"] = dt.datetime.fromtimestamp(origem.stat().st_mtime)
    # O bordero mais ANTIGO primeiro: e' o que esta parado ha mais tempo.
    # Dentro do dia, por NUMERO de bordero -- num mesmo dia saem varios (17/09
    # teve quatro), e e' por bordero que se confere com o banco. So dentro do
    # mesmo bordero e' que vale o maior saldo primeiro.
    achados.sort(key=lambda a: (a["_data"], a["num_bordero"], -a["_saldo"]))
    return achados, resumo


# ⚠ Largura conferida no navegador (18/09/2026, 61 titulos): 966px na primeira
# versao -- larga demais para o painel de leitura do Outlook. Duas mudancas
# derrubaram para ~830px sem tirar nada do que ele pediu:
#   . titulo e parcela numa coluna so ("000012576/04"), como se fala do titulo;
#   . "Tipo pgto" FORA do destaque, porque destaque = negrito + nao quebra, e
#     "AMARRAR ADIANTAMENTO" sozinho segurava 188px de coluna.
# 18/09, agrupamento por data (pedido dele): a coluna "Dt. borderô" SAIU -- a
# data virou o titulo do grupo, e repeti-la em cada linha seria gastar 91px
# dizendo o que a faixa logo acima ja diz. O NUMERO do bordero ficou: no mesmo
# dia saem borderos diferentes (em 17/09 sairam 000481, 000482, 000500, 000501).
COLUNAS_EMAIL = [
    ("Borderô", "num_bordero"),
    ("Filial", "filial"),
    ("Nº título", "titulo_parcela"),
    ("Fornecedor", "fornecedor"),
    ("Tipo pgto", "tipo_pgto"),
    ("Vencimento", "vencimento"),
    ("Vencto real", "real"),
    ("Saldo em aberto", "saldo"),
]


def faixa_detalhe(linha: dict) -> str:
    """A faixa de largura inteira embaixo de cada linha da tabela.

    Leva o CODIGO (UUID) -- pedido dele em 18/09/2026 -- e, quando o titulo
    tem boleto, a LINHA DIGITAVEL. Os dois aqui embaixo, e nao em coluna,
    porque sao palavras de 36 e 47 caracteres que nao quebram: cada uma como
    coluna joga o resto da tabela para fora da tela. Na faixa eles quebram
    sozinhos e continuam inteiros para copiar.
    """
    partes = []
    if linha.get("uuid"):
        partes.append("Código (UUID): " + linha["uuid"])
    if linha.get("linha_dig"):
        partes.append("Linha digitável: " + linha["linha_dig"])
    return "     ·     ".join(partes)


COLUNAS_EXCEL = [
    ("Dt. Borderô", "_data", "data"),
    ("Nº Borderô", "num_bordero", "texto"),
    ("Dias s/ baixa", "_dias", "inteiro"),
    ("Filial", "filial", "texto"),
    ("Prefixo", "prefixo", "texto"),
    ("Tipo", "tipo", "texto"),
    ("Nº Título", "titulo", "texto"),
    ("Parcela", "parcela", "texto"),
    ("Fornecedor", "fornecedor", "texto"),
    ("Tipo Pgto", "tipo_pgto", "texto"),
    ("Vencimento", "_venc", "data"),
    ("Vencto Real", "_real", "data"),
    ("Vlr. Título", "_valor", "moeda"),
    ("Saldo", "_saldo", "moeda"),
    ("Linha Digitável", "linha_dig", "texto"),
    ("Cod. Barras", "cod_barras", "texto"),
    ("Código (UUID)", "uuid", "texto"),
]

LARGURA = {"Dt. Borderô": 12, "Nº Borderô": 11, "Dias s/ baixa": 13,
           "Filial": 7, "Prefixo": 9,
           "Tipo": 7, "Nº Título": 13, "Parcela": 9, "Fornecedor": 38,
           # 25 e nao 22: "AMARRAR ADIANTAMENTO" em negrito nao cabia em 22 e
           # saia cortado no papel (conferido no PDF em 21/09/2026)
           "Tipo Pgto": 25, "Vencimento": 12, "Vencto Real": 12,
           "Vlr. Título": 14, "Saldo": 14, "Linha Digitável": 50,
           "Cod. Barras": 48, "Código (UUID)": 38}

CENTRALIZADAS = {"Dt. Borderô", "Nº Borderô", "Dias s/ baixa", "Filial",
                 "Prefixo", "Tipo", "Parcela", "Vencimento", "Vencto Real"}

# A MESMA paleta do corpo do e-mail, de proposito: quem abre o anexo tem de
# reconhecer na hora a lista que acabou de ler. Azul do cabecalho, cinza-azulado
# da faixa de grupo, zebrado e borda sao os mesmos codigos de alertas_comuns.
AZUL = "1F3864"
FAIXA_GRUPO = "D6DCE4"
ZEBRA = "F2F2F2"
BORDA = "BFBFBF"

# Fundo e letra por tipo de pagamento. A cor nao decora: e' por ela que se ve
# de longe que um bordero inteiro saiu como TRANSFERENCIA e um unico titulo no
# meio esta como BOLETO. Casado por PEDACO do nome porque o TOTVS escreve
# "BOLETO S/C", "TRANSFERENCIA", "TRANSF. ENTRE CONTAS", "AMARRAR ADIANTAMENTO".
CORES_TIPO_PGTO = (
    ("BOLETO S/C", "FFF2CC", "7F6000"),
    ("BOLETO", "E2EFDA", "375623"),
    ("ADIANT", "FCE4D6", "833C0C"),
    ("PIX", "E4DFEC", "5F497A"),
    ("TRANSF", "DDEBF7", "1F4E79"),
    ("DEBITO", "DDEBF7", "1F4E79"),
    ("DÉBITO", "DDEBF7", "1F4E79"),
)


def cor_do_tipo(tipo_pgto: str) -> tuple[str, str]:
    """(fundo, letra) do tipo de pagamento; cinza neutro para o desconhecido."""
    alvo = (tipo_pgto or "").upper()
    for pedaco, fundo, letra in CORES_TIPO_PGTO:
        if pedaco in alvo:
            return fundo, letra
    return "EDEDED", "3B3838"


def montar_resumo(arquivo, achados: list[dict], hoje: dt.date) -> None:
    """Uma segunda aba com as duas contas que se faz de cabeca ao abrir a lista:
    quanto parou em cada DATA de bordero e quanto parou em cada TIPO DE
    PAGAMENTO. So numero -- nenhuma instrucao, nenhum recado.
    """
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    pagina = arquivo.create_sheet("RESUMO")
    fio = Side(style="thin", color=BORDA)
    grade = Border(left=fio, right=fio, top=fio, bottom=fio)
    fundo_cab = PatternFill("solid", fgColor=AZUL)
    letra_cab = Font(name="Calibri", bold=True, color="FFFFFF")
    fundo_faixa = PatternFill("solid", fgColor=FAIXA_GRUPO)
    letra_faixa = Font(name="Calibri", bold=True, color=AZUL)

    # a coluna dos numeros de bordero e' larga porque num dia saem varios:
    # em 18/09 sairam cinco, e cortado no meio o numero nao serve para nada
    larguras = [18, 30, 10, 16, 16, 12, 12]
    for coluna, largura in enumerate(larguras, start=1):
        pagina.column_dimensions[get_column_letter(coluna)].width = largura

    def cabecalho(linha: int, rotulos: list[str]) -> int:
        for coluna, rotulo in enumerate(rotulos, start=1):
            celula = pagina.cell(linha, coluna, rotulo)
            celula.fill, celula.font, celula.border = fundo_cab, letra_cab, grade
            celula.alignment = Alignment(vertical="center", horizontal="center")
        pagina.row_dimensions[linha].height = 20
        return linha + 1

    def numeros(linha: int, valores: list, formatos: list[str],
                negrito: bool = False, fundo=None) -> int:
        for coluna, (valor, formato) in enumerate(zip(valores, formatos), start=1):
            celula = pagina.cell(linha, coluna, valor)
            celula.font = (letra_faixa if negrito else Font(name="Calibri"))
            celula.border = grade
            celula.number_format = formato
            if fundo is not None:
                celula.fill = fundo
            celula.alignment = Alignment(
                vertical="center",
                horizontal="center" if formato != '#,##0.00' else "right")
        return linha + 1

    linha = cabecalho(1, ["Dt. Borderô", "Nº Borderô", "Títulos",
                          "Vlr. Título", "Saldo", "Com boleto", "Dias parado"])
    for chave, grupo in itertools.groupby(achados, key=lambda a: a["bordero"]):
        do_grupo = list(grupo)
        numeros_bordero = []
        for l in do_grupo:
            if l["num_bordero"] and l["num_bordero"] not in numeros_bordero:
                numeros_bordero.append(l["num_bordero"])
        linha = numeros(linha, [
            chave, ", ".join(numeros_bordero) or "—", len(do_grupo),
            sum(l["_valor"] for l in do_grupo),
            sum(l["_saldo"] for l in do_grupo),
            sum(1 for l in do_grupo if l["linha_dig"]),
            (hoje - do_grupo[0]["_data"]).days,
        ], ["@", "@", "0", '#,##0.00', '#,##0.00', "0", "0"])
        pagina.cell(linha - 1, 2).alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True)
    linha = numeros(linha, [
        "TOTAL", "", len(achados),
        sum(a["_valor"] for a in achados), sum(a["_saldo"] for a in achados),
        sum(1 for a in achados if a["linha_dig"]), "",
    ], ["@", "@", "0", '#,##0.00', '#,##0.00', "0", "@"],
        negrito=True, fundo=fundo_faixa)

    linha += 2
    # o nome do tipo ocupa as duas primeiras colunas (A:B), senao "AMARRAR
    # ADIANTAMENTO" nao cabe e a coluna B fica um buraco azul no cabecalho
    cabeca_tipo = linha
    linha = cabecalho(linha, ["Tipo Pgto", "", "Títulos", "Vlr. Título",
                              "Saldo", "Com boleto", ""])
    pagina.merge_cells(start_row=cabeca_tipo, start_column=1,
                       end_row=cabeca_tipo, end_column=2)
    por_tipo: dict[str, list[dict]] = {}
    for a in achados:
        por_tipo.setdefault(a["tipo_pgto"] or "—", []).append(a)
    for tipo, do_tipo in sorted(por_tipo.items(),
                                key=lambda par: -sum(l["_saldo"] for l in par[1])):
        fundo, letra = cor_do_tipo(tipo)
        linha = numeros(linha, [
            tipo, "", len(do_tipo),
            sum(l["_valor"] for l in do_tipo), sum(l["_saldo"] for l in do_tipo),
            sum(1 for l in do_tipo if l["linha_dig"]), "",
        ], ["@", "@", "0", '#,##0.00', '#,##0.00', "0", "@"])
        marca = pagina.cell(linha - 1, 1)
        marca.fill = PatternFill("solid", fgColor=fundo)
        marca.font = Font(name="Calibri", bold=True, color=letra)
        marca.alignment = Alignment(horizontal="left", indent=1)
        vizinha = pagina.cell(linha - 1, 2)
        vizinha.fill = PatternFill("solid", fgColor=fundo)
        vizinha.border = grade
        pagina.merge_cells(start_row=linha - 1, start_column=1,
                           end_row=linha - 1, end_column=2)

    pagina.freeze_panes = "A2"


def gerar_planilha(achados: list[dict], hoje: dt.date) -> Path | None:
    """A mesma lista do e-mail, em Excel, para filtrar e somar.

    UMA linha por titulo e os valores em coluna; sem aba de instrucao e sem
    bloco de texto -- planilha e' dado, nao recado.

    21/09/2026, pedido dele: a planilha passou a SEPARAR os titulos do mesmo
    jeito que o corpo do e-mail -- uma faixa por DATA de bordero, com o mesmo
    texto (`rotulo_do_grupo`) e a mesma paleta. Cada grupo e' tambem um grupo
    de verdade do Excel: dobra e desdobra no +/- da margem. Alem disso: coluna
    "Dias s/ baixa" com escala de cor, tipo de pagamento colorido, saldo zerado
    em destaque, TOTAL no pe por formula e uma aba RESUMO com as somas por data
    e por tipo de pagamento.

    Data vai como DATA e valor como NUMERO (o e-mail manda os dois ja
    formatados; aqui isso nao serve, porque no Excel texto nao soma). Ja a
    linha digitavel, o codigo de barras e o numero do bordero vao como TEXTO
    FORCADO: sao numeros com zero na frente, e o Excel come o zero se puder.
    """
    try:
        import openpyxl
        from openpyxl.formatting.rule import ColorScaleRule
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("   AVISO: sem openpyxl -- o e-mail vai sem a planilha.")
        return None

    arquivo = openpyxl.Workbook()
    pagina = arquivo.active
    pagina.title = "BORDERO SEM BAIXA"
    quantas_colunas = len(COLUNAS_EXCEL)
    ultima_letra = get_column_letter(quantas_colunas)
    rotulos = [rotulo for rotulo, _, _ in COLUNAS_EXCEL]

    fio = Side(style="thin", color=BORDA)
    grade = Border(left=fio, right=fio, top=fio, bottom=fio)
    fundo_cab = PatternFill("solid", fgColor=AZUL)
    letra_cab = Font(name="Calibri", bold=True, color="FFFFFF")
    fundo_faixa = PatternFill("solid", fgColor=FAIXA_GRUPO)
    letra_faixa = Font(name="Calibri", bold=True, color=AZUL)
    fundo_zebra = PatternFill("solid", fgColor=ZEBRA)
    # saldo zerado e SEM baixa e' a bagunca que este alerta existe para mostrar:
    # em vez de sumir no meio dos numeros, vem com a cor do alerta
    fundo_zero = PatternFill("solid", fgColor="FCE4D6")
    letra_zero = Font(name="Calibri", bold=True, color="833C0C")

    for coluna, rotulo in enumerate(rotulos, start=1):
        celula = pagina.cell(1, coluna, rotulo)
        celula.fill, celula.font = fundo_cab, letra_cab
        celula.border = grade
        celula.alignment = Alignment(vertical="center", horizontal="center"
                                     if rotulo in CENTRALIZADAS else "left")
        pagina.column_dimensions[get_column_letter(coluna)].width = LARGURA.get(rotulo, 14)
    pagina.row_dimensions[1].height = 22

    linha = 2
    for chave, grupo in itertools.groupby(achados, key=lambda a: a["bordero"]):
        do_grupo = list(grupo)
        # A faixa do grupo e' a MESMA do corpo do e-mail, texto inclusive
        # (rotulo_do_grupo): numero do bordero, data, quantos titulos e a soma.
        for coluna in range(1, quantas_colunas + 1):
            celula = pagina.cell(linha, coluna)
            celula.fill, celula.border = fundo_faixa, grade
        banda = pagina.cell(linha, 1, rotulo_do_grupo(chave, do_grupo))
        banda.font = letra_faixa
        banda.alignment = Alignment(vertical="center", indent=1)
        pagina.merge_cells(start_row=linha, start_column=1,
                           end_row=linha, end_column=quantas_colunas)
        pagina.row_dimensions[linha].height = 20
        linha += 1

        for i, achado in enumerate(do_grupo):
            zebrada = bool(i % 2)
            for coluna, (rotulo, campo, tipo) in enumerate(COLUNAS_EXCEL, start=1):
                if campo == "_dias":
                    valor = (hoje - achado["_data"]).days
                else:
                    valor = achado.get(campo)
                if tipo == "texto":
                    valor = str(valor or "")
                celula = pagina.cell(linha, coluna,
                                     valor if valor not in ("",) else None)
                celula.font = Font(name="Calibri")
                celula.border = grade
                if zebrada:
                    celula.fill = fundo_zebra
                if tipo == "data":
                    celula.number_format = "DD/MM/YYYY"
                    celula.alignment = Alignment(horizontal="center")
                elif tipo == "moeda":
                    celula.number_format = '#,##0.00'
                elif tipo == "inteiro":
                    celula.number_format = "0"
                    celula.alignment = Alignment(horizontal="center")
                else:
                    celula.number_format = "@"
                    if rotulo in CENTRALIZADAS:
                        celula.alignment = Alignment(horizontal="center")

                if rotulo == "Tipo Pgto":
                    fundo, letra = cor_do_tipo(achado["tipo_pgto"])
                    celula.fill = PatternFill("solid", fgColor=fundo)
                    celula.font = Font(name="Calibri", bold=True, color=letra)
                elif rotulo == "Vencto Real":
                    # o mesmo destaque que ele tem no corpo do e-mail
                    celula.font = Font(name="Calibri", bold=True)
                elif rotulo == "Saldo":
                    if achado["_saldo"]:
                        celula.font = Font(name="Calibri", bold=True)
                    else:
                        celula.fill, celula.font = fundo_zero, letra_zero
                elif rotulo in ("Linha Digitável", "Cod. Barras", "Código (UUID)"):
                    # dado de copiar, nao de ler: menor e mais apagado, para
                    # nao competir com o nome e o valor
                    celula.font = Font(name="Calibri", size=9, color="404040")
            # cada data de bordero vira um grupo que dobra no +/- do Excel
            pagina.row_dimensions[linha].outlineLevel = 1
            linha += 1

    ultima_de_dados = linha - 1

    # TOTAL no pe, na cor do cabecalho: soma por formula, para continuar certo
    # se ele apagar linha na mao.
    for coluna in range(1, quantas_colunas + 1):
        celula = pagina.cell(linha, coluna)
        celula.fill, celula.font, celula.border = fundo_cab, letra_cab, grade
    pagina.cell(linha, 1, "TOTAL")
    pagina.cell(linha, rotulos.index("Fornecedor") + 1,
                f"{len(achados)} título{'s' if len(achados) > 1 else ''} "
                f"em {len({a['bordero'] for a in achados})} data(s) de borderô")
    for rotulo in ("Vlr. Título", "Saldo"):
        coluna = rotulos.index(rotulo) + 1
        letra_coluna = get_column_letter(coluna)
        celula = pagina.cell(linha, coluna,
                             f"=SUM({letra_coluna}2:{letra_coluna}{ultima_de_dados})")
        celula.number_format = '#,##0.00'
        celula.fill, celula.font, celula.border = fundo_cab, letra_cab, grade
    pagina.row_dimensions[linha].height = 20

    # verde -> amarelo -> vermelho conforme o titulo envelhece parado
    coluna_dias = get_column_letter(rotulos.index("Dias s/ baixa") + 1)
    pagina.conditional_formatting.add(
        f"{coluna_dias}2:{coluna_dias}{ultima_de_dados}",
        ColorScaleRule(start_type="min", start_color="C6EFCE",
                       mid_type="percentile", mid_value=50, mid_color="FFEB9C",
                       end_type="max", end_color="FFC7CE"))

    pagina.sheet_properties.outlinePr.summaryBelow = False
    pagina.freeze_panes = "A2"
    pagina.auto_filter.ref = f"A1:{ultima_letra}{ultima_de_dados}"

    montar_resumo(arquivo, achados, hoje)

    destino = PASTA / "DADOS" / f"BORDERO SEM BAIXA {hoje:%d-%m-%Y}.xlsx"
    destino.parent.mkdir(parents=True, exist_ok=True)
    try:
        arquivo.save(destino)
    except PermissionError:
        # ele pode estar com a planilha de ontem aberta; nome alternativo
        # em vez de derrubar o alerta inteiro
        destino = Path(tempfile.gettempdir()) / destino.name
        arquivo.save(destino)
        print(f"   (a planilha estava aberta; gravei em {destino})")
    return destino


def rotulo_do_grupo(chave: str, linhas: list[dict]) -> str:
    """O texto da faixa que abre cada data de bordero.

    Leva os NUMEROS dos borderos daquele dia (pedido dele em 18/09/2026) e a
    conta do grupo -- quantos titulos e quanto somam -- porque e' isso que se
    olha primeiro: "o que saiu no bordero do dia 17 e nao voltou".

    Os numeros vem sem repetir e na ordem, que e' a mesma das linhas embaixo.
    Se um dia tiver bordero demais para caber, os primeiros aparecem e o resto
    vira "e mais N": faixa que estoura a largura empurra a tabela.
    """
    total = sum(l["_saldo"] for l in linhas)
    quantos = len(linhas)
    numeros = []
    for l in linhas:
        if l["num_bordero"] and l["num_bordero"] not in numeros:
            numeros.append(l["num_bordero"])
    if len(numeros) > 6:
        lista = ", ".join(numeros[:6]) + f" e mais {len(numeros) - 6}"
    elif len(numeros) > 1:
        lista = ", ".join(numeros[:-1]) + " e " + numeros[-1]
    else:
        lista = numeros[0] if numeros else "sem número"
    return (f"Borderô {lista} — {chave} — {quantos} título"
            f"{'s' if quantos > 1 else ''} · {reais(total)}")


def montar_html(achados: list[dict], resumo: dict, hoje: dt.date,
                desde: dt.date) -> str:
    quantos = len(achados)
    total = sum(a["_saldo"] for a in achados)
    com_ld = sum(1 for a in achados if a["linha_dig"])
    mais_antigo = achados[0]["bordero"] if achados else ""
    frase = ("1 título está com borderô emitido"
             if quantos == 1 else
             f"{quantos} títulos estão com borderô emitido")
    plural = "esse título" if quantos == 1 else "esses títulos"
    return f"""<div style="{comuns.FONTE}font-size:11pt;color:#000000;">
  <p style="margin:0 0 12px 0;">Bom dia,</p>

  <p style="margin:0 0 12px 0;">{frase} (borderô de {desde:%d/%m/%Y} até
  hoje) e <b>continua sem data de baixa na SE2</b> — somando
  <b>{reais(total)}</b> em aberto. O borderô mais antigo da lista é de
  <b>{mais_antigo}</b>.</p>

  <p style="margin:0 0 14px 0;">O borderô já saiu, mas a baixa não voltou para
  o título: ou o pagamento não aconteceu, ou aconteceu e a baixa não foi
  lançada. Enquanto isso {plural} segue figurando em aberto e pode voltar para
  pagamento. Abaixo vão o <i>tipo de pagamento</i>, o <i>vencimento</i> e o
  <i>vencimento real</i> de cada um, <b>agrupados por data de borderô</b> (do
  mais antigo para o mais recente, com a soma de cada dia); na faixa abaixo de
  cada linha vai o
  <b>código (UUID)</b> e, quando o título tem boleto, a
  <b>linha digitável</b> vem junto ({com_ld} de {quantos}). O <b>nº do
  título</b> abre o <a href="{PAINEL_SE2}" style="color:#1F3864;">painel da
  SE2</a>, onde o <b>código (UUID)</b> da faixa serve de busca. A mesma lista
  vai <b>anexada em Excel</b>, com o código de barras e o valor do título, para
  quem preferir filtrar e somar.</p>

  {comuns.tabela(COLUNAS_EMAIL, achados,
                 direita=("saldo",),
                 destaque=("real",),
                 links={"titulo_parcela": PAINEL_SE2},
                 sublinha=faixa_detalhe,
                 grupo=lambda l: l["bordero"],
                 grupo_rotulo=rotulo_do_grupo)}

  {comuns.rodape(
      f"Recorte: Dt. borderô de {desde:%d/%m/%Y} até {hoje:%d/%m/%Y}, sem DT Baixa.",
      f"Base SE2 salva em {resumo['salva_em']:%d/%m/%Y às %H:%M} — "
      f"{resumo['com_bordero']} títulos com borderô na base, "
      f"{resumo['ate_hoje']} com borderô até hoje"
      + (f"; {resumo['antes_do_corte']} fora da lista por serem anteriores a "
         f"{desde:%d/%m/%Y}." if resumo['antes_do_corte'] else "."))}
</div>"""


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--para", default="")
    parser.add_argument("--copia", default="")
    parser.add_argument("--teste", action="store_true",
                        help="nao envia: grava a previa e mostra o resumo")
    parser.add_argument("--forcar", action="store_true",
                        help="envia mesmo se ja mandou hoje")
    parser.add_argument("--desde", default=f"{DESDE_PADRAO:%d/%m/%Y}",
                        help="so bordero a partir desta data (dd/mm/aaaa)")
    parser.add_argument("--base", type=Path, default=None)
    args = parser.parse_args()

    base = args.base or achar_base()
    if not base or not base.exists():
        print(f"   AVISO: nao achei a SE2 em {BASES} -- alerta de borderô nao rodou.")
        return 0

    hoje = dt.date.today()
    desde = como_data(args.desde) or DESDE_PADRAO

    try:
        achados, resumo = ler_se2(base, hoje, desde)
    except Exception as erro:  # noqa: BLE001
        print(f"   AVISO: nao consegui ler a SE2 ({erro}). A rodada segue.")
        return 0

    print(f"   SE2: {resumo['linhas']} títulos | {resumo['com_bordero']} com borderô | "
          f"{resumo['ate_hoje']} com borderô até {hoje:%d/%m} | "
          f"{resumo['antes_do_corte']} descartado(s) por serem anteriores a "
          f"{desde:%d/%m/%Y}")

    if not achados:
        print("   Nenhum título com borderô emitido e sem data de baixa.")
        return 0

    total = sum(a["_saldo"] for a in achados)
    print(f"   {len(achados)} título(s) com borderô e SEM baixa ({reais(total)}, "
          f"{resumo['com_linha_dig']} com linha digitável):")
    for a in achados:
        print(f"     bord {a['bordero']} nº {a['num_bordero'] or '-':<6} | {a['filial']} | "
              f"{a['titulo']}/{a['parcela'] or '-'} | venc {a['vencimento'] or '-'} | "
              f"real {a['real'] or '-'} | {a['tipo_pgto'][:14]:<14} | {a['saldo']:>15} | "
              f"{a['fornecedor'][:28]}")

    corpo = montar_html(achados, resumo, hoje, desde)
    assunto = (f"[Painel Boletos] {len(achados)} título"
               f"{'s' if len(achados) > 1 else ''} com borderô emitido e SEM baixa "
               f"na SE2 - {hoje:%d/%m/%Y}")

    planilha = gerar_planilha(achados, hoje)
    if planilha:
        print(f"   Planilha: {planilha}")

    if args.teste:
        PREVIA.parent.mkdir(parents=True, exist_ok=True)
        PREVIA.write_text(corpo, encoding="utf-8")
        print(f"\n   MODO TESTE: nada foi enviado.\n   Assunto: {assunto}")
        print(f"   Prévia:  {PREVIA}")
        return 0

    destino = args.para or comuns.destinatarios(CHAVE)
    if not destino:
        print("   AVISO: nao achei para quem mandar (.alerta_bordero_para).")
        return 0
    copia = args.copia or comuns.copias(CHAVE)

    if not args.forcar and comuns.ja_enviado_hoje(CHAVE):
        print(f"   Já enviado hoje ({comuns.quando_enviou(CHAVE)}) -- não mando de novo.")
        print("   Para mandar assim mesmo: --forcar")
        return 0

    try:
        comuns.enviar(assunto, corpo, destino, copia=copia,
                      anexos=[planilha] if planilha else None)
        comuns.marcar_enviado(CHAVE, len(achados))
        print(f"   E-mail enviado para {destino}"
              + (f" (cópia: {copia})" if copia else "")
              + (f" com a planilha {planilha.name} em anexo." if planilha else "."))
    except Exception as erro:  # noqa: BLE001
        print(f"   AVISO: o e-mail NAO foi enviado ({erro}). O painel segue.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
