# -*- coding: utf-8 -*-
"""Titulo de BOLETO em aberto cujo valor NAO bate com o valor dentro do boleto.

Pedido dele em 21/09/2026, com o print do AUTO POSTO ESTRELA DA GUIA: o titulo
estava R$ 174.392,15 e a linha digitavel terminava em 0017439220 -- R$
174.392,20. Faltavam R$ 0,05 de juros no titulo. "Precioso que voce monitore
isso."

  o recorte   Tipo Pgto de BOLETO, EM ABERTO (Saldo > 0 e sem DT Baixa),
              VENCIMENTO de 7 dias atras ate 31/12 do ano corrente.
  o que sai   so o que DIVERGE. Titulo cujo valor bate com o boleto nao vira
              e-mail: alerta que mostra o que esta certo ensina a ignorar.
  sem margem  qualquer centavo de diferenca entra. Foi um centavo -- cinco --
              que gerou o pedido, e "tolerancia de centavos" mataria
              exatamente o caso que ele quer ver.

--------------------------------------------------------------------------
DE ONDE SAI O VALOR DO BOLETO. Sao duas familias de codigo, e elas nao se
parecem em nada -- ler as duas do mesmo jeito daria numero errado, calado.

  BOLETO BANCARIO -- linha digitavel de 47 digitos
      campo1(10) campo2(11) campo3(11) DVgeral(1) campo5(14)
      o campo 5 e' fator de vencimento (4) + VALOR EM CENTAVOS (10)
      => o valor sao os 10 ULTIMOS digitos. E' o que ele viu no print.
      O codigo de barras da mesma coisa em 44 digitos, com o valor nas
      posicoes 10 a 19.

  ARRECADACAO / CONVENIO -- linha digitavel de 48 digitos (comeca com 8)
      4 blocos de 12: 11 digitos uteis + 1 DV. Tirando os DVs sobram os 44 do
      codigo de barras:
        pos1 produto(8) | pos2 segmento | pos3 TIPO DE VALOR | pos4 DV geral
        pos5..15 VALOR EM CENTAVOS | pos16..44 campo livre
      ⚠ o TIPO DE VALOR decide se aquele campo e' dinheiro:
        6 e 8 = valor efetivo (e' dinheiro, da para comparar)
        7 e 9 = quantidade de moeda (NAO e' dinheiro -- comparar daria uma
                divergencia inventada, entao esses ficam de fora)

CONFERIDO EM 21/09/2026, tres vezes e de tres jeitos diferentes:
  1. os 10 ultimos digitos da linha digitavel -- bateu com o print dele,
     digito a digito, inclusive os R$ 0,05;
  2. o CODIGO DE BARRAS, que e' outro campo da SE2 e tem o valor em outro
     lugar -- deu o mesmo valor nos 21 divergentes, sem excecao;
  3. o DV modulo 11 do codigo de barras -- os 789 boletos da janela passam,
     ou seja, nenhum codigo esta corrompido e nenhuma divergencia e' erro de
     leitura.

Uso:
    python alerta_valor_boleto.py                 # manda, no maximo uma vez por dia
    python alerta_valor_boleto.py --teste         # nao manda: grava a previa
    python alerta_valor_boleto.py --forcar        # manda de novo no mesmo dia
    python alerta_valor_boleto.py --minimo 1,00   # so diferenca acima de 1 real
    python alerta_valor_boleto.py --dias 15 --ate 31/12/2027
    python alerta_valor_boleto.py --base "C:\\caminho\\SE2.xlsx"

Codigo de saida: SEMPRE 0 quando conseguiu ler a base. Falha de e-mail nao
derruba a rodada do painel.
"""

from __future__ import annotations

import argparse
import datetime as dt
import itertools
import re
import shutil
import tempfile
from pathlib import Path

import alertas_comuns as comuns
import caminhos

PASTA = Path(__file__).resolve().parent
PREVIA = PASTA / "DADOS" / "alerta_valor_boleto_previa.html"
CHAVE = "valor_boleto"

RAIZ = caminhos.raiz_lucas()
BASES = caminhos.bases_genericos()

ABA = "SE2"

PAINEL_SE2 = "https://se2-lucas.vercel.app/"

# Quantos dias para TRAS. Para a FRENTE vai ate 31/12 do ano corrente (pedido
# dele: "todos os titulos em aberto de boleto que aparecerem no final do ano").
DIAS_PADRAO = 7

# Diferenca minima para entrar. ZERO de proposito -- ver o cabecalho.
MINIMO_PADRAO = 0.0

# Teto da diferenca. 0 = sem teto (o padrao). Serve para separar as duas
# conversas que a lista mistura: ate uns poucos reais e quase sempre JUROS OU
# MULTA a lancar no titulo -- conserto de rotina, um lancamento e acabou --,
# enquanto desvio grande e boleto errado ou titulo errado, que precisa de
# gente falando com o fornecedor. Rodar com teto entrega a primeira lista
# limpa, sem os poucos casos grandes roubando a atencao.
MAXIMO_PADRAO = 0.0

LIMITE_NO_EMAIL = 150

COLUNAS = {
    "filial": "Filial", "prefixo": "Prefixo", "tipo": "Tipo",
    "num": "No. Titulo", "parcela": "Parcela",
    "nome": "Nome Fornece", "razao": "Razão Social", "natureza": "Natureza",
    "emissao": "DT Emissao",
    "vencimento": "Vencimento", "real": "Vencto Real",
    "valor": "Vlr.Titulo", "baixa": "DT Baixa", "saldo": "Saldo",
    "tipo_pgto": "Tipo Pgto", "linha_dig": "Linha Dig.",
    "cod_barras": "Cod.Barras", "uuid": "Campo UUID",
    "usuario_inc": "Usuario Inc",
}

MESES = ("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro")


def achar_base() -> Path | None:
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
        limpo = valor.strip().replace(".", "").replace(",", ".")
        try:
            return float(limpo)
        except ValueError:
            return 0.0
    return 0.0


def reais(valor: float) -> str:
    return "R$ " + f"{valor:,.2f}".replace(",", "~").replace(".", ",").replace("~", ".")


def com_sinal(valor: float) -> str:
    """O sinal e' a informacao: + quer dizer que o boleto cobra MAIS que o
    titulo, - que cobra menos. Sem ele, os dois casos viram o mesmo numero."""
    return ("+" if valor > 0 else "−") + reais(abs(valor))


def texto(valor) -> str:
    if valor is None:
        return ""
    limpo = str(valor).strip()
    if limpo.lower() in ("none", "nan", "null"):
        return ""
    if limpo and set(limpo) <= set(" /:-"):
        return ""
    return limpo


def so_digitos(valor) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def valor_do_codigo(bruto) -> tuple[float | None, str]:
    """(valor em reais, de onde saiu). None quando o codigo nao diz um valor.

    Ver o cabecalho do arquivo para o desenho dos dois formatos. O motivo volta
    junto porque ele vai para o rodape do e-mail: "nao achei valor em N" e
    "achei e bateu" sao coisas diferentes, e quem le precisa saber a diferenca.
    """
    d = so_digitos(bruto)
    if not d:
        return None, "sem código"

    if len(d) == 47:                      # linha digitavel de boleto bancario
        return int(d[-10:]) / 100.0, "linha digitável"
    if len(d) == 44 and d[0] != "8":      # codigo de barras de boleto bancario
        return int(d[9:19]) / 100.0, "código de barras"

    if len(d) == 48 and d[0] == "8":      # linha digitavel de arrecadacao
        barras = "".join(d[i:i + 11] for i in range(0, 48, 12))
    elif len(d) == 44 and d[0] == "8":    # codigo de barras de arrecadacao
        barras = d
    else:
        return None, f"código com {len(d)} dígitos"

    if barras[2] not in "68":
        # 7 e 9 sao quantidade de moeda, nao dinheiro
        return None, "convênio sem valor em reais"
    return int(barras[4:15]) / 100.0, "convênio"


def barras_de_47(d: str) -> str:
    """Os 44 digitos do codigo de barras, remontados da linha digitavel de 47.

    Serve para conferir o DV: a linha digitavel embaralha os campos, e so no
    codigo de barras o modulo 11 fecha.
    """
    return d[0:4] + d[32] + d[33:47] + d[4:9] + d[10:20] + d[21:31]


def dv_confere(barras: str) -> bool:
    """DV geral (posicao 5) do boleto bancario, modulo 11 com pesos 2..9.

    ⚠ Isto NAO e' enfeite: e' o que separa "o boleto cobra outro valor" de
    "o campo esta corrompido e eu li lixo". So o primeiro merece e-mail.
    """
    if len(barras) != 44 or not barras.isdigit():
        return False
    sem_dv = barras[:4] + barras[5:]
    peso, soma = 2, 0
    for c in reversed(sem_dv):
        soma += int(c) * peso
        peso = peso + 1 if peso < 9 else 2
    dv = 11 - (soma % 11)
    return barras[4] == ("1" if dv in (0, 10, 11) else str(dv))


def ler_se2(caminho: Path, hoje: dt.date, desde: dt.date, ate: dt.date,
            minimo: float, maximo: float = 0.0) -> tuple[list[dict], dict]:
    """Varre a SE2 e devolve os boletos em aberto cujo valor nao bate.

    ⚠ A SE2 costuma estar ABERTA no Excel na hora da rodada, e ai o openpyxl
    leva PermissionError. Copiar para a pasta temporaria e ler a copia resolve.
    """
    import openpyxl

    origem = caminho
    temporario = None
    try:
        arquivo = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    except PermissionError:
        temporario = Path(tempfile.gettempdir()) / f"se2_valor_{dt.date.today():%Y%m%d}.xlsx"
        shutil.copy2(caminho, temporario)
        arquivo = openpyxl.load_workbook(temporario, read_only=True, data_only=True)
        print("   (a SE2 estava aberta; li uma cópia)")

    if ABA not in arquivo.sheetnames:
        arquivo.close()
        raise KeyError(f'a planilha nao tem a aba "{ABA}" (tem: {arquivo.sheetnames})')
    pagina = arquivo[ABA]
    linhas = pagina.iter_rows(values_only=True)
    cabecalho = [str(c).strip() if c is not None else "" for c in next(linhas)]

    onde, faltando = {}, []
    for chave, rotulo in COLUNAS.items():
        if rotulo in cabecalho:
            onde[chave] = cabecalho.index(rotulo)
        else:
            faltando.append(rotulo)
    for obrigatoria in ("num", "tipo_pgto", "linha_dig", "valor", "saldo", "vencimento"):
        if obrigatoria not in onde:
            arquivo.close()
            raise KeyError("colunas ausentes na SE2: " + ", ".join(faltando))

    def campo(linha, chave):
        i = onde.get(chave)
        return linha[i] if i is not None and i < len(linha) else None

    achados = []
    resumo = {"linhas": 0, "boletos": 0, "em_aberto": 0, "com_codigo": 0,
              "bateram": 0, "sem_codigo": 0, "dv_ruim": 0, "abaixo_do_minimo": 0,
              "acima_do_maximo": 0, "fora_do_teto": []}
    for linha in linhas:
        if campo(linha, "num") in (None, ""):
            continue
        resumo["linhas"] += 1

        tipo_pgto = texto(campo(linha, "tipo_pgto")).upper()
        if "BOLETO" not in tipo_pgto:
            continue
        vencimento = como_data(campo(linha, "vencimento"))
        if not vencimento or not (desde <= vencimento <= ate):
            continue
        resumo["boletos"] += 1

        # EM ABERTO = saldo de verdade E sem baixa. Os dois, e nao um so: a SE2
        # tem titulo baixado com saldo residual e titulo zerado sem baixa, e
        # nenhum dos dois e' conta a pagar.
        baixa = como_data(campo(linha, "baixa"))
        saldo = como_numero(campo(linha, "saldo"))
        if baixa or saldo <= 0:
            continue
        resumo["em_aberto"] += 1

        linha_dig = texto(campo(linha, "linha_dig"))
        cod_barras = texto(campo(linha, "cod_barras"))
        do_ld, fonte_ld = valor_do_codigo(linha_dig)
        do_cb, fonte_cb = valor_do_codigo(cod_barras)
        # A linha digitavel manda -- e' o que a pessoa digita no banco. O codigo
        # de barras entra como segunda opiniao, e so assume quando ela falta.
        lido = do_ld if do_ld is not None else do_cb
        fonte = fonte_ld if do_ld is not None else fonte_cb
        if lido is None:
            resumo["sem_codigo"] += 1
            continue
        resumo["com_codigo"] += 1

        valor = como_numero(campo(linha, "valor"))
        diferenca = round(lido - valor, 2)
        if diferenca == 0:
            resumo["bateram"] += 1
            continue
        if abs(diferenca) < minimo:
            resumo["abaixo_do_minimo"] += 1
            continue
        # ⚠ O que o teto corta NAO some: fica anotado com nome e valor, e vai
        # ESCRITO no rodape do e-mail. Lista com teto e um recorte, nao a
        # verdade inteira -- e quem le tem de saber que existe coisa maior
        # fora dela, senao o teto vira uma mentira por omissao.
        if maximo and abs(diferenca) > maximo:
            resumo["acima_do_maximo"] += 1
            resumo["fora_do_teto"].append(
                (texto(campo(linha, "razao")) or texto(campo(linha, "nome")),
                 diferenca))
            continue

        # ⚠ Codigo que nao fecha o DV NAO vira alerta: seria acusar o
        # fornecedor por causa de um campo digitado errado na base. Conta no
        # rodape para nao sumir em silencio.
        digitos = so_digitos(linha_dig)
        if len(digitos) == 47 and not dv_confere(barras_de_47(digitos)):
            resumo["dv_ruim"] += 1
            continue

        # As duas leituras discordando e' sinal de campo bagunçado, e vai
        # ESCRITO na faixa -- nao some.
        discordam = (do_ld is not None and do_cb is not None
                     and round(do_ld - do_cb, 2) != 0)
        real = como_data(campo(linha, "real"))
        emissao = como_data(campo(linha, "emissao"))
        prazo = real or vencimento
        atraso = (hoje - prazo).days if prazo else 0
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
            "natureza": texto(campo(linha, "natureza")),
            "tipo_pgto": tipo_pgto,
            "emissao": f"{emissao:%d/%m/%Y}" if emissao else "",
            "vencimento": f"{vencimento:%d/%m/%Y}" if vencimento else "",
            "real": f"{real:%d/%m/%Y}" if real else "",
            "valor": reais(valor),
            "boleto": reais(lido),
            "diferenca": com_sinal(diferenca),
            "saldo": reais(saldo),
            "linha_dig": linha_dig,
            "cod_barras": cod_barras,
            "fonte": fonte,
            "discordam": discordam,
            "usuario": texto(campo(linha, "usuario_inc")),
            "grupo": "maior" if diferenca > 0 else "menor",
            "situacao": (f"Vencido há {atraso}d" if atraso > 0 else "Em aberto"),
            "_valor": valor, "_boleto": lido, "_dif": diferenca,
            "_saldo": saldo, "_venc": vencimento, "_real": real,
            "_emissao": emissao, "_atraso": atraso,
        })
    arquivo.close()
    if temporario and temporario.exists():
        try:
            temporario.unlink()
        except OSError:
            pass
    resumo["base"] = origem
    resumo["salva_em"] = dt.datetime.fromtimestamp(origem.stat().st_mtime)
    # Primeiro os que o boleto cobra A MAIS (e o que gera pagamento a menor e
    # juros depois), e dentro de cada lado o maior desvio na frente.
    achados.sort(key=lambda a: (0 if a["grupo"] == "maior" else 1, -abs(a["_dif"]),
                                a["_venc"]))
    return achados, resumo


COLUNAS_EMAIL = [
    ("Filial", "filial"),
    ("Nº título", "titulo_parcela"),
    ("Fornecedor", "fornecedor"),
    ("Vencimento", "vencimento"),
    ("Vencto real", "real"),
    ("Vlr. título", "valor"),
    ("Vlr. boleto", "boleto"),
    ("Diferença", "diferenca"),
]


def faixa_detalhe(linha: dict) -> str:
    """A faixa de largura inteira embaixo de cada linha.

    Leva a LINHA DIGITAVEL inteira e crua -- e' o dado de copiar para o banco,
    e mascarar faria a pessoa digitar na mao. Leva tambem de onde saiu o valor,
    porque "li da linha digitavel" e "li do codigo de barras" mudam o que se
    confere primeiro.
    """
    partes = [f"Valor lido da {linha['fonte']}"]
    if linha.get("discordam"):
        partes.append("⚠ LINHA DIGITÁVEL E CÓDIGO DE BARRAS DISCORDAM ENTRE SI")
    if linha.get("tipo_pgto") and linha["tipo_pgto"] != "BOLETO":
        partes.append("Tipo pgto: " + linha["tipo_pgto"])
    if linha.get("saldo") and linha["saldo"] != linha["valor"]:
        partes.append("Saldo em aberto: " + linha["saldo"])
    if linha.get("uuid"):
        partes.append("Código (UUID): " + linha["uuid"])
    if linha.get("linha_dig"):
        partes.append("Linha digitável: " + linha["linha_dig"])
    return "     ·     ".join(partes)


COLUNAS_EXCEL = [
    ("Vencimento", "_venc", "data"),
    ("Vencto Real", "_real", "data"),
    ("Filial", "filial", "texto"),
    ("Prefixo", "prefixo", "texto"),
    ("Tipo", "tipo", "texto"),
    ("Nº Título", "titulo", "texto"),
    ("Parcela", "parcela", "texto"),
    ("Fornecedor", "fornecedor", "texto"),
    ("Natureza", "natureza", "texto"),
    ("Tipo Pgto", "tipo_pgto", "texto"),
    ("Vlr. Título", "_valor", "moeda"),
    ("Vlr. Boleto", "_boleto", "moeda"),
    ("Diferença", "_dif", "moeda"),
    ("Saldo", "_saldo", "moeda"),
    ("Situação", "situacao", "texto"),
    ("Valor lido de", "fonte", "texto"),
    ("Linha Digitável", "linha_dig", "texto"),
    ("Cod. Barras", "cod_barras", "texto"),
    ("Código (UUID)", "uuid", "texto"),
    ("Usuário Inc", "usuario", "texto"),
]

LARGURA = {"Vencimento": 12, "Vencto Real": 12, "Filial": 7, "Prefixo": 9,
           "Tipo": 7, "Nº Título": 13, "Parcela": 9, "Fornecedor": 38,
           "Natureza": 11, "Tipo Pgto": 14, "Vlr. Título": 15,
           "Vlr. Boleto": 15, "Diferença": 14, "Saldo": 15, "Situação": 14,
           "Valor lido de": 18, "Linha Digitável": 50, "Cod. Barras": 48,
           "Código (UUID)": 38, "Usuário Inc": 26}

CENTRALIZADAS = {"Vencimento", "Vencto Real", "Filial", "Prefixo", "Tipo",
                 "Parcela", "Natureza", "Situação"}

AZUL = "1F3864"
FAIXA_GRUPO = "D6DCE4"
ZEBRA = "F2F2F2"
BORDA = "BFBFBF"
# Vermelho para o boleto que cobra A MAIS (risco de pagar a menos e tomar
# juros), laranja para o que cobra a menos.
VERMELHO_FUNDO, VERMELHO_LETRA = "FCE4D6", "C00000"
LARANJA_FUNDO, LARANJA_LETRA = "FFF2CC", "7F6000"


def rotulo_do_grupo(chave: str, linhas: list[dict]) -> str:
    """A faixa que abre cada lado da divergencia.

    O lado e' a informacao que decide o que fazer: boleto MAIOR quase sempre e'
    juros ou multa que ainda nao entraram no titulo (foi o caso do AUTO POSTO
    ESTRELA DA GUIA, R$ 0,05); boleto MENOR e' desconto concedido ou titulo
    lancado a maior.
    """
    total = sum(abs(l["_dif"]) for l in linhas)
    quantos = len(linhas)
    plural = "s" if quantos > 1 else ""
    if chave == "maior":
        return (f"O boleto cobra MAIS que o título — {quantos} título{plural} "
                f"· diferença somada {reais(total)} "
                f"(normalmente juros ou multa que faltam no título)")
    return (f"O boleto cobra MENOS que o título — {quantos} título{plural} "
            f"· diferença somada {reais(total)} "
            f"(desconto concedido, ou título lançado a maior)")


def montar_resumo(arquivo, achados: list[dict]) -> None:
    """Uma segunda aba com as duas contas de cabeca: quanto desvia para cada
    lado, e quais fornecedores repetem. So numero -- nenhuma instrucao."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    pagina = arquivo.create_sheet("RESUMO")
    fio = Side(style="thin", color=BORDA)
    grade = Border(left=fio, right=fio, top=fio, bottom=fio)
    fundo_cab = PatternFill("solid", fgColor=AZUL)
    letra_cab = Font(name="Calibri", bold=True, color="FFFFFF")
    fundo_faixa = PatternFill("solid", fgColor=FAIXA_GRUPO)
    letra_faixa = Font(name="Calibri", bold=True, color=AZUL)

    for coluna, largura in enumerate([38, 14, 10, 16, 16, 16], start=1):
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
                horizontal="right" if formato == '#,##0.00' else "center")
        return linha + 1

    linha = cabecalho(1, ["Lado da divergência", "Títulos", "",
                          "Vlr. Título", "Vlr. Boleto", "Diferença"])
    for chave, rotulo in (("maior", "Boleto cobra MAIS"),
                          ("menor", "Boleto cobra MENOS")):
        do_lado = [a for a in achados if a["grupo"] == chave]
        if not do_lado:
            continue
        linha = numeros(linha, [
            rotulo, len(do_lado), "",
            sum(a["_valor"] for a in do_lado),
            sum(a["_boleto"] for a in do_lado),
            sum(a["_dif"] for a in do_lado),
        ], ["@", "0", "@", '#,##0.00', '#,##0.00', '#,##0.00'])
    linha = numeros(linha, [
        "TOTAL", len(achados), "",
        sum(a["_valor"] for a in achados),
        sum(a["_boleto"] for a in achados),
        sum(a["_dif"] for a in achados),
    ], ["@", "0", "@", '#,##0.00', '#,##0.00', '#,##0.00'],
        negrito=True, fundo=fundo_faixa)

    linha += 2
    linha = cabecalho(linha, ["Fornecedor", "Títulos", "",
                              "Vlr. Título", "Vlr. Boleto", "Diferença"])
    por_fornecedor: dict[str, list[dict]] = {}
    for a in achados:
        por_fornecedor.setdefault(a["fornecedor"] or "—", []).append(a)
    for fornecedor, deles in sorted(por_fornecedor.items(),
                                    key=lambda par: -abs(sum(l["_dif"] for l in par[1]))):
        linha = numeros(linha, [
            fornecedor, len(deles), "",
            sum(l["_valor"] for l in deles),
            sum(l["_boleto"] for l in deles),
            sum(l["_dif"] for l in deles),
        ], ["@", "0", "@", '#,##0.00', '#,##0.00', '#,##0.00'])
        pagina.cell(linha - 1, 1).alignment = Alignment(
            horizontal="left", vertical="center", indent=1)

    pagina.freeze_panes = "A2"


def gerar_planilha(achados: list[dict], hoje: dt.date) -> Path | None:
    """A mesma lista do e-mail, em Excel, para filtrar e somar.

    UMA linha por titulo e os valores em coluna; sem aba de instrucao e sem
    bloco de texto. Uma faixa por lado da divergencia, com o mesmo texto e a
    mesma paleta do corpo do e-mail, e cada faixa e' um grupo do Excel.

    Data vai como DATA e valor como NUMERO. Linha digitavel, codigo de barras e
    numero do titulo vao como TEXTO FORCADO: sao numeros com zero na frente, e
    o Excel come o zero se puder.
    """
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("   AVISO: sem openpyxl -- o e-mail vai sem a planilha.")
        return None

    arquivo = openpyxl.Workbook()
    pagina = arquivo.active
    pagina.title = "VALOR DIVERGENTE"
    quantas = len(COLUNAS_EXCEL)
    ultima_letra = get_column_letter(quantas)
    rotulos = [r for r, _, _ in COLUNAS_EXCEL]

    fio = Side(style="thin", color=BORDA)
    grade = Border(left=fio, right=fio, top=fio, bottom=fio)
    fundo_cab = PatternFill("solid", fgColor=AZUL)
    letra_cab = Font(name="Calibri", bold=True, color="FFFFFF")
    fundo_faixa = PatternFill("solid", fgColor=FAIXA_GRUPO)
    letra_faixa = Font(name="Calibri", bold=True, color=AZUL)
    fundo_zebra = PatternFill("solid", fgColor=ZEBRA)

    for coluna, rotulo in enumerate(rotulos, start=1):
        celula = pagina.cell(1, coluna, rotulo)
        celula.fill, celula.font, celula.border = fundo_cab, letra_cab, grade
        celula.alignment = Alignment(vertical="center", horizontal="center"
                                     if rotulo in CENTRALIZADAS else "left")
        pagina.column_dimensions[get_column_letter(coluna)].width = LARGURA.get(rotulo, 14)
    pagina.row_dimensions[1].height = 22

    linha = 2
    for chave, grupo in itertools.groupby(achados, key=lambda a: a["grupo"]):
        do_grupo = list(grupo)
        for coluna in range(1, quantas + 1):
            celula = pagina.cell(linha, coluna)
            celula.fill, celula.border = fundo_faixa, grade
        banda = pagina.cell(linha, 1, rotulo_do_grupo(chave, do_grupo))
        banda.font = letra_faixa
        banda.alignment = Alignment(vertical="center", indent=1)
        pagina.merge_cells(start_row=linha, start_column=1,
                           end_row=linha, end_column=quantas)
        pagina.row_dimensions[linha].height = 20
        linha += 1

        for i, achado in enumerate(do_grupo):
            zebrada = bool(i % 2)
            for coluna, (rotulo, campo, tipo) in enumerate(COLUNAS_EXCEL, start=1):
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
                else:
                    celula.number_format = "@"
                    if rotulo in CENTRALIZADAS:
                        celula.alignment = Alignment(horizontal="center")

                if rotulo == "Diferença":
                    # a coluna que motiva o e-mail: cor pelo LADO do desvio
                    fundo, letra = ((VERMELHO_FUNDO, VERMELHO_LETRA)
                                    if achado["_dif"] > 0
                                    else (LARANJA_FUNDO, LARANJA_LETRA))
                    celula.fill = PatternFill("solid", fgColor=fundo)
                    celula.font = Font(name="Calibri", bold=True, color=letra)
                elif rotulo in ("Vlr. Título", "Vlr. Boleto"):
                    celula.font = Font(name="Calibri", bold=True)
                elif rotulo in ("Linha Digitável", "Cod. Barras", "Código (UUID)"):
                    celula.font = Font(name="Calibri", size=9, color="404040")
            pagina.row_dimensions[linha].outlineLevel = 1
            linha += 1

    ultima_de_dados = linha - 1

    for coluna in range(1, quantas + 1):
        celula = pagina.cell(linha, coluna)
        celula.fill, celula.font, celula.border = fundo_cab, letra_cab, grade
    pagina.cell(linha, 1, "TOTAL")
    pagina.cell(linha, rotulos.index("Fornecedor") + 1,
                f"{len(achados)} título{'s' if len(achados) > 1 else ''} "
                f"com valor divergente")
    for rotulo in ("Vlr. Título", "Vlr. Boleto", "Diferença", "Saldo"):
        coluna = rotulos.index(rotulo) + 1
        letra_coluna = get_column_letter(coluna)
        celula = pagina.cell(linha, coluna,
                             f"=SUM({letra_coluna}2:{letra_coluna}{ultima_de_dados})")
        celula.number_format = '#,##0.00'
        celula.fill, celula.font, celula.border = fundo_cab, letra_cab, grade
    pagina.row_dimensions[linha].height = 20

    pagina.sheet_properties.outlinePr.summaryBelow = False
    pagina.freeze_panes = "A2"
    pagina.auto_filter.ref = f"A1:{ultima_letra}{ultima_de_dados}"

    montar_resumo(arquivo, achados)

    destino = PASTA / "DADOS" / f"BOLETO VALOR DIVERGENTE {hoje:%d-%m-%Y}.xlsx"
    destino.parent.mkdir(parents=True, exist_ok=True)
    try:
        arquivo.save(destino)
    except PermissionError:
        destino = Path(tempfile.gettempdir()) / destino.name
        arquivo.save(destino)
        print(f"   (a planilha estava aberta; gravei em {destino})")
    return destino


def montar_html(achados: list[dict], resumo: dict, hoje: dt.date,
                desde: dt.date, ate: dt.date, cortados: int = 0,
                total_dif: float = 0.0, quantos_geral: int = 0,
                maximo: float = 0.0) -> str:
    quantos = quantos_geral or len(achados)
    maiores = [a for a in achados if a["_dif"] > 0]
    menores = [a for a in achados if a["_dif"] < 0]
    centavos = sum(1 for a in achados if abs(a["_dif"]) <= 1)
    frase = ("1 boleto está cobrando um valor diferente do título"
             if quantos == 1 else
             f"{quantos} boletos estão cobrando um valor diferente do título")
    lados = []
    if maiores:
        lados.append(f"<b>{len(maiores)}</b> cobra"
                     + ("" if len(maiores) == 1 else "m")
                     + f" a mais (somando {reais(sum(a['_dif'] for a in maiores))})")
    if menores:
        lados.append(f"<b>{len(menores)}</b> cobra"
                     + ("" if len(menores) == 1 else "m")
                     + f" a menos (somando {reais(abs(sum(a['_dif'] for a in menores)))})")
    # O teto esconde justamente os maiores. Esconder CALADO seria o pior tipo
    # de erro num alerta de valor: quem le concluiria que nao ha nada grande.
    fora = resumo.get("fora_do_teto") or []
    maiores_fora = sorted(fora, key=lambda p: -abs(p[1]))[:3]
    aviso_teto = ("" if not fora else f"""
  <p style="margin:0 0 12px 0;padding:8px 10px;background:#FCE4D6;
  border-left:4px solid #C00000;font-size:10pt;color:#833C0C;">
  Fora desta lista, por passarem do teto de <b>{reais(maximo)}</b>, há mais
  <b>{len(fora)}</b> título{'s' if len(fora) > 1 else ''} com diferença maior —
  somando <b>{reais(sum(abs(d) for _, d in fora))}</b>. """
                  + "; ".join(f"{nome[:34]} ({com_sinal(d)})" for nome, d in maiores_fora)
                  + (f" e mais {len(fora) - 3}" if len(fora) > 3 else "") + ".</p>")
    aviso_corte = ("" if not cortados else f"""
  <p style="margin:12px 0 0 0;padding:8px 10px;background:#FFF2CC;
  border-left:4px solid #BF8F00;font-size:10pt;color:#7F6000;">
  A tabela acima mostra os <b>{len(achados)}</b> primeiros (maior diferença).
  Os outros <b>{cortados}</b> estão na <b>planilha anexada</b>, que traz sempre
  a lista completa — corpo de e-mail muito grande é cortado pelo próprio
  programa de e-mail.</p>""")
    return f"""<div style="{comuns.FONTE}font-size:11pt;color:#000000;">
  <p style="margin:0 0 12px 0;">Bom dia,</p>

  <p style="margin:0 0 12px 0;">{frase}: {" e ".join(lados)}. São títulos
  <b>em aberto</b>, com pagamento em boleto, vencendo entre
  <b>{desde:%d/%m/%Y}</b> e <b>{ate:%d/%m/%Y}</b>{f", com diferença de até <b>{reais(maximo)}</b>" if maximo else ""}.</p>
  {aviso_teto}

  <p style="margin:0 0 12px 0;">O valor que o banco vai cobrar está
  <b>dentro da própria linha digitável</b> — nos seus últimos dígitos — e é ele
  que está comparado aqui com o <i>Vlr. Título</i> da SE2. Quando o boleto cobra
  <b>a mais</b>, quase sempre são <b>juros ou multa</b> que ainda não foram
  lançados no título; pagando pelo valor do título, a diferença volta depois.
  Quando cobra <b>a menos</b>, costuma ser desconto concedido ou título lançado
  a maior.{f" <b>{centavos}</b> " + ("é de" if centavos == 1 else "são de") + " centavos — o caso clássico de juros a lançar." if centavos else ""}</p>

  <p style="margin:0 0 14px 0;">Abaixo vão o <i>vencimento</i> e o
  <i>vencimento real</i>, o <b>valor do título</b>, o <b>valor do boleto</b> e a
  <b>diferença</b>, <b>separados pelo lado do desvio</b> e do maior para o
  menor. Na faixa abaixo de cada linha vão de onde o valor foi lido, o
  <b>código (UUID)</b> e a <b>linha digitável</b> inteira, para conferir no
  banco. O <b>nº do título</b> abre o
  <a href="{PAINEL_SE2}" style="color:#1F3864;">painel da SE2</a>. A mesma lista
  vai <b>anexada em Excel</b>, com natureza, saldo, código de barras e quem
  lançou o título.</p>

  {comuns.tabela(COLUNAS_EMAIL, achados,
                 direita=("valor", "boleto", "diferenca"),
                 destaque=("diferenca",),
                 links={"titulo_parcela": PAINEL_SE2},
                 sublinha=faixa_detalhe,
                 grupo=lambda l: l["grupo"],
                 grupo_rotulo=rotulo_do_grupo)}
  {aviso_corte}

  {comuns.rodape(
      f"Recorte: Tipo Pgto de boleto, em aberto (saldo > 0 e sem DT Baixa), "
      f"vencimento de {desde:%d/%m/%Y} a {ate:%d/%m/%Y}, "
      f"valor do título diferente do valor lido no boleto"
      + (f", diferença de até {reais(maximo)}." if maximo else "."),
      f"Base SE2 salva em {resumo['salva_em']:%d/%m/%Y às %H:%M} — "
      f"{resumo['em_aberto']} boletos em aberto na janela, "
      f"{resumo['com_codigo']} com código legível, "
      f"{resumo['bateram']} com valor batendo certo"
      + (f", {resumo['sem_codigo']} sem linha digitável" if resumo["sem_codigo"] else "")
      + (f", {resumo['dv_ruim']} fora por o código não fechar o dígito verificador"
         if resumo["dv_ruim"] else "")
      + (f", {resumo['acima_do_maximo']} fora por passarem do teto"
         if resumo["acima_do_maximo"] else "")
      + (f"; {cortados} fora da tabela, só na planilha." if cortados else "."))}
</div>"""


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--para", default="")
    parser.add_argument("--copia", default="")
    parser.add_argument("--teste", action="store_true",
                        help="nao envia: grava a previa e mostra o resumo")
    parser.add_argument("--forcar", action="store_true",
                        help="envia mesmo se ja mandou hoje")
    parser.add_argument("--dias", type=int, default=DIAS_PADRAO,
                        help=f"quantos dias a janela pega para TRAS (padrao: {DIAS_PADRAO})")
    parser.add_argument("--ate", default="",
                        help="fim da janela (dd/mm/aaaa); o padrao e 31/12 deste ano")
    parser.add_argument("--minimo", default="",
                        help="diferenca minima em reais (padrao: qualquer centavo)")
    parser.add_argument("--maximo", default="",
                        help="teto da diferenca em reais, ex. 10 (padrao: sem teto)")
    parser.add_argument("--base", type=Path, default=None)
    args = parser.parse_args()

    base = args.base or achar_base()
    if not base or not base.exists():
        print(f"   AVISO: nao achei a SE2 em {BASES} -- alerta de valor nao rodou.")
        return 0

    hoje = dt.date.today()
    desde = hoje - dt.timedelta(days=max(args.dias, 0))
    ate = como_data(args.ate) or dt.date(hoje.year, 12, 31)
    minimo = como_numero(args.minimo) if args.minimo else MINIMO_PADRAO
    maximo = como_numero(args.maximo) if args.maximo else MAXIMO_PADRAO

    try:
        achados, resumo = ler_se2(base, hoje, desde, ate, minimo, maximo)
    except Exception as erro:  # noqa: BLE001
        print(f"   AVISO: nao consegui ler a SE2 ({erro}). A rodada segue.")
        return 0

    print(f"   SE2: {resumo['linhas']} títulos | {resumo['boletos']} de boleto "
          f"vencendo entre {desde:%d/%m/%Y} e {ate:%d/%m/%Y} | "
          f"{resumo['em_aberto']} em aberto | {resumo['com_codigo']} com código legível | "
          f"{resumo['bateram']} batendo certo"
          + (f" | {resumo['sem_codigo']} sem código" if resumo["sem_codigo"] else "")
          + (f" | {resumo['dv_ruim']} com DV inválido (fora)" if resumo["dv_ruim"] else "")
          + (f" | {resumo['abaixo_do_minimo']} abaixo do mínimo"
             if resumo["abaixo_do_minimo"] else "")
          + (f" | {resumo['acima_do_maximo']} acima do teto de {reais(maximo)}"
             if resumo["acima_do_maximo"] else ""))
    if resumo["fora_do_teto"]:
        print(f"   FORA pelo teto de {reais(maximo)} (continuam existindo):")
        for nome, dif in sorted(resumo["fora_do_teto"], key=lambda p: -abs(p[1])):
            print(f"     {com_sinal(dif):>16}  {nome[:40]}")

    if not achados:
        print("   Nenhum boleto em aberto com valor divergente do título.")
        return 0

    total_dif = sum(a["_dif"] for a in achados)
    print(f"   {len(achados)} título(s) com valor DIVERGENTE "
          f"(soma das diferenças {com_sinal(total_dif)}):")
    for a in achados[:40]:
        print(f"     venc {a['vencimento']} | {a['filial']} | "
              f"{a['titulo']}/{a['parcela'] or '-'} | título {a['valor']:>16} | "
              f"boleto {a['boleto']:>16} | {a['diferenca']:>14} | "
              f"{a['fornecedor'][:28]}")
        print(f"          UUID {a['uuid'] or '(sem código)'}  ·  LD {a['linha_dig']}")
    if len(achados) > 40:
        print(f"     ... e mais {len(achados) - 40} (todos vão no e-mail e na planilha)")

    na_tabela = achados[:LIMITE_NO_EMAIL]
    cortados = len(achados) - len(na_tabela)
    if cortados:
        print(f"   A tabela do e-mail mostra {len(na_tabela)}; os outros "
              f"{cortados} vão só na planilha.")
    corpo = montar_html(na_tabela, resumo, hoje, desde, ate,
                        cortados=cortados, total_dif=total_dif,
                        quantos_geral=len(achados), maximo=maximo)
    # 22/09/2026, pedido dele: o assunto passou a chamar a coisa pelo nome que
    # ela tem na operacao. Boleto com valor diferente do titulo e' exatamente o
    # que o banco REJEITA quando o bordero sobe -- quem le o assunto reconhece
    # o problema pelo efeito, nao pela causa.
    assunto = (f"[Painel Boletos] {len(achados)} REJEITADOS (NO BORDERÔ)"
               + (f" - até {reais(maximo)}" if maximo else "")
               + f" - {hoje:%d/%m/%Y}")

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
        print("   AVISO: nao achei para quem mandar (.alerta_valor_boleto_para).")
        return 0
    copia = args.copia or comuns.copias(CHAVE)

    if not args.forcar and comuns.ja_enviado_hoje(CHAVE):
        print(f"   Já enviado hoje ({comuns.quando_enviou(CHAVE)}) -- não mando de novo.")
        print("   Para mandar assim mesmo: --forcar")
        return 0

    try:
        comuns.enviar(assunto, corpo, destino, copia=copia,
                      anexos=[planilha] if planilha else None)
    except Exception as erro:  # noqa: BLE001
        print(f"   AVISO: o e-mail NAO foi enviado ({erro}). O painel segue.")
        return 0

    comuns.marcar_enviado(CHAVE, len(achados))
    print(f"   E-mail enviado para {destino}"
          + (f" (cópia: {copia})" if copia else "")
          + (f" com a planilha {planilha.name} em anexo." if planilha else "."))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
