# -*- coding: utf-8 -*-
"""Titulo que ENTROU na SE2 incluido por um usuario -- hoje, o rafael.lima.

Pedido dele em 21/09/2026: "para todos os titulos que Usuario Inc for
rafael.lima e entrarem na SE2, envie um e-mail automatico para ele com esses
titulos, da mesma forma que o do bordero para a Grazi".

  o recorte   coluna `Usuario Inc` da SE2, que o TOTVS grava como
              "rafael.lima - 17/04/2026" -- NOME e DATA DE INCLUSAO na mesma
              celula. O nome tem de ser o do usuario vigiado; a data e' o
              carimbo de quando o titulo entrou.
  o que sai   so o que AINDA NAO FOI AVISADO. "Entrarem na SE2" e' um evento,
              nao um estado: o titulo aparece uma vez e some da proxima lista.
              Quem guarda o que ja saiu e' DADOS/alerta_inclusao_vistos.json.
  a janela    inclusao nos ultimos 7 dias (--dias). Nao e' o filtro principal
              -- quem filtra e' o caderno de vistos -- e' a REDE: perdido o
              caderno, o alerta manda uma semana, e nao os 4.957 titulos que o
              rafael.lima ja incluiu desde abril.
  o detalhe   `Tipo pgto`, `Emissao`, `Vencimento`, `Vencto real` e o valor em
              coluna; o CODIGO (UUID), a LINHA DIGITAVEL e a baixa (quando o
              titulo ja nasceu baixado) numa faixa embaixo da linha.

Por que UUID e linha digitavel nao sao COLUNA: sao 36 e 47 caracteres numa
palavra so, que nao quebram. Cada um como coluna empurra a tabela para fora da
tela -- medido no alerta BOLETO S/C em 18/09/2026 e repetido no de bordero.
Na faixa de largura inteira os dois quebram sozinhos, continuam INTEIROS para
copiar, e a tabela nao mexe.

⚠ O CADERNO DE VISTOS SO E' GRAVADO DEPOIS QUE O E-MAIL SAI. Gravar antes (ou
no modo --teste) faria o titulo ser dado por avisado sem ninguem ter recebido
nada: ele sumiria da proxima lista e ninguem saberia que sumiu. Falha de envio
tem de deixar a lista intacta para a proxima rodada.

A MESMA lista vai ANEXADA em Excel (uma linha por titulo, valores em coluna),
com sete colunas a mais que o e-mail: natureza, prefixo, tipo, parcela, saldo,
data da baixa e codigo de barras. A planilha fica em DADOS/ -- que esta no
.gitignore, e por isso nao vai para o repositorio publico junto com fornecedor
e valor.

Vai para o proprio rafael.lima (`.alerta_inclusao_para`, fora do repo).

Uso:
    python alerta_inclusao.py                  # manda, no maximo uma vez por dia
    python alerta_inclusao.py --teste          # nao manda: grava a previa
    python alerta_inclusao.py --forcar         # manda de novo no mesmo dia
    python alerta_inclusao.py --tudo           # ignora o caderno de vistos
    python alerta_inclusao.py --dias 15        # alarga a janela
    python alerta_inclusao.py --usuario nicoly.silva
    python alerta_inclusao.py --base "C:\\caminho\\SE2.xlsx"

Codigo de saida: SEMPRE 0 quando conseguiu ler a base. Falha de e-mail nao
derruba a rodada do painel.
"""

from __future__ import annotations

import argparse
import calendar
import datetime as dt
import itertools
import json
import re
import shutil
import tempfile
from pathlib import Path

import alertas_comuns as comuns
import caminhos

PASTA = Path(__file__).resolve().parent
PREVIA = PASTA / "DADOS" / "alerta_inclusao_previa.html"
VISTOS = PASTA / "DADOS" / "alerta_inclusao_vistos.json"
CHAVE = "inclusao"

RAIZ = caminhos.raiz_lucas()
BASES = caminhos.bases_genericos()

ABA = "SE2"

# O painel da SE2 no ar: o numero do titulo na tabela abre ele. Endereco da
# RAIZ porque o painel nao le filtro pela URL (conferido em 18/09/2026).
PAINEL_SE2 = "https://se2-lucas.vercel.app/"

# Quem esta sendo vigiado. Escrito aqui e nao em arquivo separado porque e' o
# ASSUNTO do alerta, nao um segredo: o que nao pode ser versionado e' o
# ENDERECO de e-mail, e esse fica em .alerta_inclusao_para.
USUARIO_PADRAO = "rafael.lima"

# 21/09/2026, recorte que ele fechou depois do primeiro disparo: so contrato de
# MED/CT. Sao os titulos de seguro, plano de saude e concessionaria -- Unimed,
# Sompo, Cemig, Scania, Bradesco Auto/RE -- que tem vencimento fixo e nao podem
# passar batido. De 4.972 titulos do rafael.lima na base, 111 sao MED+CT.
PREFIXO_PADRAO = "MED"
TIPO_PADRAO = "CT"

# ---------------------------------------------------------------------------
# OS DOIS MODOS -- e por que o alerta mudou de natureza em 21/09/2026.
#
# "vencimento" (o padrao agora, pedido dele): janela que ROLA com o calendario
#   -- 7 dias para tras e ate o FIM DO PROXIMO MES. E um ESTADO, como os outros
#   quatro alertas do painel: a mesma lista sai todo dia enquanto os titulos
#   continuarem dentro da janela. Sem caderno de vistos, porque "todos os
#   titulos" foi o pedido: esconder o que ja saiu ontem seria mentir.
#
# "inclusao" (o de manha, mantido): janela pela DATA DE INCLUSAO, com caderno
#   de vistos. E um EVENTO -- o titulo entrou -- e cada um sai uma vez so.
#
# Os dois tem TRAVA DIARIA SEPARADA: chaves diferentes em alertas_enviados.json.
# Misturar as duas apagaria uma da outra -- o de manha marcaria o dia como
# avisado e o da tarde ficaria calado, sem ninguem entender por que.
# ---------------------------------------------------------------------------
# `sem_baixados` -- 21/09/2026, pedido dele depois do primeiro disparo por
# vencimento: "nao precisa enviar os que ja estiverem baixados". No modo
# VENCIMENTO a lista e' de coisa A PAGAR, e titulo com DT Baixa ja foi pago:
# deixar ele ali so faz o Rafael conferir duas vezes o que ja esta resolvido.
# No modo INCLUSAO continua entrando, porque la a lista e' a conferencia do que
# foi LANCADO -- e um titulo que nasceu ja baixado e' justamente o que merece
# uma olhada.
MODOS = {
    "vencimento": {"campo": "_venc", "rotulo": "vencimento",
                   "trava": "inclusao_vencimento", "caderno": False,
                   "sem_baixados": True},
    "inclusao": {"campo": "_data", "rotulo": "inclusão",
                 "trava": "inclusao", "caderno": True,
                 "sem_baixados": False},
}
MODO_PADRAO = "vencimento"

# Quantos dias para TRAS a janela pega. Para a FRENTE quem manda e o fim do
# proximo mes (ver `fim_do_proximo_mes`), e nao um numero de dias: ele pensa o
# vencimento por mes fechado -- "ate o fim do proximo mes" -- e uma contagem de
# dias faria a janela cortar um mes no meio, diferente a cada dia da semana.
DIAS_PADRAO = 7

# Quanto tempo uma chave fica no caderno de vistos. Seis meses e' folgado para
# a janela de 7 dias e impede o arquivo de crescer para sempre.
LEMBRAR_DIAS = 180

# Quantos titulos a TABELA DO E-MAIL mostra. A planilha anexada leva SEMPRE
# todos -- este teto e' so do corpo da mensagem.
#
# Cada linha custa ~1,7 KB de HTML (estilo em cada celula, porque o Outlook
# ignora <style>, mais a faixa de UUID e linha digitavel). Medido em
# 21/09/2026: 318 titulos deram 535 KB de corpo. Cliente de e-mail corta corpo
# grande SEM AVISAR -- e alerta cortado em silencio e' pior que alerta curto.
# Aqui o corte e' explicito e vem escrito embaixo da tabela.
#
# No dia a dia isso nunca pega: o rafael.lima inclui de 10 a 220 titulos por
# dia, e o alerta so leva os ainda nao avisados. Quem enche e' o primeiro
# disparo, que vem com a janela inteira de uma vez.
LIMITE_NO_EMAIL = 150

# Conferidas pelo NOME no cabecalho, nunca por posicao: o TOTVS ja inseriu
# coluna no meio da exportacao antes, e indice fixo erraria tudo em silencio.
COLUNAS = {
    "filial": "Filial", "prefixo": "Prefixo", "tipo": "Tipo",
    "num": "No. Titulo", "parcela": "Parcela", "fornecedor": "Fornecedor",
    "nome": "Nome Fornece", "razao": "Razão Social", "natureza": "Natureza",
    "emissao": "DT Emissao",
    "vencimento": "Vencimento", "real": "Vencto Real",
    "valor": "Vlr.Titulo", "baixa": "DT Baixa", "saldo": "Saldo",
    "tipo_pgto": "Tipo Pgto", "linha_dig": "Linha Dig.",
    "cod_barras": "Cod.Barras", "uuid": "Campo UUID",
    "usuario_inc": "Usuario Inc",
}

DIAS_SEMANA = ("segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
               "sexta-feira", "sábado", "domingo")

MESES = ("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro")

# "rafael.lima - 17/04/2026" -> nome e data. O mesmo formato que o painel da
# SE2 ja separa (splitUserDate em PAINEIS/SE2/REPOSITORIO/app.js): nome, um
# hifen cercado de espacos, e a data no fim.
SEPARA_USUARIO = re.compile(r"^(?P<nome>.+?)\s*-\s*(?P<data>\d{2}/\d{2}/\d{4})\s*$")


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
        texto_limpo = valor.strip().replace(".", "").replace(",", ".")
        try:
            return float(texto_limpo)
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


def fim_do_proximo_mes(hoje: dt.date) -> dt.date:
    """O ultimo dia do mes SEGUINTE ao de hoje.

    Em 21/09/2026 devolve 31/10/2026. Vira o ano sozinho: em dezembro aponta
    para 31/01 do ano que vem. Escrito assim, e nao como "hoje + 45 dias",
    porque e' o mes que ele fecha na cabeca -- somar dias faria a janela
    terminar no meio de um mes, num dia diferente a cada semana.
    """
    proximo = dt.date(hoje.year + (hoje.month == 12), hoje.month % 12 + 1, 1)
    ultimo = calendar.monthrange(proximo.year, proximo.month)[1]
    return dt.date(proximo.year, proximo.month, ultimo)


def separa_usuario(valor) -> tuple[str, dt.date | None]:
    """"rafael.lima - 17/04/2026" -> ("rafael.lima", date(2026, 4, 17)).

    Celula sem a data devolve (nome, None): o titulo continua sendo daquele
    usuario, so nao da para dizer QUANDO entrou -- e sem data ele nao entra na
    janela, porque nao ha como saber se e' de hoje ou de 2024.
    """
    bruto = texto(valor)
    if not bruto:
        return "", None
    achou = SEPARA_USUARIO.match(bruto)
    if not achou:
        return bruto, None
    return achou.group("nome").strip(), como_data(achou.group("data"))


def chave_do_titulo(achado: dict) -> str:
    """A identidade do titulo no caderno de vistos.

    Filial + prefixo + tipo + numero + parcela -- a chave do titulo no TOTVS,
    e nao o UUID: o UUID e' campo de apoio e pode vir vazio ou ser trocado,
    e chave que some faz o titulo ser avisado duas vezes.
    """
    return "|".join((achado["filial"], achado["prefixo"], achado["tipo"],
                     achado["titulo"], achado["parcela"]))


def ler_vistos() -> dict:
    """{chave do titulo: dia em que foi avisado}. Arquivo ilegivel = caderno
    em branco: melhor repetir um aviso do que engolir um titulo novo."""
    if not VISTOS.exists():
        return {}
    try:
        dados = json.loads(VISTOS.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        print("   (o caderno de vistos estava ilegível; comecei um novo)")
        return {}
    return dados.get("avisados", {}) if isinstance(dados, dict) else {}


def gravar_vistos(avisados: dict, novas: list[str], hoje: dt.date) -> None:
    """Fecha o dia: as chaves novas entram, e as velhas demais saem.

    ⚠ So e' chamada DEPOIS que o Outlook aceitou o e-mail. Ver o cabecalho.
    """
    corte = hoje - dt.timedelta(days=LEMBRAR_DIAS)
    guardadas = {c: d for c, d in avisados.items()
                 if (como_data(d) or hoje) >= corte}
    for chave in novas:
        guardadas[chave] = hoje.isoformat()
    VISTOS.parent.mkdir(parents=True, exist_ok=True)
    VISTOS.write_text(json.dumps({"avisados": guardadas}, ensure_ascii=False,
                                 indent=2, sort_keys=True), encoding="utf-8")


def ler_se2(caminho: Path, hoje: dt.date, usuario: str, desde: dt.date,
            ate: dt.date, modo: str = MODO_PADRAO, prefixo: str = PREFIXO_PADRAO,
            tipo: str = TIPO_PADRAO,
            sem_baixados: bool = True) -> tuple[list[dict], dict]:
    """Varre a SE2 e devolve os titulos do usuario que caem na janela.

    `modo` decide QUAL data e olhada na janela: "vencimento" (padrao) ou
    "inclusao". `prefixo` e `tipo` peneiram antes disso; "" em qualquer um
    desliga aquele filtro.

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
        temporario = Path(tempfile.gettempdir()) / f"se2_inclusao_{dt.date.today():%Y%m%d}.xlsx"
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
    for obrigatoria in ("usuario_inc", "num", "filial", "valor"):
        if obrigatoria not in onde:
            arquivo.close()
            raise KeyError("colunas ausentes na SE2: " + ", ".join(faltando))

    def campo(linha, chave):
        i = onde.get(chave)
        return linha[i] if i is not None and i < len(linha) else None

    alvo = usuario.strip().lower()
    filtro_prefixo = prefixo.strip().upper()
    filtro_tipo = tipo.strip().upper()
    achados = []
    resumo = {"linhas": 0, "do_usuario": 0, "no_recorte": 0, "sem_data_inc": 0,
              "sem_data_janela": 0, "na_janela": 0, "fora_da_janela": 0,
              "com_linha_dig": 0, "ja_baixados": 0}
    for linha in linhas:
        if campo(linha, "num") in (None, ""):
            continue
        resumo["linhas"] += 1

        quem, quando = separa_usuario(campo(linha, "usuario_inc"))
        if quem.lower() != alvo:
            continue
        resumo["do_usuario"] += 1

        # Prefixo e tipo ANTES da data: sao os filtros que mais cortam (111 de
        # 4.972 em 21/09/2026), e cortar cedo deixa os contadores do rodape
        # falando do que interessa -- "quantos MED/CT ficaram fora da janela",
        # e nao "quantos titulos do rafael.lima existem".
        if filtro_prefixo and texto(campo(linha, "prefixo")).upper() != filtro_prefixo:
            continue
        if filtro_tipo and texto(campo(linha, "tipo")).upper() != filtro_tipo:
            continue
        resumo["no_recorte"] += 1

        if not quando:
            # Celula sem o carimbo de data de inclusao.
            resumo["sem_data_inc"] += 1

        vencimento_bruto = como_data(campo(linha, "vencimento"))
        referencia = quando if modo == "inclusao" else vencimento_bruto
        if not referencia:
            # Sem a data que define a janela nao da para dizer se o titulo
            # entra ou nao -- e chutar aqui e' pior que deixar de fora, porque
            # ninguem saberia que foi chutado. O rodape do e-mail conta quantos.
            resumo["sem_data_janela"] += 1
            continue
        if referencia < desde or referencia > ate:
            resumo["fora_da_janela"] += 1
            continue
        resumo["na_janela"] += 1

        baixa = como_data(campo(linha, "baixa"))
        if baixa:
            resumo["ja_baixados"] += 1
            # ⚠ O descarte vem DEPOIS de contar, de proposito: o rodape do
            # e-mail diz quantos ficaram de fora por ja estarem pagos. Sumir em
            # silencio faria a lista encolher sem ninguem saber por que.
            if sem_baixados:
                continue
        linha_dig = texto(campo(linha, "linha_dig"))
        if linha_dig:
            resumo["com_linha_dig"] += 1

        emissao = como_data(campo(linha, "emissao"))
        vencimento = vencimento_bruto
        real = como_data(campo(linha, "real"))
        valor = como_numero(campo(linha, "valor"))
        # ⚠ VENCIDO se mede pelo VENCTO REAL, nao pelo Vencimento: o real e' a
        # data que o banco cobra (o TOTVS empurra o vencimento que cai em
        # sabado, domingo ou feriado -- e' o que o alerta_dia_util.py vigia).
        # Titulo ja baixado nunca e' vencido: foi pago.
        prazo = real or vencimento
        atraso = (hoje - prazo).days if (prazo and not baixa) else 0
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
            "usuario": quem,
            "inclusao": f"{quando:%d/%m/%Y}" if quando else "",
            "tipo_pgto": texto(campo(linha, "tipo_pgto")),
            "emissao": f"{emissao:%d/%m/%Y}" if emissao else "",
            "vencimento": f"{vencimento:%d/%m/%Y}" if vencimento else "",
            "real": f"{real:%d/%m/%Y}" if real else "",
            "baixa": f"{baixa:%d/%m/%Y}" if baixa else "",
            "situacao": ("Baixado" if baixa else
                         (f"Vencido há {atraso}d" if atraso > 0 else "Em aberto")),
            "valor": reais(valor),
            "saldo": reais(como_numero(campo(linha, "saldo"))),
            "linha_dig": linha_dig,
            "cod_barras": texto(campo(linha, "cod_barras")),
            "_data": quando,
            # a data que mandou este titulo entrar na lista, e a faixa em que
            # ele vai aparecer: as duas mudam com o modo
            "_ref": referencia,
            "grupo": (f"{referencia:%m/%Y}" if modo == "vencimento"
                      else f"{referencia:%d/%m/%Y}"),
            "_modo": modo, "_atraso": atraso,
            # crus para a planilha: no Excel data tem de ser data e valor tem
            # de ser numero, senao nao soma nem ordena
            "_valor": valor, "_saldo": como_numero(campo(linha, "saldo")),
            "_emissao": emissao, "_venc": vencimento, "_real": real,
            "_baixa": baixa,
        })
    arquivo.close()
    if temporario and temporario.exists():
        try:
            temporario.unlink()
        except OSError:
            pass
    resumo["base"] = origem
    resumo["salva_em"] = dt.datetime.fromtimestamp(origem.stat().st_mtime)
    # O mais ANTIGO primeiro -- no modo vencimento, e' o que vence antes (ou ja
    # venceu); no modo inclusao, o que esta esperando conferencia ha mais tempo.
    # Dentro do dia, por filial e numero: e' assim que se procura um titulo na
    # tela do TOTVS.
    achados.sort(key=lambda a: (a["_ref"], a["filial"], a["titulo"], a["parcela"]))
    return achados, resumo


# ⚠ Largura: oito colunas, o mesmo teto do alerta de bordero (~830px no painel
# de leitura do Outlook). Titulo e parcela numa coluna so ("000012576/04"),
# como se fala do titulo; "Tipo pgto" sem destaque, porque destaque = negrito +
# nao quebra, e "AMARRAR ADIANTAMENTO" sozinho segurava 188px de coluna.
# A DATA DE INCLUSAO nao e' coluna: ela e' o titulo do grupo -- repeti-la em
# cada linha seria gastar 91px dizendo o que a faixa logo acima ja diz.
# 21/09/2026, pedido dele: "informe o vencimento e vencimento real". Os DOIS
# em coluna, lado a lado -- eles divergem com frequencia neste recorte (o
# 399019094 vence 15/09 e o real e 21/09, porque o TOTVS empurra o que cai em
# dia nao util) e mostrar so um esconde justamente a diferenca que importa.
# A EMISSAO saiu para os dois caberem sem alargar a tabela: aqui o que se olha
# e' o que vai vencer, nao quando a nota nasceu. Ela continua na planilha.
COLUNAS_EMAIL = [
    ("Filial", "filial"),
    ("Nº título", "titulo_parcela"),
    ("Fornecedor", "fornecedor"),
    ("Tipo pgto", "tipo_pgto"),
    ("Vencimento", "vencimento"),
    ("Vencto real", "real"),
    ("Situação", "situacao"),
    ("Vlr. título", "valor"),
]


def faixa_detalhe(linha: dict) -> str:
    """A faixa de largura inteira embaixo de cada linha da tabela.

    Leva o CODIGO (UUID) e, quando o titulo tem boleto, a LINHA DIGITAVEL --
    os dois aqui embaixo, e nao em coluna, porque sao palavras de 36 e 47
    caracteres que nao quebram: cada uma como coluna joga o resto da tabela
    para fora da tela. Na faixa eles quebram sozinhos e continuam inteiros
    para copiar.

    A BAIXA entra aqui tambem, e so quando existe: titulo que ja nasceu
    baixado e' o caso raro, e caso raro nao merece uma coluna inteira em
    branco nas outras linhas -- merece uma frase onde ele acontece.
    """
    partes = []
    if linha.get("emissao"):
        partes.append("Emissão: " + linha["emissao"])
    if linha.get("inclusao"):
        # No modo vencimento a data de inclusao perdeu a faixa de grupo, mas
        # continua sendo o que liga o titulo a quem o lancou -- e o motivo de
        # este e-mail chegar nele, e nao em outra pessoa.
        partes.append(f"Incluído em {linha['inclusao']} por {linha.get('usuario', '')}".strip())
    if linha.get("natureza"):
        partes.append("Natureza: " + linha["natureza"])
    if linha.get("baixa"):
        partes.append("JÁ BAIXADO em " + linha["baixa"])
    if linha.get("uuid"):
        partes.append("Código (UUID): " + linha["uuid"])
    if linha.get("linha_dig"):
        partes.append("Linha digitável: " + linha["linha_dig"])
    return "     ·     ".join(partes)


COLUNAS_EXCEL = [
    ("Dt. Inclusão", "_data", "data"),
    ("Usuário Inc", "usuario", "texto"),
    ("Filial", "filial", "texto"),
    ("Prefixo", "prefixo", "texto"),
    ("Tipo", "tipo", "texto"),
    ("Nº Título", "titulo", "texto"),
    ("Parcela", "parcela", "texto"),
    ("Fornecedor", "fornecedor", "texto"),
    ("Natureza", "natureza", "texto"),
    ("Tipo Pgto", "tipo_pgto", "texto"),
    ("DT Emissão", "_emissao", "data"),
    ("Vencimento", "_venc", "data"),
    ("Vencto Real", "_real", "data"),
    ("Dias p/ vencer", "_dias", "inteiro"),
    ("Vlr. Título", "_valor", "moeda"),
    ("Saldo", "_saldo", "moeda"),
    ("DT Baixa", "_baixa", "data"),
    ("Situação", "situacao", "texto"),
    ("Linha Digitável", "linha_dig", "texto"),
    ("Cod. Barras", "cod_barras", "texto"),
    ("Código (UUID)", "uuid", "texto"),
]

LARGURA = {"Dt. Inclusão": 12, "Usuário Inc": 16, "Filial": 7, "Prefixo": 9,
           "Tipo": 7, "Nº Título": 13, "Parcela": 9, "Fornecedor": 38,
           "Natureza": 11,
           # 25 e nao 22: "AMARRAR ADIANTAMENTO" em negrito nao cabia em 22 e
           # saia cortado no papel (conferido no PDF em 21/09/2026)
           "Tipo Pgto": 25, "DT Emissão": 12, "Vencimento": 12,
           "Vencto Real": 12, "Dias p/ vencer": 14, "Vlr. Título": 14,
           "Saldo": 14, "DT Baixa": 12, "Situação": 12,
           "Linha Digitável": 50, "Cod. Barras": 48, "Código (UUID)": 38}

CENTRALIZADAS = {"Dt. Inclusão", "Filial", "Prefixo", "Tipo", "Parcela",
                 "Natureza", "DT Emissão", "Vencimento", "Vencto Real",
                 "Dias p/ vencer", "DT Baixa", "Situação"}

# A MESMA paleta do corpo do e-mail, de proposito: quem abre o anexo tem de
# reconhecer na hora a lista que acabou de ler. Azul do cabecalho, cinza-azulado
# da faixa de grupo, zebrado e borda sao os mesmos codigos de alertas_comuns.
AZUL = "1F3864"
FAIXA_GRUPO = "D6DCE4"
ZEBRA = "F2F2F2"
BORDA = "BFBFBF"

# Fundo e letra por tipo de pagamento. A cor nao decora: e' por ela que se ve
# de longe que um lote inteiro entrou como TRANSFERENCIA e um unico titulo no
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


def montar_resumo(arquivo, achados: list[dict]) -> None:
    """Uma segunda aba com as duas contas que se faz de cabeca ao abrir a lista:
    quanto entrou em cada DIA e quanto entrou em cada TIPO DE PAGAMENTO. So
    numero -- nenhuma instrucao, nenhum recado.
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

    larguras = [16, 18, 10, 16, 16, 13, 12]
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

    por_vencimento = achados and achados[0].get("_modo") == "vencimento"
    primeira = "Vencimento" if por_vencimento else "Dt. Inclusão"
    segunda = "Vencidos" if por_vencimento else "Dia da semana"
    # Baixado nenhum na lista = a coluna "Já baixados" seria uma coluna de
    # zeros. No lugar dela vai o SALDO VENCIDO, que e' o numero que se procura
    # numa lista de coisa a pagar: quanto ja passou da data.
    tem_baixados = any(a["baixa"] for a in achados)
    sexta = "Já baixados" if tem_baixados else "Saldo vencido"
    formato_6 = "0" if tem_baixados else '#,##0.00'

    def conta_6(linhas_do_grupo):
        if tem_baixados:
            return sum(1 for l in linhas_do_grupo if l["baixa"])
        return sum(l["_saldo"] for l in linhas_do_grupo if l["_atraso"] > 0)

    linha = cabecalho(1, [primeira, segunda, "Títulos",
                          "Vlr. Título", "Saldo", sexta, "Com boleto"])
    for chave, grupo in itertools.groupby(achados, key=lambda a: a["grupo"]):
        do_grupo = list(grupo)
        quando = do_grupo[0]["_ref"]
        if por_vencimento:
            rotulo = f"{MESES[quando.month - 1]}/{quando.year}"
            segundo = sum(1 for l in do_grupo if l["_atraso"] > 0)
            formato_2 = "0"
        else:
            rotulo = chave
            segundo = DIAS_SEMANA[quando.weekday()]
            formato_2 = "@"
        linha = numeros(linha, [
            rotulo, segundo, len(do_grupo),
            sum(l["_valor"] for l in do_grupo),
            sum(l["_saldo"] for l in do_grupo),
            conta_6(do_grupo),
            sum(1 for l in do_grupo if l["linha_dig"]),
        ], ["@", formato_2, "0", '#,##0.00', '#,##0.00', formato_6, "0"])
    linha = numeros(linha, [
        "TOTAL",
        sum(1 for a in achados if a["_atraso"] > 0) if por_vencimento else "",
        len(achados),
        sum(a["_valor"] for a in achados), sum(a["_saldo"] for a in achados),
        conta_6(achados),
        sum(1 for a in achados if a["linha_dig"]),
    ], ["@", "0" if por_vencimento else "@", "0", '#,##0.00', '#,##0.00',
        formato_6, "0"],
        negrito=True, fundo=fundo_faixa)

    linha += 2
    # o nome do tipo ocupa as duas primeiras colunas (A:B), senao "AMARRAR
    # ADIANTAMENTO" nao cabe e a coluna B fica um buraco azul no cabecalho
    cabeca_tipo = linha
    linha = cabecalho(linha, ["Tipo Pgto", "", "Títulos", "Vlr. Título",
                              "Saldo", sexta, "Com boleto"])
    pagina.merge_cells(start_row=cabeca_tipo, start_column=1,
                       end_row=cabeca_tipo, end_column=2)
    por_tipo: dict[str, list[dict]] = {}
    for a in achados:
        por_tipo.setdefault(a["tipo_pgto"] or "—", []).append(a)
    for tipo, do_tipo in sorted(por_tipo.items(),
                                key=lambda par: -sum(l["_valor"] for l in par[1])):
        fundo, letra = cor_do_tipo(tipo)
        linha = numeros(linha, [
            tipo, "", len(do_tipo),
            sum(l["_valor"] for l in do_tipo), sum(l["_saldo"] for l in do_tipo),
            conta_6(do_tipo),
            sum(1 for l in do_tipo if l["linha_dig"]),
        ], ["@", "@", "0", '#,##0.00', '#,##0.00', formato_6, "0"])
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


def gerar_planilha(achados: list[dict], hoje: dt.date, usuario: str,
                   modo: str = MODO_PADRAO) -> Path | None:
    """A mesma lista do e-mail, em Excel, para filtrar e somar.

    UMA linha por titulo e os valores em coluna; sem aba de instrucao e sem
    bloco de texto -- planilha e' dado, nao recado.

    Separa os titulos do mesmo jeito que o corpo do e-mail: uma faixa por DIA
    de inclusao, com o mesmo texto (`rotulo_do_grupo`) e a mesma paleta. Cada
    grupo e' tambem um grupo de verdade do Excel: dobra e desdobra no +/- da
    margem.

    Data vai como DATA e valor como NUMERO (o e-mail manda os dois ja
    formatados; aqui isso nao serve, porque no Excel texto nao soma). Ja a
    linha digitavel, o codigo de barras e o numero do titulo vao como TEXTO
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
    pagina.title = ("TITULOS A VENCER" if modo == "vencimento"
                    else "TITULOS INCLUIDOS")
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
    # titulo que ja nasceu baixado e' o caso que foge do comum nesta lista:
    # em vez de sumir no meio dos numeros, vem com a cor do alerta
    fundo_baixado = PatternFill("solid", fgColor="E2EFDA")
    letra_baixada = Font(name="Calibri", bold=True, color="375623")

    for coluna, rotulo in enumerate(rotulos, start=1):
        celula = pagina.cell(1, coluna, rotulo)
        celula.fill, celula.font = fundo_cab, letra_cab
        celula.border = grade
        celula.alignment = Alignment(vertical="center", horizontal="center"
                                     if rotulo in CENTRALIZADAS else "left")
        pagina.column_dimensions[get_column_letter(coluna)].width = LARGURA.get(rotulo, 14)
    pagina.row_dimensions[1].height = 22

    linha = 2
    for chave, grupo in itertools.groupby(achados, key=lambda a: a["grupo"]):
        do_grupo = list(grupo)
        # A faixa do grupo e' a MESMA do corpo do e-mail, texto inclusive
        # (rotulo_do_grupo): o mes (ou o dia), quantos titulos e a soma.
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
                    valor = ((achado["_real"] - hoje).days
                             if achado["_real"] else None)
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
                elif rotulo == "Vlr. Título":
                    celula.font = Font(name="Calibri", bold=True)
                elif rotulo == "Situação" and achado["baixa"]:
                    celula.fill, celula.font = fundo_baixado, letra_baixada
                elif rotulo in ("Linha Digitável", "Cod. Barras", "Código (UUID)"):
                    # dado de copiar, nao de ler: menor e mais apagado, para
                    # nao competir com o nome e o valor
                    celula.font = Font(name="Calibri", size=9, color="404040")
            # cada dia de inclusao vira um grupo que dobra no +/- do Excel
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
                f"em {len({a['inclusao'] for a in achados})} dia(s) de inclusão")
    for rotulo in ("Vlr. Título", "Saldo"):
        coluna = rotulos.index(rotulo) + 1
        letra_coluna = get_column_letter(coluna)
        celula = pagina.cell(linha, coluna,
                             f"=SUM({letra_coluna}2:{letra_coluna}{ultima_de_dados})")
        celula.number_format = '#,##0.00'
        celula.fill, celula.font, celula.border = fundo_cab, letra_cab, grade
    pagina.row_dimensions[linha].height = 20

    # vermelho -> amarelo -> verde conforme o vencimento se afasta: o vencido
    # (numero negativo) e' o que precisa de olho hoje.
    coluna_dias = get_column_letter(rotulos.index("Dias p/ vencer") + 1)
    pagina.conditional_formatting.add(
        f"{coluna_dias}2:{coluna_dias}{ultima_de_dados}",
        ColorScaleRule(start_type="min", start_color="FFC7CE",
                       mid_type="percentile", mid_value=50, mid_color="FFEB9C",
                       end_type="max", end_color="C6EFCE"))

    pagina.sheet_properties.outlinePr.summaryBelow = False
    pagina.freeze_panes = "A2"
    pagina.auto_filter.ref = f"A1:{ultima_letra}{ultima_de_dados}"

    montar_resumo(arquivo, achados)

    nome = "TITULOS A VENCER" if modo == "vencimento" else "TITULOS INCLUIDOS"
    destino = (PASTA / "DADOS"
               / f"{nome} {usuario.upper()} {hoje:%d-%m-%Y}.xlsx")
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
    """O texto da faixa que abre cada grupo.

    No modo vencimento a faixa e o MES ("Vencimento em outubro/2026"), e nao o
    dia: a janela cobre uns 45 dias e uma faixa por data daria mais faixa que
    linha -- em 21/09/2026 seriam 13 faixas para 19 titulos. Por mes sao duas,
    e e' assim que ele fala do vencimento ("o que vence em outubro").
    No modo inclusao continua uma faixa por dia, com o dia da semana: la o
    volume e outro (dezenas por dia) e o dia e' a propria noticia.

    Nos dois casos a faixa leva a conta do grupo -- quantos titulos e quanto
    somam -- porque e' o primeiro numero que se procura.
    """
    total = sum(l["_valor"] for l in linhas)
    quantos = len(linhas)
    quando = linhas[0]["_ref"]
    plural = "s" if quantos > 1 else ""
    if linhas[0].get("_modo") == "vencimento":
        vencidos = sum(1 for l in linhas if l["_atraso"] > 0)
        aviso = f" · {vencidos} vencido{'s' if vencidos > 1 else ''}" if vencidos else ""
        return (f"Vencimento em {MESES[quando.month - 1]}/{quando.year} — "
                f"{quantos} título{plural} · {reais(total)}{aviso}")
    return (f"Incluídos em {chave} ({DIAS_SEMANA[quando.weekday()]}) — "
            f"{quantos} título{plural} · {reais(total)}")


def montar_html(achados: list[dict], resumo: dict, hoje: dt.date,
                desde: dt.date, ate: dt.date, usuario: str, repetidos: int,
                modo: str = MODO_PADRAO, prefixo: str = PREFIXO_PADRAO,
                tipo: str = TIPO_PADRAO, sem_baixados: bool = True,
                cortados: int = 0, total_geral: float = 0.0) -> str:
    """O corpo do e-mail. `achados` ja vem cortado em LIMITE_NO_EMAIL;
    `cortados` diz quantos ficaram so na planilha, e `total_geral` e' a soma
    de TODOS -- o numero do texto de abertura tem de ser o da lista inteira,
    nao o do pedaco que coube."""
    por_vencimento = modo == "vencimento"
    quantos = len(achados) + cortados
    total = total_geral or sum(a["_valor"] for a in achados)
    com_ld = sum(1 for a in achados if a["linha_dig"])
    baixados = sum(1 for a in achados if a["baixa"])
    vencidos = sum(1 for a in achados if a["_atraso"] > 0)
    recorte = " e ".join(p for p in (f"prefixo <b>{prefixo}</b>" if prefixo else "",
                                     f"tipo <b>{tipo}</b>" if tipo else "") if p)

    if por_vencimento:
        frase = ("1 título seu vence nesta janela" if quantos == 1 else
                 f"{quantos} títulos seus vencem nesta janela")
        periodo = (f" — vencimento de <b>{desde:%d/%m/%Y}</b> até "
                   f"<b>{ate:%d/%m/%Y}</b> (fim do próximo mês)")
        aviso_vencidos = ("" if not vencidos else
                          f" <b>{vencidos}</b> "
                          + ("já está vencido" if vencidos == 1 else "já estão vencidos")
                          + " e sem baixa — aparecem marcados na coluna Situação.")
    else:
        dias = len({a["grupo"] for a in achados})
        mais_antigo = achados[0]["inclusao"] if achados else ""
        frase = ("1 título novo entrou na SE2 no seu usuário" if quantos == 1 else
                 f"{quantos} títulos novos entraram na SE2 no seu usuário")
        periodo = ("" if dias <= 1 else
                   f", em {dias} dias de inclusão (o mais antigo em <b>{mais_antigo}</b>)")
        aviso_vencidos = ""
    if sem_baixados:
        # Quantos ficaram de fora por ja estarem pagos. Dizer o numero e' o que
        # separa "nao ha mais nada" de "o resto ja foi pago": sem essa linha,
        # uma lista curta parece base incompleta.
        fora = resumo.get("ja_baixados", 0)
        ja_baixados = ("" if not fora else
                       f" Outro{'s' if fora > 1 else ''} <b>{fora}</b> "
                       + ("venceu" if fora == 1 else "venceram")
                       + " nesta janela e já "
                       + ("está pago" if fora == 1 else "estão pagos")
                       + ", e por isso não " + ("entrou" if fora == 1 else "entraram")
                       + " na lista.")
    else:
        ja_baixados = ("" if not baixados else
                       f" <b>{baixados}</b> "
                       + ("já está baixado" if baixados == 1 else "já estão baixados")
                       + " — a baixa aparece na faixa embaixo da linha.")
    # O corte fica ESCRITO, e nao escondido: quem le tem de saber que a tabela
    # nao e' a lista inteira, e onde esta o resto.
    aviso_corte = ("" if not cortados else
                   f"""
  <p style="margin:12px 0 0 0;padding:8px 10px;background:#FFF2CC;
  border-left:4px solid #BF8F00;font-size:10pt;color:#7F6000;">
  A tabela acima mostra os <b>{len(achados)}</b> primeiros (os mais antigos).
  Os outros <b>{cortados}</b> estão na <b>planilha anexada</b>, que traz
  sempre a lista completa — corpo de e-mail muito grande é cortado pelo
  próprio programa de e-mail.</p>""")
    return f"""<div style="{comuns.FONTE}font-size:11pt;color:#000000;">
  <p style="margin:0 0 12px 0;">Bom dia, Rafael,</p>

  <p style="margin:0 0 12px 0;">{frase} — lançados no seu usuário
  <b>{usuario}</b>, {recorte}{periodo} — somando
  <b>{reais(total)}</b>.{aviso_vencidos}{ja_baixados}</p>

  <p style="margin:0 0 14px 0;">Cada linha traz o <i>fornecedor</i>, o
  <i>tipo de pagamento</i>, o <b>vencimento</b> e o <b>vencimento real</b>
  lado a lado — eles divergem quando a data cai em dia não útil, e é o
  <i>real</i> que o banco cobra. Os títulos vão <b>agrupados pelo mês de
  vencimento</b>, do que vence antes para o que vence depois, com a soma de
  cada mês. Na faixa abaixo de cada linha vão a <i>emissão</i>, a <i>data de
  inclusão</i>, a <b>natureza</b>, o <b>código (UUID)</b> e, quando o título
  tem boleto, a <b>linha digitável</b> ({com_ld} de {quantos}). O <b>nº do
  título</b> abre o
  <a href="{PAINEL_SE2}" style="color:#1F3864;">painel da SE2</a>, onde o
  <b>código (UUID)</b> da faixa serve de busca. A mesma lista vai
  <b>anexada em Excel</b>, com o saldo, a data da baixa e o código de barras,
  para quem preferir filtrar e somar.</p>

  {comuns.tabela(COLUNAS_EMAIL, achados,
                 direita=("valor",),
                 destaque=("real",),
                 links={"titulo_parcela": PAINEL_SE2},
                 sublinha=faixa_detalhe,
                 grupo=lambda l: l["grupo"],
                 grupo_rotulo=rotulo_do_grupo)}
  {aviso_corte}

  {comuns.rodape(
      f"Recorte: Usuário Inc = {usuario}"
      + (f", prefixo {prefixo}" if prefixo else "")
      + (f", tipo {tipo}" if tipo else "")
      + (f", vencimento de {desde:%d/%m/%Y} a {ate:%d/%m/%Y}"
         if por_vencimento else
         f", incluídos de {desde:%d/%m/%Y} a {ate:%d/%m/%Y}, "
         "apenas os ainda não avisados")
      + (", sem os que já têm data de baixa." if sem_baixados else "."),
      f"Base SE2 salva em {resumo['salva_em']:%d/%m/%Y às %H:%M} — "
      f"{resumo['do_usuario']} títulos de {usuario} na base, "
      f"{resumo['no_recorte']} no recorte, "
      f"{resumo['na_janela']} dentro da janela"
      + (f"; {resumo['ja_baixados']} fora por já estarem baixados"
         if sem_baixados and resumo.get("ja_baixados") else "")
      + (f"; {repetidos} já avisado(s) em envio anterior" if repetidos else "")
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
    parser.add_argument("--tudo", action="store_true",
                        help="ignora o caderno de vistos e manda a janela inteira")
    parser.add_argument("--usuario", default=USUARIO_PADRAO,
                        help=f"usuario da coluna Usuario Inc (padrao: {USUARIO_PADRAO})")
    parser.add_argument("--por", choices=sorted(MODOS), default=MODO_PADRAO,
                        help=f"qual data manda na janela (padrao: {MODO_PADRAO})")
    parser.add_argument("--prefixo", default=PREFIXO_PADRAO,
                        help=f'prefixo do titulo; "" para todos (padrao: {PREFIXO_PADRAO})')
    parser.add_argument("--tipo", default=TIPO_PADRAO,
                        help=f'tipo do titulo; "" para todos (padrao: {TIPO_PADRAO})')
    parser.add_argument("--dias", type=int, default=DIAS_PADRAO,
                        help=f"quantos dias a janela pega para TRAS (padrao: {DIAS_PADRAO})")
    parser.add_argument("--ate", default="",
                        help="fim da janela (dd/mm/aaaa); o padrao e o fim do proximo mes")
    parser.add_argument("--com-baixados", action="store_true",
                        help="traz tambem os titulos que ja tem DT Baixa")
    parser.add_argument("--base", type=Path, default=None)
    args = parser.parse_args()

    base = args.base or achar_base()
    if not base or not base.exists():
        print(f"   AVISO: nao achei a SE2 em {BASES} -- alerta de inclusão nao rodou.")
        return 0

    hoje = dt.date.today()
    usuario = args.usuario.strip()
    modo = args.por
    regras = MODOS[modo]
    desde = hoje - dt.timedelta(days=max(args.dias, 0))
    # No modo INCLUSAO a janela nao pode passar de hoje: ninguem inclui titulo
    # no futuro, e uma ponta la na frente so daria a impressao de que o alerta
    # olha mais longe do que olha. No modo VENCIMENTO ela vai ate o fim do
    # proximo mes -- e' o ponto do alerta.
    if modo == "vencimento":
        ate = como_data(args.ate) or fim_do_proximo_mes(hoje)
    else:
        ate = como_data(args.ate) or hoje

    sem_baixados = regras["sem_baixados"] and not args.com_baixados

    try:
        achados, resumo = ler_se2(base, hoje, usuario, desde, ate, modo,
                                  args.prefixo, args.tipo, sem_baixados)
    except Exception as erro:  # noqa: BLE001
        print(f"   AVISO: nao consegui ler a SE2 ({erro}). A rodada segue.")
        return 0

    recorte = " ".join(p for p in (f"prefixo {args.prefixo}" if args.prefixo else "",
                                   f"tipo {args.tipo}" if args.tipo else "") if p)
    print(f"   SE2: {resumo['linhas']} títulos | {resumo['do_usuario']} de {usuario} | "
          f"{resumo['no_recorte']} com {recorte or 'qualquer prefixo/tipo'} | "
          f"{resumo['na_janela']} com {regras['rotulo']} entre {desde:%d/%m/%Y} e "
          f"{ate:%d/%m/%Y} | {resumo['fora_da_janela']} fora da janela"
          + (f" | {resumo['sem_data_janela']} sem a data da janela"
             if resumo["sem_data_janela"] else "")
          + (f" | {resumo['ja_baixados']} descartado(s) por já estarem baixados"
             if sem_baixados and resumo["ja_baixados"] else ""))

    # ⚠ O caderno de vistos SO existe no modo inclusao. No modo vencimento a
    # janela e um ESTADO que rola com o calendario, e ele pediu "TODOS os
    # titulos": esconder o que saiu ontem faria a lista mentir por omissao.
    ja_vistos = {} if (args.tudo or not regras["caderno"]) else ler_vistos()
    novos = [a for a in achados if chave_do_titulo(a) not in ja_vistos]
    repetidos = len(achados) - len(novos)
    if repetidos:
        print(f"   {repetidos} já avisado(s) em envio anterior -- fora desta lista.")

    if not novos:
        print(f"   Nenhum título de {usuario} {'vencendo' if modo == 'vencimento' else 'incluído'} "
              f"na janela.")
        return 0

    total = sum(a["_valor"] for a in novos)
    print(f"   {len(novos)} título(s) de {usuario} ({reais(total)}, "
          f"{sum(1 for a in novos if a['_atraso'] > 0)} vencido(s), "
          f"{sum(1 for a in novos if a['linha_dig'])} com linha digitável, "
          f"{sum(1 for a in novos if a['baixa'])} já baixado(s)):")
    # Lista longa nao cabe na janela do .cmd e empurra o resto do log para
    # fora do buffer: as 40 primeiras bastam para conferir de olho, e a lista
    # inteira esta na planilha anexada.
    #
    # O UUID vai JUNTO, numa segunda linha (pedido dele em 21/09/2026): e' por
    # ele que se acha o titulo no painel da SE2, e conferir a lista sem o
    # codigo obriga a abrir a planilha so para copiar um campo.
    for a in novos[:40]:
        print(f"     venc {a['vencimento'] or '-':<10} real {a['real'] or '-':<10} | "
              f"{a['filial']} | {a['titulo']}/{a['parcela'] or '-'} | "
              f"{a['tipo_pgto'][:14]:<14} | {a['valor']:>15} | "
              f"{a['situacao']:<14} | {a['fornecedor'][:28]}")
        print(f"          UUID {a['uuid'] or '(sem código)'}"
              + (f"  ·  inc {a['inclusao']}" if a["inclusao"] else "")
              + (f"  ·  LD {a['linha_dig']}" if a["linha_dig"] else ""))
    if len(novos) > 40:
        print(f"     ... e mais {len(novos) - 40} (todos vão no e-mail e na planilha)")

    # A tabela do e-mail e' cortada; a planilha nunca. Ver LIMITE_NO_EMAIL.
    na_tabela = novos[:LIMITE_NO_EMAIL]
    cortados = len(novos) - len(na_tabela)
    if cortados:
        print(f"   A tabela do e-mail mostra {len(na_tabela)}; os outros "
              f"{cortados} vão só na planilha (corpo grande é cortado pelo "
              f"programa de e-mail).")
    corpo = montar_html(na_tabela, resumo, hoje, desde, ate, usuario, repetidos,
                        modo=modo, prefixo=args.prefixo, tipo=args.tipo,
                        sem_baixados=sem_baixados,
                        cortados=cortados, total_geral=total)
    plural = "s" if len(novos) > 1 else ""
    marca = (f"{args.prefixo}/{args.tipo}" if args.prefixo and args.tipo
             else (args.prefixo or args.tipo or "").strip())
    if modo == "vencimento":
        assunto = (f"[Painel Boletos] {len(novos)} título{plural}"
                   + (f" {marca}" if marca else "")
                   + f" de {usuario} vencendo até {ate:%d/%m/%Y}"
                   + f" - {hoje:%d/%m/%Y}")
    else:
        assunto = (f"[Painel Boletos] {len(novos)} título{plural} novo{plural} "
                   f"incluído{plural} na SE2 por {usuario} - {hoje:%d/%m/%Y}")

    planilha = gerar_planilha(novos, hoje, usuario, modo)
    if planilha:
        print(f"   Planilha: {planilha}")

    if args.teste:
        PREVIA.parent.mkdir(parents=True, exist_ok=True)
        PREVIA.write_text(corpo, encoding="utf-8")
        print("\n   MODO TESTE: nada foi enviado."
              + (" O caderno de vistos ficou como estava."
                 if regras["caderno"] else "")
              + f"\n   Assunto: {assunto}")
        print(f"   Prévia:  {PREVIA}")
        return 0

    destino = args.para or comuns.destinatarios(CHAVE)
    if not destino:
        print("   AVISO: nao achei para quem mandar (.alerta_inclusao_para).")
        return 0
    copia = args.copia or comuns.copias(CHAVE)

    # A trava e POR MODO: o de vencimento e o de inclusao podem sair no mesmo
    # dia sem um calar o outro.
    trava = regras["trava"]
    if not args.forcar and comuns.ja_enviado_hoje(trava):
        print(f"   Já enviado hoje ({comuns.quando_enviou(trava)}) -- não mando de novo.")
        print("   Para mandar assim mesmo: --forcar")
        return 0

    try:
        comuns.enviar(assunto, corpo, destino, copia=copia,
                      anexos=[planilha] if planilha else None)
    except Exception as erro:  # noqa: BLE001
        print(f"   AVISO: o e-mail NAO foi enviado ({erro}). O painel segue.")
        if regras["caderno"]:
            print("   O caderno de vistos NAO foi gravado: estes títulos voltam na "
                  "próxima rodada.")
        return 0

    # So agora: e-mail aceito pelo Outlook, entao os titulos podem ser dados
    # por avisados. Nesta ordem, e nunca antes.
    comuns.marcar_enviado(trava, len(novos))
    if regras["caderno"]:
        gravar_vistos(ja_vistos, [chave_do_titulo(a) for a in novos], hoje)
    print(f"   E-mail enviado para {destino}"
          + (f" (cópia: {copia})" if copia else "")
          + (f" com a planilha {planilha.name} em anexo." if planilha else "."))
    if regras["caderno"]:
        print(f"   {len(novos)} título(s) anotados como avisados em {VISTOS.name}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
