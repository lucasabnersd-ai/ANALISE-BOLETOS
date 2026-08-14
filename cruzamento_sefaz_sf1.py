# -*- coding: utf-8 -*-
"""Aba "SEFAZ x SF1": a nota que a SEFAZ mostra ja foi lancada no TOTVS?

Cada NOTA da SEFAZ.xlsx (nao cada linha: a de NF-e vem por item x duplicata) e
procurada na SF1 -- a base inteira, TODOS os status, nao so o classificado.
O usuario pediu o cruzamento "usando numero de NFs", entao o numero da NF e a
chave da busca; o resto serve para dizer o quanto dar de fe no que foi achado.

⚠ NUMERO DE NF NAO IDENTIFICA NOTA. Dois fornecedores emitem a NF 1234 no mesmo
mes sem nenhum problema -- na SF1 sao 8.011 titulos e o numero se repete muito.
Por isso o veredito tem tres degraus, e nao dois:

    ACHADA      numero bate E o emitente confere (raiz do CNPJ) -- ou, melhor
                ainda, a CHAVE NF-e de 44 digitos e identica. Ai nao ha duvida.
    CONFERIR    numero bate e ALGO mais sustenta (valor exato ou emissao no
                mesmo dia), mas o emitente nao confirma. E a linha que precisa
                de olho humano -- e sao poucas.
    NAO ACHADA  nenhum titulo na SF1 com esse numero, OU ate ha, mas sem nada
                em comum alem do numero. A nota existe na SEFAZ e nao foi
                lancada -- que e o que esta aba serve para mostrar.

⚠ Nada aqui encosta na SE2 (pedido explicito do usuario). Esta aba compara
SEFAZ com SF1 e mais nada.
"""

from __future__ import annotations

import datetime as dt
import re
import shutil
from pathlib import Path

import cruzamento_classificacao as cc
import sefaz

PREFIXO = "SFZ1:"

# Nomes de saida.
CHAVE = "Chave"
SITUACAO = "Situação"
ORIGEM = "Origem"
NUM_NF = "Nº NF"
EMISSAO = "Emissão"
NOME_EMIT = "Emitente"
CNPJ_EMIT = "CNPJ Emitente"
VLR_SEFAZ = "Vlr SEFAZ"
CHAVE_NFE = "Chave NF-e"
# 13/08/2026, pedido do usuario: colunas da nota que a visao por NOTA descartava.
FANTASIA = "Nome Fantasia"
CFOP = "CFOP"
TIPO_OP = "Tipo Operação"
# ⚠ Vazia nas NFS-e: a coluna existe na aba `NFS` da SEFAZ.xlsx mas a exportacao
# nao a preenche (medido: 0 de 462). So as NF-e tem natureza.
NAT_OP = "Nat. Operação"
SAIDA = "Saída/Entrada"
NOME_DEST = "Destinatário"
CNPJ_DEST = "CNPJ Destinatário"
VENC_DUP = "Venc Duplicata"
VLR_DUP = "Vlr Duplicata"
# Quantas duplicatas a nota tem. So aparece quando ha MAIS DE UMA: as colunas
# acima levam a PRIMEIRA (a que vence antes), e sem este aviso a linha diria
# "vence em 10/09" para uma nota que na verdade tem 3 parcelas. Medido: 33 das
# 617 notas tem mais de uma duplicata.
QTD_DUP = "Duplicatas"
# lado da SF1
FILIAL = "Filial"
NF_SF1 = "Nº NF SF1"
SERIE = "Série"
RAZAO = "Razão Social"
VLR_TITULO = "Vlr.Título"
EMISSAO_SF1 = "Emissão SF1"
CLASSIFICADO = "Classificado em"
CRITERIO = "Critério"
CANDIDATOS = "Títulos"

ALERTA = "Alerta"

# ⚠ `Emissão SF1` e `Campo UUID` (o UUID do titulo na SF1) SAIRAM em 13/08/2026,
# a pedido do usuario. A emissao da SF1 continua sendo lida e continua valendo
# ponto no cruzamento (`_nota_do_par`) -- ela so nao vira coluna. Tirar daqui
# tambem tira da carteira que viaja para o navegador, que e o ponto.
CABECALHO = [CHAVE, SITUACAO, ALERTA, ORIGEM, NUM_NF, NF_SF1, EMISSAO,
             NOME_EMIT, RAZAO, CNPJ_EMIT, FANTASIA, NOME_DEST, CNPJ_DEST,
             TIPO_OP, NAT_OP, SAIDA, CFOP, VLR_SEFAZ, VLR_TITULO,
             VENC_DUP, VLR_DUP, QTD_DUP,
             FILIAL, SERIE, CLASSIFICADO, CANDIDATOS, CRITERIO, CHAVE_NFE]

ACHADA = "ACHADA NA SF1"
CONFERIR = "CONFERIR"
NAO_ACHADA = "NÃO ACHADA"

# Pesos: bits distintos, para que a soma ordene do mais forte para o mais fraco
# (o melhor candidato sai de um `max`). Mesma ideia do cruzamento_classificacao.
PESO_CHAVE = 32      # chave NF-e de 44 digitos identica: nao ha o que discutir
PESO_CNPJ = 8        # raiz do CNPJ do emitente = codigo do fornecedor na SF1
PESO_VALOR = 4
PESO_EMISSAO = 2


def _raiz(cnpj) -> str:
    """8 primeiros digitos do CNPJ -- e o codigo do fornecedor na SF1."""
    return re.sub(r"\D", "", str(cnpj or ""))[:8]


def _digitos(valor) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def ler_sf1(base: Path) -> list[dict]:
    """A SF1 INTEIRA, indexada depois pelo numero da NF.

    Ao contrario do `cruzamento_classificacao.ler_sf1`, aqui nao ha filtro de
    status nem de data: o usuario pediu "todas as NFs da base SF1". Uma nota
    lancada e depois cancelada continua sendo uma nota lancada para quem esta
    conferindo se a SEFAZ ja entrou no TOTVS.
    """
    wb, temporaria = cc._abrir(base)
    try:
        ws = wb[wb.sheetnames[0]]
        brutas = ws.iter_rows(values_only=True)
        next(brutas, None)                       # cabecalho
        titulos = []
        for linha in brutas:
            # ⚠ Continua sendo o SF1_UUID que mede a linha, mesmo depois de a
            # coluna do UUID ter saido do painel (13/08/2026): e o maior indice
            # usado por este modulo, entao e ele que diz "a linha veio inteira".
            if len(linha) <= cc.SF1_UUID:
                continue
            numero = cc._texto(linha[cc.SF1_NUMERO])
            if not numero:
                continue
            titulos.append({
                "nf": sefaz.nf_chave(numero),
                FILIAL: cc._texto(linha[cc.SF1_FILIAL]),
                NF_SF1: numero,
                SERIE: cc._texto(linha[cc.SF1_SERIE]),
                "fornecedor": cc._texto(linha[cc.SF1_FORNECEDOR]),
                RAZAO: cc._texto(linha[cc.SF1_RAZAO]),
                VLR_TITULO: cc._moeda(linha[cc.SF1_VLR_BRUTO]),
                EMISSAO_SF1: cc._data(linha[cc.SF1_EMISSAO]),
                CLASSIFICADO: cc._data(linha[cc.SF1_DIGITACAO]),
                "status": cc._texto(linha[cc.SF1_STATUS]),
                "chave_nfe": _digitos(cc._texto(linha[cc.SF1_CHAVE_NFE])),
            })
        return titulos
    finally:
        wb.close()
        if temporaria:
            shutil.rmtree(temporaria, ignore_errors=True)


# Campos que NAO variam dentro da nota: medido em 13/08/2026 nas 617 notas --
# Tipo Operação, Saída/Entrada, Nome Fantasia, Destinatário, CNPJ Destinatário e
# Nat. Operação sao iguais em todas as linhas da mesma nota, entao a PRIMEIRA
# linha basta. ⚠ O CFOP NAO entra aqui: ele e do ITEM e varia em 40 das 155
# notas de NF-e -- por isso e juntado, e nao pego da primeira linha.
DA_PRIMEIRA_LINHA = (
    (NOME_EMIT, sefaz.NOME_EMIT), (CNPJ_EMIT, sefaz.CNPJ_EMIT),
    (FANTASIA, sefaz.FANTASIA), (NOME_DEST, sefaz.NOME_DEST),
    (CNPJ_DEST, sefaz.CNPJ_DEST), (TIPO_OP, sefaz.TIPO_OP), (SAIDA, sefaz.SAIDA),
    (NAT_OP, sefaz.NAT_OP),
)


def notas_da_sefaz(linhas: list[dict]) -> list[dict]:
    """Uma linha por NOTA. A base de NF-e vem por item x duplicata, entao o
    cruzamento tem de subir um nivel -- senao uma nota de 51 itens viraria 51
    vezes a mesma pergunta na tela.

    ⚠ Subir de nivel PERDE dado se a agregacao for descuidada, e e por isso que
    cada campo tem um jeito proprio: o que se repete vem da primeira linha, o
    CFOP (que e do item) vem junto e a duplicata vem a que VENCE ANTES, com a
    contagem do lado para a linha nao mentir sobre as outras parcelas.
    """
    porchave: dict[str, dict] = {}
    for l in linhas:
        chave = l[sefaz.CHAVE_NFE]
        nota = porchave.get(chave)
        if nota is None:
            nota = porchave[chave] = {
                ORIGEM: l[sefaz.ORIGEM],
                CHAVE_NFE: chave,
                NUM_NF: l[sefaz.NUM_NF],
                EMISSAO: l[sefaz.EMISSAO],
                # o total da nota se repete em toda linha dela; a primeira basta
                VLR_SEFAZ: l[sefaz.VLR_TOTAL],
                "itens": 0,
                "cfops": [],
                "duplicatas": [],
            }
            for destino, origem in DA_PRIMEIRA_LINHA:
                nota[destino] = l.get(origem, "")
        nota["itens"] += 1

        cfop = str(l.get(sefaz.CFOP) or "").strip()
        if cfop and cfop not in nota["cfops"]:
            nota["cfops"].append(cfop)

        # (vencimento, valor) -- a mesma duplicata se repete em cada ITEM da nota,
        # entao o que identifica a parcela e o par, nao a linha.
        venc, valor = l.get(sefaz.VENC_DUP), l.get(sefaz.VLR_DUP)
        if (venc or valor) and (venc, valor) not in nota["duplicatas"]:
            nota["duplicatas"].append((venc, valor))

    for nota in porchave.values():
        nota[CFOP] = " · ".join(nota.pop("cfops"))
        # a que chega ANTES: e a que decide o alerta de vencimento e a que a
        # pessoa precisa ver primeiro. Sem data vai para o fim.
        parcelas = sorted(nota.pop("duplicatas"),
                          key=lambda d: (d[0] is None, d[0] or dt.date.min))
        nota[VENC_DUP] = parcelas[0][0] if parcelas else None
        nota[VLR_DUP] = parcelas[0][1] if parcelas else None
        # 1 duplicata nao e aviso de nada -- so o "tem mais de uma" informa
        nota[QTD_DUP] = len(parcelas) if len(parcelas) > 1 else None
    return list(porchave.values())


def _nota_do_par(nota: dict, titulo: dict) -> tuple[int, list[str]]:
    """Quanto este titulo da SF1 combina com esta nota, e por que."""
    pontos, porque = 0, ["Nº da NF confere"]

    chave = _digitos(nota[CHAVE_NFE])
    if len(chave) == 44 and titulo["chave_nfe"] == chave:
        pontos += PESO_CHAVE
        porque.append("Chave NF-e idêntica (44 dígitos)")

    raiz = _raiz(nota[CNPJ_EMIT])
    if raiz and _digitos(titulo["fornecedor"])[:8] == raiz:
        pontos += PESO_CNPJ
        porque.append("Emitente confere (raiz do CNPJ = cód. fornecedor)")

    valor_s, valor_t = nota[VLR_SEFAZ], titulo[VLR_TITULO]
    if valor_s and valor_t and abs(valor_s - valor_t) < 0.01:
        pontos += PESO_VALOR
        porque.append("Valor exato")

    if nota[EMISSAO] and titulo[EMISSAO_SF1] and nota[EMISSAO] == titulo[EMISSAO_SF1]:
        pontos += PESO_EMISSAO
        porque.append("Emissão no mesmo dia")

    return pontos, porque


def cruzar(notas: list[dict], titulos: list[dict]) -> list[dict]:
    """Para cada nota da SEFAZ, o melhor titulo da SF1 com o mesmo numero."""
    indice: dict[str, list[dict]] = {}
    for t in titulos:
        if t["nf"]:
            indice.setdefault(t["nf"], []).append(t)

    resultado = []
    for nota in notas:
        numero = sefaz.nf_chave(nota[NUM_NF])
        candidatos = indice.get(numero, []) if numero else []

        if not candidatos:
            linha = dict(nota)
            linha[SITUACAO] = NAO_ACHADA
            linha[CRITERIO] = ("Nenhum título na SF1 com este número de NF"
                               if numero else "Nota sem número de NF legível")
            linha[CANDIDATOS] = 0
            resultado.append(linha)
            continue

        avaliados = [(*_nota_do_par(nota, t), t) for t in candidatos]
        pontos, porque, melhor = max(avaliados, key=lambda a: a[0])
        # Confirmado = a chave de 44 digitos ou o emitente. So o numero nao basta.
        confirmado = pontos >= PESO_CNPJ

        linha = dict(nota)
        linha[CANDIDATOS] = len(candidatos)

        # ⚠ SO O NUMERO BATENDO NAO E UM ACHADO -- e o numero batendo.
        # Medido na base real: das 251 notas que caiam em CONFERIR, 242 nao
        # tinham NADA alem do numero (nem valor, nem emissao, nem emitente), e
        # 246 eram NFS-e, cuja numeracao comeca do 1 e colide com meio TOTVS.
        # Chamar isso de "conferir" criaria 242 tarefas que nao existem. Fica
        # NAO ACHADA, e o quadro do Criterio conta que havia homonimos.
        if not pontos:
            linha[SITUACAO] = NAO_ACHADA
            linha[CRITERIO] = (
                f"{len(candidatos)} título(s) na SF1 com o número {numero}, "
                "mas nenhum é deste emitente nem bate valor ou emissão — "
                "coincidência de numeração, a nota continua sem lançamento")
            resultado.append(linha)
            continue

        linha[SITUACAO] = ACHADA if confirmado else CONFERIR
        linha[FILIAL] = melhor[FILIAL]
        linha[NF_SF1] = melhor[NF_SF1]
        linha[SERIE] = melhor[SERIE]
        linha[RAZAO] = melhor[RAZAO]
        linha[VLR_TITULO] = melhor[VLR_TITULO]
        # ⚠ `Emissão SF1` nao e mais coluna (saiu do CABECALHO em 13/08/2026),
        # mas continua sendo gravada: e um `max` de pontos que decide o melhor
        # titulo, e a emissao e um dos pontos. Para devolver a coluna a tela,
        # basta reinserir EMISSAO_SF1 no CABECALHO -- nada mais muda.
        linha[EMISSAO_SF1] = melhor[EMISSAO_SF1]
        linha[CLASSIFICADO] = melhor[CLASSIFICADO]
        if not confirmado:
            porque.append("Emitente NÃO confirmado — pode ser outro fornecedor com o mesmo número")
        if melhor["status"] and melhor["status"].upper() != "A":
            porque.append(f"Status do título na SF1: {melhor['status']}")
        linha[CRITERIO] = " + ".join(porque)
        resultado.append(linha)
    return resultado


def montar_linhas(cruzadas: list[dict]) -> list[tuple]:
    """Cruzadas -> tuplas na ordem de CABECALHO."""
    ordem = {NAO_ACHADA: 0, CONFERIR: 1, ACHADA: 2}
    # o que NAO foi achado primeiro: e a fila de trabalho desta aba
    ordenadas = sorted(cruzadas, key=lambda l: (
        ordem.get(l[SITUACAO], 3),
        l.get(EMISSAO) is None,
        -(l[EMISSAO].toordinal() if l.get(EMISSAO) else 0),
    ))
    saida = []
    for l in ordenadas:
        linha = dict(l)
        linha[CHAVE] = PREFIXO + (_digitos(l[CHAVE_NFE]) or l[CHAVE_NFE])
        saida.append(tuple(linha.get(coluna, "") for coluna in CABECALHO))
    return saida


class Cruzamento:
    """O resultado do cruzamento antes de virar tabela.

    Existe porque TRES abas precisam da mesma coisa e ler a SF1 (8.011 titulos)
    tres vezes seria desperdicio: esta aba, a `Análise Base SEFAZ` (que precisa
    do `nao_lancadas` para os alertas) e a `NF x Pedido de Compra`, que mostra o
    mesmo veredito de lancamento ao lado do pedido.
    """

    def __init__(self, linhas, cruzadas, nao_lancadas, alertas, hoje, titulos_sf1):
        self.linhas = linhas              # SEFAZ crua (por item x duplicata)
        self.cruzadas = cruzadas          # uma por NOTA, ja com o veredito
        self.nao_lancadas = nao_lancadas  # chaves NF-e que a SF1 nao tem
        self.alertas = alertas            # {chave NF-e: [avisos]}
        self.hoje = hoje
        self.titulos_sf1 = titulos_sf1    # quantos titulos a SF1 tinha

    def por_chave(self) -> dict[str, dict]:
        return {l[CHAVE_NFE]: l for l in self.cruzadas}


def preparar(base_sefaz: Path, base_sf1: Path, hoje=None,
             linhas: list[dict] | None = None) -> Cruzamento:
    """Le as duas bases e cruza. Nao monta tabela nenhuma -- ver `carregar`.

    `linhas` chega pronto do `sefaz.ler_com_historico()`: e a base do dia MAIS as
    notas que sumiram dela. Ler a SEFAZ aqui de novo devolveria so o que esta no
    arquivo hoje, e as removidas nao chegariam nesta aba nem na de pedidos.
    """
    hoje = hoje or dt.date.today()
    # as linhas cruas ficam: e delas que sai o vencimento das duplicatas, que a
    # visao por nota (`notas_da_sefaz`) nao carrega
    linhas = sefaz.ler(base_sefaz) if linhas is None else linhas
    notas = notas_da_sefaz(linhas)
    titulos = ler_sf1(base_sf1)
    cruzadas = cruzar(notas, titulos)

    # Nota que a SF1 nao tem = nota nao lancada. E a base dos alertas 1 e 3.
    nao_lancadas = {l[CHAVE_NFE] for l in cruzadas if l[SITUACAO] == NAO_ACHADA}
    alertas = sefaz.alertas_por_nota(linhas, nao_lancadas, hoje)
    for l in cruzadas:
        l[ALERTA] = " · ".join(alertas.get(l[CHAVE_NFE], []))
    return Cruzamento(linhas, cruzadas, nao_lancadas, alertas, hoje, len(titulos))


def carregar(base_sefaz: Path, base_sf1: Path, hoje=None,
             pronto: "Cruzamento | None" = None) -> tuple[list[str], list[tuple], dict]:
    """Ponto de entrada: (cabecalho, linhas, resumo) para o gerador.

    ⚠ O resumo leva `nao_lancadas` (as chaves em NÃO ACHADA). E este cruzamento
    quem sabe o que falta lancar, e a aba da base SEFAZ precisa disso para poder
    mostrar os mesmos alertas -- por isso o gerador roda o cruzamento ANTES dela.

    `pronto` evita reler as bases quando o gerador ja chamou `preparar()`.
    """
    c = pronto or preparar(base_sefaz, base_sf1, hoje)
    # ⚠ `cruzadas` e uma linha por NOTA -- e por isso serve tambem de contagem de
    # notas: `cruzar()` devolve exatamente uma saida por entrada.
    cruzadas, alertas, nao_lancadas = c.cruzadas, c.alertas, c.nao_lancadas

    contagem = {ACHADA: 0, CONFERIR: 0, NAO_ACHADA: 0}
    for l in cruzadas:
        contagem[l[SITUACAO]] = contagem.get(l[SITUACAO], 0) + 1
    resumo = {
        **sefaz.contar_alertas(alertas),
        "nao_lancadas": nao_lancadas,
        "notas": len(cruzadas),
        "sf1": c.titulos_sf1,
        "achadas": contagem[ACHADA],
        "conferir": contagem[CONFERIR],
        "nao_achadas": contagem[NAO_ACHADA],
        "nfe": sum(1 for n in cruzadas if n[ORIGEM] == sefaz.NFE),
        "nfs": sum(1 for n in cruzadas if n[ORIGEM] == sefaz.NFSE),
    }
    return list(CABECALHO), montar_linhas(cruzadas), resumo
