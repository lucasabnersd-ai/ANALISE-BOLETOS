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
# lado da SF1
FILIAL = "Filial"
NF_SF1 = "Nº NF SF1"
SERIE = "Série"
RAZAO = "Razão Social"
VLR_TITULO = "Vlr.Título"
EMISSAO_SF1 = "Emissão SF1"
CLASSIFICADO = "Classificado em"
UUID_NF = "Campo UUID"
CRITERIO = "Critério"
CANDIDATOS = "Títulos"

ALERTA = "Alerta"

CABECALHO = [CHAVE, SITUACAO, ALERTA, ORIGEM, NUM_NF, NF_SF1, EMISSAO, EMISSAO_SF1,
             NOME_EMIT, RAZAO, CNPJ_EMIT, VLR_SEFAZ, VLR_TITULO,
             FILIAL, SERIE, CLASSIFICADO, CANDIDATOS, CRITERIO,
             CHAVE_NFE, UUID_NF]

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
                UUID_NF: cc._texto(linha[cc.SF1_UUID]),
            })
        return titulos
    finally:
        wb.close()
        if temporaria:
            shutil.rmtree(temporaria, ignore_errors=True)


def notas_da_sefaz(linhas: list[dict]) -> list[dict]:
    """Uma linha por NOTA. A base de NF-e vem por item x duplicata, entao o
    cruzamento tem de subir um nivel -- senao uma nota de 51 itens viraria 51
    vezes a mesma pergunta na tela."""
    porchave = {}
    for l in linhas:
        chave = l[sefaz.CHAVE_NFE]
        nota = porchave.get(chave)
        if nota is None:
            porchave[chave] = {
                ORIGEM: l[sefaz.ORIGEM],
                CHAVE_NFE: chave,
                NUM_NF: l[sefaz.NUM_NF],
                EMISSAO: l[sefaz.EMISSAO],
                NOME_EMIT: l[sefaz.NOME_EMIT],
                CNPJ_EMIT: l[sefaz.CNPJ_EMIT],
                # o total da nota se repete em toda linha dela; a primeira basta
                VLR_SEFAZ: l[sefaz.VLR_TOTAL],
                "itens": 1,
            }
        else:
            nota["itens"] += 1
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
        linha[EMISSAO_SF1] = melhor[EMISSAO_SF1]
        linha[CLASSIFICADO] = melhor[CLASSIFICADO]
        linha[UUID_NF] = melhor[UUID_NF]
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


def carregar(base_sefaz: Path, base_sf1: Path,
             hoje=None) -> tuple[list[str], list[tuple], dict]:
    """Ponto de entrada: (cabecalho, linhas, resumo) para o gerador.

    ⚠ O resumo leva `nao_lancadas` (as chaves em NÃO ACHADA). E este cruzamento
    quem sabe o que falta lancar, e a aba da base SEFAZ precisa disso para poder
    mostrar os mesmos alertas -- por isso o gerador roda o cruzamento ANTES dela.
    """
    hoje = hoje or dt.date.today()
    # as linhas cruas ficam: e delas que sai o vencimento das duplicatas, que a
    # visao por nota (`notas_da_sefaz`) nao carrega
    linhas = sefaz.ler(base_sefaz)
    notas = notas_da_sefaz(linhas)
    titulos = ler_sf1(base_sf1)
    cruzadas = cruzar(notas, titulos)

    # Nota que a SF1 nao tem = nota nao lancada. E a base dos alertas 1 e 3.
    nao_lancadas = {l[CHAVE_NFE] for l in cruzadas if l[SITUACAO] == NAO_ACHADA}
    alertas = sefaz.alertas_por_nota(linhas, nao_lancadas, hoje)
    for l in cruzadas:
        l[ALERTA] = " · ".join(alertas.get(l[CHAVE_NFE], []))

    contagem = {ACHADA: 0, CONFERIR: 0, NAO_ACHADA: 0}
    for l in cruzadas:
        contagem[l[SITUACAO]] = contagem.get(l[SITUACAO], 0) + 1
    resumo = {
        **sefaz.contar_alertas(alertas),
        "nao_lancadas": nao_lancadas,
        "notas": len(notas),
        "sf1": len(titulos),
        "achadas": contagem[ACHADA],
        "conferir": contagem[CONFERIR],
        "nao_achadas": contagem[NAO_ACHADA],
        "nfe": sum(1 for n in notas if n[ORIGEM] == sefaz.NFE),
        "nfs": sum(1 for n in notas if n[ORIGEM] == sefaz.NFSE),
    }
    return list(CABECALHO), montar_linhas(cruzadas), resumo
