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
import sc7
import sefaz

PREFIXO = "SFZ1:"

# Nomes de saida.
CHAVE = "Chave"
SITUACAO = "Situação"
ORIGEM = "Origem"
NUM_NF = "Nº NF"
EMISSAO = "Emissão"
# O status da nota NA SEFAZ (`Autorizada`/`Cancelada` na NF-e, `Normal`/
# `Cancelada`/`Substituída` na NFS-e). Ver sefaz.STATUS_NOTA.
STATUS_NOTA = "Status da Nota"
# O que a empresa respondeu sobre a nota (ciencia/confirmacao/desconhecimento/
# nao realizada). Nao e o `Status da Nota` acima, que e o que a SEFAZ diz.
MANIFESTACAO = "Manifestação"
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

# 09/09/2026, pedido do usuario: *"GERE POSSIBILIDADE USANDO O NOME DO
# FORNECEDOR VALOR E DATA TAMBEM"*. Os titulos da SF1 que PODEM ser esta nota
# ainda que o numero da NF nao bate -- ver `possibilidades()`.
PROVAVEIS = "Títulos Prováveis na SF1"
# O por que de cada provavel, um paragrafo por titulo. Texto longo: na tela e um
# botao "ver ▾" que abre o quadro (`texto_longo` no gerar_painel), do mesmo jeito
# que a Info Comp da SEFAZ -- e por isso NAO precisou de desenho novo no painel.
PORQUE_PROVAVEIS = "Por que estes títulos prováveis"

ALERTA = "Alerta"

# ⚠ `Emissão SF1` e `Campo UUID` (o UUID do titulo na SF1) SAIRAM em 13/08/2026,
# a pedido do usuario. A emissao da SF1 continua sendo lida e continua valendo
# ponto no cruzamento (`_nota_do_par`) -- ela so nao vira coluna. Tirar daqui
# tambem tira da carteira que viaja para o navegador, que e o ponto.
CABECALHO = [CHAVE, SITUACAO, ALERTA, ORIGEM, NUM_NF, NF_SF1, EMISSAO,
             STATUS_NOTA,
             NOME_EMIT, RAZAO, CNPJ_EMIT, FANTASIA, NOME_DEST, CNPJ_DEST,
             TIPO_OP, NAT_OP, SAIDA, CFOP, VLR_SEFAZ, VLR_TITULO,
             VENC_DUP, VLR_DUP, QTD_DUP,
             FILIAL, SERIE, CLASSIFICADO, CANDIDATOS,
             PROVAVEIS, PORQUE_PROVAVEIS, CRITERIO, CHAVE_NFE]

ACHADA = "ACHADA NA SF1"
CONFERIR = "CONFERIR"
NAO_ACHADA = "NÃO ACHADA"

# Pesos: bits distintos, para que a soma ordene do mais forte para o mais fraco
# (o melhor candidato sai de um `max`). Mesma ideia do cruzamento_classificacao.
PESO_CHAVE = 32      # chave NF-e de 44 digitos identica: nao ha o que discutir
PESO_CNPJ = 8        # raiz do CNPJ do emitente = codigo do fornecedor na SF1
PESO_VALOR = 4
PESO_EMISSAO = 2
# 04/09/2026, pedido do usuario: "use o nome do emissor pra cruzar". MEDIDO
# antes, nos 319 titulos cuja chave de 44 bate com uma nota: o CNPJ confere
# em 319, o nome em 249, e o nome NUNCA confere onde o CNPJ nao confere. E
# corroboracao, nao descoberta -- por isso vale 1 e entra no Criterio, mas
# nao muda sozinho um veredito. Onde ele diverge (70), e abreviacao ou corte:
# "BORRACHARIA DO JOAQUIM COM PNEUS L" x "...COMERCIO DE PNEUS LTDA".
PESO_NOME = 1


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
    # 28/08/2026: medido nas duas abas, ZERO das 813 notas tem status diferente
    # entre as proprias linhas -- ele e da nota, entao a primeira linha basta.
    (STATUS_NOTA, sefaz.STATUS_NOTA),
    # 01/09/2026. A manifestacao e da NOTA por construcao, e nao por medicao: o
    # manifestacao.py grava o mesmo valor em TODAS as linhas de uma chave (o
    # de-para dele e por chave NF-e). A primeira linha nao so basta como e a
    # unica resposta possivel. Ver manifestacao.py.
    (MANIFESTACAO, sefaz.MANIFESTACAO),
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


def _mesmo_cnpj(nota: dict, titulo: dict) -> bool:
    """Raiz do CNPJ do emitente = codigo do fornecedor na SF1."""
    raiz = _raiz(nota[CNPJ_EMIT])
    return bool(raiz) and _digitos(titulo["fornecedor"])[:8] == raiz


def _mesmo_fornecedor(nota: dict, titulo: dict) -> bool:
    """O titulo e DESTE fornecedor -- pelo CNPJ ou pela razao social.

    ⚠ E este par (CNPJ **ou** nome) que responde "de quem e este titulo?" tanto
    no veredito quanto nas `possibilidades()`. Duas respostas diferentes para a
    mesma pergunta poriam um titulo na lista de provaveis e o deixariam fora do
    cruzamento -- ou o contrario.
    """
    return _mesmo_cnpj(nota, titulo) or _mesmo_emissor(nota.get(NOME_EMIT),
                                                       titulo.get(RAZAO))


def _mesmo_valor(nota: dict, titulo: dict) -> bool:
    valor_s, valor_t = nota[VLR_SEFAZ], titulo[VLR_TITULO]
    return bool(valor_s and valor_t) and abs(valor_s - valor_t) < 0.01


def _mesma_emissao(nota: dict, titulo: dict) -> bool:
    return bool(nota[EMISSAO] and titulo[EMISSAO_SF1]) and nota[EMISSAO] == titulo[EMISSAO_SF1]


def _nota_do_par(nota: dict, titulo: dict, nf_porque: str = "") -> tuple[int, list[str]]:
    """Quanto este titulo da SF1 combina com esta nota, e por que.

    `nf_porque` troca a primeira frase quando o numero SO bate depois de tirar o
    ano grudado (ver `sefaz.nf_variantes`). Ele nao vale ponto: o numero e a
    CHAVE da busca, e chave nao pontua a si mesma -- o que pontua e o que
    sustenta o achado. Mas a frase tem de dizer a verdade, senao o quadro do
    Criterio afirma "Nº da NF confere" para uma nota `2026000000009` que casou
    com o titulo `000000009`.
    """
    pontos, porque = 0, [nf_porque or "Nº da NF confere"]

    chave = _digitos(nota[CHAVE_NFE])
    if len(chave) == 44 and titulo["chave_nfe"] == chave:
        pontos += PESO_CHAVE
        porque.append("Chave NF-e idêntica (44 dígitos)")

    if _mesmo_cnpj(nota, titulo):
        pontos += PESO_CNPJ
        porque.append("Emitente confere (raiz do CNPJ = cód. fornecedor)")

    if _mesmo_valor(nota, titulo):
        pontos += PESO_VALOR
        porque.append("Valor exato")

    if _mesma_emissao(nota, titulo):
        pontos += PESO_EMISSAO
        porque.append("Emissão no mesmo dia")

    if _mesmo_emissor(nota.get(NOME_EMIT), titulo.get(RAZAO)):
        pontos += PESO_NOME
        porque.append("Nome do emissor = razão social do fornecedor na SF1")

    return pontos, porque


def _mesmo_emissor(nome_nota, razao_sf1) -> bool:
    """O emissor da nota e o fornecedor do titulo tem o mesmo nome?

    Normaliza os dois com `sc7.normalizar_nome` (acento, pontuacao, sufixo) e
    aceita, alem da igualdade/contencao de `sc7.mesmo_nome`, o caso em que um
    lado e PREFIXO do outro: a SEFAZ corta o nome do emitente em ~34 caracteres
    ("MUNDIAL ALTERNADORES IMPORTACAO E ") e a SF1 traz inteiro.

    ⚠ PREFIXO POR PALAVRA INTEIRA, e com pelo menos DUAS. Testado antes de
    entrar: `longo.startswith(curto)` casava "BORRACHA" com "BORRACHARIA X" -- e
    um ponto de nome basta para virar CONFERIR uma nota que so tinha homonimo
    de numero (`if not pontos`). A ultima palavra do lado curto pode estar
    cortada pela metade (e o corte da SEFAZ), as anteriores tem de ser iguais.
    Abreviacao ("COM" x "COMERCIO") NAO casa, de proposito: e o CNPJ que decide.
    """
    a, b = sc7.normalizar_nome(nome_nota), sc7.normalizar_nome(razao_sf1)
    if not a or not b:
        return False
    if sc7.mesmo_nome(a, b):
        return True
    curto, longo = (a.split(" "), b.split(" ")) if len(a) <= len(b) else (b.split(" "), a.split(" "))
    if len(curto) < 2 or len(curto) > len(longo):
        return False
    if curto[:-1] != longo[:len(curto) - 1]:
        return False
    return longo[len(curto) - 1].startswith(curto[-1])


def _amarrar(linha: dict, melhor: dict) -> None:
    """Copia o titulo escolhido para as colunas da SF1 da linha."""
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
    linha["_titulo"] = melhor


def _titulo_id(titulo: dict) -> tuple:
    """Identidade do titulo, para saber se ele ja e de OUTRA nota.

    Filial + numero + serie + valor: a `ler_sf1` nao carrega chave primaria e o
    mesmo numero existe em series diferentes (JAIR PINTO, series "001" e "1").
    """
    return (titulo[FILIAL], titulo[NF_SF1], titulo[SERIE], titulo[VLR_TITULO])


def _por_numero(titulos: list[dict]) -> tuple[dict, dict]:
    """Dois indices: pelo numero cru e por TODAS as formas (com e sem o ano)."""
    exato: dict[str, list[dict]] = {}
    formas: dict[str, list[tuple[dict, str]]] = {}
    for t in titulos:
        if t["nf"]:
            exato.setdefault(t["nf"], []).append(t)
        ano = t[EMISSAO_SF1].year if t[EMISSAO_SF1] else None
        for numero, como in sefaz.nf_variantes(t[NF_SF1], ano).items():
            formas.setdefault(numero, []).append((t, como))
    return exato, formas


def _por_fornecedor(titulos: list[dict]) -> dict[str, list[dict]]:
    """Titulos indexados pela raiz do CNPJ E pela ancora da razao social.

    Dois baldes no MESMO dicionario porque a pergunta e uma so ("quais titulos
    podem ser deste fornecedor?") e quem decide no fim e `_mesmo_fornecedor`:
    o indice existe so para nao varrer 8.662 titulos por nota. A ancora
    (`sc7.ancora_nome`) e o que deixa a contencao/prefixo de nome ainda achar --
    indexar pelo nome inteiro exigiria texto identico.
    """
    indice: dict[str, list[dict]] = {}
    for t in titulos:
        raiz = _digitos(t["fornecedor"])[:8]
        if raiz:
            indice.setdefault("c:" + raiz, []).append(t)
        ancora = sc7.ancora_nome(sc7.normalizar_nome(t[RAZAO]))
        if ancora:
            indice.setdefault("n:" + ancora, []).append(t)
    return indice


def _titulos_do_fornecedor(nota: dict, indice: dict[str, list[dict]]) -> list[dict]:
    """Os titulos da SF1 que sao deste fornecedor (CNPJ ou razao social)."""
    pool, vistos = [], set()
    raiz = _raiz(nota[CNPJ_EMIT])
    ancora = sc7.ancora_nome(sc7.normalizar_nome(nota.get(NOME_EMIT)))
    for balde in (("c:" + raiz) if raiz else "", ("n:" + ancora) if ancora else ""):
        for t in indice.get(balde, ()):
            marca = id(t)
            if marca in vistos:
                continue
            vistos.add(marca)
            if _mesmo_fornecedor(nota, t):
                pool.append(t)
    return pool


# Quantos titulos provaveis cabem na celula, com "+N" no fim para o resto.
# Medido em 09/09/2026: das 233 notas em NAO ACHADA, 64 tem provavel e a maior
# lista tem 2 -- o teto e rede para o dia em que um fornecedor emitir muita nota
# de mesmo valor, nao corte do dia a dia. Corte calado pareceria a lista inteira.
MAX_PROVAVEIS = 8
JUNTA = " · "
# separador ENTRE titulos no quadro (o " · " separa os campos DENTRO de um)
BLOCO = "   ▪ "


def _moeda(valor) -> str:
    if valor in (None, ""):
        return ""
    return f"R$ {valor:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")


def _dia(data) -> str:
    return data.strftime("%d/%m/%Y") if data else ""


def _sinais_do_provavel(nota: dict, titulo: dict) -> list[str]:
    """O que sustenta ESTE titulo como provavel -- em ordem de peso."""
    sinais = []
    if _mesmo_cnpj(nota, titulo):
        sinais.append("mesmo CNPJ")
    elif _mesmo_emissor(nota.get(NOME_EMIT), titulo.get(RAZAO)):
        sinais.append("mesmo nome")
    if _mesmo_valor(nota, titulo):
        sinais.append("valor exato")
    if _mesma_emissao(nota, titulo):
        sinais.append("mesma emissão")
    return sinais


def possibilidades(cruzadas: list[dict], indice_forn: dict[str, list[dict]],
                   ocupados: set) -> int:
    """Preenche `Títulos Prováveis na SF1` -- e, no caso mais forte, o veredito.

    09/09/2026, ele: *"GERE POSSIBILIDADE USANDO O NOME DO FORNECEDOR VALOR E
    DATA TAMBEM"*. E a rede para a nota lancada com numero de NF que nao tem
    NADA a ver com o da SEFAZ -- nem depois de tirar o ano. O caso que ele
    mandou: NFS-e `2026000000009` da ABEL, R$ 4.822,58 de 31/08, lancada na SF1
    como `043001/000020269`, R$ 4.822,58 de 31/08. Mesmo CNPJ, mesmo centavo,
    mesmo dia -- e numero de NF sem parentesco nenhum.

    **ADMISSAO: fornecedor + (valor exato OU emissao no mesmo dia).** Fornecedor
    sozinho NAO entra: 150 das 233 notas em NAO ACHADA tem algum titulo do
    fornecedor e nada mais, e listar 20 titulos de um fornecedor nao responde
    "esta nota foi lancada?" -- seria a coluna de possibilidades dos pedidos
    (que e so por fornecedor) resolvendo outra pergunta.

    ⚠ O VEREDITO SO MUDA NO TRIPLO, E SO ATE `CONFERIR`. Fornecedor + valor
    exato + emissao no mesmo dia, candidato UNICO e titulo que nao e de outra
    nota: ai a nota deixa de ser NAO ACHADA e vira CONFERIR -- o degrau que este
    modulo criou justamente para "algo forte sustenta, mas o numero nao
    confirma". Nao vira ACHADA: quem afirma lancamento aqui e o numero da NF ou
    a chave de 44 digitos, e nenhum dos dois bate.

    ⚠ E POR QUE `CONFERIR` E NAO "DEIXA EM NAO ACHADA COM A LISTA DO LADO":
    porque NAO ACHADA alimenta o `nao_lancadas`, que dispara os alertas de
    "SEM LANCAMENTO NA SF1 HA MAIS DE 3 DIAS" -- inclusive os dois diarios da
    Graziela. Nota que esta lancada nao pode ficar cobrando lancamento.

    ⚠ TITULO CONTESTADO NAO E DE NINGUEM. As notas 564 e 565 da SRS (mesmo
    valor, mesmo dia) apontam as duas para o titulo `000000566`: promover as
    duas diria que o mesmo lancamento e de duas notas. Titulo pedido por 2+
    notas fica na LISTA das duas e nao promove nenhuma.

    ⚠ E TITULO CUJO NUMERO E DE OUTRA NOTA DA SEFAZ TAMBEM NAO PROMOVE. A NF-e
    12766 da JAIR PINTO (R$ 32,00 de 04/08) casaria no triplo com o titulo
    `000012769` -- mas 12769 e o numero de OUTRA nota da propria base, do mesmo
    fornecedor, do mesmo dia e do mesmo valor. O titulo tem dono, e nao e esta
    nota; promover aqui apagaria o alerta de uma nota que talvez nunca tenha
    sido lancada. Na base de 09/09/2026 essa guarda barra 1 promocao e mantem as
    4 legitimas. O titulo continua na LISTA, marcado.

    Devolve quantas linhas ganharam veredito novo.
    """
    # De quem e cada numero de NF da SEFAZ -- para a guarda do "titulo de outra
    # nota". Nota sem numero legivel nao entra (nao reivindica numero nenhum).
    dono_do_numero: dict[str, set] = {}
    for linha in cruzadas:
        numero = sefaz.nf_chave(linha[NUM_NF])
        if numero:
            dono_do_numero.setdefault(numero, set()).add(linha[CHAVE_NFE])

    # 1a volta: quem quer qual titulo (so o triplo pode promover)
    querem: dict[tuple, list[dict]] = {}
    for linha in cruzadas:
        if linha.get("_com_fornecedor"):
            continue
        pool = _titulos_do_fornecedor(linha, indice_forn)
        provaveis = [t for t in pool
                     if _mesmo_valor(linha, t) or _mesma_emissao(linha, t)]
        # ordem: os dois sinais > valor > emissao > numero do titulo
        provaveis.sort(key=lambda t: (
            -(2 * _mesmo_valor(linha, t) + _mesma_emissao(linha, t)), t[NF_SF1]))
        linha["_provaveis"] = provaveis
        if linha[SITUACAO] != NAO_ACHADA:
            continue
        triplo = [t for t in provaveis
                  if _mesmo_valor(linha, t) and _mesma_emissao(linha, t)
                  and _titulo_id(t) not in ocupados
                  and not (dono_do_numero.get(t["nf"], set()) - {linha[CHAVE_NFE]})]
        if len(triplo) == 1:
            querem.setdefault(_titulo_id(triplo[0]), []).append(linha)
            linha["_triplo"] = triplo[0]

    promovidas = 0
    for linha in cruzadas:
        provaveis = linha.get("_provaveis") or []
        if not provaveis:
            continue

        alvo = linha.get("_triplo")
        # ⚠ `amarrado` e o titulo que REALMENTE foi para a linha, e nao o
        # candidato: titulo contestado por 2+ notas (SRS 564/565 x 000000566)
        # nao promove ninguem, e a lista nao pode dizer "amarrado a esta nota".
        amarrado = None
        if alvo is not None and len(querem.get(_titulo_id(alvo), ())) == 1:
            amarrado = alvo
            linha[SITUACAO] = CONFERIR
            _amarrar(linha, alvo)
            porque = [
                "Nº da NF NÃO confere com nenhum título da SF1 — "
                "casou pelo fornecedor + valor exato + emissão no mesmo dia",
                f"Título na SF1: {alvo[FILIAL]}/{alvo[NF_SF1]} "
                f"({_moeda(alvo[VLR_TITULO])}, emissão {_dia(alvo[EMISSAO_SF1])})",
                "Confirme no TOTVS antes de dar a nota por lançada — "
                "o número da NF foi digitado diferente do da SEFAZ",
            ]
            if alvo["status"] and alvo["status"].upper() != "A":
                porque.append(f"Status do título na SF1: {alvo['status']}")
            linha[CRITERIO] = " + ".join(porque)
            promovidas += 1

        mostrados = provaveis[:MAX_PROVAVEIS]
        celula, blocos = [], []
        for t in mostrados:
            rotulo = f"{t[FILIAL]}/{t[NF_SF1]}"
            if rotulo not in celula:
                celula.append(rotulo)
            # ⚠ O AVISO DE TITULO OCUPADO E O QUE IMPEDE LANCAMENTO EM DOBRO.
            # Metade dos provaveis desta base sao titulos que JA sao de outra
            # nota (o `000012769` da JAIR PINTO e da NF 12769, nao da 12766):
            # sem dizer isso, a lista convidaria a amarrar a nota errada.
            if amarrado is not None and _titulo_id(t) == _titulo_id(amarrado):
                marca = "  ·  ← amarrado a esta nota (Situação CONFERIR)"
            elif _titulo_id(t) in ocupados:
                marca = "  ·  ⚠ este título já está amarrado a OUTRA nota da SEFAZ"
            elif dono_do_numero.get(t["nf"], set()) - {linha[CHAVE_NFE]}:
                marca = ("  ·  ⚠ o nº deste título é o da NF "
                         f"{t['nf']} da SEFAZ — provavelmente é dela")
            else:
                marca = ""
            # ⚠ A SERIE ENTRA AQUI. A JAIR PINTO tem DOIS titulos `000012769`,
            # series "001" e "1" -- sem a serie o quadro mostrava a mesma linha
            # duas vezes com marcas diferentes e parecia erro de montagem.
            blocos.append(
                f"{rotulo}"
                + (f" série {t[SERIE]}" if t[SERIE] else "")
                + f"  {_moeda(t[VLR_TITULO])}  emissão {_dia(t[EMISSAO_SF1])}"
                + f"  ·  {t[RAZAO]}  ·  " + ", ".join(_sinais_do_provavel(linha, t))
                + marca)
        sobra = len(provaveis) - len(mostrados)
        linha[PROVAVEIS] = JUNTA.join(celula) + (f" +{sobra}" if sobra else "")
        linha[PORQUE_PROVAVEIS] = (
            f"{len(provaveis)} título{'s' if len(provaveis) > 1 else ''} da SF1 "
            f"pode{'m' if len(provaveis) > 1 else ''} ser esta nota. O cruzamento "
            "é pelo FORNECEDOR (CNPJ ou razão social) mais valor exato ou emissão "
            "no mesmo dia — o número da NF não entra, porque é justamente ele que "
            "não bate. " + ("O mais forte já foi amarrado à linha e a Situação "
                            "virou CONFERIR."
                            if amarrado is not None
                            else "Nenhum foi escolhido: é material de conferência.")
            # ⚠ `BLOCO` E NAO "\n": o `gerar_painel.texto()` colapsa QUALQUER
            # espaco em branco (re.sub de \s+ por " ") antes de a coluna virar
            # celula, entao a quebra de linha chegava ao quadro como espaco e os
            # titulos saiam num paragrafo unico -- mesmo com o `white-space:
            # pre-wrap` do `.cartao.texto .frase`. O marcador sobrevive.
            + BLOCO + BLOCO.join(blocos))
    return promovidas


def cruzar(notas: list[dict], titulos: list[dict]) -> list[dict]:
    """Para cada nota da SEFAZ, o titulo da SF1 que a lancou -- em tres passes.

    1. **numero da NF identico** (o de sempre): a chave da busca e o numero, e o
       resto -- chave de 44 digitos, CNPJ, valor, emissao, nome -- diz o quanto
       dar de fe. Ver `_nota_do_par`.
    2. **numero da NF sem o ano grudado** (09/09/2026, pedido dele), so para as
       notas que o passe 1 NAO confirmou. Ver `sefaz.nf_variantes`.
    3. **sem numero nenhum**: fornecedor + valor + emissao. Ver `possibilidades`.

    ⚠ OS PASSES 2 E 3 NUNCA MEXEM EM QUEM JA ESTA `ACHADA`. Nao e cautela vaga:
    o passe 2 acha candidato para 6 notas que ja estavam ACHADA (G28, DELTA,
    AGR...) e algumas com valor E emissao batendo -- em cima de um titulo que
    provavelmente e lancamento em duplicidade na SF1. Deixar o passe 2 escolher
    ali trocaria o titulo mostrado numa linha que ja estava certa, e a foto
    antes/depois acusaria mudanca em linha que ninguem pediu para mudar.
    """
    exato, formas = _por_numero(titulos)
    indice_forn = _por_fornecedor(titulos)

    # ------------------------------------------------------- passe 1: numero igual
    resultado = []
    for nota in notas:
        numero = sefaz.nf_chave(nota[NUM_NF])
        candidatos = exato.get(numero, []) if numero else []
        linha = dict(nota)
        linha[CANDIDATOS] = len(candidatos)

        if not candidatos:
            linha[SITUACAO] = NAO_ACHADA
            linha[CRITERIO] = ("Nenhum título na SF1 com este número de NF"
                               if numero else "Nota sem número de NF legível")
            resultado.append(linha)
            continue

        avaliados = [(*_nota_do_par(nota, t), t) for t in candidatos]
        pontos, porque, melhor = max(avaliados, key=lambda a: a[0])
        # Confirmado = a chave de 44 digitos ou o emitente. So o numero nao basta.
        confirmado = pontos >= PESO_CNPJ

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
        _amarrar(linha, melhor)
        if not confirmado:
            porque.append("Emitente NÃO confirmado — pode ser outro fornecedor com o mesmo número")
        if melhor["status"] and melhor["status"].upper() != "A":
            porque.append(f"Status do título na SF1: {melhor['status']}")
        linha[CRITERIO] = " + ".join(porque)
        resultado.append(linha)

    # -------------------------------------------- passe 2: numero sem o ano grudado
    for linha in resultado:
        if linha[SITUACAO] == ACHADA:
            continue
        ja = {id(t) for t in exato.get(sefaz.nf_chave(linha[NUM_NF]), ())}
        ano = linha[EMISSAO].year if linha[EMISSAO] else None
        candidatos = []
        for numero, como_nota in sefaz.nf_variantes(linha[NUM_NF], ano).items():
            for titulo, como_sf1 in formas.get(numero, ()):
                if id(titulo) in ja:
                    continue
                # os dois lados com o numero cru e o caminho do passe 1
                if not como_nota and not como_sf1:
                    continue
                # ⚠ A GUARDA QUE FAZ O PASSE 2 SER SEGURO. Tirar o ano AFROUXA a
                # chave, e chave frouxa sozinha nao pode virar achado: exige o
                # fornecedor E (valor exato OU mesma emissao). Medido: 1.692
                # pares caem no fornecedor e 3 no valor/emissao -- entre eles o
                # titulo de R$ 15.073,20 de 2025 (JOSE ESTEVAO, numero 20258)
                # que a nota 2026000000008 casaria por "2025 no inicio".
                if not _mesmo_fornecedor(linha, titulo):
                    continue
                if not (_mesmo_valor(linha, titulo) or _mesma_emissao(linha, titulo)):
                    continue
                onde = JUNTA.join(
                    x for x in (f"na SEFAZ, {como_nota}" if como_nota else "",
                                f"na SF1, {como_sf1}" if como_sf1 else "") if x)
                candidatos.append((titulo, f"Nº da NF confere sem o ano grudado ({onde})"))
        if not candidatos:
            continue
        avaliados = [(*_nota_do_par(linha, t, como), t) for t, como in candidatos]
        pontos, porque, melhor = max(avaliados, key=lambda a: a[0])
        linha[CANDIDATOS] = linha.get(CANDIDATOS, 0) + len(candidatos)
        linha[SITUACAO] = ACHADA if pontos >= PESO_CNPJ else CONFERIR
        _amarrar(linha, melhor)
        if pontos < PESO_CNPJ:
            porque.append("Emitente confere pelo NOME, não pelo CNPJ — confira no TOTVS")
        if melhor["status"] and melhor["status"].upper() != "A":
            porque.append(f"Status do título na SF1: {melhor['status']}")
        linha[CRITERIO] = " + ".join(porque)

    # ------------------------- passe 3: sem numero -- fornecedor + valor + emissao
    for linha in resultado:
        titulo = linha.get("_titulo")
        linha["_com_fornecedor"] = bool(titulo is not None
                                        and _mesmo_fornecedor(linha, titulo))
    ocupados = {_titulo_id(l["_titulo"]) for l in resultado
                if l.get("_titulo") is not None and l[SITUACAO] != NAO_ACHADA}
    possibilidades(resultado, indice_forn, ocupados)
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
