# -*- coding: utf-8 -*-
"""Aba "NF x Pedido de Compra": qual PC deu origem a esta nota da SEFAZ?

Junta tres bases: a nota (SEFAZ.xlsx), o lancamento dela no TOTVS (SF1, pelo
`cruzamento_sefaz_sf1`) e o pedido de compra (SC7 + SC1, pelo `sc7`).

O QUE MEDIMOS ANTES DE ESCREVER A REGRA (13/08/2026, base real)
---------------------------------------------------------------
617 notas. 609 tem algum texto em `Info Comp` / `Descrição do Serviço`. O numero
do pedido aparece nesse texto -- quando aparece -- em quatro formatos, e essa foi
a primeira tentativa: extrair o numero e conferir se ele existe no SC7.

⚠ ESSA CONFERENCIA QUASE NAO CONFERE NADA, e foi o achado que mudou o desenho.
O SC7 tem os PCs de 1 a 3492 quase sem buraco: QUALQUER numero de 3 ou 4 digitos
abaixo disso "existe no SC7". Prova medida: das notas em que o texto deu um PC
existente, 64 apontavam para um pedido de OUTRO fornecedor -- a nota da KOMATSU
apontando para um pedido da AUTO TRUCK, a da TECSOLOS para um do SUPERMERCADO
JONAS CAMARA. Nao eram pedidos: eram numeros soltos no texto.

Por isso o veredito NAO sai do texto. Ele sai do texto MAIS a corroboracao:

    sinal                  no pedido certo    num pedido errado do mesmo forn.
    valor total identico        23,5%                  0,1%
    valor de um item            (ver abaixo)
    emissao a <= 30 dias        61,7%                 37,4%
    vencimento a <= 7 dias      22,8%                 16,0%

So o VALOR discrimina de verdade. Data de emissao e vencimento praticamente nao
separam -- ha 9,1 pedidos por fornecedor em media (178 no maior), e um terco
deles cai dentro de 30 dias por acaso. Emissao entra como desempate e como
sinal fraco, nunca sozinha.

O ROTULO TAMBEM PESA
--------------------
`PEDIDO DE COMPRA`, `ORDEM DE COMPRA`, `PC` e `OC` sao nossos (73 com o CNPJ
conferindo). `PEDIDO:` sozinho e quase sempre o pedido DO FORNECEDOR
("Nota fiscal referente ao pedido: 4468", "Nro.Pedido: 916616"), por isso vale
menos -- mas nao pode ser jogado fora: 12 notas so tem esse rotulo e mesmo assim
o CNPJ confere.

⚠ NUMERO ACIMA DO MAIOR PC DA BASE NAO E PEDIDO NOSSO. Medido: dos 45 numeros
citados que nao estao no SC7, 27 estao ACIMA de 3492 (o maior PC que existe) --
5692, 35409, 290274, 916616. Sao pedidos do fornecedor. A faixa vem da propria
SC7 (`resumo["maior_pc"]`), com folga, e nao de um numero escrito aqui.
"""

from __future__ import annotations

import datetime as dt
import re
from pathlib import Path

import cruzamento_sefaz_sf1 as csf1
import sc7
import sefaz

PREFIXO = "PC:"

# --------------------------------------------------------------- nomes de saida
CHAVE = "Chave"
SITUACAO = "Situação"
ALERTA = "Alerta"
LANCADA = "Lançada na SF1"
ORIGEM = "Origem"
NUM_NF = "Nº NF"
NF_TOTVS = "NF no TOTVS"
EMISSAO = "Emissão"
NOME_EMIT = "Emitente"
CNPJ_EMIT = "CNPJ Emitente"
FANTASIA = "Nome Fantasia"
NOME_DEST = "Destinatário"
CNPJ_DEST = "CNPJ Destinatário"
TIPO_OP = "Tipo Operação"
NAT_OP = "Nat. Operação"
SAIDA = "Saída/Entrada"
CFOP = "CFOP"
VLR_SEFAZ = "Vlr SEFAZ"
VENC_DUP = "Venc Duplicata"
VLR_DUP = "Vlr Duplicata"
QTD_DUP = "Duplicatas"
# lado do pedido
PC = "Numero PC"
NUM_SC = "Numero da SC"
SOLICITANTE = "Solicitante"
COMPRADOR = "Comprador"
# ⚠ `Usuário SC` (SC7, coluna GU) NAO e coluna desta aba: esta preenchida em 83
# das 24.401 linhas da SC7 e em ZERO dos pedidos que casam com nota -- nasceria
# vazia para sempre. O `sc7.py` continua lendo; para devolver a coluna, basta
# reinserir aqui e no COMPACTA_PEDIDOS.
QTD_CLASSI = "Qtd.a Classi"
# O total do pedido, ao lado do "a classificar" (pedido do usuario, 14/08/2026).
# ⚠ As duas so aparecem quando ESTA NOTA tem pedido amarrado, e ai aparecem
# ate quando valem ZERO: "0 de 23" e a resposta, celula vazia nao e.
QTD_PC = "Quantidade"
CONTROLE = "Controle Ap."
ENCERRADO = "Ped. Encerr."
ENTREGA = "Dt. Entrega"
VENC_PC = "1° Venc"
VLR_PC = "Vlr Pedido"
FORNEC_PC = "Fornecedor do PC"
ITENS_PC = "Itens do PC"
CRITERIO = "Critério"
INFO = "Informações"
CHAVE_NFE = "Chave NF-e"

CABECALHO = [
    CHAVE, SITUACAO, ALERTA, LANCADA, ORIGEM,
    NUM_NF, NF_TOTVS, EMISSAO,
    NOME_EMIT, CNPJ_EMIT, FANTASIA, NOME_DEST, CNPJ_DEST,
    TIPO_OP, NAT_OP, SAIDA, CFOP,
    VLR_SEFAZ, VENC_DUP, VLR_DUP, QTD_DUP,
    PC, NUM_SC, SOLICITANTE, COMPRADOR,
    QTD_CLASSI, QTD_PC, CONTROLE, ENCERRADO, ENTREGA, VENC_PC,
    VLR_PC, FORNEC_PC, ITENS_PC,
    CRITERIO, INFO, CHAVE_NFE,
]

# ------------------------------------------------------------------- vereditos
# ⚠ Sao CINCO e nao tres, e a diferenca entre eles e o que a pessoa faz depois:
#   CONFIRMADO      nada a fazer, o pedido esta amarrado
#   PROVÁVEL        so confirmar (o painel achou, o texto nao dizia)
#   CONFERIR        a nota cita um pedido, mas nada sustenta -- olho humano
#   SEM PEDIDO      a nota nao cita pedido e nada foi achado
#   FORA DO SC7     cita um numero que a base de pedidos nao tem (lacuna de base)
CONFIRMADO = "PEDIDO CONFIRMADO"
PROVAVEL = "PEDIDO PROVÁVEL"
CONFERIR = "CONFERIR PEDIDO"
FORA = "PEDIDO FORA DA BASE"
SEM = "SEM PEDIDO"
# Os DOIS ultimos degraus (14/08 -> 24/08/2026, pedido do usuario): a analise
# AGRUPADA, que roda por ULTIMO e so sobre o que sobrou sem pedido. Uma nota
# pode ser composta por VARIOS pedidos, e um pedido pode gerar VARIAS notas --
# ver `analise_agrupada`. Sao sempre "conferir" (olho humano): casam so por
# valor, dentro do mesmo fornecedor.
AGRUPADO_PC = "PEDIDO AGRUPADO"   # 1 NF = soma de varios PCs
AGRUPADO_NF = "NOTAS AGRUPADAS"   # varias NFs = 1 PC

ORDEM_SITUACAO = {SEM: 0, FORA: 1, CONFERIR: 2, AGRUPADO_NF: 3, AGRUPADO_PC: 4,
                  PROVAVEL: 5, CONFIRMADO: 6}

# ------------------------------------------------------------------- extracao
# Rotulo FORTE: e a nossa ordem de compra, escrita por extenso ou abreviada.
ROTULO_FORTE = (r"P\.?\s?C\.?"
                r"|O\.?\s?C\.?"
                r"|PEDIDO\s+DE\s+COMPRA"
                r"|ORDEM\s+DE\s+COMPRA"
                r"|ORDEM")
# Rotulo FRACO: "PEDIDO" sozinho -- costuma ser o pedido do fornecedor.
ROTULO_FRACO = r"PEDIDO"
# Palavras que, perto do numero, dizem que ele NAO e pedido de compra.
RUIM = re.compile(r"\b(GDS|VENDAS?|NFC-?E|NOTA|NF-?E|SERIE|SÉRIE)\b")

RE_FORTE = re.compile(rf"\b(?:{ROTULO_FORTE})\s*[:\-\.\s]?\s*(\d{{3,6}})(?!\d)")
RE_FRACO = re.compile(rf"\b(?:{ROTULO_FRACO})\s*[:\-\.\s]?\s*(\d{{3,6}})(?!\d)")

# Quanto texto olhar em volta do numero para o veto acima. 22 antes cobre
# "Nota fiscal referente ao pedido:", 8 depois cobre o "/1" e o separador.
ANTES, DEPOIS = 22, 8

# Folga sobre o maior PC da base: uma exportacao um pouco velha nao pode fazer o
# painel rejeitar pedido novo e legitimo.
FOLGA_FAIXA = 100

# ---------------------------------------------------------------------- pesos
# Bits distintos: a soma ordena do mais forte para o mais fraco e o melhor
# candidato sai de um `max`. Mesma ideia do cruzamento_sefaz_sf1.
PESO_TEXTO_FORTE = 32
PESO_TEXTO_FRACO = 16
PESO_CNPJ = 8
PESO_VALOR = 4        # total do pedido = total da nota
PESO_VALOR_ITEM = 2   # total da nota = valor de um item (entrega parcial)
PESO_EMISSAO = 1

DIAS_EMISSAO = 45     # pedido e nota costumam ficar dentro disso
CENTAVO = 0.01

# ---------------------------------------------------- analise agrupada
# Casa por SOMA de valores, dentro do MESMO fornecedor, so no que sobrou sem
# pedido. Combinacao de ate 3 elementos: mais que isso, qualquer valor "fecha"
# com algum somatorio e o sinal deixa de valer (⚠ o mesmo motivo de o 1:1 so
# confiar no valor). Comparacao em CENTAVOS inteiros -- somar float acumula erro
# e "R$ 0,01 de diferenca" viraria falso negativo (ou pior, falso positivo).
MAX_COMBINACAO = 3
# Guarda de custo: fornecedor com muitos itens no balde vira busca combinatoria
# grande. Acima de LIMITE_TRIPLA so tenta pares; acima de LIMITE_TOTAL nem tenta
# (a S&D LOGISTICA sozinha tem centenas de notas). O teto de nos corta o resto.
LIMITE_POOL_TRIPLA = 25
LIMITE_POOL_TOTAL = 400
LIMITE_NOS = 200_000


def _digitos(valor) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def _raiz(cnpj) -> str:
    """8 primeiros digitos do CNPJ -- e o codigo do fornecedor no TOTVS."""
    return _digitos(cnpj)[:8]


def nf_totvs(valor) -> str:
    """A NF como se cola no TOTVS: so digitos, com zeros a frente ate 9.

    Pedido do usuario (13/08/2026). ⚠ E formato de EXIBICAO/copia, nao de
    cruzamento: quem casa NF com a SF1 continua sendo `sefaz.nf_chave()`, que
    tira os zeros dos dois lados.
    """
    numero = _digitos(valor).lstrip("0")
    return numero.rjust(9, "0") if numero else ""


def citados(texto: str, teto: int) -> list[tuple[str, int]]:
    """[(numero normalizado, peso do rotulo)] achados no texto, sem repetir.

    `teto` e o maior PC que a base de pedidos conhece (com folga): numero acima
    disso nao e pedido nosso e nem entra na lista -- e o que separa o nosso PC
    do "Nro.Pedido: 916616" do fornecedor.
    """
    alto = (texto or "").upper()
    achados: dict[str, int] = {}
    for regex, peso in ((RE_FORTE, PESO_TEXTO_FORTE), (RE_FRACO, PESO_TEXTO_FRACO)):
        for m in regex.finditer(alto):
            vizinhanca = alto[max(0, m.start() - ANTES):m.end() + DEPOIS]
            if RUIM.search(vizinhanca):
                continue
            numero = m.group(1).lstrip("0")
            if not numero or (teto and int(numero) > teto):
                continue
            # o mesmo numero achado pelos dois rotulos vale o rotulo mais forte
            achados[numero] = max(achados.get(numero, 0), peso)
    return list(achados.items())


def pontuar(nota: dict, pedido: sc7.Pedido, peso_texto: int = 0) -> tuple[int, list[str]]:
    """Quanto este pedido combina com esta nota, e por que.

    O `porque` vira o texto do quadro que abre no selo -- e o que permite
    discordar do painel sem ter de reabrir as planilhas.
    """
    pontos, porque = peso_texto, []
    if peso_texto == PESO_TEXTO_FORTE:
        porque.append("Nº do pedido citado na nota")
    elif peso_texto == PESO_TEXTO_FRACO:
        porque.append("Nº citado como “pedido” (rótulo fraco: pode ser o pedido do fornecedor)")

    raiz = _raiz(nota.get(CNPJ_EMIT))
    if raiz and raiz in pedido.fornecedores:
        pontos += PESO_CNPJ
        porque.append("Fornecedor confere (raiz do CNPJ = cód. do fornecedor no PC)")

    valor = nota.get(VLR_SEFAZ)
    if valor and pedido.vlr and abs(valor - pedido.vlr) < CENTAVO:
        pontos += PESO_VALOR
        porque.append("Valor total idêntico ao do pedido")
    elif valor and any(abs(valor - v) < CENTAVO for v in pedido.valores_item):
        pontos += PESO_VALOR_ITEM
        porque.append("Valor da nota = valor de um item do pedido (entrega parcial)")

    emissao, emissao_pc = nota.get(EMISSAO), pedido.emissao
    if emissao and emissao_pc:
        dias = abs((emissao - emissao_pc).days)
        if dias <= DIAS_EMISSAO:
            pontos += PESO_EMISSAO
            porque.append(f"Pedido emitido {dias} dia(s) antes/depois da nota")
    return pontos, porque


def _distancia(nota: dict, pedido: sc7.Pedido) -> int:
    """Dias entre a emissao da nota e a do pedido -- so para desempate."""
    if not (nota.get(EMISSAO) and pedido.emissao):
        return 10 ** 6
    return abs((nota[EMISSAO] - pedido.emissao).days)


def escolher(nota: dict, pedidos: dict[str, sc7.Pedido],
             por_fornecedor: dict[str, list[str]], teto: int) -> dict:
    """O pedido desta nota: quem e, com que forca e por que.

    Duas frentes, nesta ordem -- o texto manda, a busca so entra quando o texto
    nao resolve:
      1. os numeros citados no texto da nota;
      2. se nenhum se sustentar, os pedidos DO MESMO FORNECEDOR cujo valor bate.
    """
    achados = citados(nota.get(INFO), teto)
    fora_da_base = [n for n, _p in achados if n not in pedidos]

    avaliados = [(*pontuar(nota, pedidos[n], peso), pedidos[n])
                 for n, peso in achados if n in pedidos]

    # --- 2) sem numero util no texto: procurar pelo fornecedor + valor
    buscados = []
    if not avaliados:
        raiz = _raiz(nota.get(CNPJ_EMIT))
        for numero in por_fornecedor.get(raiz, []):
            pontos, porque = pontuar(nota, pedidos[numero])
            # ⚠ A BARRA E O VALOR. Sem ele sobra "mesmo fornecedor + emissao
            # proxima", que 37% dos pedidos ERRADOS do mesmo fornecedor tambem
            # tem -- viraria um palpite com cara de resposta.
            if pontos >= PESO_CNPJ + PESO_VALOR_ITEM:
                buscados.append((pontos, porque, pedidos[numero]))
        avaliados = buscados

    if not avaliados:
        return {
            SITUACAO: FORA if fora_da_base else SEM,
            CRITERIO: (
                f"A nota cita o pedido {' e '.join(fora_da_base)}, que não existe na "
                f"base de pedidos (SC7) — pedido cancelado, de outra filial ou "
                f"exportação incompleta" if fora_da_base else
                "A nota não cita pedido de compra e nenhum pedido deste fornecedor "
                "bate com o valor dela"),
        }

    melhor_pontos = max(a[0] for a in avaliados)
    empatados = [a for a in avaliados if a[0] == melhor_pontos]
    # Desempate pela emissao mais proxima. Sinal fraco -- por isso NUNCA escolhe
    # sozinho, so decide entre iguais.
    pontos, porque, pedido = min(empatados, key=lambda a: _distancia(nota, a[2]))

    # Os pesos sao bits distintos de proposito: da para perguntar "tem valor?"
    # sem refazer a comparacao e sem depender da ordem em que foram somados.
    tem_texto = bool(melhor_pontos & (PESO_TEXTO_FORTE | PESO_TEXTO_FRACO))
    tem_cnpj = bool(melhor_pontos & PESO_CNPJ)
    tem_valor = bool(melhor_pontos & (PESO_VALOR | PESO_VALOR_ITEM))

    if tem_texto and tem_cnpj:
        situacao = CONFIRMADO
    elif tem_texto and tem_valor:
        situacao = CONFIRMADO if melhor_pontos & PESO_VALOR else PROVAVEL
    elif not tem_texto:
        # veio da busca: sempre precisou de CNPJ + valor para chegar aqui
        situacao = PROVAVEL if melhor_pontos & PESO_VALOR else CONFERIR
    else:
        situacao = CONFERIR

    if len(empatados) > 1:
        outros = " · ".join(sorted(p.numero_totvs for _pt, _pq, p in empatados
                                   if p is not pedido))
        porque = list(porque) + [f"⚠ {len(empatados)} pedidos empataram ({outros}); "
                                 "ficou o de emissão mais próxima"]
        if situacao == CONFIRMADO:
            situacao = PROVAVEL
    if not tem_texto:
        porque = ["Pedido NÃO citado na nota — achado pelo fornecedor e pelo valor"] + list(porque)
    if fora_da_base:
        porque = list(porque) + [f"A nota também cita {' e '.join(fora_da_base)}, "
                                 "que não existe na base de pedidos"]

    return {SITUACAO: situacao, CRITERIO: " + ".join(porque), "pedido": pedido,
            "pontos": melhor_pontos}


def _preencher_pedido(linha: dict, pedido: sc7.Pedido,
                      solicitantes: dict[str, list[str]]) -> None:
    """Escreve as colunas do bloco PEDIDO a partir de um Pedido inteiro."""
    linha.update({
        PC: pedido.numero_totvs,
        NUM_SC: pedido.sc,
        COMPRADOR: pedido.comprador,
        SOLICITANTE: " · ".join(solicitantes.get(pedido.numero, [])),
        QTD_CLASSI: pedido.qtd_classi,
        QTD_PC: pedido.quantidade,
        CONTROLE: pedido.controle,
        ENCERRADO: pedido.encerrado,
        ENTREGA: pedido.entrega,
        VENC_PC: pedido.venc1,
        VLR_PC: pedido.vlr or None,
        FORNEC_PC: pedido.razao,
        ITENS_PC: pedido.itens,
    })


def _centavos(valor) -> int:
    """Valor em centavos inteiros -- e como as somas sao comparadas."""
    return int(round((valor or 0) * 100))


def _combinacoes_soma(itens: list, alvo_cent: int, max_tam: int) -> list[list]:
    """Combinacoes (tam 2..max_tam) de `itens` cujos valores somam `alvo_cent`.

    `itens` = [(ref, centavos)], ordenado do maior para o menor. Devolve listas de
    `ref`. Busca em profundidade, podando quando a parcela nao cabe e parando no
    teto de nos -- o que a torna barata quando NAO ha combinacao (o caso comum).
    """
    if max_tam < 2 or alvo_cent <= 0:
        return []
    n = len(itens)
    achados: list[list] = []
    nos = [0]

    def rec(inicio: int, restante: int, faltam: int, combo: list) -> None:
        if restante == 0:
            if len(combo) >= 2:
                achados.append([itens[i][0] for i in combo])
            return
        if faltam == 0:
            return
        for i in range(inicio, n):
            cent = itens[i][1]
            if cent > restante:
                continue          # ordenado desc: um menor adiante ainda cabe
            nos[0] += 1
            if nos[0] > LIMITE_NOS or len(achados) >= 50:
                return
            combo.append(i)
            rec(i + 1, restante - cent, faltam - 1, combo)
            combo.pop()

    rec(0, alvo_cent, max_tam, [])
    return achados


def _cap_tam(tamanho_pool: int) -> int:
    """Ate quantos elementos combinar, conforme o tamanho do balde."""
    if tamanho_pool > LIMITE_POOL_TOTAL:
        return 0
    if tamanho_pool > LIMITE_POOL_TRIPLA:
        return 2
    return MAX_COMBINACAO


def _moeda(valor) -> str:
    return f"R$ {(valor or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def analise_agrupada(linhas: list[dict], pedidos: dict[str, sc7.Pedido],
                     solicitantes: dict[str, list[str]],
                     pcs_usados: set) -> dict:
    """ULTIMO passo: casa por SOMA o que sobrou em SEM PEDIDO / FORA DA BASE.

    Duas direcoes, nesta ordem (a mais ancorada primeiro):
      B) 1 PC : N NFs  -- varias notas do mesmo fornecedor somam um pedido inteiro;
      A) N PCs : 1 NF  -- uma nota vale a soma de varios pedidos do fornecedor.

    So mexe nas linhas ainda sem pedido; nada do 1:1 e reescrito. `pcs_usados` sao
    os pedidos que o 1:1 ja atribuiu -- ficam de fora dos dois lados para nao
    aparecer o mesmo pedido em dois lugares dizendo coisas diferentes.
    """
    contagem = {AGRUPADO_NF: 0, AGRUPADO_PC: 0}

    # baldes por fornecedor (raiz do CNPJ do emitente da nota)
    por_forn_notas: dict[str, list[dict]] = {}
    for linha in linhas:
        if linha.get(SITUACAO) not in (SEM, FORA):
            continue
        if not linha.get(VLR_SEFAZ):
            continue
        por_forn_notas.setdefault(_raiz(linha.get(CNPJ_EMIT)), []).append(linha)

    por_forn_pedidos: dict[str, list[sc7.Pedido]] = {}
    for numero, pedido in pedidos.items():
        if numero in pcs_usados or not pedido.vlr:
            continue
        for raiz in pedido.fornecedores:
            por_forn_pedidos.setdefault(raiz, []).append(pedido)

    for raiz, notas in por_forn_notas.items():
        peds = por_forn_pedidos.get(raiz, [])
        if not peds:
            continue
        pendentes = {id(l): l for l in notas}   # notas ainda nao agrupadas

        # --- B) 1 PC : N NFs -------------------------------------------------
        # do maior pedido para o menor: o grande e o que mais precisa de varias
        # notas para fechar, e consumir cedo evita casar a nota no pedido errado.
        for pedido in sorted(peds, key=lambda p: p.vlr or 0, reverse=True):
            disponiveis = [(l, _centavos(l.get(VLR_SEFAZ))) for l in pendentes.values()]
            disponiveis = [(l, c) for l, c in disponiveis if c > 0]
            tam = _cap_tam(len(disponiveis))
            if tam < 2:
                continue
            disponiveis.sort(key=lambda ic: ic[1], reverse=True)
            combos = _combinacoes_soma(disponiveis, _centavos(pedido.vlr), tam)
            if not combos:
                continue
            combos.sort(key=len)               # a mais simples ganha
            escolha = combos[0]
            ambiguo = sum(1 for c in combos if len(c) == len(escolha)) > 1
            nfs = " + ".join(f"NF {l.get(NUM_NF) or '?'} ({_moeda(l.get(VLR_SEFAZ))})"
                             for l in escolha)
            for l in escolha:
                _preencher_pedido(l, pedido, solicitantes)
                l[SITUACAO] = AGRUPADO_NF
                l[CRITERIO] = (
                    f"Esta nota faz parte de {len(escolha)} notas do mesmo fornecedor "
                    f"que somam o pedido {pedido.numero_totvs} "
                    f"({_moeda(pedido.vlr)}): {nfs}"
                    + ("  ⚠ há outra combinação possível — conferir" if ambiguo else "")
                    + "  — casado por soma de valores, confira")
                contagem[AGRUPADO_NF] += 1
                pendentes.pop(id(l), None)

        # --- A) N PCs : 1 NF -------------------------------------------------
        itens_pc = [(p, _centavos(p.vlr)) for p in peds if _centavos(p.vlr) > 0]
        tam = _cap_tam(len(itens_pc))
        if tam < 2:
            continue
        itens_pc.sort(key=lambda ic: ic[1], reverse=True)
        for l in list(pendentes.values()):
            combos = _combinacoes_soma(itens_pc, _centavos(l.get(VLR_SEFAZ)), tam)
            if not combos:
                continue
            combos.sort(key=len)
            escolha = combos[0]
            ambiguo = sum(1 for c in combos if len(c) == len(escolha)) > 1
            escolha.sort(key=lambda p: p.numero_totvs)
            l[SITUACAO] = AGRUPADO_PC
            l[PC] = " · ".join(p.numero_totvs for p in escolha)
            l[NUM_SC] = " · ".join(dict.fromkeys(
                s for p in escolha for s in (p.sc.split(" · ") if p.sc else []) if s))
            l[VLR_PC] = round(sum(p.vlr for p in escolha), 2)
            l[FORNEC_PC] = escolha[0].razao
            pcs = " + ".join(f"PC {p.numero_totvs} ({_moeda(p.vlr)})" for p in escolha)
            l[CRITERIO] = (
                f"O valor da nota ({_moeda(l.get(VLR_SEFAZ))}) = soma de {len(escolha)} "
                f"pedidos do mesmo fornecedor: {pcs}"
                + ("  ⚠ há outra combinação possível — conferir" if ambiguo else "")
                + "  — casado por soma de valores, confira")
            contagem[AGRUPADO_PC] += 1

    return contagem


# ---------------------------------------------------------------------- saida
def montar(cruzamento: csf1.Cruzamento, pedidos: dict[str, sc7.Pedido],
           solicitantes: dict[str, list[str]], teto: int,
           texto_por_nota: dict[str, str]) -> list[dict]:
    """Uma linha por NOTA, com o lancamento e o pedido do lado.

    ⚠ `texto_por_nota` chega ANTES do veredito e nao depois: e desse texto que o
    numero do pedido e lido, e preenche-lo so na saida faria a extracao rodar
    sempre contra string vazia -- a aba inteira nasceria em SEM PEDIDO, sem erro
    nenhum aparecer.
    """
    por_fornecedor: dict[str, list[str]] = {}
    for numero, pedido in pedidos.items():
        for raiz in pedido.fornecedores:
            por_fornecedor.setdefault(raiz, []).append(numero)

    # pedidos que o 1:1 ja atribuiu -- a analise agrupada nao os reaproveita
    pcs_usados: set = set()
    saida = []
    for nota in cruzamento.cruzadas:
        # A nota ja vem da visao por NOTA do cruzamento com a SF1: as colunas da
        # SEFAZ ja estao agregadas ali (CFOP junto, 1a duplicata, etc.).
        linha = {
            ORIGEM: nota.get(csf1.ORIGEM),
            CHAVE_NFE: nota.get(csf1.CHAVE_NFE),
            NUM_NF: nota.get(csf1.NUM_NF),
            NF_TOTVS: nf_totvs(nota.get(csf1.NUM_NF)),
            EMISSAO: nota.get(csf1.EMISSAO),
            NOME_EMIT: nota.get(csf1.NOME_EMIT),
            CNPJ_EMIT: nota.get(csf1.CNPJ_EMIT),
            FANTASIA: nota.get(csf1.FANTASIA),
            NOME_DEST: nota.get(csf1.NOME_DEST),
            CNPJ_DEST: nota.get(csf1.CNPJ_DEST),
            TIPO_OP: nota.get(csf1.TIPO_OP),
            NAT_OP: nota.get(csf1.NAT_OP),
            SAIDA: nota.get(csf1.SAIDA),
            CFOP: nota.get(csf1.CFOP),
            VLR_SEFAZ: nota.get(csf1.VLR_SEFAZ),
            VENC_DUP: nota.get(csf1.VENC_DUP),
            VLR_DUP: nota.get(csf1.VLR_DUP),
            QTD_DUP: nota.get(csf1.QTD_DUP),
            # o veredito da OUTRA aba, do lado: "achei o pedido mas a nota nunca
            # foi lancada" e exatamente o caso que interessa
            LANCADA: nota.get(csf1.SITUACAO),
            # Os MESMOS tres alertas das abas da SEFAZ (pedido do usuario,
            # 13/08/2026): vence em ate 5 dias sem lancamento, sem lancamento ha
            # mais de 3 dias, prazo curto entre emissao e vencimento. A regra
            # mora em `sefaz.alertas_da_nota()` -- um lugar so para as tres abas.
            ALERTA: nota.get(csf1.ALERTA, ""),
            INFO: texto_por_nota.get(nota.get(csf1.CHAVE_NFE), ""),
        }
        veredito = escolher(linha, pedidos, por_fornecedor, teto)
        linha[SITUACAO] = veredito[SITUACAO]
        linha[CRITERIO] = veredito[CRITERIO]

        pedido = veredito.get("pedido")
        if pedido is not None:
            # ⚠ o helper preenche QTD_CLASSI/QTD_PC SEM `or None`: aqui o ZERO e
            # informacao ("nada a classificar neste pedido") e some se virar
            # celula vazia -- e vazio, nesta aba, ja quer dizer outra coisa:
            # nota que nao tem pedido nenhum.
            _preencher_pedido(linha, pedido, solicitantes)
            if veredito[SITUACAO] in (CONFIRMADO, PROVAVEL, CONFERIR):
                pcs_usados.add(pedido.numero)
        linha[CHAVE] = PREFIXO + (_digitos(linha[CHAVE_NFE]) or str(linha[CHAVE_NFE]))
        saida.append(linha)

    # ULTIMO passo: o que sobrou sem pedido ainda pode casar por SOMA de valores
    # -- varias notas para um pedido, ou varios pedidos para uma nota. Roda depois
    # de tudo e so mexe nas linhas SEM PEDIDO / FORA DA BASE. Ver `analise_agrupada`.
    analise_agrupada(saida, pedidos, solicitantes, pcs_usados)
    return saida


def montar_linhas(linhas: list[dict]) -> list[tuple]:
    """Linhas -> tuplas na ordem de CABECALHO.

    Ordem: o que ainda nao tem pedido primeiro (e a fila de trabalho), e dentro
    de cada grupo a nota mais recente na frente.
    """
    ordenadas = sorted(linhas, key=lambda l: (
        ORDEM_SITUACAO.get(l[SITUACAO], 9),
        l.get(EMISSAO) is None,
        -(l[EMISSAO].toordinal() if l.get(EMISSAO) else 0),
    ))
    return [tuple(l.get(coluna, "") for coluna in CABECALHO) for l in ordenadas]


def carregar(base_sefaz: Path, base_sf1: Path, base_sc7: Path | None = None,
             base_sc1: Path | None = None, hoje=None,
             pronto: csf1.Cruzamento | None = None,
             texto_por_nota: dict[str, str] | None = None
             ) -> tuple[list[str], list[tuple], dict]:
    """Ponto de entrada: (cabecalho, linhas, resumo) para o gerador."""
    cruzamento = pronto or csf1.preparar(base_sefaz, base_sf1, hoje)
    pedidos, solicitantes, resumo_sc7 = sc7.carregar(base_sc7, base_sc1)
    teto = resumo_sc7["maior_pc"] + FOLGA_FAIXA

    # O texto de onde o pedido e lido nao existe na visao por nota (o cruzamento
    # com a SF1 nao precisa dele): vem das linhas cruas, a primeira de cada nota.
    # ⚠ `Info Comp` se repete em toda linha da nota -- a primeira basta, e e a
    # unica que cabe: o texto chega a 5.000 caracteres.
    if texto_por_nota is None:
        texto_por_nota = {}
        for l in cruzamento.linhas:
            texto_por_nota.setdefault(l[sefaz.CHAVE_NFE], l.get(sefaz.INFO) or "")

    linhas = montar(cruzamento, pedidos, solicitantes, teto, texto_por_nota)

    contagem: dict[str, int] = {}
    for l in linhas:
        contagem[l[SITUACAO]] = contagem.get(l[SITUACAO], 0) + 1
    resumo = {
        **sefaz.contar_alertas(cruzamento.alertas),
        "notas": len(linhas),
        "com_pedido": sum(1 for l in linhas if l.get(PC)),
        "confirmados": contagem.get(CONFIRMADO, 0),
        "provaveis": contagem.get(PROVAVEL, 0),
        "conferir": contagem.get(CONFERIR, 0),
        "agrupado_nf": contagem.get(AGRUPADO_NF, 0),
        "agrupado_pc": contagem.get(AGRUPADO_PC, 0),
        "fora_da_base": contagem.get(FORA, 0),
        "sem_pedido": contagem.get(SEM, 0),
        "com_solicitante": sum(1 for l in linhas if l.get(SOLICITANTE)),
        **{f"sc7_{k}": v for k, v in resumo_sc7.items()},
    }
    return list(CABECALHO), montar_linhas(linhas), resumo
