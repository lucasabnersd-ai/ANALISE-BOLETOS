# -*- coding: utf-8 -*-
"""CT-e: a aba "PREMISSA 2 SEFAZ" da SEFAZ.xlsx -- conhecimentos de transporte.

10/09/2026, pedido do usuario: *"INCLUI A ABA PREMISSA 2 SEFAZ, SEMELHANTE A DA
BASE. FACA ANALISE DELA DE MANEIRA CRITICA, ANALITICA E PESADA PARA CRUZAR AS
INFORMACOES E TRAZER AS MESMAS ANALISES DE CRUZAMENTO. CRUZE OS DADOS DE
PRESTADOR PARA VER CRUZAMENTO POR FORNECEDOR. NELA CONSIDERE O NUMERO DE CT-e
(Num CT-e) PARA CRUZAR COM OS NUMEROS DAS NFS E VER SE ESTA NA SF1"*.

O QUE A ABA E
-------------
A exportacao de CT-e (modelo 57, o documento fiscal do FRETE) que o Transmite
gera: 221 colunas, UMA LINHA POR COMPONENTE DO FRETE (FRETE PESO, PEDAGIO,
GRIS/ADEME, IMPOSTO...) -- em 10/09/2026, 19 linhas para 3 CT-e. O
transportador (PAULINERIS, RODONAVES) e o PRESTADOR; a nossa empresa e o
TOMADOR; a mercadoria transportada vem como `Chave NF-e Inf. NFe` (a NF-e do
fornecedor da peca, ex.: INNOVA TRACTOR).

O CABECALHO ORIGINAL ESTA DESALINHADO DOS DADOS a partir da coluna EY:
`CNPJ Emitente` traz "000", `Nome Fantasia Emitente` traz "0.10". O usuario
acrescentou a mao dois rotulos nas colunas certas -- `CNPJ` (GO) e `PRESTADOR`
(GQ) -- e o tomador esta em colunas SEM NOME (HI/HK). Por isso quem identifica
o prestador e a CHAVE DE ACESSO, e nao o cabecalho: a chave de 44 digitos
carrega o CNPJ do emitente nas posicoes 7-20 e o numero nas 26-34
(UF 2 + AAMM 4 + CNPJ 14 + modelo 2 + serie 3 + numero 9 + tpEmis 1 + cNF 8 +
DV 1). Conferido nos 3 CT-e da base contra as colunas GO e G. O nome do
prestador vem do rotulo `PRESTADOR` quando existe; o tomador e a coluna cujos
CNPJs estao TODOS na LISTAGEM EMPRESAS BIOFLOR (a mais a direita) -- e o nome
dele, a primeira coluna de texto a direita dela.

COMO ENTRA NO PAINEL
--------------------
Uma linha por CT-e (os componentes sao SOMADOS no texto `Informações`), no
vocabulario das linhas da SEFAZ, com `Origem = "CT-e"`:
  Nº NF = Num CT-e            (e ele que cruza com o `Numero` da SF1)
  Chave NF-e = Chave CT-e     (a SF1 guarda a chave do CT-e em `Chave NFe`:
                               81 titulos com chave de modelo 57 na base)
  Emitente/CNPJ = prestador   (cruza com o codigo do fornecedor e a razao)
  Destinatário/CNPJ = tomador (o corte do grupo BIOFLOR, como as outras)
  Vlr Total = Vlr. Total Prest.
Dali segue o caminho de todas: `cruzamento_sefaz_sf1` (esta lancada?), pedidos
(SC7) e alertas. Medido em 10/09/2026: dos 3 CT-e, 2 estao na SF1 com o MESMO
numero, fornecedor, valor e dia (especie CTE); o terceiro esta `Cancelado` na
SEFAZ e nao esta na SF1 -- que e o certo.

ENTRA DEPOIS do historico e da manifestacao, como a PREMISSA 2: CT-e nao pode
ir para o `sefaz_historico.json` (que acusa nota que SUMIU da SEFAZ) nem ser
marcado "NAO ENCONTRADA" pela manifestacao (que so existe para NF-e).
A ABA E OPCIONAL: sem ela a rodada avisa e segue.
"""
from __future__ import annotations

import hashlib
import re
import shutil
from pathlib import Path

import sefaz

ORIGEM = "CT-e"

# Colunas lidas pelo CABECALHO (comparadas por `sefaz._rotulo`).
COL_CHAVE = "Chave CT-e"
COL_NUM = "Num CT-e"
COL_STATUS = "Status da Nota"
COL_EMISSAO = "Dt. Emissão"
COL_CFOP = "CFOP"
COL_NAT = "Nat. Operação"
COL_TIPO = "Tp. CT-e"
COL_SERIE = "Série CT-e"
COL_VLR_TOTAL = "Vlr. Total Prest."
COL_VLR_RECEBER = "Vlr. Receber Prestador"
COL_COMP_NOME = "Nome Comp. Prest."
COL_COMP_VLR = "Vlr. Comp. Prest."
COL_NFE_TRANSP = "Chave NF-e Inf. NFe"
COL_VLR_CARGA = "Vlr. Tot. Carga"
COL_PROD_CARGA = "Prod. Pred. Carga"
COL_MUN_INI = "Desc. Mun. Início"
COL_UF_INI = "UF Início"
COL_MUN_FIM = "Desc. Mun. Término"
COL_UF_FIM = "UF Término"
COL_VENC_DUP = "Dt. Venc. Duplicata"
COL_VLR_DUP = "Vlr. Duplicata"
# os dois rotulos que o usuario escreveu a mao; podem nao existir
COL_CNPJ_PREST = "CNPJ"
COL_NOME_PREST = "PRESTADOR"
IDENTIDADE = (COL_CHAVE, COL_NUM)

_DIGITOS = re.compile(r"\D")


def _digitos(valor) -> str:
    return _DIGITOS.sub("", str(valor or ""))


def cnpj_da_chave(chave: str) -> str:
    """CNPJ do emitente embutido na chave de acesso (posicoes 7-20)."""
    chave = _digitos(chave)
    return chave[6:20] if len(chave) == 44 else ""


def numero_da_chave(chave: str) -> str:
    """Numero do documento embutido na chave de acesso (posicoes 26-34)."""
    chave = _digitos(chave)
    return chave[25:34].lstrip("0") if len(chave) == 44 else ""


def _achar_aba(wb):
    procurados = {sefaz._rotulo(c) for c in IDENTIDADE}
    for ws in wb.worksheets:
        cabecalho = {sefaz._rotulo(c) for c in sefaz._cabecalho_de(ws) if c}
        if procurados <= cabecalho:
            return ws
    raise RuntimeError(
        "A SEFAZ.xlsx nao tem a aba de CT-e (PREMISSA 2 SEFAZ): procurei uma aba "
        "com " + " / ".join(repr(c) for c in IDENTIDADE) + " no cabecalho.\n"
        "  a planilha tem: " + " / ".join(repr(n) for n in wb.sheetnames))


def _coluna_tomador(tabela: list[tuple], permitidos) -> tuple[int | None, int | None]:
    """(indice do CNPJ do tomador, indice do nome dele) -- por CONTEUDO.

    O tomador e a coluna em que TODO valor preenchido e um CNPJ de 14 digitos e
    a MAIOR PARTE deles esta na LISTAGEM (empate: a mais a direita). O nome e a
    primeira coluna a direita com texto.

    ⚠ 18/09/2026: era "TODOS na listagem", e isso derrubou a aba inteira: dos 6
    CT-e da exportacao, UM tinha tomador fora da listagem (32293283000284) e a
    coluna certa (HI, 28 de 29 na listagem) foi rejeitada -- zero CT-e no
    painel, so um AVISO de "sem coluna de tomador". A coluna e escolhida pela
    PROPORCAO; o CT-e de fora e cortado linha a linha pelo `linhas_novas`, como
    qualquer nota fora do grupo. O recebedor (HE, 13 de 19) perde por proporcao,
    que e o que se quer: la a nossa empresa aparece so em parte das linhas.
    """
    if not tabela:
        return None, None
    largura = max(len(l) for l in tabela)
    escolhida, melhor = None, 0.0
    for i in range(largura):
        valores = [_digitos(l[i]) for l in tabela if i < len(l) and l[i] not in (None, "")]
        if not valores or not all(len(v) == 14 for v in valores):
            continue
        proporcao = sum(1 for v in valores if v in permitidos) / len(valores)
        if proporcao > 0 and proporcao >= melhor:
            escolhida, melhor = i, proporcao
    if escolhida is None:
        return None, None
    nome = None
    for j in range(escolhida + 1, largura):
        textos = [str(l[j]) for l in tabela if j < len(l) and l[j] not in (None, "")]
        if textos and all(re.search(r"[A-Za-z]{3}", t) for t in textos):
            nome = j
            break
    return escolhida, nome


def _coluna_prestador(tabela: list[tuple], i_chave: int | None) -> tuple[int | None, int | None]:
    """(indice do CNPJ do prestador, indice do nome dele) -- por CONTEUDO.

    18/09/2026: os rotulos `CNPJ`/`PRESTADOR` que o usuario tinha escrito a mao
    nas colunas GO/GQ nao vieram na exportacao nova, e o nome do transportador
    saia VAZIO. O CNPJ do prestador ja vem da chave (posicoes 7-20); aqui se
    procura a coluna em que TODA linha repete exatamente esse CNPJ (a GO), e o
    nome e a primeira coluna a direita dela com texto (a GQ). Sem rotulo, sem
    contagem de niveis: se a exportacao mudar as colunas de lugar, a busca
    acompanha.
    """
    if not tabela or i_chave is None:
        return None, None
    largura = max(len(l) for l in tabela)
    escolhida = None
    for i in range(largura):
        if i == i_chave:
            continue
        pares = [(_digitos(l[i]), cnpj_da_chave(l[i_chave]))
                 for l in tabela if i < len(l) and l[i] not in (None, "")]
        if pares and all(c and v == c for v, c in pares):
            escolhida = i
            break
    if escolhida is None:
        return None, None
    nome = None
    for j in range(escolhida + 1, largura):
        textos = [str(l[j]) for l in tabela if j < len(l) and l[j] not in (None, "")]
        if textos and all(re.search(r"[A-Za-z]{3}", t) for t in textos):
            nome = j
            break
    return escolhida, nome


def ler(base: Path | None = None, permitidos=None) -> tuple[list[dict], dict]:
    """Uma linha por CT-e, no vocabulario da SEFAZ. Devolve (linhas, resumo).

    `permitidos` ({CNPJ: razao} da listagem) e usado para ACHAR a coluna do
    tomador; o corte por grupo e feito em `linhas_novas`.
    """
    base = base or sefaz.BASE_PADRAO
    permitidos = sefaz.ler_cnpjs_empresas() if permitidos is None else permitidos
    wb, temporaria = sefaz._abrir(base)
    try:
        ws = _achar_aba(wb)
        titulo = ws.title
        brutas = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
        if temporaria:
            shutil.rmtree(temporaria, ignore_errors=True)
    if not brutas:
        return [], {"aba": titulo, "linhas": 0, "ctes": 0}
    cabecalho = [sefaz._rotulo(c) for c in brutas[0]]
    dados = [b for b in brutas[1:] if not all(v in (None, "") for v in b)]

    def indice(nome):
        rot = sefaz._rotulo(nome)
        return cabecalho.index(rot) if rot in cabecalho else None

    pos = {nome: indice(nome) for nome in (
        COL_CHAVE, COL_NUM, COL_STATUS, COL_EMISSAO, COL_CFOP, COL_NAT, COL_TIPO,
        COL_SERIE, COL_VLR_TOTAL, COL_VLR_RECEBER, COL_COMP_NOME, COL_COMP_VLR,
        COL_NFE_TRANSP, COL_VLR_CARGA, COL_PROD_CARGA, COL_MUN_INI, COL_UF_INI,
        COL_MUN_FIM, COL_UF_FIM, COL_VENC_DUP, COL_VLR_DUP, COL_CNPJ_PREST,
        COL_NOME_PREST)}
    i_tom, i_tom_nome = _coluna_tomador(dados, permitidos)
    # Sem os rotulos escritos a mao, o prestador e achado pelo conteudo.
    if pos[COL_CNPJ_PREST] is None or pos[COL_NOME_PREST] is None:
        i_prest, i_prest_nome = _coluna_prestador(dados, pos[COL_CHAVE])
        if pos[COL_CNPJ_PREST] is None:
            pos[COL_CNPJ_PREST] = i_prest
        if pos[COL_NOME_PREST] is None:
            pos[COL_NOME_PREST] = i_prest_nome

    def campo(b, nome):
        i = pos[nome]
        return b[i] if i is not None and i < len(b) else None

    def cel(b, i):
        return b[i] if i is not None and i < len(b) else None

    por_chave: dict[str, dict] = {}
    ordem: list[str] = []
    for b in dados:
        chave = _digitos(campo(b, COL_CHAVE))
        if len(chave) != 44:
            continue
        cte = por_chave.get(chave)
        if cte is None:
            # o CNPJ da chave e a verdade; o rotulo escrito a mao so confirma
            cnpj_prest = cnpj_da_chave(chave) or _digitos(campo(b, COL_CNPJ_PREST))
            nome_prest = sefaz._texto(campo(b, COL_NOME_PREST))
            fantasia = ""
            if pos[COL_NOME_PREST] is not None:
                # a coluna logo a direita de PRESTADOR e a fantasia ("PTE MATRIZ")
                prox = cel(b, pos[COL_NOME_PREST] + 1)
                if prox not in (None, "") and re.search(r"[A-Za-z]{2}", str(prox)):
                    fantasia = sefaz._texto(prox)
            numero = sefaz._texto(campo(b, COL_NUM)) or numero_da_chave(chave)
            cnpj_tom = _digitos(cel(b, i_tom)) if i_tom is not None else ""
            nome_tom = sefaz._texto(cel(b, i_tom_nome)) if i_tom_nome is not None else ""
            cte = por_chave[chave] = {
                sefaz.ORIGEM: ORIGEM,
                sefaz.CHAVE_NFE: chave,
                sefaz.NUM_NF: numero,
                sefaz.EMISSAO: sefaz._data(campo(b, COL_EMISSAO)),
                sefaz.STATUS_NOTA: sefaz._texto(campo(b, COL_STATUS)),
                sefaz.SAIDA: None,
                sefaz.TIPO_OP: sefaz._texto(campo(b, COL_TIPO)),
                sefaz.NAT_OP: sefaz._texto(campo(b, COL_NAT)),
                sefaz.CFOP: sefaz._texto(campo(b, COL_CFOP)),
                sefaz.VLR_ITEM: None,
                sefaz.VLR_TOTAL: sefaz._numero(campo(b, COL_VLR_TOTAL)),
                sefaz.VENC_DUP: sefaz._data(campo(b, COL_VENC_DUP)),
                sefaz.VLR_DUP: sefaz._numero(campo(b, COL_VLR_DUP)),
                sefaz.CNPJ_EMIT: cnpj_prest,
                sefaz.NOME_EMIT: nome_prest,
                sefaz.FANTASIA: fantasia,
                sefaz.CNPJ_DEST: cnpj_tom,
                sefaz.NOME_DEST: nome_tom or permitidos.get(cnpj_tom, ""),
                sefaz.MANIFESTACAO: "— não se aplica (CT-e)",
                "_serie": sefaz._texto(campo(b, COL_SERIE)),
                "_receber": sefaz._numero(campo(b, COL_VLR_RECEBER)),
                "_vlr_carga": sefaz._numero(campo(b, COL_VLR_CARGA)),
                "_prod_carga": sefaz._texto(campo(b, COL_PROD_CARGA)),
                "_trecho": " → ".join(x for x in (
                    "/".join(t for t in (sefaz._texto(campo(b, COL_MUN_INI)),
                                          sefaz._texto(campo(b, COL_UF_INI))) if t),
                    "/".join(t for t in (sefaz._texto(campo(b, COL_MUN_FIM)),
                                          sefaz._texto(campo(b, COL_UF_FIM))) if t)) if x),
                "_componentes": [],
                "_nfes": [],
                "_linhas": 0,
            }
            ordem.append(chave)
        cte["_linhas"] += 1
        comp = sefaz._texto(campo(b, COL_COMP_NOME))
        vlr = sefaz._numero(campo(b, COL_COMP_VLR))
        if comp and (comp, vlr) not in cte["_componentes"]:
            cte["_componentes"].append((comp, vlr))
        nfe = _digitos(campo(b, COL_NFE_TRANSP))
        if len(nfe) == 44 and nfe not in cte["_nfes"]:
            cte["_nfes"].append(nfe)

    linhas = []
    for chave in ordem:
        cte = por_chave[chave]
        cte[sefaz.INFO] = _informacoes(cte)
        cte[sefaz.CHAVE] = sefaz.PREFIXO + hashlib.sha1(
            f"{ORIGEM}|{chave}".encode("utf-8")).hexdigest()[:16]
        linhas.append(cte)
    resumo = {
        "aba": titulo, "linhas": len(dados), "ctes": len(linhas),
        "coluna_tomador": i_tom, "coluna_prestador": pos[COL_CNPJ_PREST],
        "sem_tomador": sum(1 for l in linhas if not l[sefaz.CNPJ_DEST]),
        "status": _contar(l[sefaz.STATUS_NOTA] for l in linhas),
    }
    return linhas, resumo


def _moeda(valor) -> str:
    if valor in (None, ""):
        return ""
    return f"R$ {valor:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")


def _informacoes(cte: dict) -> str:
    """O texto da coluna `Informações`: o que o CT-e e, componente a componente."""
    partes = [f"CT-e nº {cte[sefaz.NUM_NF]}"
              + (f" série {cte['_serie']}" if cte.get("_serie") else "")
              + f" de {cte[sefaz.NOME_EMIT] or cte[sefaz.CNPJ_EMIT]}"
              + f" (lido da aba de CT-e da SEFAZ.xlsx, {cte['_linhas']} linhas de componentes)"]
    if cte.get("_trecho"):
        partes.append(f"Trecho: {cte['_trecho']}")
    if cte["_componentes"]:
        partes.append("Componentes do frete: " + " · ".join(
            f"{nome} {_moeda(vlr)}".strip() for nome, vlr in cte["_componentes"]))
    if cte.get("_receber") is not None:
        partes.append(f"Valor a receber pelo prestador: {_moeda(cte['_receber'])}")
    if cte.get("_prod_carga") or cte.get("_vlr_carga") is not None:
        partes.append("Carga: " + " ".join(x for x in (
            cte.get("_prod_carga") or "",
            _moeda(cte["_vlr_carga"]) if cte.get("_vlr_carga") is not None else "") if x))
    for nfe in cte["_nfes"]:
        partes.append(f"NF-e transportada: nº {numero_da_chave(nfe)} do CNPJ "
                      f"{cnpj_da_chave(nfe)} (chave {nfe})")
    return " · ".join(partes)


def _contar(valores) -> dict:
    contagem: dict[str, int] = {}
    for v in valores:
        contagem[v or "(vazio)"] = contagem.get(v or "(vazio)", 0) + 1
    return dict(sorted(contagem.items()))


def linhas_novas(linhas_sefaz: list[dict], permitidos: dict,
                 base: Path | None = None) -> tuple[list[dict], dict]:
    """Os CT-e que ENTRAM no painel: do grupo e ausentes das linhas ja lidas.

    Tambem anota, no texto de cada CT-e, se a NF-e transportada esta na base
    da SEFAZ (e de quem e) -- e o elo entre o frete e a compra.
    """
    todas, resumo = ler(base, permitidos)
    ja_no_painel = {_digitos(l.get(sefaz.CHAVE_NFE)) for l in linhas_sefaz}
    emitentes_nfe = {}
    for l in linhas_sefaz:
        ch = _digitos(l.get(sefaz.CHAVE_NFE))
        if len(ch) == 44 and ch not in emitentes_nfe:
            emitentes_nfe[ch] = (l.get(sefaz.NOME_EMIT) or "", l.get(sefaz.NUM_NF) or "")
    do_grupo = [l for l in todas if l[sefaz.CNPJ_DEST] in permitidos]
    novas = []
    for l in do_grupo:
        if l[sefaz.CHAVE_NFE] in ja_no_painel:
            continue
        extras = []
        for nfe in l.get("_nfes", []):
            if nfe in emitentes_nfe:
                nome, num = emitentes_nfe[nfe]
                extras.append(f"a NF-e {num} de {nome} ESTÁ na base da SEFAZ (aba NF-e)")
            else:
                extras.append(f"a NF-e {numero_da_chave(nfe)} transportada NÃO está na base da SEFAZ")
        if extras:
            l[sefaz.INFO] = l[sefaz.INFO] + " · " + " · ".join(extras)
        novas.append(l)
    resumo.update({
        "arquivo": str(base or sefaz.BASE_PADRAO),
        "do_grupo": len(do_grupo),
        "fora_do_grupo": len(todas) - len(do_grupo),
        "ja_no_painel": len(do_grupo) - len(novas),
        "novas": len(novas),
    })
    return novas, resumo
