# -*- coding: utf-8 -*-
"""A manifestação do destinatário, colada nas notas da SEFAZ.

O que esta base acrescenta ao painel e uma coisa que ele nao tinha: o que a
empresa DISSE sobre a nota que originou o boleto -- deu ciencia, confirmou,
desconheceu, ou nao manifestou nada. Boleto a pagar de nota que a propria
empresa marcou como `DesconhecimentoOperacao` ou `NaoRealizada` e o caso que
esta base existe para pegar.

⚠ ESTE MODULO SO ENRIQUECE, NUNCA ACRESCENTA NOTA AO PAINEL. Quem acrescenta e
o premissa2.py, lendo a MESMA base. Medido em 01/09/2026: das 1.380 linhas do
arquivo, 643 sao notas de TERCEIROS para a BRACELL, em que a S&D so aparece
como transportadora -- nota que o grupo nao paga. Como o cruzamento e um
de-para em cima das notas que a SEFAZ ja trouxe, essas se excluem sozinhas.

DE ONDE VEM (18/09/2026)
------------------------
Ate 17/09 a base era o `MESMA PREMISSA.xlsx` (CSV embrulhado em .xlsx, aba de
nome GUID). Em 18/09/2026 o usuario a trouxe para DENTRO da SEFAZ.xlsx, como
aba `PRODUTO 2`, ao lado da `PRODUTO 1` (as NF-e): *"as abas PRODUTO 1 e
PRODUTO 2 se complementam"*. A leitura passou a ser a do premissa2.py
(`premissa2.ler`), que acha a aba pelo cabecalho em qualquer dos dois arquivos;
o `MESMA PREMISSA.xlsx` ficou de reserva (`escolher_base`).

⚠ NA ABA PRODUTO 2 A CHAVE VEIO COMO NUMERO ('4.1260927755427e+43') e perdeu 29
digitos -- ver o topo do premissa2.py. Ate aqui a chave de 44 era a UNICA ponte
entre a manifestacao e a nota da SEFAZ, e chave torta parava a rodada. Agora ha
DUAS pontes, nesta ordem: a chave de 44 quando os dois lados a tem, e a
IDENTIDADE (CNPJ do emitente, numero, serie) -- medida sem colisao nas duas
abas -- com o float da chave de TRAVA (identidade igual e float incompativel
= outra nota). A serie da NF-e entrou na leitura da PRODUTO 1 para isso
(`sefaz.COLUNAS_NFE`, campo interno `_serie`); linha antiga do historico, sem
serie, casa pelo par (CNPJ, numero) so quando ele e unico na base.

A chave era confiavel, e isso foi medido, nao suposto (01/09/2026):
  - 481 das 483 notas NF-e da SEFAZ estavam la; as 2 de fora eram de 01/09,
    fora do recorte do arquivo (que ia de 01/08 a 31/08);
  - `ValorNF` batia com o `Vlr Total NF` da SEFAZ em 413 de 413 notas cruzadas,
    zero divergencia -- prova independente de que a chave casa a nota certa;
  - zero intersecao com NFS-e, e o correto: manifestacao nao existe para nota
    de servico.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import caminhos
import premissa2
import sefaz

# ⚠ Derivado da pasta ANALISES BOLETOS procurada por marca (caminhos.py), e nao
# do `sefaz.BASE_PADRAO`: aquele tem um fallback para a maquina do Lucas.
# E a RESERVA: a fonte preferida e a aba PRODUTO 2 da SEFAZ.xlsx (escolher_base).
BASE_PADRAO = caminhos.pasta_analises() / "MESMA PREMISSA.xlsx"

# As 14 colunas da base (vocabulario do CSV cru); a leitura em si aceita tambem
# o vocabulario da versao organizada -- ver premissa2.SINONIMOS.
COLUNAS = ("NomeFilial", "Numero", "Serie", "DataEmissao", "Chave", "Emissor",
           "CnpjCpfEmitente", "Destinatario", "CnpjCpfDestinatario", "ValorNF",
           "Ambiente", "StatusManifestacao", "NaturezaOperacao",
           "JustificativaManifestacao")

CHAVE = "Chave"
STATUS = "StatusManifestacao"
FILIAL = "NomeFilial"
EMISSAO = "DataEmissao"
JUSTIFICATIVA = "JustificativaManifestacao"

# ⚠ O vocabulario e o da ORIGEM, sem traducao -- mesma regra do `Status da Nota`
# no sefaz.py. Inventar rotulo bonito aqui faria um status novo do TOTVS chegar
# na tela como se fosse conhecido, ou sumir num `else`.
SEM_MANIFESTACAO = "SemManifestacao"
CIENCIA = "CienciaOperacao"
CONFIRMACAO = "ConfirmacaoOperacao"
DESCONHECIMENTO = "DesconhecimentoOperacao"
NAO_REALIZADA = "NaoRealizada"

# Os dois que interessam ao painel de boletos: a empresa disse que nao reconhece
# a operacao, e mesmo assim existe titulo/boleto para pagar.
GRAVES = (DESCONHECIMENTO, NAO_REALIZADA)

# ⚠ NAO SAO STATUS: sao o que NOS dizemos quando a base nao fala da nota. Ficam
# graficamente diferentes dos valores da origem (travessao/caixa alta) de
# proposito -- quem olha a coluna tem de saber, sem perguntar, se aquilo veio do
# TOTVS ou e uma conclusao nossa.
NAO_SE_APLICA = "— não se aplica (NFS-e)"
FORA_DO_PERIODO = "— fora do período do arquivo"
NAO_ENCONTRADA = "NÃO ENCONTRADA"
SEM_BASE = "— base de manifestação não lida"


def escolher_base(explicita: Path | None, base_sefaz: Path | None) -> Path:
    """De onde ler: o que o usuario mandou > a aba PRODUTO 2 da SEFAZ.xlsx > reserva.

    `explicita` igual ao BASE_PADRAO nao conta como mandado: e o default do
    argparse do gerar_painel.py, e ele nao sabe se a SEFAZ.xlsx tem a aba.
    """
    if explicita and Path(explicita) != BASE_PADRAO:
        return Path(explicita)
    if base_sefaz and premissa2.tem_aba(base_sefaz):
        return Path(base_sefaz)
    return BASE_PADRAO


def _data(texto):
    """'04-08-2026T11:03:48AM-03:00' -> date(2026, 8, 4). Reaproveita o premissa2."""
    return premissa2._data(texto)


def _registro(linha: dict) -> dict:
    """Uma linha do premissa2.ler no vocabulario deste modulo."""
    return {
        CHAVE: linha.get(sefaz.CHAVE_NFE) or "",
        STATUS: linha.get(sefaz.MANIFESTACAO) or "",
        FILIAL: linha.get("_filial_nome") or "",
        EMISSAO: linha.get(sefaz.EMISSAO),
        JUSTIFICATIVA: linha.get("_justificativa") or "",
        "_identidade": linha["_identidade"],
        "_chave_float": linha.get("_chave_float"),
        "_chave_origem": linha.get("_chave_origem"),
    }


def ler(base: Path | None = None) -> tuple[dict, dict]:
    """(indices, resumo). Uma entrada por NOTA em cada indice.

    `indices` tem tres de-paras para a mesma colecao de registros:
      por_chave        {chave de 44: registro}  -- so quem tem a chave inteira
      por_identidade   {(cnpj, numero, serie): registro}
      por_cnpj_numero  {(cnpj, numero): registro, ou None quando ambiguo}

    ⚠ A MESMA NOTA VEM MAIS DE UMA VEZ. Medido em 01/09/2026: 56 chaves vinham
    duas vezes, variando `StatusManifestacao` E `NomeFilial` -- a mesma nota
    vista por duas filiais, uma que ja manifestou e outra que nao (em
    18/09/2026, 90 das 2.442 notas da PRODUTO 2). Se o join fosse feito sem
    resolver isso, essas notas duplicariam as linhas da SEFAZ.
    A regra: qualquer status vence `SemManifestacao`, porque `SemManifestacao` e
    ausencia de resposta, e nao uma resposta. Dois status REAIS em conflito
    param a rodada -- nunca aconteceu, e no dia em que acontecer e para alguem
    olhar, nao para o codigo escolher calado.
    """
    base = Path(base or BASE_PADRAO)
    if not base.exists():
        raise RuntimeError(
            f"A base de manifestação não foi encontrada:\n  {base}\n"
            "Sem ela a coluna de manifestação não tem como ser preenchida. "
            "Confira se o OneDrive sincronizou (a aba PRODUTO 2 da SEFAZ.xlsx ou "
            "o MESMA PREMISSA.xlsx), ou corrija BASE_PADRAO no manifestacao.py."
        )
    linhas, leitura = premissa2.ler(base)
    if not linhas:
        raise RuntimeError(f"{base.name} (aba {leitura.get('aba')!r}) está vazio — "
                           "nada para cruzar.")

    por_identidade: dict[tuple, dict] = {}
    conflitos: list[str] = []
    for linha in linhas:
        r = _registro(linha)
        atual = por_identidade.get(r["_identidade"])
        if atual is None:
            por_identidade[r["_identidade"]] = r
            continue
        a, b = atual[STATUS], r[STATUS]
        if a == b:
            continue
        if a == SEM_MANIFESTACAO:
            por_identidade[r["_identidade"]] = r
        elif b != SEM_MANIFESTACAO:
            conflitos.append(f"{r[CHAVE]}: {a!r} ({atual[FILIAL].strip()}) x "
                             f"{b!r} ({r[FILIAL].strip()})")
    if conflitos:
        raise RuntimeError(
            f"{base.name}: {len(conflitos)} nota(s) com DOIS status de "
            "manifestação diferentes, e nenhum deles é 'SemManifestacao':\n  "
            + "\n  ".join(conflitos[:10])
            + "\nAté 18/09/2026 isso nunca aconteceu (todo conflito era um status "
              "real contra 'SemManifestacao'). Olhe a nota antes de escolher: "
              "qualquer regra automática aqui esconderia a divergência."
        )

    por_chave = {premissa2._digitos(r[CHAVE]): r for r in por_identidade.values()
                 if len(premissa2._digitos(r[CHAVE])) == 44}
    por_cnpj_numero: dict[tuple, dict | None] = {}
    for ident, r in por_identidade.items():
        par = ident[:2]
        por_cnpj_numero[par] = None if par in por_cnpj_numero else r

    datas = [r[EMISSAO] for r in por_identidade.values() if r[EMISSAO]]
    status = sorted({r[STATUS] for r in por_identidade.values()})
    resumo = {
        "arquivo": str(base),
        "aba": leitura.get("aba"),
        "linhas": leitura["linhas"],
        "notas": len(por_identidade),
        "repetidas": leitura["linhas"] - len(por_identidade),
        "chaves_44": len(por_chave),
        "chaves_float": leitura.get("chaves_float", 0),
        "recuperadas": leitura.get("recuperadas", 0),
        "de": min(datas) if datas else None,
        "ate": max(datas) if datas else None,
        "por_status": {s: sum(1 for r in por_identidade.values() if r[STATUS] == s)
                       for s in status},
    }
    indices = {"por_chave": por_chave, "por_identidade": por_identidade,
               "por_cnpj_numero": por_cnpj_numero}
    return indices, resumo


def _registro_da_linha(linha: dict, indices: dict) -> tuple[dict | None, str]:
    """(registro da manifestacao desta linha ou None, como foi achado).

    Primeiro pela chave de 44 (os dois lados inteiros); depois pela identidade
    (CNPJ do emitente, numero, serie) -- so para NF-e, que e a unica origem com
    identidade comparavel. O float da chave e a trava nos dois sentidos: se a
    PRODUTO 2 so tem o float e ele nao bate com a chave da SEFAZ, nao e a mesma
    nota; se os dois lados tem a chave inteira e sao diferentes, idem.
    """
    chave = premissa2._digitos(linha.get(sefaz.CHAVE_NFE))
    if len(chave) == 44:
        r = indices["por_chave"].get(chave)
        if r is not None:
            return r, "chave"
    if linha.get(sefaz.ORIGEM) != sefaz.NFE:
        return None, ""
    ident = premissa2.identidade(linha.get(sefaz.CNPJ_EMIT), linha.get(sefaz.NUM_NF),
                                 linha.get("_serie"))
    if not ident[0] or not ident[1]:
        return None, ""
    r = indices["por_identidade"].get(ident) if ident[2] else None
    como = "identidade"
    if r is None and not ident[2]:
        r = indices["por_cnpj_numero"].get(ident[:2])
        como = "cnpj+numero"
    if r is None:
        return None, ""
    if len(chave) == 44:
        if not premissa2.compativel(chave, r.get("_chave_float")):
            return None, ""
        chave_r = premissa2._digitos(r[CHAVE])
        if len(chave_r) == 44 and chave_r != chave:
            return None, ""
    return r, como


def _valor(registro: dict | None, emissao, de, ate, e_servico: bool) -> str:
    """O que a coluna mostra para UMA nota. Quatro respostas possíveis."""
    if registro is not None:
        return registro[STATUS]
    # ⚠ NFS-e nao tem manifestacao do destinatario -- nao e omissao do arquivo, e
    # o instituto que nao existe para nota de servico. Medido: 0 das 366 NFS-e
    # do painel estao na base, e e o esperado.
    if e_servico:
        return NAO_SE_APLICA
    # ⚠ ESTE E O CASO QUE FAZ A COLUNA NAO MENTIR. A SEFAZ.xlsx e ACUMULATIVA e
    # a manifestacao pode vir por MES. Em 01/09/2026 as duas cobriam agosto e a
    # diferenca quase nao aparecia -- 2 notas. Se um dia a PRODUTO 1 tiver
    # ago+set+out e a PRODUTO 2 so outubro, sem esta distincao o painel diria
    # "sem manifestacao" para centenas de notas de agosto que JA foram
    # manifestadas. O periodo sai da PROPRIA base, e nao de um mes escrito aqui.
    if emissao is None or de is None or not (de <= emissao <= ate):
        return FORA_DO_PERIODO
    # Dentro do periodo e mesmo assim ausente: isso e anomalia de verdade.
    return NAO_ENCONTRADA


def anotar(linhas: list[dict], base: Path | None = None) -> dict:
    """Preenche `sefaz.MANIFESTACAO` em cada linha e devolve o resumo.

    Roda DEPOIS do `sefaz.ler_com_historico()` e ANTES do `sefaz.carregar()`:
    as tres abas que leem a SEFAZ enxergam a coluna, e o `sefaz_historico.json`
    continua guardando a linha pura. Manifestacao e fato de HOJE, relido a cada
    rodada -- congelar no historico faria a nota removida carregar para sempre o
    status que ela tinha no dia em que sumiu.
    """
    indices, resumo = ler(base)
    de, ate = resumo["de"], resumo["ate"]
    contagem: dict[str, int] = {}
    como_achou: dict[str, int] = {}
    notas_vistas: set[str] = set()
    for linha in linhas:
        chave = linha[sefaz.CHAVE_NFE]
        registro, como = _registro_da_linha(linha, indices)
        valor = _valor(registro, linha.get(sefaz.EMISSAO), de, ate,
                       linha.get(sefaz.ORIGEM) != sefaz.NFE)
        linha[sefaz.MANIFESTACAO] = valor
        # O alerta e da NOTA e vai pela propria linha, como o `_removida_em`:
        # assim o `alertas_por_nota` o encontra sem que ninguem precise passar
        # mais um parametro pelos tres caminhos que chamam aquela funcao.
        if registro is not None and registro[STATUS] in GRAVES:
            justificativa = (registro.get(JUSTIFICATIVA) or "").strip()
            linha[sefaz.MANIFESTO_GRAVE] = (
                f"{registro[STATUS]}"
                + (f" — {justificativa}" if justificativa else "")
                + f" (filial {registro[FILIAL].strip()})")
        if chave not in notas_vistas:
            notas_vistas.add(chave)
            contagem[valor] = contagem.get(valor, 0) + 1
            if como:
                como_achou[como] = como_achou.get(como, 0) + 1
    resumo["anotadas"] = {k: contagem[k] for k in sorted(contagem)}
    resumo["como_achou"] = como_achou
    resumo["notas_no_painel"] = len(notas_vistas)
    resumo["graves"] = sum(n for s, n in contagem.items() if s in GRAVES)
    return resumo


def marcar_sem_base(linhas: list[dict]) -> None:
    """Sem a base, a coluna diz que NAO FOI LIDA -- e nao fica vazia.

    Celula vazia seria lida como "esta nota nao tem manifestacao", que e uma
    afirmacao que ninguem fez: o arquivo nem foi aberto.
    """
    for linha in linhas:
        linha[sefaz.MANIFESTACAO] = SEM_BASE
