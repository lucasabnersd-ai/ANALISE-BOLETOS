# -*- coding: utf-8 -*-
"""PREMISSA 2: as notas de `MESMA PREMISSA - ORGANIZADA.xlsx` que a SEFAZ.xlsx nao tem.

04/09/2026, pedido do usuario: "inclua essa base no cruzamento... use o nome do
emissor pra cruzar e verifique se tem na SF1. Chame essa base de PREMISSA 2 no
painel".

O QUE ESTE MODULO FAZ E O QUE NAO FAZ
-------------------------------------
Ao contrario do `manifestacao.py` -- que le a MESMA PREMISSA.xlsx so para
ENRIQUECER as notas da SEFAZ com a manifestacao --, este e uma FONTE DE LINHAS:
as notas que estao aqui e NAO estao na SEFAZ.xlsx entram no painel como notas,
com `Origem = "PREMISSA 2"`, e seguem o mesmo caminho das outras: cruzamento com
a SF1 (esta lancada?) e, para as que nao estao, a aba NF x Pedido de Compra com
os pedidos provaveis do fornecedor.

MEDIDO ANTES DE ESCREVER (04/09/2026, arquivo de 01/08 a 31/08):
  - 1.380 linhas, 1.324 chaves distintas (56 repetidas: a mesma nota vista por
    duas filiais, como no manifestacao.py), todas com 44 digitos limpos;
  - 858 sao notas de TERCEIROS (a Bracell e outros, em que a S&D so aparece como
    transportadora) -- o DF-e do TOTVS puxa tudo que cita o CNPJ do grupo em
    qualquer papel. Elas NAO entram: o corte e o mesmo da SEFAZ, CNPJ do
    destinatario na LISTAGEM EMPRESAS BIOFLOR;
  - das 522 do grupo, **467 ja estao na SEFAZ.xlsx** (pela chave de 44) e ficam
    de fora daqui -- ja estao no painel, com mais dado do que este arquivo tem;
  - sobram **55 notas novas**. Contra a SF1: 32 tem a chave de 44 num titulo, 1
    casa por NF + CNPJ, 4 so tem homonimo de numero e **16 nao estao** -- essas
    16 (mais as 4) sao o que ele quer ver.

⚠ POR QUE A CHAVE DE 44 E A IDENTIDADE, E NAO O NUMERO. A mesma nota pode estar
nas duas bases com o numero escrito diferente ("12765" x "000012765"), e o
numero da NFS-e comeca do 1 em cada prestador. A chave nao: `_digitos(chave)`
casa a mesma nota nos dois lados sem falso positivo, e foi assim que as 467
repetidas foram achadas.

⚠ O QUE ESTE ARQUIVO NAO TEM, E A LINHA NASCE SEM: itens, CFOP, duplicatas
(vencimento/valor da parcela), nome fantasia, tipo de operacao e o texto
`Informações` de onde a aba de pedidos le o numero do PC citado. Sem o texto, o
pedido dessas notas so pode ser achado por fornecedor + valor -- e e por isso
que a coluna `Numero PC Prováveis` importa mais ainda para elas. O `Status da
Nota` (Autorizada/Cancelada) tambem nao existe aqui e fica VAZIO -- inventar
"Autorizada" seria afirmar o que ninguem conferiu.

⚠ A MANIFESTACAO VEM DO PROPRIO ARQUIVO (coluna `Status Manifestação`), porque
ele E o arquivo de manifestacao. E gravada direto na linha; o `manifestacao.anotar`
roda ANTES desta injecao e por isso nao a sobrescreve.
"""
from __future__ import annotations

import hashlib
import re
import shutil
import tempfile
from pathlib import Path

from openpyxl import load_workbook

import sefaz

_AQUI = Path(__file__).resolve()
# Mesma pasta da SEFAZ.xlsx e da MESMA PREMISSA.xlsx: ANALISES BOLETOS.
BASE_PADRAO = _AQUI.parents[1] / "MESMA PREMISSA - ORGANIZADA.xlsx"

# O que a coluna `Origem` mostra para estas linhas. ⚠ E um TERCEIRO valor ao lado
# de sefaz.NFE / sefaz.NFSE. Quem separa as linhas por origem (`sefaz.carregar`
# conta nfe/nfs, `manifestacao` pergunta "e servico?") continua funcionando: uma
# linha PREMISSA 2 nao e nem uma coisa nem outra, e nao entra nessas contas.
ORIGEM = "PREMISSA 2"

ABA = "NOTAS"
# (nome no cabecalho da aba NOTAS, campo de destino na linha do painel)
COLUNAS = {
    "Filial":                 "_filial_nome",
    "Número NF":              sefaz.NUM_NF,
    "Série":                  "_serie",
    "Data Emissão":           sefaz.EMISSAO,
    "Chave de Acesso":        sefaz.CHAVE_NFE,
    "Emissor":                sefaz.NOME_EMIT,
    "CNPJ/CPF Emitente":      sefaz.CNPJ_EMIT,
    "Destinatário":           sefaz.NOME_DEST,
    "CNPJ/CPF Destinatário":  sefaz.CNPJ_DEST,
    "Valor NF (R$)":          sefaz.VLR_TOTAL,
    "Status Manifestação":    sefaz.MANIFESTACAO,
    "Natureza da Operação":   sefaz.NAT_OP,
}
OBRIGATORIAS = ("Chave de Acesso", "Emissor", "CNPJ/CPF Emitente",
                "CNPJ/CPF Destinatário", "Número NF", "Data Emissão", "Valor NF (R$)")


def _digitos(valor) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def _abrir(caminho: Path):
    """Abre a planilha; se o Excel/OneDrive estiver segurando, usa uma copia.

    O mesmo remedio de sefaz.py / manifestacao.py: este arquivo vive aberto na
    tela de quem acabou de organiza-lo.
    """
    try:
        return load_workbook(caminho, read_only=True, data_only=True), None
    except (PermissionError, OSError):
        pasta = tempfile.mkdtemp(prefix="premissa2_")
        shutil.copy2(caminho, Path(pasta) / caminho.name)
        return load_workbook(Path(pasta) / caminho.name, read_only=True,
                             data_only=True), pasta


def ler(base: Path | None = None) -> list[dict]:
    """Todas as linhas da aba NOTAS, ja no vocabulario do painel. Sem filtro."""
    base = base or BASE_PADRAO
    if not base.exists():
        raise RuntimeError(
            f"A base PREMISSA 2 nao foi encontrada:\n  {base}\n"
            "Confira se o OneDrive sincronizou, ou corrija BASE_PADRAO no premissa2.py."
        )
    wb, temporaria = _abrir(base)
    try:
        if ABA not in wb.sheetnames:
            raise RuntimeError(
                f"{base.name} nao tem a aba {ABA!r} (tem: {wb.sheetnames}). "
                "Corrija ABA no premissa2.py.")
        brutas = wb[ABA].iter_rows(values_only=True)
        cabecalho = [str(c).strip() if c is not None else "" for c in next(brutas, ())]
        faltando = [c for c in OBRIGATORIAS if c not in cabecalho]
        if faltando:
            raise RuntimeError(
                f"{base.name} mudou de layout: faltam as colunas "
                + " / ".join(repr(c) for c in faltando) + "\n  o arquivo tem: "
                + " / ".join(repr(c) for c in cabecalho) + "\nCorrija COLUNAS no premissa2.py.")
        posicoes = [(cabecalho.index(nome), destino)
                    for nome, destino in COLUNAS.items() if nome in cabecalho]
        linhas = []
        for bruta in brutas:
            if all(v in (None, "") for v in bruta):
                continue
            linha = {sefaz.ORIGEM: ORIGEM}
            for i, destino in posicoes:
                valor = bruta[i] if i < len(bruta) else None
                if destino == sefaz.EMISSAO:
                    linha[destino] = sefaz._data(valor)
                elif destino == sefaz.VLR_TOTAL:
                    linha[destino] = sefaz._numero(valor)
                else:
                    linha[destino] = sefaz._texto(valor)
            chave = _digitos(linha.get(sefaz.CHAVE_NFE))
            if len(chave) != 44:
                # linha de rodape, celula vazia, ou chave que virou notacao
                # cientifica ao passar pelo Excel -- nada disso e nota
                continue
            linha[sefaz.CHAVE_NFE] = chave
            # ⚠ O CNPJ do destinatario vem "45108676000131" e a filial vem como
            # "003001 - BRASMAC..." -- o codigo da filial e util para conferir,
            # e fica num campo interno (comeca com "_", nao vai para a tela).
            linha["_filial_totvs"] = (linha.get("_filial_nome") or "")[:6]
            # o que o arquivo NAO tem, e a linha nasce declarando que nao tem
            for vazio in (sefaz.STATUS_NOTA, sefaz.TIPO_OP, sefaz.FANTASIA,
                          sefaz.CFOP, sefaz.INFO):
                linha.setdefault(vazio, "")
            for nenhum in (sefaz.SAIDA, sefaz.VENC_DUP, sefaz.VLR_DUP, sefaz.VLR_ITEM):
                linha.setdefault(nenhum, None)
            # ⚠ Uma linha por NOTA: e o `Chave` que segura check e tratativa, e
            # `sefaz.PREFIXO` e o mesmo das outras para o painel tratar igual.
            linha[sefaz.CHAVE] = sefaz.PREFIXO + hashlib.sha1(
                f"{ORIGEM}|{chave}".encode("utf-8")).hexdigest()[:16]
            linhas.append(linha)
        return linhas
    finally:
        wb.close()
        if temporaria:
            shutil.rmtree(temporaria, ignore_errors=True)


def linhas_novas(linhas_sefaz: list[dict], permitidos: dict,
                 base: Path | None = None) -> tuple[list[dict], dict]:
    """As notas PREMISSA 2 que ENTRAM no painel: do grupo e ausentes da SEFAZ.

    `permitidos` e o {CNPJ: razao} da LISTAGEM EMPRESAS BIOFLOR (o mesmo corte
    da SEFAZ). `linhas_sefaz` sao as linhas que o painel ja tem -- e delas que
    saem as chaves a NAO repetir.

    ⚠ A CHAVE REPETIDA DENTRO DO PROPRIO ARQUIVO (a mesma nota vista por duas
    filiais) vira UMA linha: vale a primeira, e a manifestacao segue a regra do
    manifestacao.py -- qualquer status vence `SemManifestacao`.
    """
    todas = ler(base)
    ja_no_painel = {_digitos(l.get(sefaz.CHAVE_NFE)) for l in linhas_sefaz}
    do_grupo = [l for l in todas if _digitos(l.get(sefaz.CNPJ_DEST)) in permitidos]

    por_chave: dict[str, dict] = {}
    for l in do_grupo:
        atual = por_chave.get(l[sefaz.CHAVE_NFE])
        if atual is None:
            por_chave[l[sefaz.CHAVE_NFE]] = l
        elif (atual.get(sefaz.MANIFESTACAO) or "") == "SemManifestacao" \
                and (l.get(sefaz.MANIFESTACAO) or "") not in ("", "SemManifestacao"):
            por_chave[l[sefaz.CHAVE_NFE]] = l

    novas = [l for ch, l in por_chave.items() if ch not in ja_no_painel]
    resumo = {
        "arquivo": str(base or BASE_PADRAO),
        "linhas": len(todas),
        "do_grupo": len(por_chave),
        "terceiros": len(todas) - len(do_grupo),
        "ja_na_sefaz": len(por_chave) - len(novas),
        "novas": len(novas),
    }
    return novas, resumo
