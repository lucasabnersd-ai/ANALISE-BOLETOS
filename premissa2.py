# -*- coding: utf-8 -*-
"""PRODUTO 2 (ex-PREMISSA 2): as notas da manifestacao que a aba de NF-e nao tem.

04/09/2026, pedido do usuario: "inclua essa base no cruzamento... use o nome do
emissor pra cruzar e verifique se tem na SF1. Chame essa base de PREMISSA 2 no
painel".

O QUE ESTE MODULO FAZ E O QUE NAO FAZ
-------------------------------------
Ao contrario do `manifestacao.py` -- que le a mesma base so para ENRIQUECER as
notas da SEFAZ com a manifestacao --, este e uma FONTE DE LINHAS: as notas que
estao aqui e NAO estao na aba de NF-e entram no painel como notas, com
`Origem = "PRODUTO 2"`, e seguem o mesmo caminho das outras: cruzamento com a
SF1 (esta lancada?) e, para as que nao estao, a aba NF x Pedido de Compra com
os pedidos provaveis do fornecedor.

18/09/2026 -- A FONTE MUDOU DE ARQUIVO E PERDEU A CHAVE
-------------------------------------------------------
Pedido do usuario: *"AS ABAS PRODUTO 1 E PRODUTO 2 SE COMPLEMENTAM, AS VEZES UMA
NF QUE ESTA NO PRODUTO 1 NAO ESTA NA PRODUTO 2, MAS AS VEZES UMA NF ESTA EM UMA
NAO ESTA EM OUTRA, VOCE SEGUIRA CRUZANDO COM O PEDIDO DE COMPRA ELAS, MAS
CONSOLIDE ELAS PARA ANALISE"*. A base de manifestacao passou a morar DENTRO da
SEFAZ.xlsx, como aba `PRODUTO 2` (o mesmo layout de 14 colunas do CSV do DF-e),
ao lado da `PRODUTO 1` (a exportacao de NF-e por item x duplicata). Este modulo
le essa aba -- achada pelo CABECALHO, como as outras --, e os arquivos
`MESMA PREMISSA*.xlsx` da pasta ficaram de reserva: sao o fallback quando a aba
nao existe e o DICIONARIO DE CHAVES descrito abaixo.

⚠ NA ABA PRODUTO 2 A CHAVE VEIO COMO NUMERO. As 2.532 linhas de 18/09/2026 tem
`Chave` gravada como float ('4.1260927755427e+43'): o Excel guarda 15 digitos
significativos e os outros 29 sumiram -- e nao ha como recupera-los do proprio
arquivo (o cNF de 8 digitos da chave e aleatorio). Consequencias, e o que se faz:
  - a IDENTIDADE da nota passa a ser (CNPJ do emitente, numero, serie).
    Medido em 18/09/2026: zero colisoes na PRODUTO 1 (825 chaves) e zero na
    PRODUTO 2 (2.442 notas). E por ela que se decide se a nota "ja esta na
    PRODUTO 1" -- 591 estavam, 1.851 nao (a maior parte de terceiros);
  - o float NAO e jogado fora: ele conserva cUF + AAMM + os primeiros digitos
    do CNPJ e serve de TRAVA -- identidade igual com float incompativel com a
    chave de 44 da PRODUTO 1 e nota diferente (`compativel()`). Medido: 200 de
    200 pares compativeis;
  - a chave de 44 e RECUPERADA quando algum `MESMA PREMISSA*.xlsx` da pasta a
    tem para a mesma identidade (`chaves_conhecidas()`), conferida pelo float.
    E o que mantem o uuid -- e a marcacao ja feita -- das notas que entraram
    pelo CONSOLIDADA desde 04/09;
  - sem chave em lugar nenhum, a `Chave NF-e` sai como um identificador
    LEGIVEL e ESTAVEL ("PRODUTO 2 <cnpj> <numero> <serie>"), mesmo desenho do
    nfs_status.py. O cruzamento com a SF1 segue por fornecedor + numero (a
    chave so pontua quando tem 44 digitos) e a aba de pedidos por fornecedor +
    valor.
⚠ O CNPJ TAMBEM VIROU NUMERO e perdeu o zero a esquerda: 533 emitentes com 13
digitos e 41 com 12; 141 destinatarios com 13 (a 09515262000163). Sem o
`_cnpj()` que recompoe os 14, essas 141 notas do grupo cairiam como "de
terceiros" em silencio. O corte do grupo continua sendo pelo destinatario.
⚠ `Origem` mostra "PRODUTO 2" (o nome que a aba tem agora), mas o uuid continua
semeado com "PREMISSA 2" (`_SEMENTE_UUID`): trocar a semente trocaria o uuid de
toda linha e apagaria as marcacoes feitas desde 04/09.

⚠ QUAL ARQUIVO, E EM QUE FORMA -- as duas coisas mudam sem aviso.
No mesmo 04/09/2026 em que este modulo nasceu apontando para o
`MESMA PREMISSA - ORGANIZADA.xlsx` (aba `NOTAS`, 1.380 linhas), ele foi
reexportado e VOLTOU A SER o CSV cru de uma coluna so (aba com nome GUID, 584
notas), enquanto o organizado com 1.466 notas passou a ser o
`MESMA PREMISSA - CONSOLIDADA.xlsx`, que nao existia de manha. A licao virou
codigo em tres lugares:
  - `escolher_base()`: a SEFAZ.xlsx vem primeiro, se tiver a aba; senao os
    `CANDIDATOS`, do mais completo para o menos;
  - `_tabela()`: a aba e escolhida pelo CABECALHO (a que tem mais colunas
    conhecidas), nunca pelo nome -- mesma regra do `sefaz._achar_aba()`;
  - `SINONIMOS`: os dois vocabularios das mesmas 14 colunas convivem
    (`Chave de Acesso` x `Chave`, `Número NF` x `Numero`...).

MEDIDO NA PRIMEIRA VERSAO (arquivo de 01/08 a 31/08, 1.380 linhas):
  - 56 chaves repetidas (a mesma nota vista por duas filiais, como no
    manifestacao.py), todas com 44 digitos limpos;
  - 858 sao notas de TERCEIROS (a Bracell e outros, em que a S&D so aparece como
    transportadora) -- o DF-e do TOTVS puxa tudo que cita o CNPJ do grupo em
    qualquer papel. Elas NAO entram: o corte e o mesmo da SEFAZ, CNPJ do
    destinatario na LISTAGEM EMPRESAS BIOFLOR;
  - das 466 do grupo, 413 ja estavam na SEFAZ.xlsx e ficaram de fora -- ja
    estao no painel, com mais dado do que este arquivo tem;
  - sobraram 53 notas novas, 20 delas SEM lancamento na SF1.

⚠ O QUE ESTA BASE NAO TEM, E A LINHA NASCE SEM: itens, CFOP, duplicatas
(vencimento/valor da parcela), nome fantasia, tipo de operacao e o texto
`Informações` de onde a aba de pedidos le o numero do PC citado. Sem o texto, o
pedido dessas notas so pode ser achado por fornecedor + valor -- e e por isso
que a coluna `Numero PC Prováveis` importa mais ainda para elas. O `Status da
Nota` (Autorizada/Cancelada) tambem nao existe aqui e fica VAZIO -- inventar
"Autorizada" seria afirmar o que ninguem conferiu.

⚠ A MANIFESTACAO VEM DA PROPRIA BASE (coluna `StatusManifestacao`), porque ela E
a base de manifestacao. E gravada direto na linha; o `manifestacao.anotar` roda
ANTES desta injecao e por isso nao a sobrescreve. Status grave (desconhecimento
/ nao realizada) tambem poe o `sefaz.MANIFESTO_GRAVE` na linha, para o alerta
"NOTA NAO RECONHECIDA" sair igual ao das notas da PRODUTO 1.
"""
from __future__ import annotations

import csv
import datetime as dt
import hashlib
import io
import re
import shutil
import tempfile
from pathlib import Path

from openpyxl import load_workbook

import caminhos
import sefaz

_AQUI = Path(__file__).resolve()
# 10/09/2026: era contagem de niveis a partir deste arquivo (parents[N]).
# Contar amarra o script a UM lugar da arvore -- ao mudar a pasta do painel
# a conta aponta para o pai errado e a aba nasce vazia com um AVISO que
# ninguem le. Agora a pasta e' PROCURADA por marca. Ver caminhos.py.
_ANALISES = caminhos.pasta_analises()

# Os arquivos de reserva, "o mais completo primeiro". Desde 18/09/2026 a fonte
# preferida e a aba PRODUTO 2 da propria SEFAZ.xlsx (ver `escolher_base`); estes
# so valem quando ela nao existe -- e como dicionario de chaves de 44 digitos.
CANDIDATOS = ("MESMA PREMISSA - CONSOLIDADA.xlsx", "MESMA PREMISSA - ORGANIZADA.xlsx",
              "MESMA PREMISSA.xlsx")


def _primeira_que_existe() -> Path:
    for nome in CANDIDATOS:
        caminho = _ANALISES / nome
        if caminho.is_file():
            return caminho
    return _ANALISES / CANDIDATOS[0]


BASE_PADRAO = _primeira_que_existe()

# O que a coluna `Origem` mostra para estas linhas. ⚠ E um valor a mais ao lado
# de sefaz.NFE / sefaz.NFSE. Quem separa as linhas por origem (`sefaz.carregar`
# conta nfe/nfs, `manifestacao` pergunta "e NF-e?") continua funcionando: uma
# linha PRODUTO 2 nao e nem uma coisa nem outra, e nao entra nessas contas.
ORIGEM = "PRODUTO 2"
# ⚠ A semente do uuid NAO acompanhou o nome (ver o topo do arquivo).
_SEMENTE_UUID = "PREMISSA 2"

# ⚠ DOIS VOCABULARIOS PARA AS MESMAS 14 COLUNAS. O export cru chama
# `NomeFilial/Numero/Chave/CnpjCpfEmitente`; a versao organizada chama
# `Filial/Número NF/Chave de Acesso/CNPJ/CPF Emitente`. Aqui os dois apontam
# para o mesmo destino -- e por isso a leitura nao se importa com qual veio.
SINONIMOS = {
    "_filial_nome":       ("Filial", "NomeFilial"),
    sefaz.NUM_NF:         ("Número NF", "Numero"),
    "_serie":             ("Série", "Serie"),
    sefaz.EMISSAO:        ("Data Emissão", "DataEmissao"),
    sefaz.CHAVE_NFE:      ("Chave de Acesso", "Chave"),
    sefaz.NOME_EMIT:      ("Emissor", "Emissor"),
    sefaz.CNPJ_EMIT:      ("CNPJ/CPF Emitente", "CnpjCpfEmitente"),
    sefaz.NOME_DEST:      ("Destinatário", "Destinatario"),
    sefaz.CNPJ_DEST:      ("CNPJ/CPF Destinatário", "CnpjCpfDestinatario"),
    sefaz.VLR_TOTAL:      ("Valor NF (R$)", "ValorNF"),
    sefaz.MANIFESTACAO:   ("Status Manifestação", "StatusManifestacao"),
    sefaz.NAT_OP:         ("Natureza da Operação", "NaturezaOperacao"),
    "_justificativa":     ("Justificativa Manifestação", "JustificativaManifestacao"),
}
# Sem estas nao ha nota; o resto pode faltar que a linha nasce vazia.
OBRIGATORIOS = (sefaz.CHAVE_NFE, sefaz.NOME_EMIT, sefaz.CNPJ_EMIT,
                sefaz.CNPJ_DEST, sefaz.NUM_NF, sefaz.EMISSAO, sefaz.VLR_TOTAL)
# Quantas colunas conhecidas uma aba precisa ter para ser "a aba": a PRODUTO 1
# tem UMA em comum ("Data Emissão") e a SERVICO 2 duas ("Filial", "Série") --
# nenhuma pode passar por PRODUTO 2 so por isso.
MIN_CONHECIDAS = len(OBRIGATORIOS)

# Os dois status que o painel trata como graves -- mesma lista do manifestacao.py
# (repetida aqui para nao importar aquele modulo, que importa este).
GRAVES = ("DesconhecimentoOperacao", "NaoRealizada")
SEM_MANIFESTACAO = "SemManifestacao"

# De onde saiu a chave de cada linha (campo interno `_chave_origem`).
CHAVE_DA_ABA = "aba"            # veio com 44 digitos
CHAVE_RECUPERADA = "recuperada"  # veio como float; a de 44 saiu de um MESMA PREMISSA*.xlsx
CHAVE_SINTETICA = "sintetica"    # veio como float e ninguem tem a de 44


def _digitos(valor) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def _cnpj(valor) -> str:
    """CNPJ/CPF so digitos, com o zero a esquerda que o Excel comeu de volta.

    A aba grava o CNPJ como numero: 09515262000163 chega como 9515262000163
    (13 digitos) e um 00xxx como 12. Nada com 12 ou 13 digitos e CPF nem CNPJ
    valido, entao completar para 14 nao inventa nada; abaixo de 11, completa
    para CPF. 11 e 14 ficam como estao.
    """
    d = _digitos(valor)
    if 11 < len(d) < 14:
        return d.zfill(14)
    if 0 < len(d) < 11:
        return d.zfill(11)
    return d


def _numero_nf(valor) -> str:
    return _digitos(valor).lstrip("0")


def identidade(cnpj_emitente, numero, serie) -> tuple[str, str, str]:
    """(CNPJ do emitente com 14 digitos, numero sem zeros, serie sem zeros).

    E o que identifica a nota quando a chave de 44 nao esta disponivel -- e
    tambem o que casa uma linha da PRODUTO 2 com a mesma nota na PRODUTO 1.
    """
    return (_cnpj(cnpj_emitente), _numero_nf(numero), _numero_nf(serie))


def chave_float(valor):
    """O float que o Excel fez da chave, ou None se o campo nao e isso.

    Uma chave de 44 digitos vale entre 1e43 e 1e44; qualquer coisa abaixo de
    1e42 nao e chave virada em numero (e um numero de NF, uma celula vazia...).
    """
    texto = str(valor or "").strip()
    if not texto or len(_digitos(texto)) == 44:
        return None
    try:
        f = float(texto)
    except ValueError:
        return None
    return f if f >= 1e42 else None


def compativel(chave44: str, f: float | None) -> bool:
    """A chave de 44 digitos e o float sao a MESMA chave?

    O float guarda ~15 digitos significativos; comparar com folga relativa de
    1e-13 exige que os ~13 primeiros batam (UF, ano/mes, boa parte do CNPJ).
    Sem float nao ha o que contradizer: True.
    """
    if f is None:
        return True
    try:
        return abs(float(chave44) - f) <= abs(f) * 1e-13
    except (TypeError, ValueError):
        return False


def _data(valor):
    """A emissao, venha ela como for.

    Sao TRES formatos no mesmo campo: `datetime` de verdade (a aba NOTAS traz a
    coluna formatada), "2026-08-04 11:06:38" (a mesma aba quando o Excel a
    guardou como texto) e "04-08-2026T11:03:48AM-03:00" (o CSV cru e a aba
    PRODUTO 2). ⚠ O `sefaz._data` sozinho nao cobre os dois ultimos: ele tenta
    a string INTEIRA, e a hora colada faz todo strptime falhar -- a emissao
    viraria None e a nota sairia sem data, sem erro nenhum aparecer.
    """
    if isinstance(valor, (dt.datetime, dt.date)):
        return sefaz._data(valor)
    bruto = str(valor or "").strip()
    if not bruto:
        return None
    achado = sefaz._data(bruto)
    if achado:
        return achado
    cabeca = bruto[:10]                      # so a data; o resto e hora/fuso
    for formato in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(cabeca, formato).date()
        except ValueError:
            pass
    return None


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


def _tabela(wb) -> tuple[list[list[str]], str]:
    """(linhas da aba certa, nome dela), venha ela como planilha ou como CSV.

    ⚠ A ABA NAO E ESCOLHIDA PELO NOME -- e a que tem MAIS COLUNAS conhecidas no
    cabecalho, e pelo menos MIN_CONHECIDAS delas. Foi o nome que quebrou a
    rodada de 04/09 ('NOTAS' virou um GUID), e a mesma licao ja esta no
    `sefaz._achar_aba`: o nome nao e identidade. Dentro da SEFAZ.xlsx e assim
    que a PRODUTO 2 e achada no meio de PRODUTO 1 / SERVICO / SERVICO 2 / CTE.

    ⚠ O CSV E LIDO COM `csv.reader`, NUNCA COM `split(',')` -- mesma razao do
    manifestacao.py: ha virgula DENTRO de campo entre aspas, e o split parte a
    linha em pedacos a mais, jogando todos os campos seguintes uma casa para o
    lado. A chave passaria a ser o nome do emissor e o cruzamento erraria calado.
    """
    conhecidos = {n for nomes in SINONIMOS.values() for n in nomes}
    melhor, melhor_nome, melhor_nota = [], "", -1
    for ws in wb.worksheets:
        cabecalho = sefaz._cabecalho_de(ws)
        if len(cabecalho) <= 1:
            # CSV embrulhado: uma coluna so, o cabecalho esta DENTRO do texto
            cruas = [c[0] for c in ws.iter_rows(values_only=True) if c]
            texto = "\n".join(str(c) for c in cruas
                              if c is not None and str(c).strip())
            tabela = [l for l in csv.reader(io.StringIO(texto)) if l]
            if not tabela:
                continue
            nota = sum(1 for c in tabela[0] if str(c).strip() in conhecidos)
        else:
            nota = sum(1 for c in cabecalho if str(c or "").strip() in conhecidos)
            if nota <= melhor_nota or nota < MIN_CONHECIDAS:
                continue          # nao vale ler 3.600 linhas x 386 colunas a toa
            tabela = [["" if v is None else str(v).strip() for v in bruta]
                      for bruta in ws.iter_rows(values_only=True)
                      if not all(v in (None, "") for v in bruta)]
        if nota > melhor_nota and nota >= MIN_CONHECIDAS:
            melhor, melhor_nome, melhor_nota = tabela, ws.title, nota
    return melhor, melhor_nome


def tem_aba(base: Path | None) -> bool:
    """A planilha tem uma aba com o layout da PRODUTO 2 / MESMA PREMISSA?"""
    if not base or not Path(base).is_file():
        return False
    try:
        wb, temporaria = _abrir(Path(base))
    except Exception:  # noqa: BLE001 - arquivo corrompido/aberto: nao tem
        return False
    try:
        tabela, _ = _tabela(wb)
    finally:
        wb.close()
        if temporaria:
            shutil.rmtree(temporaria, ignore_errors=True)
    return bool(tabela)


def escolher_base(explicita: Path | None, base_sefaz: Path | None) -> Path:
    """De onde ler: o que o usuario mandou > a aba PRODUTO 2 da SEFAZ.xlsx > reserva.

    `explicita` igual ao BASE_PADRAO nao conta como mandado: e o default do
    argparse do gerar_painel.py, e ele nao sabe se a SEFAZ.xlsx tem a aba.
    """
    if explicita and Path(explicita) != BASE_PADRAO:
        return Path(explicita)
    if base_sefaz and tem_aba(base_sefaz):
        return Path(base_sefaz)
    return BASE_PADRAO


def _ler_bruto(base: Path) -> tuple[list[dict], dict]:
    """Todas as notas da base, no vocabulario do painel, SEM resolver a chave.

    Cada linha sai com `_identidade`, `_chave_float` e a `Chave NF-e` so quando
    ela veio inteira (44 digitos) -- vazia caso contrario. Quem resolve e o
    `ler()`; separar os dois e o que permite ao dicionario de chaves ler os
    MESMA PREMISSA*.xlsx sem cair em recursao.
    """
    base = Path(base)
    if not base.exists():
        raise RuntimeError(
            f"A base PRODUTO 2 nao foi encontrada:\n  {base}\n"
            f"Procurei a aba PRODUTO 2 na SEFAZ.xlsx e, na pasta, "
            f"{' / '.join(CANDIDATOS)}. Confira se o OneDrive sincronizou, ou "
            "aponte com --base-premissa2.")
    wb, temporaria = _abrir(base)
    try:
        tabela, aba = _tabela(wb)
        abas = list(wb.sheetnames)
    finally:
        wb.close()
        if temporaria:
            shutil.rmtree(temporaria, ignore_errors=True)
    if not tabela:
        raise RuntimeError(
            f"{base.name} nao tem nenhuma aba com o layout da PRODUTO 2 / MESMA "
            f"PREMISSA (procurei {MIN_CONHECIDAS}+ das colunas "
            + " / ".join(repr(n[1]) for n in SINONIMOS.values())
            + f").\n  a planilha tem: " + " / ".join(repr(a) for a in abas))

    cabecalho = [str(c).strip() if c is not None else "" for c in tabela[0]]
    # {indice da coluna: destino} -- o primeiro sinonimo que aparecer vence
    posicoes = {}
    for destino, nomes in SINONIMOS.items():
        for nome in nomes:
            if nome in cabecalho:
                posicoes[cabecalho.index(nome)] = destino
                break
    faltando = [d for d in OBRIGATORIOS if d not in posicoes.values()]
    if faltando:
        raise RuntimeError(
            f"{base.name} (aba {aba!r}) mudou de layout: nao achei coluna para "
            + " / ".join(repr(f) for f in faltando) + "\n  o arquivo tem: "
            + " / ".join(repr(c) for c in cabecalho if c)
            + "\nAcrescente o nome novo em SINONIMOS no premissa2.py.")

    linhas = []
    sem_identidade = 0
    for bruta in tabela[1:]:
        linha = {sefaz.ORIGEM: ORIGEM}
        for i, destino in posicoes.items():
            valor = bruta[i] if i < len(bruta) else None
            if destino == sefaz.EMISSAO:
                linha[destino] = _data(valor)
            elif destino == sefaz.VLR_TOTAL:
                linha[destino] = sefaz._numero(valor)
            elif destino in (sefaz.CNPJ_EMIT, sefaz.CNPJ_DEST):
                linha[destino] = _cnpj(valor)
            else:
                linha[destino] = sefaz._texto(valor)
        ident = identidade(linha.get(sefaz.CNPJ_EMIT), linha.get(sefaz.NUM_NF),
                           linha.get("_serie"))
        if not ident[0] or not ident[1]:
            # rodape, celula vazia, linha de titulo -- nada disso e nota
            sem_identidade += 1
            continue
        bruto_chave = linha.get(sefaz.CHAVE_NFE)
        chave = _digitos(bruto_chave)
        linha["_identidade"] = ident
        linha["_chave_float"] = chave_float(bruto_chave)
        linha[sefaz.CHAVE_NFE] = chave if len(chave) == 44 else ""
        # ⚠ A filial vem como "003001 - BRASMAC..." -- o codigo e util para
        # conferir e fica num campo interno ("_", nao vai para a tela).
        linha["_filial_totvs"] = (linha.get("_filial_nome") or "").strip()[:6]
        # o que a base NAO tem, e a linha nasce declarando que nao tem
        for vazio in (sefaz.STATUS_NOTA, sefaz.TIPO_OP, sefaz.FANTASIA,
                      sefaz.CFOP, sefaz.INFO):
            linha.setdefault(vazio, "")
        for nenhum in (sefaz.SAIDA, sefaz.VENC_DUP, sefaz.VLR_DUP, sefaz.VLR_ITEM):
            linha.setdefault(nenhum, None)
        linhas.append(linha)
    resumo = {"arquivo": str(base), "nome": base.name, "aba": aba,
              "linhas": len(tabela) - 1, "sem_identidade": sem_identidade,
              "chaves_44": sum(1 for l in linhas if l[sefaz.CHAVE_NFE]),
              "chaves_float": sum(1 for l in linhas if l["_chave_float"] is not None)}
    return linhas, resumo


# {pasta: {identidade: chave de 44}} -- lido uma vez por rodada
_CONHECIDAS: dict[str, dict] = {}


def chaves_conhecidas(pasta: Path | None = None) -> dict[tuple, str]:
    """{identidade: chave de 44 digitos} de todo MESMA PREMISSA*.xlsx da pasta.

    E o que devolve a chave verdadeira as notas que a aba PRODUTO 2 trouxe como
    float: os exports anteriores (CSV cru) guardaram a chave inteira. Arquivo
    que nao abre ou nao tem o layout e pulado -- este dicionario e um extra, e
    um extra nao pode derrubar a leitura principal.
    """
    pasta = Path(pasta or _ANALISES)
    marca = str(pasta)
    if marca in _CONHECIDAS:
        return _CONHECIDAS[marca]
    mapa: dict[tuple, str] = {}
    for arquivo in sorted(pasta.glob("MESMA PREMISSA*.xlsx")):
        try:
            linhas, _ = _ler_bruto(arquivo)
        except Exception:  # noqa: BLE001 - reserva: pula e segue
            continue
        for l in linhas:
            if l[sefaz.CHAVE_NFE]:
                mapa.setdefault(l["_identidade"], l[sefaz.CHAVE_NFE])
    _CONHECIDAS[marca] = mapa
    return mapa


def chave_sintetica(ident: tuple[str, str, str]) -> str:
    """Identificador legivel e ESTAVEL para nota sem chave de 44 digitos.

    Numero e serie com largura fixa (9 e 3, as larguras da propria chave NF-e)
    para os digitos que o `cruzamento_sefaz_sf1`/`pedidos_sefaz` tiram daqui
    nunca colidirem entre "NF 6981911 serie 0" e "NF 698191 serie 1".
    """
    cnpj, numero, serie = ident
    return f"{ORIGEM} {cnpj} {numero.zfill(9)} {(serie or '0').zfill(3)}"


def ler(base: Path | None = None) -> tuple[list[dict], dict]:
    """Todas as notas da base, ja com a chave resolvida. Sem filtro de grupo.

    Devolve (linhas, resumo). Cada linha tem `_chave_origem` dizendo de onde a
    chave saiu (CHAVE_DA_ABA / CHAVE_RECUPERADA / CHAVE_SINTETICA) e o uuid
    (`sefaz.CHAVE`) montado -- por NOTA, porque a base e por nota.
    """
    base = Path(base or BASE_PADRAO)
    linhas, resumo = _ler_bruto(base)
    conhecidas = chaves_conhecidas() if resumo["chaves_float"] else {}
    recuperadas = sinteticas = incompativeis = 0
    for linha in linhas:
        if linha[sefaz.CHAVE_NFE]:
            linha["_chave_origem"] = CHAVE_DA_ABA
            continue
        ident = linha["_identidade"]
        lembrada = conhecidas.get(ident)
        if lembrada and compativel(lembrada, linha["_chave_float"]):
            linha[sefaz.CHAVE_NFE] = lembrada
            linha["_chave_origem"] = CHAVE_RECUPERADA
            recuperadas += 1
            continue
        if lembrada:
            # mesma identidade, chave lembrada NAO bate com o float: e outra
            # nota (ou o arquivo de reserva esta errado) -- nao afirma nada
            incompativeis += 1
        linha[sefaz.CHAVE_NFE] = chave_sintetica(ident)
        linha["_chave_origem"] = CHAVE_SINTETICA
        sinteticas += 1
    for linha in linhas:
        # ⚠ Uma linha por NOTA: e o `Chave` que segura check e tratativa, e
        # `sefaz.PREFIXO` e o mesmo das outras para o painel tratar igual.
        # Semeado com a chave de 44 quando ha (uuid identico ao de antes) e com
        # a sintetica quando nao ha.
        linha[sefaz.CHAVE] = sefaz.PREFIXO + hashlib.sha1(
            f"{_SEMENTE_UUID}|{linha[sefaz.CHAVE_NFE]}".encode("utf-8")).hexdigest()[:16]
    emissoes = [l[sefaz.EMISSAO] for l in linhas if l.get(sefaz.EMISSAO)]
    resumo.update({
        "recuperadas": recuperadas, "sinteticas": sinteticas,
        "incompativeis": incompativeis,
        "de": min(emissoes) if emissoes else None,
        "ate": max(emissoes) if emissoes else None,
    })
    return linhas, resumo


def _indice_nfe(linhas_sefaz: list[dict]) -> tuple[set, dict, dict]:
    """O que o painel JA tem, visto de tres jeitos.

    (chaves so digitos de TODA linha; {identidade: chave} das NF-e;
    {(cnpj, numero): chave ou None se ambiguo} das NF-e). Os dois ultimos so
    das NF-e (`sefaz.NFE`): NFS-e e CT-e nao tem como ser a mesma nota que uma
    linha da PRODUTO 2. O terceiro existe para as linhas antigas do historico,
    que nasceram antes da `_serie` entrar na leitura da PRODUTO 1.
    """
    ja = {_digitos(l.get(sefaz.CHAVE_NFE)) for l in linhas_sefaz}
    por_identidade: dict[tuple, str] = {}
    por_cnpj_numero: dict[tuple, str | None] = {}
    for l in linhas_sefaz:
        if l.get(sefaz.ORIGEM) != sefaz.NFE:
            continue
        chave = _digitos(l.get(sefaz.CHAVE_NFE))
        ident = identidade(l.get(sefaz.CNPJ_EMIT), l.get(sefaz.NUM_NF), l.get("_serie"))
        if not ident[0] or not ident[1]:
            continue
        if ident[2]:
            por_identidade.setdefault(ident, chave)
        par = ident[:2]
        atual = por_cnpj_numero.get(par, chave)
        por_cnpj_numero[par] = chave if atual == chave else None
    return ja, por_identidade, por_cnpj_numero


def chave_no_painel(linha: dict, por_identidade: dict, por_cnpj_numero: dict) -> str | None:
    """A chave de 44 da MESMA nota na PRODUTO 1, se ela estiver la -- ou None.

    Casa pela identidade (CNPJ, numero, serie); sem serie do lado da PRODUTO 1
    (linha antiga do historico), pelo par (CNPJ, numero) quando ele e unico la.
    O float da PRODUTO 2 e a trava: identidade igual com float incompativel e
    nota diferente, e devolve None.
    """
    ident = linha["_identidade"]
    chave = por_identidade.get(ident)
    if chave is None and ident[2]:
        chave = por_cnpj_numero.get(ident[:2])
    if not chave:
        return None
    if len(chave) == 44 and not compativel(chave, linha.get("_chave_float")):
        return None
    if linha["_chave_origem"] == CHAVE_DA_ABA and len(chave) == 44 \
            and chave != linha[sefaz.CHAVE_NFE]:
        return None      # as duas tem chave inteira e sao diferentes
    return chave


def linhas_novas(linhas_sefaz: list[dict], permitidos: dict,
                 base: Path | None = None) -> tuple[list[dict], dict]:
    """As notas da PRODUTO 2 que ENTRAM no painel: do grupo e ausentes da PRODUTO 1.

    `permitidos` e o {CNPJ: razao} da LISTAGEM EMPRESAS BIOFLOR (o mesmo corte
    da SEFAZ). `linhas_sefaz` sao as linhas que o painel ja tem -- e delas que
    saem as notas a NAO repetir, pela chave de 44 quando ha e pela identidade
    (CNPJ do emitente, numero, serie) quando a PRODUTO 2 so trouxe o float.

    ⚠ A NOTA REPETIDA DENTRO DA PROPRIA ABA (a mesma nota vista por duas
    filiais) vira UMA linha: vale a primeira, e a manifestacao segue a regra do
    manifestacao.py -- qualquer status vence `SemManifestacao`.
    """
    base = Path(base or BASE_PADRAO)
    todas, resumo = ler(base)
    ja_no_painel, por_identidade, por_cnpj_numero = _indice_nfe(linhas_sefaz)
    do_grupo = [l for l in todas if l.get(sefaz.CNPJ_DEST) in permitidos]

    por_chave: dict[str, dict] = {}
    for l in do_grupo:
        atual = por_chave.get(l[sefaz.CHAVE_NFE])
        if atual is None:
            por_chave[l[sefaz.CHAVE_NFE]] = l
        elif (atual.get(sefaz.MANIFESTACAO) or "") == SEM_MANIFESTACAO \
                and (l.get(sefaz.MANIFESTACAO) or "") not in ("", SEM_MANIFESTACAO):
            por_chave[l[sefaz.CHAVE_NFE]] = l

    novas: list[dict] = []
    ja_por_chave = ja_por_identidade = 0
    for chave, l in por_chave.items():
        if _digitos(chave) in ja_no_painel:
            ja_por_chave += 1
            continue
        if chave_no_painel(l, por_identidade, por_cnpj_numero):
            ja_por_identidade += 1
            continue
        filial = (l.get("_filial_nome") or "").strip()
        nome_filial = filial.split(" - ", 1)[1] if " - " in filial else filial
        # ⚠ Sem a palavra "pedido"/"PC"/"ordem" seguida de numero: e deste texto
        # que o pedidos_sefaz le o numero do PC citado na nota (RE_FORTE/RE_FRACO).
        partes = [f"Nota lida da aba PRODUTO 2 da SEFAZ.xlsx (não está na aba PRODUTO 1)."]
        if nome_filial:
            partes.append(f"Filial: {nome_filial}.")
        if l.get(sefaz.MANIFESTACAO):
            partes.append(f"Manifestação do destinatário: {l[sefaz.MANIFESTACAO]}.")
        if l["_chave_origem"] == CHAVE_SINTETICA:
            partes.append("A chave de 44 dígitos não veio na aba (a coluna foi gravada "
                          "como número); o cruzamento com a SF1 vai por fornecedor + nº da NF.")
        elif l["_chave_origem"] == CHAVE_RECUPERADA:
            partes.append("A chave de 44 dígitos foi recuperada de um MESMA PREMISSA*.xlsx "
                          "da pasta (a aba a trouxe como número).")
        l[sefaz.INFO] = " ".join(partes)
        if (l.get(sefaz.MANIFESTACAO) or "") in GRAVES:
            justificativa = (l.get("_justificativa") or "").strip()
            l[sefaz.MANIFESTO_GRAVE] = (
                f"{l[sefaz.MANIFESTACAO]}"
                + (f" — {justificativa}" if justificativa else "")
                + (f" (filial {filial})" if filial else ""))
        novas.append(l)

    resumo.update({
        "do_grupo": len(por_chave),
        "terceiros": len(todas) - len(do_grupo),
        "ja_na_sefaz": ja_por_chave + ja_por_identidade,
        "ja_por_chave": ja_por_chave,
        "ja_por_identidade": ja_por_identidade,
        "novas": len(novas),
        "novas_sinteticas": sum(1 for l in novas if l["_chave_origem"] == CHAVE_SINTETICA),
        "novas_recuperadas": sum(1 for l in novas if l["_chave_origem"] == CHAVE_RECUPERADA),
    })
    return novas, resumo
