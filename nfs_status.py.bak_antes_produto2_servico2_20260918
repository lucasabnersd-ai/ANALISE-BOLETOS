# -*- coding: utf-8 -*-
"""NFS STATUS: a aba de SITUACAO das NFS-e dentro da SEFAZ.xlsx.

10/09/2026, pedido do usuario: *"NA BASE SEFAZ INCLUI AS ABAS NFS STATUS. CRUZE
AS INFORMACOES PELO NUMERO DA NF E FORNECEDOR, E USANDO OUTRA ANALISE QUE ACHAR
VALIDA, TRAZENDO TAMBEM SE FOI CANCELADA OU SUBSTITUIDA. HA CASOS DE NF COMECAR
COM 26 OU 2026 -- BASTA DESCONSIDERAR E CONSIDERAR OS DIGITOS FINAIS PARA CRUZAR
E VER SE ESTA NO TOTVS. USE CNPJ PRESTADOR TAMBEM PARA CRUZAR"*.

O QUE A ABA E
-------------
Um relatorio de situacao das NFS-e tomadas pelas nossas filiais (15 colunas:
`Situação NFSe`, `Filial`, `Série`, `RPS`, `N NF`, `Emissão`, `CNPJ/CPF
Prestador`, `Prestador`, ..., `Integração ERP`, `Valor dos Serviços`). Ele NAO
tem chave de 44 digitos nem CNPJ do tomador -- a filial vem como
"043001 - S&D FLORESTAL LOGISTICA LTDA", e o CNPJ dela sai da LISTAGEM EMPRESAS
BIOFLOR (coluna FILIAL TOTVS).

O QUE ESTE MODULO FAZ
---------------------
1. **Enriquece** cada NFS-e do painel com duas coisas do relatorio, na coluna
   `NFS STATUS`: a `Integração ERP` (**Exportada** = o portal ja mandou a nota
   para o TOTVS; **Disponível** = ainda nao) e, quando a situacao do relatorio
   DIVERGE do `Status da Nota` da SEFAZ, a situacao da prefeitura por extenso.
   A situacao do relatorio tambem vai para a linha numa chave interna
   (`sefaz.SITUACAO_PREFEITURA`): e ela, junto com o `Status da Nota`, que diz
   se a nota esta cancelada/substituida -- ver `sefaz.esta_cancelada`.
2. **Acrescenta** ao painel as notas do relatorio que a aba de NFS-e da
   SEFAZ.xlsx NAO tem (Origem `NFS STATUS`), para que passem pelo mesmo
   cruzamento com a SF1 ("esta no TOTVS?") e pelos pedidos. Medido em
   10/09/2026: as 112 notas unicas do relatorio ja estavam TODAS na aba de
   NFS-e -- zero novas, zero divergencias de situacao. O caminho existe para a
   proxima exportacao, que pode trazer nota que a SEFAZ ainda nao trouxe.

COMO CRUZA -- E POR QUE ASSIM
-----------------------------
A chave e **CNPJ do prestador (14 digitos) + numero da NF** em TODAS as formas
comparaveis de `sefaz.nf_variantes`: `2600000010412` (ano em 2 + sequencial, o
formato da prefeitura) casa com `10412`, e `2026000000009` com `9`. O CNPJ
entra INTEIRO, nao pela raiz: a NFS-e comeca do 1 em cada prestador e o mesmo
numero existe em dezenas de fornecedores -- so o numero cruzaria a nota errada.
O RPS fica de reserva: casa quando o numero nao casou.

O RELATORIO VEM COM LINHAS 100% REPETIDAS (o usuario cola exportacoes em cima
da anterior): em 10/09/2026 eram 223 linhas para 112 notas. Aqui elas sao
tiradas na leitura, e a contagem vai para o resumo -- tirar calado esconderia
que o arquivo precisa de limpeza.

O VALOR VEM EM FORMATO AMERICANO: "R$5,400.00". O `sefaz._numero` le a virgula
como decimal e devolveria 5.4 -- por isso o `_valor` daqui.

A ABA E OPCIONAL: exportacao sem ela nao derruba a rodada. A coluna sai como
"— aba NFS STATUS não lida" (e nao vazia), porque vazio seria lido como "esta
nota nao esta no relatorio", afirmacao que ninguem fez.
"""
from __future__ import annotations

import hashlib
import re
import shutil
from pathlib import Path

import sefaz

# O que a coluna `Origem` mostra para as notas que SO existem no relatorio.
ORIGEM = "NFS STATUS"

# Cabecalho da aba (comparado por `sefaz._rotulo`: sem acento, espacos
# colapsados, minusculo). Os tres primeiros identificam a aba.
COL_SITUACAO = "Situação NFSe"
COL_NNF = "N NF"
COL_CNPJ = "CNPJ/CPF Prestador"
COL_FILIAL = "Filial"
COL_SERIE = "Série"
COL_RPS = "RPS"
COL_EMISSAO = "Emissão"
COL_PRESTADOR = "Prestador"
COL_MUNICIPIO = "Município"
COL_INTEGRACAO = "Integração ERP"
COL_VALOR = "Valor dos Serviços"
IDENTIDADE = (COL_SITUACAO, COL_NNF, COL_CNPJ)

# Valores da coluna `NFS STATUS` quando NAO ha registro para a nota. Comecam
# com travessao como os da manifestacao, para ficarem juntos no filtro.
NAO_SE_APLICA = "— não se aplica (não é NFS-e)"
FORA_DO_PERIODO = "— fora do período do relatório"
NAO_CONSTA = "NÃO CONSTA NO RELATÓRIO"
SEM_BASE = "— aba NFS STATUS não lida"


def _digitos(valor) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def _valor(valor):
    """Valor do servico: "R$5,400.00" (americano) ou "R$ 5.400,00" (BR)."""
    if valor in (None, ""):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    limpo = re.sub(r"[^0-9,.\-]", "", str(valor))
    if re.fullmatch(r"-?\d{1,3}(,\d{3})*(\.\d+)?", limpo) or re.fullmatch(r"-?\d+\.\d+", limpo):
        try:
            return float(limpo.replace(",", ""))
        except ValueError:
            return None
    return sefaz._numero(valor)


def _sem_acento_min(texto) -> str:
    return sefaz._sem_acento(str(texto or "")).strip().lower()


def _achar_aba(wb):
    """A aba do relatorio, pelo CABECALHO -- o nome ("NFS  STATUS", com dois
    espacos) e detalhe de digitacao e pode mudar na proxima exportacao."""
    procurados = {sefaz._rotulo(c) for c in IDENTIDADE}
    for ws in wb.worksheets:
        cabecalho = {sefaz._rotulo(c) for c in sefaz._cabecalho_de(ws) if c}
        if procurados <= cabecalho:
            return ws
    raise RuntimeError(
        "A SEFAZ.xlsx nao tem a aba NFS STATUS (procurei uma aba com as colunas "
        + " / ".join(repr(c) for c in IDENTIDADE) + " no cabecalho).\n"
        "  a planilha tem: " + " / ".join(repr(n) for n in wb.sheetnames))


def ler(base: Path | None = None) -> tuple[list[dict], dict]:
    """Os registros do relatorio, sem repetidos. Devolve (registros, resumo)."""
    base = base or sefaz.BASE_PADRAO
    wb, temporaria = sefaz._abrir(base)
    try:
        ws = _achar_aba(wb)
        titulo = ws.title
        brutas = ws.iter_rows(values_only=True)
        cabecalho = [sefaz._rotulo(c) for c in (next(brutas, ()) or ())]
        pos = {}
        for nome in (COL_SITUACAO, COL_NNF, COL_CNPJ, COL_FILIAL, COL_SERIE, COL_RPS,
                     COL_EMISSAO, COL_PRESTADOR, COL_MUNICIPIO, COL_INTEGRACAO, COL_VALOR):
            rot = sefaz._rotulo(nome)
            pos[nome] = cabecalho.index(rot) if rot in cabecalho else None

        def campo(bruta, nome):
            i = pos[nome]
            return bruta[i] if i is not None and i < len(bruta) else None

        vistas: set = set()
        registros, lidas, repetidas = [], 0, 0
        for bruta in brutas:
            if all(v in (None, "") for v in bruta):
                continue
            lidas += 1
            marca = tuple(bruta)
            if marca in vistas:
                repetidas += 1
                continue
            vistas.add(marca)
            cnpj = _digitos(campo(bruta, COL_CNPJ))
            nnf = sefaz._texto(campo(bruta, COL_NNF))
            if not cnpj or not nnf:
                continue
            filial = sefaz._texto(campo(bruta, COL_FILIAL))
            registros.append({
                "situacao": sefaz._texto(campo(bruta, COL_SITUACAO)),
                "filial": filial,
                "filial_codigo": _digitos(filial[:6]) if filial else "",
                "serie": sefaz._texto(campo(bruta, COL_SERIE)),
                "rps": _digitos(campo(bruta, COL_RPS)).lstrip("0"),
                "nnf": nnf,
                "emissao": sefaz._data(campo(bruta, COL_EMISSAO)),
                "cnpj": cnpj,
                "prestador": sefaz._texto(campo(bruta, COL_PRESTADOR)),
                "municipio": sefaz._texto(campo(bruta, COL_MUNICIPIO)),
                "integracao": sefaz._texto(campo(bruta, COL_INTEGRACAO)),
                "valor": _valor(campo(bruta, COL_VALOR)),
            })
    finally:
        wb.close()
        if temporaria:
            shutil.rmtree(temporaria, ignore_errors=True)
    emissoes = [r["emissao"] for r in registros if r["emissao"]]
    resumo = {
        "aba": titulo, "linhas": lidas, "repetidas": repetidas,
        "registros": len(registros),
        "de": min(emissoes) if emissoes else None,
        "ate": max(emissoes) if emissoes else None,
        "situacoes": _contar(r["situacao"] for r in registros),
        "integracao": _contar(r["integracao"] for r in registros),
    }
    return registros, resumo


def _contar(valores) -> dict:
    contagem: dict[str, int] = {}
    for v in valores:
        contagem[v or "(vazio)"] = contagem.get(v or "(vazio)", 0) + 1
    return dict(sorted(contagem.items()))


def _formas(nnf, ano) -> set:
    return set(sefaz.nf_variantes(nnf, ano).keys())


def _indice(registros: list[dict]) -> tuple[dict, dict]:
    """Dois indices: (cnpj, forma do numero) e (cnpj, rps)."""
    por_nf: dict[tuple, dict] = {}
    por_rps: dict[tuple, dict] = {}
    for r in registros:
        ano = r["emissao"].year if r["emissao"] else None
        for forma in _formas(r["nnf"], ano):
            por_nf.setdefault((r["cnpj"], forma), r)
        if r["rps"]:
            por_rps.setdefault((r["cnpj"], r["rps"]), r)
    return por_nf, por_rps


def _registro_da_linha(linha: dict, por_nf: dict, por_rps: dict):
    cnpj = _digitos(linha.get(sefaz.CNPJ_EMIT))
    if len(cnpj) != 14:
        return None
    emissao = linha.get(sefaz.EMISSAO)
    ano = emissao.year if emissao else None
    for forma in _formas(linha.get(sefaz.NUM_NF), ano):
        achado = por_nf.get((cnpj, forma))
        if achado is not None:
            return achado
    rps = _digitos(linha.get("_rps")).lstrip("0")
    if rps:
        return por_rps.get((cnpj, rps))
    return None


def _e_servico(linha: dict) -> bool:
    return linha.get(sefaz.ORIGEM) in (sefaz.NFSE, ORIGEM)


def marcar_sem_base(linhas: list[dict]) -> None:
    """Sem a aba, a coluna DIZ que nao foi lida -- nunca fica vazia."""
    for linha in linhas:
        linha[sefaz.NFS_STATUS] = SEM_BASE if _e_servico(linha) else NAO_SE_APLICA


def _linha_nova(r: dict, permitidos: dict, filiais_por_codigo: dict) -> dict | None:
    """Uma nota que SO existe no relatorio, no vocabulario das linhas da SEFAZ.

    NAO tem chave de 44 digitos -- o relatorio nao traz. A `Chave NF-e` sai
    como um identificador legivel e ESTAVEL ("NFS STATUS <cnpj> <numero>"), e
    dele saem os uuids das outras abas (elas usam os digitos da chave). Tomador
    = a filial do relatorio, traduzida pela LISTAGEM (FILIAL TOTVS -> CNPJ);
    filial fora da listagem nao entra, como qualquer outra nota fora do grupo.
    """
    cnpj_dest = filiais_por_codigo.get(r["filial_codigo"])
    if not cnpj_dest or cnpj_dest not in permitidos:
        return None
    numero = _digitos(r["nnf"])
    chave = f"{ORIGEM} {r['cnpj']} {numero}"
    linha = {
        sefaz.ORIGEM: ORIGEM,
        sefaz.CHAVE_NFE: chave,
        sefaz.NUM_NF: r["nnf"],
        sefaz.EMISSAO: r["emissao"],
        sefaz.STATUS_NOTA: r["situacao"],
        sefaz.SITUACAO_PREFEITURA: r["situacao"],
        sefaz.NFS_STATUS: r["integracao"] or NAO_CONSTA,
        sefaz.MANIFESTACAO: "— não se aplica (NFS-e)",
        sefaz.NAT_OP: "",
        sefaz.CFOP: "",
        sefaz.TIPO_OP: "",
        sefaz.SAIDA: None,
        sefaz.VLR_ITEM: None,
        sefaz.VLR_TOTAL: r["valor"],
        sefaz.VENC_DUP: None,
        sefaz.VLR_DUP: None,
        sefaz.INFO: (f"Nota lida da aba NFS STATUS da SEFAZ.xlsx (não está na aba de NFS-e). "
                     f"Filial {r['filial']} · série {r['serie']} · RPS {r['rps']} · "
                     f"município {r['municipio']} · Integração ERP: {r['integracao']}"),
        sefaz.CNPJ_EMIT: r["cnpj"],
        sefaz.NOME_EMIT: r["prestador"],
        sefaz.FANTASIA: "",
        sefaz.CNPJ_DEST: cnpj_dest,
        sefaz.NOME_DEST: permitidos.get(cnpj_dest, ""),
        "_rps": r["rps"],
    }
    linha[sefaz.CHAVE] = sefaz.PREFIXO + hashlib.sha1(
        f"{ORIGEM}|{r['cnpj']}|{numero}".encode("utf-8")).hexdigest()[:16]
    return linha


def anotar(linhas: list[dict], permitidos: dict, base: Path | None = None,
           filiais: dict | None = None) -> tuple[list[dict], dict]:
    """Preenche `sefaz.NFS_STATUS` nas linhas e devolve as notas NOVAS.

    Roda DEPOIS do `sefaz.ler_com_historico()` (as linhas novas nao podem ir
    para o historico, que existe para acusar nota que sumiu da SEFAZ) e do
    `manifestacao.anotar()` (a coluna de manifestacao das novas e posta aqui).
    Devolve (linhas novas, resumo).
    """
    registros, resumo = ler(base)
    por_nf, por_rps = _indice(registros)
    de, ate = resumo["de"], resumo["ate"]
    if filiais is None:
        try:
            filiais = sefaz.ler_filiais_empresas()
        except Exception:  # noqa: BLE001 - sem a filial nao ha como saber o tomador
            filiais = {}
    filiais_por_codigo = {codigo: cnpj for cnpj, codigo in filiais.items()}

    usados: set[int] = set()
    divergentes = 0
    contagem: dict[str, int] = {}
    notas_vistas: set[str] = set()
    for linha in linhas:
        if not _e_servico(linha):
            linha[sefaz.NFS_STATUS] = NAO_SE_APLICA
            continue
        r = _registro_da_linha(linha, por_nf, por_rps)
        if r is None:
            emissao = linha.get(sefaz.EMISSAO)
            dentro = bool(de and ate and emissao and de <= emissao <= ate)
            valor = NAO_CONSTA if dentro else FORA_DO_PERIODO
        else:
            usados.add(id(r))
            valor = r["integracao"] or NAO_CONSTA
            linha[sefaz.SITUACAO_PREFEITURA] = r["situacao"]
            status_sefaz = _sem_acento_min(linha.get(sefaz.STATUS_NOTA))
            situacao = _sem_acento_min(r["situacao"])
            if situacao and situacao != status_sefaz \
                    and {situacao, status_sefaz} != {"normal", "autorizada"}:
                valor += (f" · {r['situacao'].upper()} NA PREFEITURA "
                          f"(a SEFAZ diz {linha.get(sefaz.STATUS_NOTA) or 'sem status'})")
                divergentes += 1
        linha[sefaz.NFS_STATUS] = valor
        chave = linha.get(sefaz.CHAVE_NFE)
        if chave not in notas_vistas:
            notas_vistas.add(chave)
            rotulo = valor.split(" · ")[0]
            contagem[rotulo] = contagem.get(rotulo, 0) + 1

    novas, sem_filial = [], 0
    for r in registros:
        if id(r) in usados:
            continue
        nova = _linha_nova(r, permitidos, filiais_por_codigo)
        if nova is None:
            sem_filial += 1
            continue
        novas.append(nova)

    resumo.update({
        "anotadas": dict(sorted(contagem.items())),
        "divergentes": divergentes,
        "ja_na_sefaz": len(usados),
        "novas": len(novas),
        "fora_do_grupo": sem_filial,
        "de": de.isoformat() if de else None,
        "ate": ate.isoformat() if ate else None,
    })
    return novas, resumo
