# -*- coding: utf-8 -*-
"""Titulo com BORDERO ja emitido (hoje ou antes) e SEM data de baixa na SE2.

O bordero saiu -- o pagamento foi mandado para o banco -- mas a baixa nunca
voltou para o titulo. Ou o pagamento nao aconteceu, ou aconteceu e ninguem deu
baixa: nos dois casos o titulo fica figurando em aberto, entra de novo em
relatorio de pendencia e pode virar pagamento em dobro. O e-mail so mostra o
fato; quem olha decide qual dos dois lados esta errado.

  o recorte   `Dt. Bordero` preenchida e MENOR OU IGUAL a hoje, e `DT Baixa`
              vazia. Foi o pedido dele, ao pe da letra, em 18/09/2026: "data de
              bordero do dia atual para tras, e sem data de baixa".
  o corte     bordero a partir de 01/01/2026. Pedido dele em 18/09/2026,
              depois do primeiro disparo: "traga somente as datas de bordero a
              partir de 01/01/2026, descarte 2025". Sairam os 13 titulos de
              2025 que estavam na primeira lista.
  o detalhe   `Tipo Pgto`, `Vencimento` e `Vencto Real` em coluna, e o CODIGO
              (UUID) mais a LINHA DIGITAVEL numa faixa embaixo da linha
              (pedido dele: "coloque o codigo UUID" e "se for boleto mande
              linha digitavel").

Por que UUID e linha digitavel nao sao COLUNA: sao 36 e 47 caracteres numa
palavra so, que nao quebram. Cada um como coluna empurra a tabela para fora da
tela -- medido no alerta BOLETO S/C em 18/09/2026, e medido de novo aqui: a
tabela sem eles fica em 869px. Na faixa de largura inteira os dois quebram
sozinhos, continuam INTEIROS para copiar, e a tabela nao mexe.

Ela sai CRUA, sem mascara de pontos e espacos: o uso e' copiar e colar no banco,
e separador inventado aqui pode ser recusado do outro lado.

Vai para o Lucas e a Graziela (`.alerta_bordero_para`, fora do repo).

Uso:
    python alerta_bordero.py                 # manda, no maximo uma vez por dia
    python alerta_bordero.py --teste         # nao manda: grava a previa
    python alerta_bordero.py --forcar        # manda de novo no mesmo dia
    python alerta_bordero.py --desde 01/01/2025  # muda o corte
    python alerta_bordero.py --base "C:\\caminho\\SE2.xlsx"

Codigo de saida: SEMPRE 0 quando conseguiu ler a base. Falha de e-mail nao
derruba a rodada do painel.
"""

from __future__ import annotations

import argparse
import datetime as dt
import shutil
import sys
import tempfile
from pathlib import Path

import alertas_comuns as comuns
import caminhos

PASTA = Path(__file__).resolve().parent
PREVIA = PASTA / "DADOS" / "alerta_bordero_previa.html"
CHAVE = "bordero"

RAIZ = caminhos.raiz_lucas()
BASES = caminhos.bases_genericos()

ABA = "SE2"

# O painel da SE2 no ar: o numero do titulo na tabela abre ele. Endereco da
# RAIZ porque o painel nao le filtro pela URL (conferido em 18/09/2026).
PAINEL_SE2 = "https://se2-lucas.vercel.app/"

# O corte por baixo: bordero a partir daqui. Pedido dele em 18/09/2026,
# "descarte 2025". Data ESCRITA, e nao "ano corrente" calculado: virando
# 2027 um ano corrente esvaziaria o e-mail sozinho e calado, e o alerta
# passaria meses dizendo que nao ha nada.
DESDE_PADRAO = dt.date(2026, 1, 1)

# Conferidas pelo NOME no cabecalho, nunca por posicao: o TOTVS ja inseriu
# coluna no meio da exportacao antes, e indice fixo erraria tudo em silencio.
COLUNAS = {
    "filial": "Filial", "prefixo": "Prefixo", "tipo": "Tipo",
    "num": "No. Titulo", "parcela": "Parcela", "fornecedor": "Fornecedor",
    "nome": "Nome Fornece", "razao": "Razão Social",
    "vencimento": "Vencimento", "real": "Vencto Real",
    "valor": "Vlr.Titulo", "baixa": "DT Baixa", "saldo": "Saldo",
    "tipo_pgto": "Tipo Pgto", "linha_dig": "Linha Dig.",
    "cod_barras": "Cod.Barras", "uuid": "Campo UUID",
    "dt_bordero": "Dt. Bordero", "num_bordero": "Num Bordero",
}


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
        texto = valor.strip().replace(".", "").replace(",", ".")
        try:
            return float(texto)
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


def ler_se2(caminho: Path, hoje: dt.date,
            desde: dt.date) -> tuple[list[dict], dict]:
    """Varre a SE2 e devolve os titulos com bordero emitido e sem baixa.

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
        temporario = Path(tempfile.gettempdir()) / f"se2_bordero_{dt.date.today():%Y%m%d}.xlsx"
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
    for obrigatoria in ("dt_bordero", "baixa", "num", "tipo_pgto"):
        if obrigatoria not in onde:
            arquivo.close()
            raise KeyError("colunas ausentes na SE2: " + ", ".join(faltando))

    def campo(linha, chave):
        i = onde.get(chave)
        return linha[i] if i is not None and i < len(linha) else None

    achados = []
    resumo = {"linhas": 0, "com_bordero": 0, "ate_hoje": 0, "com_linha_dig": 0,
              "antes_do_corte": 0}
    for linha in linhas:
        if campo(linha, "num") in (None, ""):
            continue
        resumo["linhas"] += 1

        data_bordero = como_data(campo(linha, "dt_bordero"))
        if not data_bordero:
            continue
        resumo["com_bordero"] += 1
        if data_bordero > hoje:
            continue
        resumo["ate_hoje"] += 1
        if como_data(campo(linha, "baixa")):
            continue
        # O corte por baixo vem DEPOIS da baixa, de proposito: assim
        # "antes_do_corte" conta so o que ficou de fora por ser velho, e nao o
        # que ja estava baixado. E' esse numero que vai no rodape do e-mail --
        # ele precisa saber QUANTO esta sendo escondido pelo corte.
        if data_bordero < desde:
            resumo["antes_do_corte"] += 1
            continue

        # ⚠ Saldo NAO entra no filtro, de proposito: ele pediu "sem data de
        # baixa", e so isso. Titulo com saldo zerado e sem baixa e' exatamente
        # o tipo de bagunca que este alerta existe para mostrar. O saldo vai
        # como COLUNA, para quem le julgar.
        saldo = como_numero(campo(linha, "saldo"))
        linha_dig = texto(campo(linha, "linha_dig"))
        if linha_dig:
            resumo["com_linha_dig"] += 1

        vencimento = como_data(campo(linha, "vencimento"))
        real = como_data(campo(linha, "real"))
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
            "bordero": f"{data_bordero:%d/%m/%Y}",
            "num_bordero": texto(campo(linha, "num_bordero")),
            "tipo_pgto": texto(campo(linha, "tipo_pgto")),
            "vencimento": f"{vencimento:%d/%m/%Y}" if vencimento else "",
            "real": f"{real:%d/%m/%Y}" if real else "",
            "valor": reais(como_numero(campo(linha, "valor"))),
            "saldo": reais(saldo),
            "linha_dig": linha_dig,
            "cod_barras": texto(campo(linha, "cod_barras")),
            "_data": data_bordero, "_saldo": saldo,
        })
    arquivo.close()
    if temporario and temporario.exists():
        try:
            temporario.unlink()
        except OSError:
            pass
    resumo["base"] = origem
    resumo["salva_em"] = dt.datetime.fromtimestamp(origem.stat().st_mtime)
    # o bordero mais ANTIGO primeiro: e' o que esta parado ha mais tempo
    achados.sort(key=lambda a: (a["_data"], -a["_saldo"]))
    return achados, resumo


# ⚠ Largura conferida no navegador (18/09/2026, 61 titulos): 966px na primeira
# versao -- larga demais para o painel de leitura do Outlook. Duas mudancas
# derrubaram para ~830px sem tirar nada do que ele pediu:
#   . titulo e parcela numa coluna so ("000012576/04"), como se fala do titulo;
#   . "Tipo pgto" FORA do destaque, porque destaque = negrito + nao quebra, e
#     "AMARRAR ADIANTAMENTO" sozinho segurava 188px de coluna.
COLUNAS_EMAIL = [
    ("Dt. borderô", "bordero"),
    ("Borderô", "num_bordero"),
    ("Filial", "filial"),
    ("Nº título", "titulo_parcela"),
    ("Fornecedor", "fornecedor"),
    ("Tipo pgto", "tipo_pgto"),
    ("Vencimento", "vencimento"),
    ("Vencto real", "real"),
    ("Saldo em aberto", "saldo"),
]


def faixa_detalhe(linha: dict) -> str:
    """A faixa de largura inteira embaixo de cada linha da tabela.

    Leva o CODIGO (UUID) -- pedido dele em 18/09/2026 -- e, quando o titulo
    tem boleto, a LINHA DIGITAVEL. Os dois aqui embaixo, e nao em coluna,
    porque sao palavras de 36 e 47 caracteres que nao quebram: cada uma como
    coluna joga o resto da tabela para fora da tela. Na faixa eles quebram
    sozinhos e continuam inteiros para copiar.
    """
    partes = []
    if linha.get("uuid"):
        partes.append("Código (UUID): " + linha["uuid"])
    if linha.get("linha_dig"):
        partes.append("Linha digitável: " + linha["linha_dig"])
    return "     ·     ".join(partes)


def montar_html(achados: list[dict], resumo: dict, hoje: dt.date,
                desde: dt.date) -> str:
    quantos = len(achados)
    total = sum(a["_saldo"] for a in achados)
    com_ld = sum(1 for a in achados if a["linha_dig"])
    mais_antigo = achados[0]["bordero"] if achados else ""
    frase = ("1 título está com borderô emitido"
             if quantos == 1 else
             f"{quantos} títulos estão com borderô emitido")
    plural = "esse título" if quantos == 1 else "esses títulos"
    return f"""<div style="{comuns.FONTE}font-size:11pt;color:#000000;">
  <p style="margin:0 0 12px 0;">Bom dia,</p>

  <p style="margin:0 0 12px 0;">{frase} (borderô de {desde:%d/%m/%Y} até
  hoje) e <b>continua sem data de baixa na SE2</b> — somando
  <b>{reais(total)}</b> em aberto. O borderô mais antigo da lista é de
  <b>{mais_antigo}</b>.</p>

  <p style="margin:0 0 14px 0;">O borderô já saiu, mas a baixa não voltou para
  o título: ou o pagamento não aconteceu, ou aconteceu e a baixa não foi
  lançada. Enquanto isso {plural} segue figurando em aberto e pode voltar para
  pagamento. Abaixo vão o <i>tipo de pagamento</i>, o <i>vencimento</i> e o
  <i>vencimento real</i> de cada um; na faixa abaixo de cada linha vai o
  <b>código (UUID)</b> e, quando o título tem boleto, a
  <b>linha digitável</b> vem junto ({com_ld} de {quantos}). O <b>nº do
  título</b> abre o <a href="{PAINEL_SE2}" style="color:#1F3864;">painel da
  SE2</a>, onde o <b>código (UUID)</b> da faixa serve de busca.</p>

  {comuns.tabela(COLUNAS_EMAIL, achados,
                 direita=("saldo",),
                 destaque=("bordero", "real"),
                 links={"titulo_parcela": PAINEL_SE2},
                 sublinha=faixa_detalhe)}

  {comuns.rodape(
      f"Recorte: Dt. borderô de {desde:%d/%m/%Y} até {hoje:%d/%m/%Y}, sem DT Baixa.",
      f"Base SE2 salva em {resumo['salva_em']:%d/%m/%Y às %H:%M} — "
      f"{resumo['com_bordero']} títulos com borderô na base, "
      f"{resumo['ate_hoje']} com borderô até hoje"
      + (f"; {resumo['antes_do_corte']} fora da lista por serem anteriores a "
         f"{desde:%d/%m/%Y}." if resumo['antes_do_corte'] else "."))}
</div>"""


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--para", default="")
    parser.add_argument("--copia", default="")
    parser.add_argument("--teste", action="store_true",
                        help="nao envia: grava a previa e mostra o resumo")
    parser.add_argument("--forcar", action="store_true",
                        help="envia mesmo se ja mandou hoje")
    parser.add_argument("--desde", default=f"{DESDE_PADRAO:%d/%m/%Y}",
                        help="so bordero a partir desta data (dd/mm/aaaa)")
    parser.add_argument("--base", type=Path, default=None)
    args = parser.parse_args()

    base = args.base or achar_base()
    if not base or not base.exists():
        print(f"   AVISO: nao achei a SE2 em {BASES} -- alerta de borderô nao rodou.")
        return 0

    hoje = dt.date.today()
    desde = como_data(args.desde) or DESDE_PADRAO

    try:
        achados, resumo = ler_se2(base, hoje, desde)
    except Exception as erro:  # noqa: BLE001
        print(f"   AVISO: nao consegui ler a SE2 ({erro}). A rodada segue.")
        return 0

    print(f"   SE2: {resumo['linhas']} títulos | {resumo['com_bordero']} com borderô | "
          f"{resumo['ate_hoje']} com borderô até {hoje:%d/%m} | "
          f"{resumo['antes_do_corte']} descartado(s) por serem anteriores a "
          f"{desde:%d/%m/%Y}")

    if not achados:
        print("   Nenhum título com borderô emitido e sem data de baixa.")
        return 0

    total = sum(a["_saldo"] for a in achados)
    print(f"   {len(achados)} título(s) com borderô e SEM baixa ({reais(total)}, "
          f"{resumo['com_linha_dig']} com linha digitável):")
    for a in achados:
        print(f"     bord {a['bordero']} nº {a['num_bordero'] or '-':<6} | {a['filial']} | "
              f"{a['titulo']}/{a['parcela'] or '-'} | venc {a['vencimento'] or '-'} | "
              f"real {a['real'] or '-'} | {a['tipo_pgto'][:14]:<14} | {a['saldo']:>15} | "
              f"{a['fornecedor'][:28]}")

    corpo = montar_html(achados, resumo, hoje, desde)
    assunto = (f"[Painel Boletos] {len(achados)} título"
               f"{'s' if len(achados) > 1 else ''} com borderô emitido e SEM baixa "
               f"na SE2 - {hoje:%d/%m/%Y}")

    if args.teste:
        PREVIA.parent.mkdir(parents=True, exist_ok=True)
        PREVIA.write_text(corpo, encoding="utf-8")
        print(f"\n   MODO TESTE: nada foi enviado.\n   Assunto: {assunto}")
        print(f"   Prévia:  {PREVIA}")
        return 0

    destino = args.para or comuns.destinatarios(CHAVE)
    if not destino:
        print("   AVISO: nao achei para quem mandar (.alerta_bordero_para).")
        return 0
    copia = args.copia or comuns.copias(CHAVE)

    if not args.forcar and comuns.ja_enviado_hoje(CHAVE):
        print(f"   Já enviado hoje ({comuns.quando_enviou(CHAVE)}) -- não mando de novo.")
        print("   Para mandar assim mesmo: --forcar")
        return 0

    try:
        comuns.enviar(assunto, corpo, destino, copia=copia)
        comuns.marcar_enviado(CHAVE, len(achados))
        print(f"   E-mail enviado para {destino}"
              + (f" (cópia: {copia})" if copia else "") + ".")
    except Exception as erro:  # noqa: BLE001
        print(f"   AVISO: o e-mail NAO foi enviado ({erro}). O painel segue.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
