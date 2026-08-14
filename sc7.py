# -*- coding: utf-8 -*-
"""Pedidos de compra (SC7) e quem os solicitou (SC1).

Este modulo NAO monta aba nenhuma: ele so entrega os pedidos ja agregados, para
o `pedidos_sefaz.py` cruzar com as notas da SEFAZ.

⚠ A SC7 E POR ITEM, NAO POR PEDIDO
----------------------------------
24.402 itens para 3.377 pedidos -- um PC chega a ter 119 linhas. Toda leitura
tem de subir de nivel, e cada campo sobe de um jeito diferente (ver `Pedido`):
somar `Qtd.a Classi` de 119 itens e certo, somar `Dt. Entrega` nao quer dizer
nada. Pegar "o primeiro item" seria pior: medido nos 149 pedidos que casaram com
alguma nota, 61 tem MAIS DE UM comprador e 41 mais de uma SC -- em 41% dos casos
o primeiro valor esconderia o resto sem avisar. Por isso o padrao e juntar os
distintos com " · " (escolha do usuario, 13/08/2026).

⚠ O SOLICITANTE NAO EXISTE NA SC7
---------------------------------
A coluna `Solicitante` (CP) esta VAZIA nas 24.402 linhas -- medido, nao suposto.
Quem tem o nome e a SC1, e o caminho e o numero do pedido: `Num. Pedido` (Z)
casa com o `Numero PC`, e o nome sai de `Solicitante` (Y). Cobertura medida:
2.634 dos 3.377 pedidos. `Usuário SC` (GU) da SC7 tambem quase nao existe (83 de
24.402, e ZERO nos pedidos que casaram com nota) -- ela vem junto porque o
usuario pediu, mas nascendo vazia.

⚠ COLUNAS LIDAS POR POSICAO, COM O NOME COMO TRAVA
--------------------------------------------------
Mesma regra do `sefaz.py`: a SC7 tem 207 colunas e 3 nomes repetidos, a SC1 tem
119 e 2 repetidos. Ler por nome escolhe a errada calado. A letra e o que vale; o
nome e conferido na leitura e a leitura PARA se ele saiu do lugar.
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.utils import column_index_from_string

import sefaz

BASE_GENERICOS = Path(
    r"C:\Users\lucas\OneDrive - Grupo S&D\Arquivos de Gabriella Karla Oliveira Milas - FINANCEIRO COMPARTILHADO"
    r"\LUCAS ABNER ARAUJO\BASES GENERICOS"
)
BASE_SC7 = BASE_GENERICOS / "SC7.xlsx"
BASE_SC1 = BASE_GENERICOS / "SC1.xlsx"

# (letra, nome esperado no cabecalho, apelido interno)
COLUNAS_SC7 = [
    ("AI", "Numero PC",     "pc"),
    ("AB", "Numero da SC",  "sc"),
    ("BH", "Qtd.a Classi",  "qtd_classi"),
    # ⚠ O TOTAL do pedido, que da tamanho ao "a classificar" (pedido do usuario,
    # 14/08/2026: "para saber o total a ser classificado e o total do pedido").
    # Sozinho, "20 a classificar" nao diz nada: 20 de 20 e um pedido intocado,
    # 20 de 289 e um pedido quase todo classificado.
    ("I",  "Quantidade",    "quantidade"),
    ("BM", "Controle Ap.",  "controle"),
    ("GU", "Usuário SC",    "usuario_sc"),
    ("GV", "Comprador",     "comprador"),
    ("AU", "Ped. Encerr.",  "encerrado"),
    ("L",  "Dt. Entrega",   "entrega"),
    ("GM", "1° Venc",       "venc1"),
    ("A",  "Filial",        "filial"),
    ("B",  "Razão Social",  "razao"),
    ("AG", "Fornecedor",    "fornecedor"),
    ("AH", "DT Emissao",    "emissao"),
    ("V",  "Vlr.Total",     "vlr"),
]

COLUNAS_SC1 = [
    ("Z", "Num. Pedido",  "pc"),
    ("Y", "Solicitante",  "solicitante"),
    ("B", "Numero da SC", "sc"),
    ("DI", "Usuário",     "usuario"),
]

JUNTA = " · "

# Pedido com PARTE dos itens encerrada. ⚠ O painel conhece este texto pelo nome
# (`PC_PARCIAL` no painel_modelo.html): e ele que manda o pedido para o botao
# PENDENTE do filtro "Pedido encerrado". Mudar aqui e mudar la.
PARCIAL = "PARCIAL"


def normalizar_pc(valor) -> str:
    """A chave do pedido: so digitos, sem zeros a esquerda.

    ⚠ `003100` e `3100` sao O MESMO pedido -- a SC7 grava com 6 digitos, a nota
    escreve dos dois jeitos, e a SC1 grava de um terceiro. Sem normalizar os
    tres lados, o cruzamento nao acha nada e ninguem descobre por que.
    """
    return re.sub(r"\D", "", str(valor or "")).lstrip("0")


def formatar_pc(valor) -> str:
    """Como o pedido se escreve no TOTVS: 6 digitos com zeros a frente."""
    numero = normalizar_pc(valor)
    return numero.rjust(6, "0") if numero else ""


@dataclass
class Pedido:
    """Um PC inteiro, com os itens ja somados/juntados.

    Cada campo sobe de nivel do jeito que faz sentido para ELE -- esta e a parte
    do modulo que nao da para deduzir olhando a planilha.
    """
    numero: str                                   # normalizado (chave)
    itens: int = 0
    # somas: as unicas tres em que somar responde a pergunta certa
    qtd_classi: float = 0.0
    # ⚠ Quantidade PODE SER FRACIONARIA -- medido na SC7: 889 dos 24.516 itens
    # (0,954 / 1,4 / 0,062) e 316 dos 3.382 pedidos somam quebrado. Quem exibir
    # isso nao pode arredondar para inteiro.
    quantidade: float = 0.0
    vlr: float = 0.0
    # menor data: e a que chega antes, e e o que a pessoa precisa ver
    entrega: dt.date | None = None
    venc1: dt.date | None = None
    # distintos, na ordem em que aparecem (juntados por " · " na saida)
    scs: list = field(default_factory=list)
    compradores: list = field(default_factory=list)
    usuarios_sc: list = field(default_factory=list)
    controles: list = field(default_factory=list)
    encerrados: list = field(default_factory=list)
    # quantos itens vieram com a marca de encerrado: e o que separa o pedido
    # encerrado do encerrado PELA METADE (ver a propriedade `encerrado`)
    itens_encerrados: int = 0
    filiais: list = field(default_factory=list)
    razoes: list = field(default_factory=list)
    # raizes de CNPJ do fornecedor: e por aqui que a nota encontra o pedido
    fornecedores: set = field(default_factory=set)
    emissoes: list = field(default_factory=list)
    # valor de cada item, para reconhecer entrega PARCIAL (a nota traz um item
    # do pedido, nao o pedido inteiro)
    valores_item: list = field(default_factory=list)

    def somar(self, linha: dict) -> None:
        self.itens += 1
        for campo in ("qtd_classi", "quantidade", "vlr"):
            valor = sefaz._numero(linha.get(campo))
            if valor:
                setattr(self, campo, getattr(self, campo) + valor)
        valor_item = sefaz._numero(linha.get("vlr"))
        if valor_item:
            self.valores_item.append(round(valor_item, 2))

        for campo in ("entrega", "venc1"):
            data = sefaz._data(linha.get(campo))
            atual = getattr(self, campo)
            if data and (atual is None or data < atual):
                setattr(self, campo, data)

        emissao = sefaz._data(linha.get("emissao"))
        if emissao:
            self.emissoes.append(emissao)

        for campo, lista in (("sc", self.scs), ("comprador", self.compradores),
                             ("usuario_sc", self.usuarios_sc),
                             ("controle", self.controles),
                             ("encerrado", self.encerrados),
                             ("filial", self.filiais), ("razao", self.razoes)):
            valor = sefaz._texto(linha.get(campo))
            if valor and valor not in lista:
                lista.append(valor)
        if sefaz._texto(linha.get("encerrado")):
            self.itens_encerrados += 1

        raiz = re.sub(r"\D", "", sefaz._texto(linha.get("fornecedor")))[:8]
        if raiz:
            self.fornecedores.add(raiz)

    # --- saida ---------------------------------------------------------
    @property
    def numero_totvs(self) -> str:
        return formatar_pc(self.numero)

    @property
    def sc(self) -> str:
        return JUNTA.join(self.scs)

    @property
    def comprador(self) -> str:
        return JUNTA.join(self.compradores)

    @property
    def usuario_sc(self) -> str:
        return JUNTA.join(self.usuarios_sc)

    @property
    def controle(self) -> str:
        return JUNTA.join(self.controles)

    @property
    def encerrado(self) -> str:
        """"E" so quando TODOS os itens estao encerrados; `PARCIAL` quando alguns.

        ⚠ Este e o UNICO campo que nao pode juntar os distintos como os outros.
        Medido em 14/08/2026: o PC 002890 tem 15 itens, **2** com "E" e 13 em
        aberto com 284 a classificar -- e a celula saia "E", dizendo que o pedido
        estava encerrado. Como o painel ganhou um filtro que pergunta exatamente
        isso ("encerrado ou pendente?"), o meio-termo precisa aparecer: pedido
        com item em aberto e trabalho pendente, nao pedido fechado.
        Nao e caso isolado: sao **178 dos 3.380** pedidos da SC7 (2.948 fechados
        de verdade, 254 sem marca nenhuma), e 9 deles estao amarrados a alguma
        nota da aba -- 13 linhas que, sem isto, entrariam em ENCERRADO.
        """
        if not self.encerrados:
            return ""
        if self.itens_encerrados < self.itens:
            return PARCIAL
        return JUNTA.join(self.encerrados)

    @property
    def razao(self) -> str:
        return JUNTA.join(self.razoes)

    @property
    def filial(self) -> str:
        return JUNTA.join(self.filiais)

    @property
    def emissao(self) -> dt.date | None:
        return min(self.emissoes) if self.emissoes else None


def _ler(caminho: Path, colunas: list) -> list[dict]:
    """Uma planilha de coluna larga, lida por posicao e conferida por nome."""
    wb, temporaria = sefaz._abrir(caminho)
    try:
        ws = wb[wb.sheetnames[0]]
        brutas = ws.iter_rows(values_only=True)
        cabecalho = list(next(brutas, ()) or ())
        sefaz._conferir(cabecalho, colunas, f"{caminho.name}/{ws.title}")
        posicoes = [(column_index_from_string(l) - 1, apelido)
                    for l, _n, apelido in colunas]
        linhas = []
        for bruta in brutas:
            if all(v in (None, "") for v in bruta):
                continue
            linhas.append({apelido: (bruta[i] if i < len(bruta) else None)
                           for i, apelido in posicoes})
        return linhas
    finally:
        wb.close()
        if temporaria:
            import shutil
            shutil.rmtree(temporaria, ignore_errors=True)


def ler_pedidos(caminho: Path | None = None) -> dict[str, Pedido]:
    """{numero normalizado: Pedido} -- a SC7 inteira, agregada por PC."""
    linhas = _ler(caminho or BASE_SC7, COLUNAS_SC7)
    pedidos: dict[str, Pedido] = {}
    for linha in linhas:
        numero = normalizar_pc(linha.get("pc"))
        if not numero:
            continue
        pedidos.setdefault(numero, Pedido(numero=numero)).somar(linha)
    return pedidos


def ler_solicitantes(caminho: Path | None = None) -> dict[str, list[str]]:
    """{numero do pedido: solicitantes distintos} -- vem da SC1.

    ⚠ Um PC pode nascer de varias SCs, logo de varios solicitantes (21 dos 149
    pedidos casados). Devolve TODOS, na ordem em que aparecem.
    """
    linhas = _ler(caminho or BASE_SC1, COLUNAS_SC1)
    por_pc: dict[str, list[str]] = {}
    for linha in linhas:
        numero = normalizar_pc(linha.get("pc"))
        if not numero:
            continue
        # `Usuário` (DI) e a reserva: nos exemplos vistos traz o mesmo valor, e
        # sem ela um pedido cuja SC veio sem solicitante ficaria sem ninguem.
        nome = (sefaz._texto(linha.get("solicitante"))
                or sefaz._texto(linha.get("usuario")))
        if not nome:
            continue
        lista = por_pc.setdefault(numero, [])
        if nome not in lista:
            lista.append(nome)
    return por_pc


def carregar(base_sc7: Path | None = None,
             base_sc1: Path | None = None) -> tuple[dict[str, Pedido], dict[str, list[str]], dict]:
    """(pedidos, solicitantes por pedido, resumo).

    A SC1 e OPCIONAL: sem ela os pedidos continuam valendo e so o solicitante
    fica vazio. Faltar a SC7, ao contrario, e nao ter aba nenhuma.
    """
    caminho7 = base_sc7 or BASE_SC7
    caminho1 = base_sc1 or BASE_SC1
    pedidos = ler_pedidos(caminho7)
    solicitantes = ler_solicitantes(caminho1) if caminho1.exists() else {}
    numeros = [int(n) for n in pedidos if n.isdigit()]
    resumo = {
        "pedidos": len(pedidos),
        "itens": sum(p.itens for p in pedidos.values()),
        "com_solicitante": len(set(pedidos) & set(solicitantes)),
        # a faixa de PCs que a exportacao cobre: e ela que diz se um numero
        # citado numa nota tem chance de estar aqui (ver pedidos_sefaz)
        "menor_pc": min(numeros) if numeros else 0,
        "maior_pc": max(numeros) if numeros else 0,
        "base_sc7": str(caminho7),
        "base_sc1": str(caminho1) if caminho1.exists() else "",
        "salva_em": dt.datetime.fromtimestamp(
            caminho7.stat().st_mtime).strftime("%d/%m/%Y às %H:%M"),
    }
    return pedidos, solicitantes, resumo
