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

⚠ O MESMO NUMERO DE PEDIDO EXISTE EM VARIAS FILIAIS
----------------------------------------------------
`001522` nao identifica pedido nenhum: e um pedido da CAPIVARA (filial 004001,
LACERDA DINIZ, R$ 115,02, comprador rafael.lima), OUTRO da BIOENERGIA (038001,
TELEFONICA, fabio.costa) e OUTRO da LOGISTICA (043001, ALANA CAROLINE,
otoniel.cruz). Medido na base de 31/08/2026: 2.541 dos 3.669 numeros estao em
mais de uma filial (1.405 em duas, 1.136 em tres) -- os 3.669 numeros sao 7.346
pedidos de verdade.

Por isso a chave deste modulo e `chave_pedido()` = FILIAL + numero, e nunca o
numero sozinho. Ate 31/08/2026 era so o numero, e o efeito era colar tres pedidos
num so: a linha da Capivara mostrava "rafael.lima · fabio.costa · otoniel.cruz",
somava R$ 803,44 de tres fornecedores diferentes num pedido de R$ 115,02, herdava
a SC e o solicitante do pedido da Logistica e as datas do da Bioenergia. Como o
valor somado nao batia com o da nota, a linha ainda caia de PEDIDO CONFIRMADO
para CONFERIR PEDIDO. Eram 31 das 106 notas com pedido (29%).

⚠ QUEM CASA NOTA COM PEDIDO NAO PODE MAIS BUSCAR "O PEDIDO 1522": o numero
citado numa nota da uma LISTA de candidatos (um por filial), e quem escolhe entre
eles e o CNPJ do fornecedor e o valor -- os mesmos criterios de sempre, que antes
eram aplicados a um Frankenstein das tres filiais. Ver `pedidos_sefaz.escolher()`.

⚠ O SOLICITANTE NAO EXISTE NA SC7
---------------------------------
A coluna `Solicitante` (CP) esta VAZIA nas 24.402 linhas -- medido, nao suposto.
Quem tem o nome e a SC1, e o caminho e o pedido inteiro: `Filial` (A) +
`Num. Pedido` (Z) casam com a chave, e o nome sai de `Solicitante` (Y). ⚠ A
filial entra aqui pelo mesmo motivo de sempre: casando so pelo numero, 2.031
pedidos ganhavam um solicitante de OUTRA filial (era daniela.fernandes, da
Logistica, aparecendo no pedido da Capivara). A SC1 concorda com a SC7 sobre a
filial em 15.626 das 15.631 linhas -- a solicitacao nasce na filial que faz o
pedido. Cobertura medida:
2.634 dos 3.377 pedidos. `Usuário SC` (GU) da SC7 tambem quase nao existe (83 de
24.402, e ZERO nos pedidos que casaram com nota) -- ela vem junto porque o
usuario pediu, mas nascendo vazia.

⚠ COLUNAS LIDAS POR POSICAO, COM O NOME COMO TRAVA
--------------------------------------------------
Mesma regra do `sefaz.py`: a SC7 tem 207 colunas e 3 nomes repetidos, a SC1 tem
119 e 2 repetidos. Ler por nome escolhe a errada calado. A letra e o que vale; o
nome e conferido na leitura e a leitura PARA se ele saiu do lugar.

⚠ SO OS PEDIDOS EM ABERTO SAO ANALISADOS (28/08/2026)
-----------------------------------------------------
Pedido do usuario: "na coluna AU `Ped. Encerr.`, traga e analise SOMENTE os
titulos em aberto -- considere so o que NAO tem `E` preenchido". Quem faz o corte
e o `carregar()`: pedido com TODOS os itens marcados sai da carteira que a analise
enxerga. Medido na base de 28/08/2026: dos 3.624 pedidos, **3.117 saem** (todos os
itens com "E") e ficam **507** -- 323 sem marca nenhuma e 184 PARCIAL.

⚠ PARCIAL FICA, E FICA INTEIRO. Pedido com parte dos itens encerrada ainda tem
item em aberto: ele NAO esta encerrado, entao entra. E entra com todos os itens
somados, inclusive os encerrados -- as colunas do painel leem "X a classificar de
Y do pedido", e jogar fora metade dos itens faria o Y mentir sobre o tamanho do
pedido.

⚠ DUAS COISAS NAO PODEM SAIR JUNTO COM ELES, e as duas custam caro se saem:
  1. **a FAIXA de numeros de PC** (`menor_pc`/`maior_pc`) continua vindo da SC7
     INTEIRA. E ela que diz se um numero citado numa nota tem chance de ser
     pedido nosso (ver `pedidos_sefaz`); calculada so sobre os abertos, ela
     encolheria e todo pedido antigo citado numa nota viraria "acima da faixa",
     isto e, pedido do fornecedor.
  2. **os pedidos cortados** (`encerrados`), porque "nao esta na carteira" e "nao
     existe" sao coisas diferentes. Sem eles, a nota que cita um pedido encerrado
     sairia como PEDIDO FORA DA BASE -- cujo texto diz "nao existe na base de
     pedidos (SC7) -- pedido cancelado, de outra filial ou exportacao incompleta".
     Seriam 326 notas mandando alguem procurar um pedido que esta na base,
     inteiro, so fechado.
     ⚠ Sao os `Pedido` INTEIROS, e nao so os numeros: metade dos casos nao cita
     numero nenhum -- o pedido era achado por fornecedor + valor --, e para dizer
     "existe um encerrado deste fornecedor com este valor" o `pedidos_sefaz`
     precisa comparar valor, o que a lista de numeros nao permite.

⚠ COLUNA INSERIDA: O DESLOCAMENTO EM BLOCO E RECUPERADO
--------------------------------------------------------
24/08/2026 a exportacao ganhou UMA coluna nova e TUDO andou uma casa para a
direita (`Numero PC` saiu de AI para AJ, `Vlr.Total` de V para W...). Pela regra
acima a leitura parava, e a aba de pedidos ficava sem atualizar -- calada, porque
o gerar_painel so mostra um AVISO nesse caso.

O `_posicoes()` agora resolve isso: se TODOS os nomes declarados aparecem juntos
a uma mesma distancia (ate DESLOCAMENTO_MAXIMO colunas), o deslocamento e aplicado em bloco e
sai um aviso dizendo quanto andou. Nao e "ler por nome": exige que o bloco INTEIRO
bata deslocado do mesmo tanto -- coluna trocada de lugar sozinha, ou nome que
sumiu, continua sendo ERRO. E o que separa "inseriram uma coluna" (rotina, dá para
seguir) de "a planilha virou outra" (parar e olhar).
"""

from __future__ import annotations

import datetime as dt
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.utils import column_index_from_string

import sefaz

# O painel roda em DUAS maquinas: esta pasta e a mesma biblioteca do OneDrive,
# montada com um nome diferente em cada uma. Este arquivo mora DENTRO dela
# (...\LUCAS ABNER ARAUJO\<AUTOMACOES LUCAS>\ANALISES BOLETOS\PAINEL ANALISE
# BOLETOS\sc7.py), entao quatro niveis acima esta a raiz -- e isso vale nos dois
# PCs. Ao contrario das outras bases, esta aqui nao tem argumento de linha de
# comando no gerar_painel.py: o .cmd nao teria como corrigir o caminho, e na
# outra maquina a aba de pedidos de compra nascia vazia, so com um AVISO.
# O caminho fixo fica de ultimo recurso, para quem tirar o script da biblioteca.
BASE_GENERICOS = Path(__file__).resolve().parents[3] / "BASES GENERICOS"
if not BASE_GENERICOS.is_dir():
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
    # ⚠ A filial faz parte da CHAVE do pedido (ver a docstring): sem ela o
    # solicitante do 001522 da Logistica vazava para o 001522 da Capivara.
    ("A", "Filial",       "filial"),
    ("Z", "Num. Pedido",  "pc"),
    ("Y", "Solicitante",  "solicitante"),
    ("B", "Numero da SC", "sc"),
    ("DI", "Usuário",     "usuario"),
]

JUNTA = " · "

# Ate quantas colunas de deslocamento em bloco aceitar antes de desistir. Uma ou
# duas colunas inseridas e rotina de exportacao; dez e outra planilha.
DESLOCAMENTO_MAXIMO = 6

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


def normalizar_filial(valor) -> str:
    """A filial como a SC7 a escreve -- texto, sem espacos ("004001").

    ⚠ NAO passa por `normalizar_pc`: os zeros a esquerda da filial sao o nome
    dela, e "004001" sem eles ("4001") nao existe em lugar nenhum do TOTVS.
    """
    return str(valor or "").strip()


def chave_pedido(filial, numero) -> str:
    """A IDENTIDADE de um pedido: filial + numero, nunca o numero sozinho.

    ⚠ E a peca central deste modulo (ver a docstring): o mesmo numero de pedido
    existe em varias filiais, com fornecedor, comprador, valor e solicitante
    diferentes em cada uma. Indexar pelo numero cola todos eles num pedido que
    nao existe.

    O separador "|" nao aparece em filial nem em numero (as duas coisas sao so
    digito), entao a chave nunca fica ambigua.
    """
    return f"{normalizar_filial(filial)}|{normalizar_pc(numero)}"


@dataclass
class Pedido:
    """Um PC inteiro, com os itens ja somados/juntados.

    Cada campo sobe de nivel do jeito que faz sentido para ELE -- esta e a parte
    do modulo que nao da para deduzir olhando a planilha.
    """
    numero: str                                   # normalizado ("1522")
    # ⚠ A filial e parte da CHAVE, e por isso e um campo e nao uma lista de
    # distintos como as de baixo: todos os itens deste objeto sao, por
    # construcao, da mesma filial. Ver `chave_pedido()`.
    filial: str = ""
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

        # ⚠ `filial` NAO entra aqui: ela e a chave, nao um distinto a juntar.
        for campo, lista in (("sc", self.scs), ("comprador", self.compradores),
                             ("usuario_sc", self.usuarios_sc),
                             ("controle", self.controles),
                             ("encerrado", self.encerrados),
                             ("razao", self.razoes)):
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
    def chave(self) -> str:
        """Como este pedido e indexado em todo lugar. Ver `chave_pedido()`."""
        return chave_pedido(self.filial, self.numero)

    @property
    def rotulo(self) -> str:
        """Como o pedido se identifica para gente: "004001/001522".

        ⚠ Todo texto de tela que NOMEIA um pedido usa isto, e nao
        `numero_totvs`: "confira o pedido 001522" manda a pessoa procurar em
        tres filiais.
        """
        return f"{self.filial}/{self.numero_totvs}" if self.filial else self.numero_totvs

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
    def aberto(self) -> bool:
        """O pedido NAO esta encerrado -- e o corte de 28/08/2026 (ver docstring).

        ⚠ A pergunta e feita na CONTAGEM de itens marcados, e nao no texto que a
        propriedade `encerrado` devolve. Os dois concordam hoje (a coluna AU so
        tem "E" e vazio nas 26.377 linhas), mas o texto e um relatorio -- passa
        por `PARCIAL` e por `JUNTA.join()` -- e uma marca nova na planilha ("F",
        "ENC") mudaria o texto sem mudar a contagem. Item em aberto e a definicao
        de pedido em aberto; o resto e como isso se escreve na tela.
        """
        return self.itens_encerrados < self.itens

    @property
    def razao(self) -> str:
        return JUNTA.join(self.razoes)

    @property
    def emissao(self) -> dt.date | None:
        return min(self.emissoes) if self.emissoes else None


def _posicoes(cabecalho: list, colunas: list, onde: str) -> list[tuple]:
    """[(indice, apelido)] -- as posicoes REAIS das colunas declaradas.

    Tenta, nesta ordem:
      1. a letra declarada (o caso normal: o nome bate onde deveria);
      2. o mesmo nome PERTO dali (ate DESLOCAMENTO_MAXIMO colunas), quando alguem
         inseriu/removeu coluna na exportacao;
      3. desiste e levanta o erro de sempre, dizendo o que estava em cada letra.

    ⚠ A COLUNA NOVA ENTRA NO MEIO, entao o deslocamento e PARCIAL: em 24/08/2026
    `Filial` (A), `Razão Social` (B) e `Quantidade` (I) ficaram onde estavam e
    todas as de `Dt. Entrega` (L) em diante andaram uma casa. Por isso o teste nao
    e "todas deslocadas do mesmo tanto", e sim, por coluna, o nome achado numa
    JANELA em volta da posicao declarada.

    ⚠ Duas travas seguram o passo 2, para ele nao virar "ler por nome" -- que e
    justamente o bug que a leitura por posicao existe para evitar (a SC7 tem 3
    nomes repetidos, e o repetido escolhido calado da coluna errada):
      - o nome tem de ser UNICO dentro da janela; dois candidatos = erro;
      - os deslocamentos, na ordem das colunas, so podem CRESCER (nao decrescer).
        Coluna inserida empurra tudo que vem depois; se uma coluna "andou para
        tras" enquanto a vizinha andou para frente, nao foi insercao -- e outra
        planilha, e ai para.
    """
    rotulos = [sefaz._rotulo(c) for c in cabecalho]
    declaradas = [(column_index_from_string(letra) - 1, sefaz._rotulo(nome), apelido)
                  for letra, nome, apelido in colunas]

    achadas = []          # (indice_declarado, indice_real, apelido)
    for i, esperado, apelido in declaradas:
        if i < len(rotulos) and rotulos[i] == esperado:
            achadas.append((i, i, apelido))
            continue
        janela = [j for j in range(max(0, i - DESLOCAMENTO_MAXIMO),
                                   min(len(rotulos), i + DESLOCAMENTO_MAXIMO + 1))
                  if rotulos[j] == esperado]
        if len(janela) != 1:
            # ambiguo ou ausente: erro de sempre, com o diagnostico por letra
            sefaz._conferir(cabecalho, colunas, onde)
        achadas.append((i, janela[0], apelido))

    # deslocamentos so podem crescer na ordem das colunas (ver docstring)
    ordenadas = sorted(achadas)
    anterior = None
    for declarado, real, _apelido in ordenadas:
        desloc = real - declarado
        if anterior is not None and desloc < anterior:
            sefaz._conferir(cabecalho, colunas, onde)
        anterior = desloc

    movidas = [(d, r) for d, r, _a in achadas if d != r]
    if movidas:
        maior = max(abs(r - d) for d, r in movidas)
        print(f"AVISO: {onde} mudou de layout -- {len(movidas)} de {len(declaradas)} "
              f"colunas andaram ate {maior} casa(s) (coluna inserida/removida na "
              f"exportacao). Foram reencontradas pelo nome e a leitura seguiu.\n"
              f"       Se isso virar o novo normal, corrija as letras em "
              f"COLUNAS_SC7/COLUNAS_SC1 no sc7.py.", file=sys.stderr)
    return [(real, apelido) for _d, real, apelido in achadas]


def _ler(caminho: Path, colunas: list) -> list[dict]:
    """Uma planilha de coluna larga, lida por posicao e conferida por nome."""
    wb, temporaria = sefaz._abrir(caminho)
    try:
        ws = wb[wb.sheetnames[0]]
        brutas = ws.iter_rows(values_only=True)
        cabecalho = list(next(brutas, ()) or ())
        posicoes = _posicoes(cabecalho, colunas, f"{caminho.name}/{ws.title}")
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
    """{chave_pedido(): Pedido} -- a SC7 inteira, agregada por FILIAL + numero.

    ⚠ A chave NAO e o numero do pedido (ver a docstring do modulo): 2.541 dos
    3.669 numeros existem em mais de uma filial, e agrupa-los pelo numero cria um
    pedido que nao existe -- com o comprador de uma filial, o fornecedor de outra
    e o valor das tres somado.
    """
    linhas = _ler(caminho or BASE_SC7, COLUNAS_SC7)
    pedidos: dict[str, Pedido] = {}
    for linha in linhas:
        numero = normalizar_pc(linha.get("pc"))
        if not numero:
            continue
        filial = normalizar_filial(linha.get("filial"))
        pedidos.setdefault(chave_pedido(filial, numero),
                           Pedido(numero=numero, filial=filial)).somar(linha)
    return pedidos


def ler_solicitantes(caminho: Path | None = None) -> dict[str, list[str]]:
    """{chave_pedido(): solicitantes distintos} -- vem da SC1.

    ⚠ Um PC pode nascer de varias SCs, logo de varios solicitantes (21 dos 149
    pedidos casados). Devolve TODOS, na ordem em que aparecem.

    ⚠ A chave e a MESMA do `ler_pedidos()` -- filial + numero. Casando so pelo
    numero, 2.031 pedidos recebiam o solicitante de outra filial, e a coluna
    parecia preenchida (5.710 pedidos "com solicitante" em vez de 3.679): errado
    com cara de completo, que e o pior jeito de errar.
    """
    linhas = _ler(caminho or BASE_SC1, COLUNAS_SC1)
    por_pc: dict[str, list[str]] = {}
    for linha in linhas:
        numero = normalizar_pc(linha.get("pc"))
        if not numero:
            continue
        chave = chave_pedido(linha.get("filial"), numero)
        # `Usuário` (DI) e a reserva: nos exemplos vistos traz o mesmo valor, e
        # sem ela um pedido cuja SC veio sem solicitante ficaria sem ninguem.
        nome = (sefaz._texto(linha.get("solicitante"))
                or sefaz._texto(linha.get("usuario")))
        if not nome:
            continue
        lista = por_pc.setdefault(chave, [])
        if nome not in lista:
            lista.append(nome)
    return por_pc


def carregar(base_sc7: Path | None = None, base_sc1: Path | None = None
             ) -> tuple[dict[str, Pedido], dict[str, list[str]],
                        dict[str, Pedido], dict]:
    """(pedidos EM ABERTO, solicitantes por pedido, pedidos ENCERRADOS, resumo).

    ⚠ Os tres dicionarios sao indexados por `chave_pedido()` (filial + numero),
    e nao pelo numero do pedido. Quem receber uma `chave` e for mostra-la para
    gente tem de usar `Pedido.rotulo`: a chave e interna ("004001|1522").

    ⚠ E AQUI que o corte de 28/08/2026 mora, e de proposito: `ler_pedidos()`
    continua devolvendo a SC7 inteira (e a leitura da planilha, nao a politica) e
    este e o unico lugar que decide o que a analise enxerga. Quem quiser a
    carteira completa -- uma conferencia, um script novo -- chama `ler_pedidos()`.

    A SC1 e OPCIONAL: sem ela os pedidos continuam valendo e so o solicitante
    fica vazio. Faltar a SC7, ao contrario, e nao ter aba nenhuma.
    """
    caminho7 = base_sc7 or BASE_SC7
    caminho1 = base_sc1 or BASE_SC1
    todos = ler_pedidos(caminho7)
    solicitantes = ler_solicitantes(caminho1) if caminho1.exists() else {}

    pedidos = {n: p for n, p in todos.items() if p.aberto}
    # Os que ficaram fora, INTEIROS. NAO sao "o que nao existe": sao o que existe
    # e esta fechado, e a diferenca e o que separa um aviso certo de 326 falsos
    # (ver a docstring do modulo).
    encerrados = {n: p for n, p in todos.items() if not p.aberto}
    parciais = sum(1 for p in pedidos.values() if p.encerrado == PARCIAL)

    # ⚠ A FAIXA SAI DA BASE INTEIRA (`todos`), nunca dos abertos -- ver o item 1
    # da docstring. Calculada sobre 507 pedidos em vez de 3.624, ela desceria de
    # 3.6xx para o maior PC aberto e transformaria pedido antigo citado numa nota
    # em "numero acima da faixa", isto e, pedido do fornecedor.
    numeros = [int(p.numero) for p in todos.values() if p.numero.isdigit()]
    resumo = {
        # `pedidos` = os que a analise usa (abertos). O total da base vem ao lado,
        # senao o numero da tela pareceria uma base que encolheu.
        "pedidos": len(pedidos),
        "pedidos_na_base": len(todos),
        # quantos NUMEROS distintos os pedidos da base usam: e menor que o total
        # de pedidos justamente porque o numero se repete entre filiais
        "numeros_na_base": len({p.numero for p in todos.values()}),
        "filiais": len({p.filial for p in todos.values() if p.filial}),
        "encerrados_fora": len(encerrados),
        "parciais": parciais,
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
    return pedidos, solicitantes, encerrados, resumo
