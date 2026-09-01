# -*- coding: utf-8 -*-
"""Gera o painel HTML da aba 'Associacoes Encontradas'.

Le TRATAMENTO PYTHON BOLETOS.xlsx e escreve um index.html unico, sem
dependencias externas: a mesma planilha, com as mesmas colunas e na mesma
ordem. O painel so acrescenta o que a planilha nao consegue mostrar -- o
destaque quando o boleto diverge do titulo e o botao de copiar UUID e linha
digitavel.

Uso:
    python gerar_painel.py [--base CAMINHO] [--saida CAMINHO]
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import shutil
import sys
import tempfile
import unicodedata
from pathlib import Path

from openpyxl import load_workbook

import cruzamento_classificacao
import cruzamento_sefaz_sf1
import manifestacao
import pedidos_sefaz
import pendentes_itau
import planilha_excel_js
import sc7
import sefaz
import verificacao_se2

# Motor .xlsx do painel SE2, reaproveitado aqui (pedido do usuario: "use o
# modelo de planilha que e gerado na base desse painel"). A FONTE e a copia do
# SE2; esta aqui e vendorizada para o painel nao depender de uma pasta em
# Downloads. `conferir_motor_planilha()` avisa quando as duas divergem.
MOTOR_PLANILHA_FONTE = Path(r"C:\Users\lucas\Downloads\se2 - sistema\planilha_excel_js.py")

BASE_PADRAO = Path(
    r"C:\Users\lucas\OneDrive - Grupo S&D\Arquivos de Gabriella Karla Oliveira Milas - FINANCEIRO COMPARTILHADO"
    r"\LUCAS ABNER ARAUJO\TRATAMENTO PYTHON BOLETOS.xlsx"
)
PASTA = Path(__file__).resolve().parent
SAIDA_PADRAO = PASTA / "PUBLICAR" / "index.html"
MODELO = PASTA / "painel_modelo.html"
# Fora do repositorio (.gitignore): a carteira para subir e o painel de uso local.
DADOS_JSON = PASTA / "DADOS" / "analise_boletos.json"
DEV = PASTA / "dev.html"

# A aba de pendentes do Itau vem de OUTRA planilha (a do DDA, copiada e colada)
# e de um historico proprio -- ver pendentes_itau.py. O historico e o que faz o
# boleto entrar uma vez so; apagar esse arquivo faz todos voltarem como novos.
HISTORICO_PENDENTES = PASTA / "DADOS" / "pendentes_itau_historico.json"

# Aba "Titulos Associados na Classificacao": vem da SF1 (notas classificadas no
# TOTVS) cruzada com os boletos que ainda nao acharam titulo. Mesmo desenho de
# historico do Itau -- apagar este arquivo faz todos os titulos voltarem como
# novos, e ele tambem guarda ate que DIA a SF1 ja foi lida.
HISTORICO_CRUZAMENTO = PASTA / "DADOS" / "cruzamento_classificacao_historico.json"

# A base da SEFAZ e ACUMULATIVA (traz as antigas mais as novas), entao nota que
# ESTAVA e deixou de vir e anomalia -- ver sefaz.py. Este arquivo e quem lembra
# quais notas o painel ja viu; apagar faz o painel esquecer todas e parar de
# acusar as removidas (as novas nao sao afetadas: elas vem da base).
HISTORICO_SEFAZ = PASTA / "DADOS" / "sefaz_historico.json"

ABA_ASSOCIACOES = "Associações Encontradas"
ABA_CORRESPONDENCIAS = "Tratar Correspondências"
ABA_FUTUROS = "Futuros NF Não Associada"
ABA_RESUMO = "Resumo"
ABA_PARCELAS = "NFs Múltiplas Parcelas"

# So a linha digitavel entra; o codigo de barras fica de fora (pedido do usuario).
IGNORAR = {"Código de Barras"}

# Coluna de marcacao: fica no painel mesmo vazia na planilha, virando caixa de
# selecao. O que ja estiver preenchido na base entra marcado.
COLUNA_CHECK = "CHECK (FEITO)"

# Fontes que sao a MESMA coisa para quem usa o painel. Os tres CENTRAL saem do
# mesmo lugar -- o e-mail da central de boletos --, mudando so quem recebeu.
# Consolidados a pedido do usuario (13/08/2026): tres botoes de filtro para uma
# origem so obrigavam a clicar nos tres para ver o mesmo conjunto.
# ⚠ Isto troca o TEXTO da celula, entao vale para o botao de filtro, para a
# tela e para o .xlsx exportado. A busca continua achando pelo nome antigo: ela
# e montada com os valores CRUS da base, de proposito.
COLUNA_FONTE = "Fonte Boleto"
FONTES = {
    "CENTRAL_LUCAS": "EMAIL CENTRAL DE BOLETOS",
    "CENTRAL_GRAZIELA": "EMAIL CENTRAL DE BOLETOS",
    "CENTRAL_BH": "EMAIL CENTRAL DE BOLETOS",
}

MOEDA = {"Vlr.Titulo", "Valor Boleto", "Saldo", "Desconto", "Multa", "Juros",
         "Correcao", "Val Liq Baix",
         # nomes usados na aba Tratar Correspondencias
         "Valor Título (R$)", "Valor Boleto (R$)", "Diferença (R$)",
         # aba Futuros NF Nao Associada
         "Valor (R$)",
         # aba NFs pendentes Itau
         "Até Venc.", "A Pagar",
         # aba Titulos Associados na Classificacao. ⚠ "Vlr.Título" tem ACENTO e
         # e outra chave: o "Vlr.Titulo" acima e o nome da coluna na base do
         # painel. Faltando aqui, o valor sai cru ("29307.68") no lugar de R$.
         "Vlr.Título",
         # abas da SEFAZ
         "Vlr Item", "Vlr Total", "Vlr Duplicata", "Vlr SEFAZ",
         # aba NF x Pedido de Compra: o total do PC (soma dos itens dele)
         "Vlr Pedido"}
DATA = {"Vencimento", "Vencto Real", "Vencimento Boleto", "DT Emissao",
        "DT Baixa", "DT Contab.",
        # aba NFs pendentes Itau
        "Entrou em", "Visto em",
        # aba Titulos Associados na Classificacao
        "Classificado em", "Emissão",
        # abas da SEFAZ
        "Saída/Entrada", "Venc Duplicata", "Emissão SF1",
        # aba NF x Pedido de Compra: datas do PEDIDO (a menor de cada uma)
        "Dt. Entrega", "1° Venc"}
INTEIRO = {"Score Match", "2º Score", "Margem", "Score",
           # aba SEFAZ x SF1: quantas duplicatas a nota tem (vazio = uma so)
           "Duplicatas",
           # aba NF x Pedido de Compra
           "Itens do PC"}

# Quantidade de mercadoria do pedido (SC7). ⚠ NAO entra em INTEIRO: medido na
# base, 889 dos 24.516 itens tem quantidade quebrada (0,954 / 1,4 / 0,062) e 316
# dos 3.382 pedidos somam quebrado -- o `int()` do tipo inteiro truncaria 0,954
# para 0 sem erro nenhum aparecer. Sai formatada em pt-BR e ordena por numero.
# `Qtd.a Classi` veio junto (era inteiro) porque as duas ficam LADO A LADO: uma
# truncar e a outra nao faria a comparacao mentir na coluna do meio.
QUANTIDADE = {"Quantidade", "Qtd.a Classi"}

# Colunas com botao de copiar (o usuario cola no SE2 / no banco). O rotulo so
# aparece nas que sao SO_BOTAO; nas outras o proprio valor e o botao.
COPIAVEIS = {"Campo UUID": "UUID", "Linha Digitável": "LINHA",
             "Fornecedor": "", "No. Titulo": "",
             # 44 digitos: ninguem le, so cola (no TOTVS, no portal da SEFAZ)
             "Chave NF-e": "CHAVE"}

# Dessas o valor nem aparece: a coluna e so o botao de copiar (sao longas demais
# e o usuario nunca le, so cola).
SO_BOTAO = {"Campo UUID", "Linha Digitável", "Chave NF-e"}

# Tipos de alerta que ganham selo proprio na coluna Alerta. A ordem importa:
# o texto de fatura tambem fala em "parcelas", entao FATURA e testado primeiro.
# ALTERADO e o aviso do proprio banco na aba de pendentes do Itau ("este boleto
# sofreu alteracao/instrucao por parte do beneficiario") -- sem ele o selo sairia
# como "ALERTA" generico, que nao diz nada.
ALERTAS = ((("FATURA",), "FATURA"),
           (("MAIS DE UMA PARCELA", "OUTRAS PARCELAS", "PARCELA"), "PARCELA"),
           (("ALTERAÇÃO", "ALTERACAO", "INSTRUÇÃO", "INSTRUCAO"), "ALTERADO"),
           # Alertas das abas da SEFAZ (13/08/2026). ⚠ A ORDEM decide qual selo
           # aparece quando a nota cai em mais de uma condicao -- e por isso vao
           # do mais urgente para o menos. O texto de TODOS continua no quadro
           # que abre no selo, nenhum se perde.
           # ⚠ ANTES dos outros tres: a nota removida pode estar tambem atrasada,
           # e das duas coisas e esta que muda o que fazer -- as outras dizem que
           # a nota esta parada, esta diz que ela sumiu da base. O texto das
           # demais continua inteiro no quadro que abre no selo.
           ((sefaz.PREFIXO_REMOVIDA,), "REMOVIDA DA BASE SEFAZ"),
           # 01/09/2026. Depois da removida e antes dos de prazo: os de prazo
           # dizem que a nota esta parada, este diz que talvez ela nao devesse
           # ser paga -- a empresa manifestou desconhecimento ou nao realizada.
           ((sefaz.PREFIXO_MANIFESTO,), "NÃO RECONHECIDA"),
           ((sefaz.PREFIXO_VENCE,), "VENCE ≤ 5D"),
           # 14/08/2026, pedido do usuario: era "NÃO LANÇADA", que dizia o efeito
           # mas nao o que o alerta cobra -- o tempo parado. O numero sai da
           # constante da regra, e nao escrito aqui: mudar o corte em `sefaz.py`
           # mudaria o criterio e deixaria o selo mentindo.
           ((sefaz.PREFIXO_SEM_LANCAR,), f"PENDENTE A MAIS DE {sefaz.DIAS_SEM_LANCAR} DIAS"),
           ((sefaz.PREFIXO_PRAZO,), "PRAZO CURTO"))

# Confrontos que viram destaque na celula do boleto.
#   coluna destacada -> (coluna do titulo, tipo)
CONFRONTOS = {
    "Valor Boleto": ("Vlr.Titulo", "moeda"),
    "Vencimento Boleto": ("Vencimento", "data"),
}

# Visao Conferencia (unica visao do painel): subconjunto das colunas NA MESMA
# ORDEM DA BASE. Como a base ja poe titulo e boleto lado a lado, a unica coisa
# acrescentada e a coluna de diferenca (@delta_*), logo depois do par.
#   (rotulo, lado, grupo, largura relativa)
# Empresa pagadora e contraparte NAO entram aqui: elas sao colunas do drill
# (o usuario corrigiu isso em 04/08/2026). A tabela principal segue enxuta.
# ------------------------------------------------------- fusao SE2 x SF1
# 14/08/2026, pedido do usuario: "as analises sao as mesmas, iremos levar a aba
# titulos associados na classificacao para junto da aba titulos associados". As
# NFs recem classificadas passaram a viver DENTRO da aba de titulos associados,
# para a conferencia acontecer num lugar so. Ver `fundir_classificacao`.
COLUNA_ORIGEM = "Origem"
# ⚠ Nomes do NEGOCIO, nao das tabelas do TOTVS (14/08/2026, pedido do usuario:
# "SE2 = CONTAS A PAGAR, SF1 = CLASSIFICAÇÃO"). Trocar aqui basta: e o TEXTO da
# celula, entao o botao de filtro, a grade e o .xlsx exportado mudam juntos.
# ⚠ O navegador NAO depende destes textos para saber de que lado a linha veio --
# ele olha o prefixo "SF1:" da chave (ver `daClassificacao` no painel_modelo).
# Foi de proposito: rotulo e coisa que o usuario renomeia, chave nao.
ORIGEM_SE2 = "CONTAS A PAGAR"
ORIGEM_SF1 = "CLASSIFICAÇÃO"

# ⚠ Coluna interna (nao desenhada): guarda a chave que amarra check e tratativa
# ao registro. NAO da para usar o "Campo UUID" para isso depois da fusao -- as
# linhas da SF1 se identificam pela chave com prefixo "SF1:", e trocar por o
# uuid cru desligaria as marcacoes que ja existem no banco. O Campo UUID, por
# sua vez, precisa continuar cru: e ele que o botao de copiar leva ao TOTVS.
COLUNA_CHAVE_PAINEL = "Chave Painel"

# Os dois lados dizem a mesma coisa com nomes diferentes. Sem este mapa a fusao
# criaria colunas GEMEAS -- "Status" vazio nas linhas da SF1 e "Situação" vazio
# nas da SE2 -- e nenhum filtro, ordenacao ou selo pegaria as duas metades.
SF1_PARA_BASE = {
    "Nº NF": "No. Titulo",
    "Vlr.Título": "Vlr.Titulo",    # ⚠ o acento e a UNICA diferenca entre os dois
    "Situação": "Status",
    "Critério": "Critério Match",
    "Emissão": "DT Emissao",
}

COMPACTA = [
    ("CHECK (FEITO)",     "",       "",           2.5),
    # De qual base a linha veio (14/08/2026, pedido do usuario). A aba passou a
    # somar os titulos ja associados (SE2) com as NFs recem classificadas (SF1);
    # sem esta coluna as duas metades ficariam indistinguiveis na mesma grade --
    # e elas nao querem dizer a mesma coisa (uma tem titulo, a outra ainda nao).
    # ⚠ Largura medida contra os valores REAIS: eram "SE2"/"SF1" e cabiam em
    # 3.5, mas em 14/08/2026 viraram "CONTAS A PAGAR"/"CLASSIFICAÇÃO" (14
    # caracteres). Na largura antiga sairia "CONTAS A …" -- a celula corta com
    # ellipsis e NAO acusa erro nenhum.
    # ⚠ 9 -> 11.5 em 14/08/2026: a celula passou a levar tambem o selo do elo da
    # chave (o mesmo titulo dos dois lados, ver `aglutinar()` no modelo). O selo
    # ocupa ~26px; nos 9 pontos de antes ele empurraria o texto e sairia
    # "CONTAS A P…" -- de novo sem erro nenhum.
    ("Origem",            "",       "",           11.5),
    ("Campo UUID",        "",       "",           4),
    ("Filial",            "",       "",           3.5),
    ("Prefixo",           "",       "",           3.5),
    # A NF do boleto sai do lugar dela na base para ficar colada na do titulo
    # (pedido do usuario): e essa comparacao que ele faz o tempo todo.
    ("No. Titulo",        "titulo", "NF",         6.5),
    ("NF/Doc Boleto",     "boleto", "NF",         6),
    ("Parcela",           "titulo", "",           3),
    ("Fornecedor",        "titulo", "",           5),
    ("Razão Social",      "titulo", "",          13),
    ("Vlr.Titulo",        "titulo", "VALOR",      6.5),
    ("Valor Boleto",      "boleto", "VALOR",      6.5),
    ("@delta_valor",      "delta",  "VALOR",      7.5),
    ("Vencimento",        "titulo", "VENCIMENTO", 5.5),
    ("Vencto Real",       "titulo", "VENCIMENTO", 5.5),
    ("Vencimento Boleto", "boleto", "VENCIMENTO", 5.5),
    ("@delta_dias",       "delta",  "VENCIMENTO", 4.5),
    # Tipo de pagamento (13/08/2026, pedido do usuario). Tres colunas coladas,
    # com a mesma logica dos pares titulo x boleto do resto da aba: o que a SE2
    # tem gravado (coluna IK), o que a pessoa informa que foi de verdade, e o
    # alerta quando os dois nao batem. ⚠ Nenhuma das tres vem da base do painel:
    # a primeira vem da SE2 (pelo `_meta` da aba) e as outras duas nascem no
    # navegador -- por isso todas sao virtuais (@).
    # ⚠ Largura medida contra os valores REAIS do campo: "BOLETO S/C" e o mais
    # comum desta aba e nao cabia em 5 -- saia "BOLETO S...", que nao distingue
    # de "BOLETO S/C" nenhum outro valor, mas obriga a passar o mouse.
    ("@tipo_pgto",        "titulo", "PAGAMENTO",  6),
    ("@tipo_real",        "boleto", "PAGAMENTO",  6),
    ("@alerta_tipo",      "delta",  "PAGAMENTO",  5),
    ("Status",            "",       "",           7),
    ("Alerta Fatura",     "",       "",           5),
    ("@tratativa",        "",       "",           4),
    ("Linha Digitável",   "",       "",           4),
]

# Aba "Tratar Correspondencias": MESMAS funcoes, colunas adaptadas ao que ela
# tem. Nao existe `Status` ali -- o que manda e o `Motivo Revisão`, e valor,
# diferenca e vencimento vem em colunas proprias.
COMPACTA_CORRESP = [
    ("CHECK (FEITO)",       "",       "",           2.5),
    ("Campo UUID",          "",       "",           4),
    ("Filial",              "",       "",           3.5),
    ("Prefixo",             "",       "",           3.5),
    ("No. Titulo",          "titulo", "NF",         6.5),
    ("NF/Doc Boleto",       "boleto", "NF",         6),
    ("Parcela",             "titulo", "",           3),
    ("Fornecedor",          "titulo", "",           5),
    ("Razão Social",        "titulo", "",          12),
    ("Fornecedor Boleto",   "boleto", "",          11),
    ("Valor Título (R$)",   "titulo", "VALOR",      6.5),
    ("Valor Boleto (R$)",   "boleto", "VALOR",      6.5),
    ("@delta_valor",        "delta",  "VALOR",      7.5),
    ("Vencimento",          "titulo", "VENCIMENTO", 5.5),
    ("Vencto Real",         "titulo", "VENCIMENTO", 5.5),
    ("Vencimento Boleto",   "boleto", "VENCIMENTO", 5.5),
    ("@delta_dias",         "delta",  "VENCIMENTO", 4.5),
    ("Motivo Revisão",      "",       "",           7),
    ("Alerta Fatura",       "",       "",           5),
    ("@tratativa",          "",       "",           4),
    ("Linha Digitável",     "",       "",           4),
]

# Aba "Futuros NF Nao Associada": e o outro lado da moeda -- boletos do DDA que
# NAO acharam titulo. Nao ha o que confrontar (nao existe titulo), entao ela nao
# tem Δ, nem par valor/vencimento, nem Campo UUID.
COMPACTA_FUTUROS = [
    ("CHECK (FEITO)",       "",       "",           3),
    ("Situação",            "",       "",           9),
    ("NF/Doc Original",     "boleto", "NF",         7),
    ("NF Normalizada",      "boleto", "NF",         7),
    ("Parcela",             "boleto", "NF",         4),
    ("Total Parcelas",      "boleto", "NF",         4.5),
    ("Fornecedor",          "boleto", "CONTRAPARTE", 15),
    ("CNPJ/CPF",            "boleto", "CONTRAPARTE", 9),
    ("Valor (R$)",          "boleto", "",           8),
    ("Vencimento",          "boleto", "",           7),
    ("Empresa/Origem",      "boleto", "",          12),
    ("Fonte Boleto",        "boleto", "",           7),
    ("Arquivo/Observação",  "boleto", "",          12),
    ("@tratativa",          "",       "",           4.5),
    ("Linha Digitável",     "",       "",           4.5),
]

# Aba "Boletos em Aberto DDA": vem da planilha do DDA (copiar e colar), nao da base
# do resto do painel. Nao ha titulo do TOTVS para confrontar; o que se compara e
# o valor ATE o vencimento contra o valor A PAGAR hoje -- a diferenca e o
# juros/multa que ja correu, e e isso que a coluna Δ mostra aqui.
COMPACTA_PENDENTES = [
    ("CHECK (FEITO)",    "",       "",           3),
    # As duas colunas de selo sao mais largas que nas outras abas de proposito:
    # "SAIU DA BASE" e "ALTERADO" sao textos maiores que "MATCH FORTE"/"PARCELA"
    # e, apertados, saiam como "SAIU D..." / "ALTE..." -- selo cortado nao serve
    # para nada, que o ponto dele e ser lido de relance.
    ("Situação",         "",       "",           9),
    ("Entrou em",        "",       "",           5.5),
    ("Pagador/Agregado", "titulo", "",          13),
    ("Beneficiário",     "boleto", "CONTRAPARTE", 14),
    ("CPF/CNPJ",         "boleto", "CONTRAPARTE", 9),
    ("Nº Doc.",          "boleto", "",           7),
    ("Vencimento",       "boleto", "",           6),
    ("Até Venc.",        "titulo", "VALOR",      7),
    ("A Pagar",          "boleto", "VALOR",      7),
    ("@delta_valor",     "delta",  "VALOR",      7),
    ("Tipo de Boleto",   "boleto", "",           9),
    ("Alerta Fatura",    "",       "",           7),
    ("@tratativa",       "",       "",           4.5),
]

# Cabecalho da EXPORTACAO desta aba. Aqui o Δ nao compara boleto com titulo --
# compara o valor ate o vencimento com o valor a pagar hoje. Sem isto o .xlsx
# sairia com um cabecalho que mente sobre a propria coluna.
ROTULOS_PLANILHA_PENDENTES = {
    "@delta_valor": "Juros/Multa já Acumulados (A Pagar - Até Venc.)",
    "Alerta Fatura": "Observações do Banco",
    "Situação": "Situação na Base do DDA",
    "Entrou em": "Entrou no Painel em",
}

# Aba "Titulos Associados na Classificacao": a NF recem-classificada no TOTVS
# (SF1) de um lado, o boleto que o cruzamento achou do outro. Segue o desenho
# das abas de associacao -- titulo e boleto lado a lado, com o Δ logo depois do
# par -- porque e a mesma conferencia que a pessoa ja faz nas outras.
COMPACTA_CRUZAMENTO = [
    ("CHECK (FEITO)",     "",       "",           3),
    # larga: "MATCH PROVÁVEL" e "SEM BOLETO" sao textos longos e selo cortado
    # nao serve para nada (a licao da aba do Itau)
    ("Situação",          "",       "",          10),
    # data por extenso + o Nº NF, que leva botao de copiar disputando a largura
    ("Classificado em",   "",       "",           7),
    # Emissao da NF na SF1 (coluna I / DT Emissao). Entrou no lugar da Filial em
    # 13/08/2026, a pedido do usuario: quando a nota foi emitida diz mais do que
    # em qual filial ela caiu. O modulo ja lia essa data; so nao a mostrava.
    ("Emissão",           "titulo", "",           6),
    ("Nº NF",             "titulo", "NF",         8),
    ("NF/Doc Boleto",     "boleto", "NF",         6.5),
    ("Razão Social",      "titulo", "CONTRAPARTE", 13),
    ("Fornecedor Boleto", "boleto", "CONTRAPARTE", 12),
    ("Vlr.Título",        "titulo", "VALOR",      6.5),
    ("Valor Boleto",      "boleto", "VALOR",      6.5),
    ("@delta_valor",      "delta",  "VALOR",      7.5),
    ("Vencimento",        "titulo", "VENCIMENTO", 5.5),
    ("Vencimento Boleto", "boleto", "VENCIMENTO", 5.5),
    ("@delta_dias",       "delta",  "VENCIMENTO", 4.5),
    ("Parcela",           "boleto", "",           4),
    ("Critério",          "",       "",           9),
    ("Fonte Boleto",      "",       "",           7),
    ("@tratativa",        "",       "",           4.5),
    ("Linha Digitável",   "",       "",           4),
    ("Campo UUID",        "",       "",           4),
]

# Aba "Analise Base SEFAZ": as colunas que o usuario pintou de verde nas duas
# abas da SEFAZ.xlsx, empilhadas numa tela so. Aqui nao ha nada para confrontar
# -- e a base como ela e, so que legivel. Ver sefaz.py.
COMPACTA_SEFAZ = [
    ("CHECK (FEITO)",     "",       "",           3),
    # o flag de qual aba da planilha a linha veio; vira selo e botao de filtro
    ("Origem",            "",       "",           5),
    # Alerta da NOTA (venc perto sem lancamento / sem lancar / prazo curto): o
    # selo mostra o rotulo curto e o texto inteiro abre no quadro.
    ("Alerta",            "",       "",          15),
    ("Nº NF",             "titulo", "NF",         4.5),
    ("Emissão",           "titulo", "",           5),
    # 28/08/2026, pedido do usuario: o status que a SEFAZ da a nota. Vem colado
    # na emissao (e leitura da NOTA, nao do pedido nem do lancamento) e virou
    # botao de filtro -- ver `pills` das abas.
    ("Status da Nota",    "titulo", "",           5.5),
    ("Saída/Entrada",     "titulo", "",           5),
    ("Tipo Operação",     "titulo", "",           4.5),
    ("CFOP",              "titulo", "",           3.5),
    ("Emitente",          "boleto", "EMITENTE",  11),
    ("CNPJ Emitente",     "boleto", "EMITENTE",   7),
    ("Nome Fantasia",     "boleto", "EMITENTE",   8.5),
    ("Destinatário",      "titulo", "DESTINATÁRIO", 9.5),
    ("CNPJ Destinatário", "titulo", "DESTINATÁRIO", 7),
    ("Vlr Item",          "titulo", "VALOR",      5.5),
    ("Vlr Total",         "boleto", "VALOR",      5.5),
    ("Venc Duplicata",    "titulo", "DUPLICATA",  5),
    ("Vlr Duplicata",     "boleto", "DUPLICATA",  5),
    # texto de ate 5.000 caracteres: na celula so cabe o botao que abre o quadro
    ("Informações",       "",       "",           3.5),
    ("@tratativa",        "",       "",           4),
    ("Chave NF-e",        "",       "",           3.5),
]

ROTULOS_PLANILHA_SEFAZ = {
    "Origem": "Aba de origem na SEFAZ.xlsx",
    "Status da Nota": "Status da Nota na SEFAZ",
    "Informações": "Informações Complementares / Descrição do Serviço",
    "Vlr Item": "Valor Total do Item (só NF-e)",
    "Vlr Total": "Valor Total da Nota",
    "Emitente": "Emitente (NF-e) / Prestador (NFS-e)",
    "Destinatário": "Destinatário (NF-e) / Tomador (NFS-e)",
    "CNPJ Emitente": "CNPJ do Emitente/Prestador",
    "CNPJ Destinatário": "CNPJ do Destinatário/Tomador",
}

# Aba "SEFAZ x SF1": a nota da SEFAZ de um lado, o titulo que o cruzamento achou
# na SF1 do outro -- mesmo desenho das outras abas de par. Ver
# cruzamento_sefaz_sf1.py. ⚠ Nada aqui encosta na SE2 (pedido do usuario).
# ⚠ 13/08/2026: entraram 8 colunas da nota (duplicata, fantasia, CFOP, tipo e
# natureza da operacao, destinatario) e sairam 3 (Emissão SF1, Campo UUID e a
# gutter "#"). As colunas da nota estavam sendo LIDAS e jogadas fora pela visao
# por nota -- ver `notas_da_sefaz`, que agora agrega cada uma do seu jeito.
COMPACTA_SEFAZ_SF1 = [
    ("CHECK (FEITO)",     "",       "",           3),
    ("Situação",          "",       "",           9.5),
    ("Alerta",            "",       "",          15),
    ("Origem",            "",       "",           4.5),
    ("Nº NF",             "titulo", "NF",         5),
    ("Nº NF SF1",         "boleto", "NF",         5.5),
    ("Emissão",           "titulo", "",           5.5),
    # 28/08/2026: entra aqui tambem para as tres abas da SEFAZ falarem a mesma
    # lingua. ⚠ Esta aba esta OCULTA -- so o cabecalho dela viaja --, entao o
    # efeito pratico e para o dia em que ela for religada.
    ("Status da Nota",    "titulo", "",           5.5),
    ("Emitente",          "titulo", "CONTRAPARTE", 12),
    ("Razão Social",      "boleto", "CONTRAPARTE", 12),
    ("CNPJ Emitente",     "titulo", "CONTRAPARTE", 7),
    ("Nome Fantasia",     "titulo", "CONTRAPARTE", 8.5),
    # O destinatario e a NOSSA empresa: e por ele que se sabe de qual filial a
    # nota e. Nome e CNPJ copiaveis em 1 clique (pedido do usuario).
    ("Destinatário",      "boleto", "DESTINATÁRIO", 10),
    ("CNPJ Destinatário", "boleto", "DESTINATÁRIO", 7),
    ("Tipo Operação",     "titulo", "OPERAÇÃO",   4.5),
    ("Nat. Operação",     "titulo", "OPERAÇÃO",   9),
    ("Saída/Entrada",     "titulo", "OPERAÇÃO",   5),
    # ⚠ CFOP e do ITEM e varia em 40 das 155 notas de NF-e: a celula traz TODOS
    # os CFOPs da nota juntos por " · ", nao o do primeiro item.
    ("CFOP",              "titulo", "OPERAÇÃO",   5),
    ("Vlr SEFAZ",         "titulo", "VALOR",      6),
    ("Vlr.Título",        "boleto", "VALOR",      6),
    ("@delta_valor",      "delta",  "VALOR",      6),
    # A duplicata que vence ANTES; `Duplicatas` so mostra numero quando ha mais
    # de uma -- senao a linha prometeria parcela unica onde ha 3.
    ("Venc Duplicata",    "titulo", "DUPLICATA",  5),
    ("Vlr Duplicata",     "boleto", "DUPLICATA",  5.5),
    ("Duplicatas",        "delta",  "DUPLICATA",  3.5),
    ("Série",             "boleto", "",           3.5),
    ("Classificado em",   "boleto", "",           5.5),
    ("Títulos",           "",       "",           3.5),
    # ⚠ O Criterio NAO tem coluna propria aqui: ele ja abre no quadro do selo de
    # Situação (o mesmo desenho do MATCH FORTE das outras abas). Repetir a
    # coluna custava 78 KB de carteira para mostrar duas vezes a mesma frase.
    ("@tratativa",        "",       "",           4),
    ("Chave NF-e",        "",       "",           3.5),
]

ROTULOS_PLANILHA_SEFAZ_SF1 = {
    "Situação": "A nota da SEFAZ está lançada na SF1?",
    "Nº NF": "Nº da NF (SEFAZ)",
    "Nº NF SF1": "Nº da NF (SF1/TOTVS)",
    "Emissão": "Emissão (SEFAZ)",
    "Emitente": "Emitente/Prestador (SEFAZ)",
    "Nome Fantasia": "Nome Fantasia do Emitente",
    "Razão Social": "Razão Social do Fornecedor (SF1)",
    "Destinatário": "Destinatário/Tomador (a nossa empresa)",
    "CNPJ Destinatário": "CNPJ do Destinatário/Tomador",
    "Tipo Operação": "Tipo de Operação (entrada/saída)",
    "Nat. Operação": "Natureza da Operação (só NF-e)",
    "Saída/Entrada": "Data de Saída/Entrada",
    "CFOP": "CFOP dos itens da nota",
    "Vlr SEFAZ": "Valor Total da Nota (SEFAZ)",
    "Venc Duplicata": "Vencimento da 1ª Duplicata",
    "Vlr Duplicata": "Valor da 1ª Duplicata",
    "Duplicatas": "Quantas duplicatas a nota tem (vazio = uma só)",
    "Títulos": "Quantos títulos da SF1 têm este mesmo nº de NF",
    "Critério": "O que bateu no cruzamento",
    "Classificado em": "Data de Classificação no TOTVS (Dt.Digitacao)",
}

# Aba "NF x Pedido de Compra" (13/08/2026): a nota da SEFAZ, o lancamento dela no
# TOTVS e o PEDIDO que a originou, tudo na mesma linha. E um superconjunto da
# SEFAZ x SF1 -- as duas convivem de proposito: aquela responde "ja foi
# lancada?", esta responde "de qual pedido veio?". Ver pedidos_sefaz.py.
COMPACTA_PEDIDOS = [
    ("CHECK (FEITO)",     "",       "",             3),
    # A chave NF-e vem LOGO NO COMECO (pedido do usuario, 13/08/2026): ela e so
    # o botao de copiar (SO_BOTAO, 44 digitos que ninguem le) e e a primeira
    # coisa que se cola no TOTVS/portal da SEFAZ. No fim da linha obrigava a
    # atravessar 36 colunas para chegar nela.
    ("Chave NF-e",        "",       "",             3.5),
    ("Situação",          "",       "",            10),
    # ⚠ O MESMO `Numero PC` aparece DUAS vezes de propósito (pedido do usuario,
    # 13/08/2026: "coloque o numero do pedido depois da coluna situação
    # também"). Aqui ele fica colado no veredito -- que e a leitura de sempre,
    # "qual pedido e o quanto confiar" -- e continua no bloco PEDIDO, junto do
    # resto do pedido. Duas colunas com a mesma `chave` sao inofensivas: a
    # celula e a mesma e ordenar por qualquer uma das duas ordena igual.
    ("Numero PC",         "",       "",             5),
    # ⚠ 15 pontos (≈187px) porque o selo mais longo desta coluna passou a ser
    # "PENDENTE A MAIS DE 3 DIAS": `.selo` e nowrap + ellipsis, entao com os 6,5
    # de antes ele sairia cortado ("PENDENTE A M…") sem erro nenhum aparecer.
    ("Alerta",            "",       "",            15),
    # o veredito da OUTRA aba, de proposito: "achei o pedido MAS a nota nunca
    # foi lancada" e justamente o caso que esta aba serve para achar
    ("Lançada na SF1",    "",       "",             6.5),
    ("Origem",            "",       "",             4),
    ("Nº NF",             "titulo", "NF",           4.5),
    # a NF com zeros a frente ate 9 digitos: e o que se cola no TOTVS
    ("NF no TOTVS",       "titulo", "NF",           5.5),
    ("Emissão",           "titulo", "",             5),
    # 28/08/2026, pedido do usuario. ⚠ NAO se confunde com `Situação` (o veredito
    # desta aba sobre o pedido): este e o status da nota na SEFAZ, e nota
    # `Cancelada` pode ter PEDIDO CONFIRMADO. Filtravel -- ver `pills` abaixo.
    ("Status da Nota",    "titulo", "",             5.5),
    # 01/09/2026. O que a EMPRESA respondeu sobre a nota (MESMA PREMISSA.xlsx).
    # Fica colada no `Status da Nota` de proposito: uma nota `Autorizada` na
    # SEFAZ e `DesconhecimentoOperacao` aqui e boleto a pagar de nota que a
    # propria empresa nao reconhece -- e o par que interessa, e ler os dois
    # separados por meia tela esconderia isso. Filtravel, ver `pills`.
    ("Manifestação",      "titulo", "",             6),
    ("Emitente",          "titulo", "EMITENTE",    11),
    ("CNPJ Emitente",     "titulo", "EMITENTE",     7),
    ("Nome Fantasia",     "titulo", "EMITENTE",     8),
    ("Destinatário",      "boleto", "DESTINATÁRIO", 9.5),
    ("CNPJ Destinatário", "boleto", "DESTINATÁRIO", 7),
    ("Tipo Operação",     "titulo", "OPERAÇÃO",     4.5),
    ("Nat. Operação",     "titulo", "OPERAÇÃO",     8.5),
    ("Saída/Entrada",     "titulo", "OPERAÇÃO",     5),
    ("CFOP",              "titulo", "OPERAÇÃO",     5),
    ("Vlr SEFAZ",         "titulo", "VALOR",        6),
    ("Vlr Pedido",        "boleto", "VALOR",        6),
    ("@delta_valor",      "delta",  "VALOR",        6),
    ("Venc Duplicata",    "titulo", "DUPLICATA",    5),
    ("Vlr Duplicata",     "boleto", "DUPLICATA",    5.5),
    ("Duplicatas",        "delta",  "DUPLICATA",    3.5),
    # --- o pedido de compra (SC7) e quem o pediu (SC1) ---
    # ⚠ A FILIAL ABRE O BLOCO porque o numero do pedido NAO identifica pedido
    # nenhum sozinho: 2.541 dos 3.669 numeros da SC7 existem em mais de uma
    # empresa (o 001522 e um pedido da Capivara, outro da Bioenergia e outro da
    # Logistica). Sem esta coluna, "Numero PC 001522" na tela nao diz onde
    # conferir. Ver sc7.chave_pedido().
    ("Filial do PC",      "boleto", "PEDIDO",       4.5),
    ("Numero PC",         "boleto", "PEDIDO",       5),
    ("Numero da SC",      "boleto", "PEDIDO",       5.5),
    ("Solicitante",       "boleto", "PEDIDO",       8),
    ("Comprador",         "boleto", "PEDIDO",       8),
    # ⚠ `Usuário SC` SAIU em 13/08/2026: nascia vazia (83 de 24.401 linhas da
    # SC7, e ZERO nos pedidos que casam com nota). O `sc7.py` continua lendo.
    ("Qtd.a Classi",      "boleto", "PEDIDO",       4),
    # ⚠ COLADA na de cima de proposito (pedido do usuario, 14/08/2026): e o
    # TOTAL do pedido, e e ele que da tamanho ao "a classificar" ao lado.
    # Separar as duas por qualquer coluna desfaz a leitura "X de Y".
    ("Quantidade",        "boleto", "PEDIDO",       4),
    ("Controle Ap.",      "boleto", "PEDIDO",       4),
    ("Ped. Encerr.",      "boleto", "PEDIDO",       4),
    ("Dt. Entrega",       "boleto", "PEDIDO",       5),
    ("1° Venc",           "boleto", "PEDIDO",       5),
    # a razao social do fornecedor NO PEDIDO: quando ela nao parece com o
    # emitente da nota, o cruzamento provavelmente errou -- e da para ver isso
    # sem abrir planilha nenhuma
    ("Fornecedor do PC",  "boleto", "PEDIDO",      10),
    ("Itens do PC",       "boleto", "PEDIDO",       4),
    # o texto de onde o numero do pedido foi lido; abre no quadro
    ("Informações",       "",       "",             3.5),
    ("@tratativa",        "",       "",             4),
]

ROTULOS_PLANILHA_PEDIDOS = {
    # ⚠ SEM esta linha o .xlsx desta aba herdava o rotulo GLOBAL do delta,
    # "Diferença de Valor (Boleto - Título)" -- e aqui nao ha boleto nem titulo:
    # o confronto da aba e `Vlr Pedido` contra `Vlr SEFAZ`. A coluna saia com o
    # numero certo e o nome de outra aba (achado em 24/08/2026, conferindo a
    # exportacao coluna a coluna contra a tela).
    "@delta_valor": "Diferença de Valor (Pedido - Nota)",
    "Situação": "O pedido de compra desta nota foi encontrado?",
    "Lançada na SF1": "A nota está lançada no TOTVS (SF1)?",
    "Nº NF": "Nº da NF (SEFAZ)",
    "NF no TOTVS": "Nº da NF no padrão TOTVS (9 dígitos)",
    "Manifestação": "Manifestação do destinatário (MESMA PREMISSA.xlsx)",
    "Emissão": "Emissão (SEFAZ)",
    "Emitente": "Emitente/Prestador (SEFAZ)",
    "Nome Fantasia": "Nome Fantasia do Emitente",
    "Destinatário": "Destinatário/Tomador (a nossa empresa)",
    "CNPJ Destinatário": "CNPJ do Destinatário/Tomador",
    "Tipo Operação": "Tipo de Operação (entrada/saída)",
    "Nat. Operação": "Natureza da Operação (só NF-e)",
    "Saída/Entrada": "Data de Saída/Entrada",
    "CFOP": "CFOP dos itens da nota",
    "Vlr SEFAZ": "Valor Total da Nota (SEFAZ)",
    "Vlr Pedido": "Valor Total do Pedido (soma dos itens, SC7)",
    "Venc Duplicata": "Vencimento da 1ª Duplicata",
    "Vlr Duplicata": "Valor da 1ª Duplicata",
    "Duplicatas": "Quantas duplicatas a nota tem (vazio = uma só)",
    "Filial do PC": "Filial dona do Pedido (SC7) — o nº do pedido se repete entre filiais",
    "Numero PC": "Nº do Pedido de Compra (SC7)",
    "Numero da SC": "Nº da Solicitação de Compra (SC7)",
    "Solicitante": "Solicitante da SC (SC1)",
    "Comprador": "Comprador do Pedido (SC7)",
    "Qtd.a Classi": "Quantidade a Classificar (soma dos itens do PC)",
    "Quantidade": "Quantidade Total do Pedido (soma dos itens, SC7)",
    "Controle Ap.": "Controle de Aprovação (SC7)",
    "Ped. Encerr.": "Pedido Encerrado (SC7)",
    "Dt. Entrega": "Data de Entrega do Pedido (a mais próxima)",
    "1° Venc": "1º Vencimento do Pedido (o mais próximo)",
    "Fornecedor do PC": "Razão Social do Fornecedor no Pedido",
    "Itens do PC": "Quantos itens o Pedido tem",
    "Informações": "Informações Complementares / Descrição do Serviço",
    "Critério": "O que sustentou (ou não) o pedido encontrado",
}

ROTULOS_PLANILHA_CRUZAMENTO = {
    "Situação": "Resultado do Cruzamento",
    "Classificado em": "Data de Classificação no TOTVS (Dt.Digitacao)",
    "Emissão": "Data de Emissão da NF (SF1, DT Emissao)",
    "Critério": "O que bateu no cruzamento",
    "Nº NF": "Nº da NF (SF1)",
    "Entrou em": "Entrou no Painel em",
}

# Confrontos e nomes de coluna mudam entre as abas; o resto do gerador e igual.
ABAS = {
    "associacoes": {
        "planilha": ABA_ASSOCIACOES,
        # O nome da ABA DE ORIGEM continua "Associações Encontradas" (e a
        # planilha, nao muda); aqui e so como ela se chama no painel.
        "nome": "Títulos Associados",
        "cor": "azul",
        "compacta": COMPACTA,
        "confrontos": {"Valor Boleto": ("Vlr.Titulo", "moeda"),
                       "Vencimento Boleto": ("Vencimento", "data")},
        "col_status": "Status",
        "col_criterio": "Critério Match",
        "col_score": "Score Match",
        # ⚠ NAO e o "Campo UUID": depois da fusao com a SF1 a aba tem duas
        # especies de linha, e a da SF1 se identifica pela chave com prefixo
        # "SF1:". Ver COLUNA_CHAVE_PAINEL.
        "col_uuid": COLUNA_CHAVE_PAINEL,
        "prefixo_uuid": "",
        # Botao de filtro por base de origem. O grupo de status entra sozinho
        # (ver grupos_de_pills), entao a aba fica com as duas barras.
        "pills": COLUNA_ORIGEM,
        # Confere cada titulo na SE2 a cada rodada: quem ja tem boleto lancado
        # (ou ja foi baixado) sai da fila. Ver verificacao_se2.py.
        # ⚠ As linhas da SF1 nao estao na SE2 (sao NF, nao titulo): elas caem no
        # `sem_se2` do conferir_na_se2 e FICAM -- que e o fail-open de sempre.
        "conferir_se2": True,
    },
    "correspondencias": {
        "planilha": ABA_CORRESPONDENCIAS,
        "nome": "Títulos Não Associados",
        "cor": "roxo",
        "compacta": COMPACTA_CORRESP,
        "confrontos": {"Valor Boleto (R$)": ("Valor Título (R$)", "moeda"),
                       "Vencimento Boleto": ("Vencimento", "data")},
        "col_status": "Motivo Revisão",
        "col_criterio": "Critério",
        "col_score": "Score",
        "col_uuid": "Campo UUID",
        "prefixo_uuid": "",
        "conferir_se2": True,
    },
    "futuros": {
        "planilha": ABA_FUTUROS,
        "nome": "Boletos Não Associados a Vencer",
        "cor": "vermelho",
        # Botoes de filtro por fonte do boleto, so nesta aba.
        "pills": "Fonte Boleto",
        # Todo mundo aqui e "sem par encontrado" (decisao do usuario): a base
        # ainda traz "CANDIDATO EM REVISÃO" em alguns, mas no painel nao vale
        # a distincao -- nenhum deles achou titulo.
        "forcar_status": "SEM PAR ENCONTRADO",
        "compacta": COMPACTA_FUTUROS,
        "confrontos": {},          # nao ha titulo para confrontar
        "col_status": "Situação",
        "col_criterio": "Arquivo/Observação",
        "col_score": "",
        # Sem Campo UUID nesta aba. A linha digitavel serve de chave: conferido,
        # sao 817 unicas, todas com 47 digitos. O prefixo evita colidir com os
        # UUIDs de verdade das outras abas na tabela de marcacoes.
        "col_uuid": "Linha Digitável",
        "prefixo_uuid": "BOL:",
        "rotulos": {"Fornecedor": "Fornecedor", "Parcela": "Parc."},
        "nao_copiar": {"Fornecedor"},
        # As duas NFs viram campo de copiar: e o que se cola no TOTVS para
        # procurar o titulo que ainda nao foi associado.
        "copiar_extra": {"NF Normalizada": "", "NF/Doc Original": ""},
        # Esta aba e FILTRADA pela data de vencimento no navegador, nao aqui: so
        # assim a NF sai sozinha quando o dia vira -- se o corte fosse feito na
        # geracao, ela so mudaria de estado quando alguem regerasse o painel.
        # Regra (a mesma escrita no painel): vence hoje ou depois -> aparece;
        # venceu antes de hoje -> NAO aparece, A NAO SER que ja tenha sido
        # tratada enquanto ainda estava a vencer.
        #
        # ⚠ Ate 12/08/2026 havia uma SEGUNDA particao aqui ("visao passado"),
        # com as vencidas sem tratamento. O usuario mandou apagar a aba e as
        # vencidas saírem do painel. Restou UMA particao -- que continua
        # existindo justamente porque e ela quem aplica o filtro: sem
        # `particoes` a aba passaria inteira, vencidas incluidas.
        # As linhas nao somem do banco, so deixam de ser exibidas; e o que
        # torna a decisao reversivel (basta devolver a particao "passado").
        # `guia` e o nome da aba no .xlsx exportado -- o Excel corta em 31
        # caracteres e o nome da aba do painel nao cabe. (Nao confundir com
        # `planilha`, que aqui significa a aba da PLANILHA DE ORIGEM.)
        "particoes": [
            {"id": "futuros", "modo": "futura", "cor": "vermelho",
             "nome": "Boletos Não Associados a Vencer",
             "guia": "Boletos nao assoc a vencer"},
        ],
    },
    # Unica aba que NAO vem da base do painel: a planilha do DDA do Itau, lida
    # pelo pendentes_itau.py, que tambem guarda o historico. Aqui ela chega ja
    # no formato das outras -- cabecalho + linhas -- e passa pelo mesmo codigo.
    "pendentes_itau": {
        "planilha": None,                 # nao existe aba de origem: e outro arquivo
        "fonte": "pendentes_itau",
        "nome": "Boletos em Aberto DDA",
        "guia": "Boletos em Aberto DDA",
        "cor": "verde",
        "compacta": COMPACTA_PENDENTES,
        # O unico confronto possivel aqui: quanto ja correu de juros/multa.
        "confrontos": {"A Pagar": ("Até Venc.", "moeda")},
        "col_status": "Situação",
        "col_criterio": "Tipo de Boleto",
        "col_score": "",
        # A chave e calculada pelo pendentes_itau (CNPJ + doc + venc + valor) e
        # chega numa coluna propria; o prefixo "ITAU:" ja vem embutido nela.
        "col_uuid": "Chave",
        "prefixo_uuid": "",
        # SEM botoes de filtro nesta aba (13/08/2026, pedido do usuario: "remova
        # esses filtros de botoes da aba boletos abertos dda"). Eram por empresa
        # pagadora, e o Itau trunca o pagador: dois viravam botao fantasma de 1
        # linha ("S" = SOMPO, "S D FLORESTAL" = EQUIPO) sem dizer de qual empresa
        # sao. "Pagador/Agregado" segue coluna visivel, entao a busca acha por
        # empresa do mesmo jeito -- o que sai e so a barra de botoes.
        "rotulos": {"Beneficiário": "Beneficiário"},
        # Beneficiario e o NOME, nao codigo: nao vira botao de copiar (e o mesmo
        # motivo pelo qual "Fornecedor" nao e copiavel na aba de nao associadas).
        "nao_copiar": {"Beneficiário"},
        # O que se cola no TOTVS para procurar o titulo deste boleto.
        "copiar_extra": {"Nº Doc.": "", "CPF/CNPJ": ""},
        # Os rotulos do resumo mudam: aqui nao ha titulo, entao "boleto diferente
        # do titulo" nao quer dizer nada. `None` esconde o item.
        "resumo": {"alerta": "com aviso do banco",
                   "divergentes": "já com juros/multa"},
        "rotulos_planilha": ROTULOS_PLANILHA_PENDENTES,
    },
    # Tambem nao vem da base do painel: sai da SF1 (BASES GENERICOS), cruzada
    # pelo cruzamento_classificacao.py com os boletos que ainda nao acharam
    # titulo -- os do DDA (aba "Futuros NF Nao Associada") e os do e-mail do
    # Itau. Chega aqui no formato das outras e passa pelo mesmo codigo.
    "cruzamento": {
        "planilha": None,
        "fonte": "cruzamento",
        "nome": "Títulos Associados na Classificação",
        "guia": "Titulos Assoc Classificacao",
        # ⚠ DESLIGADA em 14/08/2026: as linhas passaram a viver na aba de
        # titulos associados (ver fundir_classificacao). A aba continua aqui,
        # e continua sendo ENVIADA na carga, porque a lista de abas do painel
        # sai das linhas `#meta:` que estao no Postgres -- parar de mandar nao
        # a apaga, so deixa o cabecalho velho no ar. E este `oculta` que a tira
        # da tela. O resto do cfg fica intacto: e o caminho de volta se ele
        # quiser as duas abas separadas outra vez.
        "oculta": True,
        # ambar: cor que ficou livre com a saida da "visao passado" (12/08/2026)
        "cor": "ambar",
        "compacta": COMPACTA_CRUZAMENTO,
        "confrontos": {"Valor Boleto": ("Vlr.Título", "moeda"),
                       "Vencimento Boleto": ("Vencimento", "data")},
        "col_status": "Situação",
        "col_criterio": "Critério",
        "col_score": "",
        # O prefixo "SF1:" ja vem embutido na chave montada pelo modulo.
        "col_uuid": "Chave",
        "prefixo_uuid": "",
        "pills": "Fonte Boleto",
        # Razao Social e nome, nao codigo -- nao vira botao de copiar (mesma
        # regra do Fornecedor nas outras abas).
        "nao_copiar": {"Razão Social", "Fornecedor Boleto"},
        # O que se cola no TOTVS para achar a NF e o titulo dela.
        "copiar_extra": {"Nº NF": "", "Campo UUID": ""},
        "resumo": {"alerta": "com aviso do banco",
                   "divergentes": "boleto difere do título"},
        "rotulos_planilha": ROTULOS_PLANILHA_CRUZAMENTO,
    },
    # A SEFAZ.xlsx, so com as colunas que o usuario pintou de verde. Nao cruza
    # com nada: e a base da SEFAZ ficando legivel. Ver sefaz.py.
    "sefaz": {
        "planilha": None,
        "fonte": "sefaz",
        "nome": "Análise Base SEFAZ",
        "guia": "Analise Base SEFAZ",
        "cor": "azul",
        "compacta": COMPACTA_SEFAZ,
        "confrontos": {},          # nao ha par para confrontar
        # A Origem faz o papel de Status: e o flag de qual aba da planilha a
        # linha veio. O painel decide o desenho pelo PAPEL, nunca pelo nome.
        "col_status": "Origem",
        # ⚠ VAZIO de proposito. Com uma coluna aqui, o `montar_match` monta um
        # quadro para cada uma das 1.188 linhas ("NF-e / Saída") -- 122 KB de
        # carteira para dizer o que o proprio selo ja diz.
        "col_criterio": "",
        "col_score": "",
        # a chave ja vem montada (chave NF-e + item + duplicata), com prefixo
        "col_uuid": "Chave",
        "prefixo_uuid": "",
        # 28/08/2026: "permita filtrar o status da nota". Passa na guarda de 2 a 8
        # valores (na base de hoje sao 4 somando as duas abas: Autorizada,
        # Cancelada, Normal, Substituída), entao vira barra pelo caminho normal.
        "pills": ["Origem", "Status da Nota"],
        # Selo por valor exato: sem isto "NF-e" e "NFS-e" sairiam da mesma cor,
        # e o flag nao flagaria nada.
        "selos": {"NF-e": "nfe", "NFS-e": "nfse"},
        # Colunas de texto longo: a celula vira botao e o texto inteiro abre no
        # quadro flutuante, o mesmo do critério de match.
        "texto_longo": {"Informações"},
        "nao_copiar": {"Emitente", "Destinatário", "Nome Fantasia"},
        "copiar_extra": {"Nº NF": "", "CNPJ Emitente": ""},
        # A coluna de alerta aqui se chama so "Alerta" (nas abas de boleto e
        # "Alerta Fatura"); e o `col_alerta` que diz qual tem o papel.
        "col_alerta": "Alerta",
        "resumo": {"alerta": "com alerta", "divergentes": None},
        "rotulos_planilha": ROTULOS_PLANILHA_SEFAZ,
        # ABA DESLIGADA em 24/08/2026, pedido do usuario ("desabilite essa aba").
        # Mesmo desenho da SEFAZ x SF1 logo abaixo: `oculta` NAO e apagar a
        # entrada. A leitura da SEFAZ continua rodando (a aba de PEDIDOS depende
        # dela) e o CABECALHO continua sendo publicado -- e ele que sobrescreve,
        # com a marca de oculta, o cabecalho que ja esta no Postgres. Sem isso a
        # aba voltaria sozinha: a lista de abas do painel sai das linhas `#meta:`
        # do banco, e nao do que este arquivo gerou hoje.
        # Para trazer de volta: apagar esta linha.
        "oculta": True,
    },
    # A nota da SEFAZ ja foi lancada no TOTVS? Cruza SO com a SF1, pelo numero
    # da NF. Ver cruzamento_sefaz_sf1.py.
    "sefaz_sf1": {
        "planilha": None,
        "fonte": "sefaz_sf1",
        "nome": "SEFAZ x SF1",
        "guia": "SEFAZ x SF1",
        "cor": "verde",
        "compacta": COMPACTA_SEFAZ_SF1,
        "confrontos": {"Vlr.Título": ("Vlr SEFAZ", "moeda")},
        "col_status": "Situação",
        "col_criterio": "Critério",
        "col_score": "",
        "col_uuid": "Chave",
        "prefixo_uuid": "",
        "pills": "Origem",
        "selos": {"ACHADA NA SF1": "forte", "CONFERIR": "provavel",
                  "NÃO ACHADA": "grave"},
        # Nome de gente/empresa nao vira botao de copiar -- MENOS o destinatario,
        # que o usuario pediu explicitamente copiavel em 1 clique junto do CNPJ.
        "nao_copiar": {"Emitente", "Razão Social", "Nome Fantasia"},
        "copiar_extra": {"Nº NF": "", "Destinatário": "", "CNPJ Destinatário": ""},
        "col_alerta": "Alerta",
        "resumo": {"alerta": "com alerta", "divergentes": "valor difere da SF1"},
        "rotulos_planilha": ROTULOS_PLANILHA_SEFAZ_SF1,
        # ⚠ DESLIGADA em 14/08/2026, a pedido do usuario: a aba `pedidos` e um
        # superconjunto dela (leva o mesmo veredito de lancamento na coluna
        # `Lançada na SF1`, as mesmas colunas da nota e os mesmos 3 alertas).
        # `oculta` NAO e o mesmo que apagar a entrada: o cruzamento continua
        # rodando (a aba de pedidos depende dele) e o CABECALHO continua sendo
        # publicado -- e ele que sobrescreve, com a marca de oculta, o cabecalho
        # que ja esta no banco. Sem isso a aba voltaria sozinha: desde o
        # carregamento sob demanda, a lista de abas sai das linhas `#meta:` do
        # Postgres, e nao do que este arquivo gerou hoje.
        # Para trazer de volta: apagar esta linha.
        "oculta": True,
    },
    # De qual PEDIDO DE COMPRA veio esta nota? Junta SEFAZ + SF1 + SC7/SC1.
    # Ver pedidos_sefaz.py -- inclusive por que o veredito tem cinco degraus.
    "pedidos": {
        "planilha": None,
        "fonte": "pedidos",
        "nome": "NF x Pedido de Compra",
        "guia": "NF x Pedido de Compra",
        "cor": "indigo",
        "compacta": COMPACTA_PEDIDOS,
        # Δ = pedido - nota. Positivo quer dizer pedido maior que a nota, que e
        # o normal na entrega parcial; negativo pede explicacao.
        "confrontos": {"Vlr Pedido": ("Vlr SEFAZ", "moeda")},
        "col_status": "Situação",
        "col_criterio": "Critério",
        "col_score": "",
        "col_uuid": "Chave",
        "prefixo_uuid": "",
        # ⚠ LISTA de colunas (o resto do painel declara uma so). Cada uma so vira
        # botao se tiver de 2 a 8 valores na base do dia -- a guarda de sempre.
        # `Origem` (NF-e / NFS-e) foi pedida em 13/08/2026.
        # ⚠ `Ped. Encerr.` SAIU daqui em 14/08/2026 e virou filtro DERIVADO no
        # navegador (`PILLS_DERIVADAS` no painel_modelo.html), junto com o da
        # quantidade a classificar. Pela regra daqui ela nunca virava barra: tem
        # um valor preenchido so ("E"), e a guarda pede 2. Ver o comentario la
        # -- inclusive por que a nota SEM PEDIDO nao entra em nenhum dos dois.
        # 24/08/2026: "Controle Ap." entrou na lista a pedido do usuario. Ela
        # passa na guarda de 2 a 8 valores (na base de hoje sao 5: L, B, B · L,
        # L · R, R), entao vira barra pelo caminho normal -- ao contrario de
        # `Ped. Encerr.` e `Qtd.a Classi`, que precisaram nascer derivadas no
        # navegador. Comprador NAO entra aqui: sao 51 valores, muito acima da
        # guarda -- filtro de muito valor virou busca com dropdown, no campo de
        # texto (ver LISTAS_FILTRO no painel_modelo.html).
        # 28/08/2026: `Status da Nota` entrou a pedido do usuario ("permita
        # filtrar o status da nota"). Mesma guarda de 2 a 8 valores.
        # 31/08/2026: `Filial do PC` entrou junto com a chave por filial -- e a
        # pergunta "de qual empresa e este pedido?", que passou a existir. Cai na
        # mesma guarda de 2 a 8 valores das outras; se um dia a base trouxer mais
        # de 8 filiais amarradas a nota, ela simplesmente nao vira barra.
        # 01/09/2026: `Manifestação` entrou na lista. Cai na mesma guarda de 2 a 8
        # valores das outras -- na base de hoje sao 6 (os 4 status do TOTVS mais
        # "nao se aplica (NFS-e)" e "fora do periodo do arquivo").
        "pills": ["Origem", "Lançada na SF1", "Controle Ap.", "Status da Nota",
                  "Manifestação", "Filial do PC"],
        "selos": {pedidos_sefaz.CONFIRMADO: "forte",
                  # verde junto do CONFIRMADO: as duas sao "nada a fazer" -- uma
                  # porque o pedido esta amarrado, a outra porque ele ja fechou
                  pedidos_sefaz.ENCERRADO_PC: "forte",
                  pedidos_sefaz.PROVAVEL: "provavel",
                  pedidos_sefaz.CONFERIR: "provavel",
                  # a analise agrupada (por soma de valores) e sempre conferir
                  pedidos_sefaz.AGRUPADO_NF: "provavel",
                  pedidos_sefaz.AGRUPADO_PC: "provavel",
                  pedidos_sefaz.FORA: "grave",
                  pedidos_sefaz.SEM: "grave"},
        "texto_longo": {"Informações"},
        # nome de gente/empresa nao vira botao de copiar; o destinatario e a
        # excecao pedida pelo usuario, e o PC/NF sao o que se cola no TOTVS
        "nao_copiar": {"Emitente", "Nome Fantasia", "Fornecedor do PC",
                       "Solicitante", "Comprador"},
        "copiar_extra": {"Nº NF": "", "NF no TOTVS": "", "Numero PC": "",
                         "Numero da SC": "", "Destinatário": "",
                         "CNPJ Destinatário": ""},
        "col_alerta": "Alerta",
        "resumo": {"alerta": "com alerta", "divergentes": "valor difere do pedido"},
        "rotulos_planilha": ROTULOS_PLANILHA_PEDIDOS,
    },
}

# Cabecalho mais curto na Conferencia (o painel casa a coluna pela chave, nao
# pelo rotulo -- renomear aqui e seguro).
# Rotulos curtos: o cabecalho e uma linha so, entao nome comprido seria cortado.
ROTULOS_COMPACTA = {"@delta_valor": "Δ valor", "@delta_dias": "Δ dias",
                    "Alerta Fatura": "Alerta", "Campo UUID": "UUID",
                    "Linha Digitável": "Linha", "@tratativa": "Nota",
                    # ⚠ era "Tipo (SE2)". Depois da fusao com a SF1 (14/08/2026)
                    # a coluna carrega DUAS origens -- o tipo do titulo (SE2,
                    # coluna IK) e o da NF (SF1, coluna K) -- e o nome antigo
                    # mentia na metade das linhas.
                    "@tipo_pgto": "Tipo Pgto", "@tipo_real": "Tipo real",
                    "@alerta_tipo": "Confere?",
                    "CHECK (FEITO)": "OK", "No. Titulo": "Nº título",
                    "NF/Doc Boleto": "NF boleto", "Prefixo": "Pref.",
                    "Parcela": "Parc.", "Razão Social": "Razão social",
                    "Fornecedor": "Cód.", "Vlr.Titulo": "Vlr. título",
                    "Valor Boleto": "Vlr. boleto", "Vencimento": "Vencto",
                    "Vencto Real": "Vencto real", "Vencimento Boleto": "Venc. boleto",
                    # nomes da aba Tratar Correspondencias
                    "Valor Título (R$)": "Vlr. título", "Valor Boleto (R$)": "Vlr. boleto",
                    "Motivo Revisão": "Motivo", "Fornecedor Boleto": "Forn. boleto",
                    # nomes da aba Futuros NF Nao Associada
                    "NF/Doc Original": "NF original", "NF Normalizada": "NF norm.",
                    "Total Parcelas": "de", "Valor (R$)": "Valor",
                    "Empresa/Origem": "Empresa", "Arquivo/Observação": "Observação",
                    "Fonte Boleto": "Fonte", "CNPJ/CPF": "CNPJ/CPF",
                    "Situação": "Situação",
                    # nomes da aba NFs pendentes Itau
                    "Pagador/Agregado": "Empresa", "Nº Doc.": "Nº doc.",
                    "Tipo de Boleto": "Tipo", "Até Venc.": "Até venc.",
                    "A Pagar": "A pagar", "Entrou em": "Entrou",
                    # nomes da aba Titulos Associados na Classificacao
                    "Classificado em": "Classif.", "Nº NF": "Nº NF",
                    "Vlr.Título": "Vlr. título", "Critério": "Critério",
                    "Série": "Série", "Boletos": "Cand.",
                    # abas da SEFAZ (o cabecalho e uma linha so)
                    "Venc Duplicata": "Venc. dupl.", "Vlr Duplicata": "Vlr. dupl.",
                    "Duplicatas": "Qtd.", "CNPJ Destinatário": "CNPJ dest.",
                    "Nome Fantasia": "Fantasia",
                    "Tipo Operação": "Tipo op.", "Nat. Operação": "Natureza",
                    "Saída/Entrada": "Saída/ent."}

# Cabecalho da planilha exportada: aqui vale o nome por extenso, nao a
# abreviacao da tela. Sem entrada, sai o proprio nome da coluna da base.
ROTULOS_PLANILHA = {"@delta_valor": "Diferença de Valor (Boleto - Título)",
                    "@delta_dias": "Diferença de Dias (Boleto - Título)",
                    "@tratativa": "Tratativa",
                    "@tipo_pgto": "Tipo de Pagamento no Título (SE2)",
                    "@tipo_real": "Tipo de Pagamento Real (informado no painel)",
                    "@alerta_tipo": "Tipo Real Confere com a SE2?",
                    "CHECK (FEITO)": "Tratado",
                    "Campo UUID": "Campo UUID",
                    "No. Titulo": "No. Título",
                    "Vlr.Titulo": "Valor Título",
                    "Situação": "Situação do Boleto"}

LARGURAS = {
    "Campo UUID": 296, "Razão Social": 250, "Fornecedor Boleto": 200,
    "Nome Fornece": 175, "Critério Match": 280, "Alerta Fatura": 280,
    "Linha Digitável": 290, "Historico": 300, "CNPJ Boleto": 145,
    "No. Titulo": 104, "Status": 122, "Fonte Boleto": 135,
    "Beneficiário": 250, "Pagador/Agregado": 200, "Tipo de Boleto": 150,
    # aba Titulos Associados na Classificacao: "Situação" cabe nos 122 herdados
    # de Status, mas "MATCH PROVÁVEL" e "SEM BOLETO" sao selos longos e cortados
    # nao servem de nada; "Critério" carrega frases como "NF + CNPJ + valor".
    "Critério": 230, "Classificado em": 118,
}
LARGURA_PADRAO = 108


def chave_de(rotulo: str) -> str:
    texto = unicodedata.normalize("NFKD", rotulo)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"_+", "_", re.sub(r"[^0-9a-z]+", "_", texto.lower())).strip("_")


def abrir_planilha(caminho: Path):
    """Abre a planilha; se estiver travada pelo Excel/OneDrive, usa uma copia."""
    try:
        return load_workbook(caminho, read_only=True, data_only=True), None
    except (PermissionError, OSError):
        pasta = tempfile.mkdtemp(prefix="painel_boletos_")
        shutil.copy2(caminho, Path(pasta) / caminho.name)
        return load_workbook(Path(pasta) / caminho.name, read_only=True, data_only=True), pasta


def texto(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, (dt.datetime, dt.date)):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    limpo = re.sub(r"\s+", " ", str(valor).replace("_x000D_", " ")).strip()
    return "" if limpo in ("/ /", "//") else limpo


def numero(valor):
    if valor in (None, ""):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    limpo = re.sub(r"[^0-9,.\-]", "", str(valor))
    if "," in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")
    try:
        return float(limpo)
    except ValueError:
        return None


def data_iso(valor):
    if isinstance(valor, (dt.datetime, dt.date)):
        return valor.strftime("%Y-%m-%d")
    for formato in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(texto(valor), formato).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None


def moeda_br(valor) -> str:
    if valor is None:
        return ""
    inteiro = f"{valor:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return f"R$ {inteiro}"


def quantidade_br(valor) -> str:
    """Quantidade em pt-BR, sem casa decimal que nao existe: 20, 1.890.000, 0,954.

    ⚠ A parte decimal e cortada SEPARADAMENTE, e nao com um `rstrip("0")` no
    numero inteiro: em 1.000 o strip comeria os zeros da casa do milhar e
    devolveria "1" -- um erro de mil vezes, calado.
    """
    if valor is None:
        return ""
    inteiro, _, fracao = f"{valor:,.3f}".partition(".")
    fracao = fracao.rstrip("0")
    inteiro = inteiro.replace(",", ".")
    return f"{inteiro},{fracao}" if fracao else inteiro


def formatar_linha_digitavel(valor) -> str:
    """Deixa a linha no formato lido pelo banco: 5.5 5.6 5.6 1 14."""
    digitos = re.sub(r"\D", "", texto(valor))
    if len(digitos) != 47:
        return texto(valor)
    return " ".join([
        f"{digitos[0:5]}.{digitos[5:10]}", f"{digitos[10:15]}.{digitos[15:21]}",
        f"{digitos[21:26]}.{digitos[26:32]}", digitos[32], digitos[33:47],
    ])


def tipo_da_coluna(rotulo: str) -> str:
    if rotulo in MOEDA:
        return "moeda"
    if rotulo in DATA:
        return "data"
    if rotulo in INTEIRO:
        return "inteiro"
    if rotulo in QUANTIDADE:
        return "quantidade"
    return "texto"


def montar_colunas(cabecalho: list[str], usadas: set[str]) -> list[dict]:
    return [
        {
            "chave": chave_de(rotulo),
            "rotulo": rotulo,
            "tipo": tipo_da_coluna(rotulo),
            "copiar": COPIAVEIS.get(rotulo, ""),
            "check": rotulo == COLUNA_CHECK,
            "largura": 58 if rotulo == COLUNA_CHECK else LARGURAS.get(rotulo, LARGURA_PADRAO),
        }
        for rotulo in cabecalho
        if rotulo and rotulo not in IGNORAR and (rotulo in usadas or rotulo == COLUNA_CHECK)
    ]


def delta_bruto(origem: dict, coluna_boleto: str, confrontos: dict):
    """Diferenca numerica boleto - titulo (reais ou dias); None se nao da para comparar."""
    coluna_titulo, tipo = confrontos[coluna_boleto]
    if tipo == "moeda":
        valor_t, valor_b = numero(origem.get(coluna_titulo)), numero(origem.get(coluna_boleto))
        return None if None in (valor_t, valor_b) else round(valor_b - valor_t, 2)

    data_t, data_b = data_iso(origem.get(coluna_titulo)), data_iso(origem.get(coluna_boleto))
    if not (data_t and data_b):
        return None
    return (dt.date.fromisoformat(data_b) - dt.date.fromisoformat(data_t)).days


def diferenca(origem: dict, coluna_boleto: str, confrontos: dict) -> str:
    """Texto curto da divergencia entre o boleto e o titulo ('' se conferem)."""
    delta = delta_bruto(origem, coluna_boleto, confrontos)
    if not delta:
        return ""
    if confrontos[coluna_boleto][1] == "moeda":
        return ("+" if delta > 0 else "−") + moeda_br(abs(delta))
    return f"{delta:+d} {'dia' if abs(delta) == 1 else 'dias'}"


def ler_aba(wb, cfg: dict):
    """Cabecalho + linhas CRUAS de uma aba da base.

    ⚠ Devolve o PAR, e nao o resultado do `montar_aba`. A aba de titulos
    associados so pode ser montada depois da SF1 -- as duas viraram uma aba so
    (14/08/2026) e o `montar_aba` precisa do cabecalho ja unificado. Quem
    monta agora e o `gerar`, que sabe quando as duas fontes estao prontas.
    """
    ws = wb[cfg["planilha"]]
    linhas = list(ws.iter_rows(values_only=True))
    cabecalho = [texto(c) for c in linhas[0]]
    brutas = [b for b in linhas[1:] if not all(v in (None, "") for v in b)]
    return cabecalho, brutas


def fundir_classificacao(cab_base: list[str], brutas_base: list,
                         cab_sf1: list[str], brutas_sf1: list):
    """Une as linhas da SF1 a aba dos titulos associados.

    Devolve (cabecalho, linhas) no mesmo formato que o `montar_aba` espera de
    qualquer aba: uma lista de colunas so, com `Origem` dizendo de qual base
    cada linha veio.
    """
    renomeado = [SF1_PARA_BASE.get(c, c) for c in cab_sf1]

    # Ordem: Origem primeiro, depois a base do associador inteira, e por fim o
    # que so a SF1 traz (na ordem dela). A chave interna fecha a lista.
    unico = [COLUNA_ORIGEM] + list(cab_base)
    for coluna in renomeado:
        if coluna and coluna not in unico:
            unico.append(coluna)
    unico.append(COLUNA_CHAVE_PAINEL)
    posicao = {c: i for i, c in enumerate(unico)}

    linhas = []
    for cabecalho, brutas, marca, coluna_chave in (
            (cab_base, brutas_base, ORIGEM_SE2, "Campo UUID"),
            (renomeado, brutas_sf1, ORIGEM_SF1, "Chave")):
        for bruta in brutas:
            linha = [""] * len(unico)
            linha[posicao[COLUNA_ORIGEM]] = marca
            for coluna, valor in zip(cabecalho, bruta):
                if coluna:
                    linha[posicao[coluna]] = valor
            # SE2: o proprio Campo UUID. SF1: a "Chave", que ja vem com o
            # prefixo "SF1:" embutido pelo cruzamento_classificacao.
            linha[posicao[COLUNA_CHAVE_PAINEL]] = texto(
                dict(zip(cabecalho, bruta)).get(coluna_chave))
            linhas.append(tuple(linha))
    return unico, linhas


def ler_cruas(wb, nome_aba: str) -> list[dict]:
    """Uma aba da base como dicionarios CRUS ({coluna: valor}), sem passar pelo
    tratamento do painel.

    Serve ao cruzamento da classificacao, que precisa dos valores como estao na
    planilha (numero, data, CNPJ) -- o registro montado para a tela ja vem
    formatado em texto, e reconverter seria dar duas chances de errar.
    """
    if nome_aba not in wb.sheetnames:
        return []
    linhas = list(wb[nome_aba].iter_rows(values_only=True))
    if not linhas:
        return []
    cabecalho = [texto(c) for c in linhas[0]]
    return [dict(zip(cabecalho, b)) for b in linhas[1:]
            if not all(v in (None, "") for v in b)]


def montar_aba(cabecalho: list[str], brutas: list, cfg: dict):
    """Cabecalho + linhas -> colunas, visao Conferencia e registros do painel.

    Separado do `ler_aba` porque a aba de pendentes do Itau nao sai de uma aba
    da planilha: ela vem do `pendentes_itau.py`, que junta a planilha do DDA com
    o historico. Dali para frente o tratamento e exatamente o mesmo.
    """
    compacta = montar_compacta(cabecalho, cfg["compacta"], cfg)
    confrontos = cfg["confrontos"]
    # Futuros NF nao tem titulo para confrontar -- fica sem Δ.
    moeda_boleto = next((c for c, (_, t) in confrontos.items() if t == "moeda"), None)

    # Uma coluna 100% vazia nas linhas associadas nao ajuda ninguem a conferir.
    usadas = {
        rotulo for i, rotulo in enumerate(cabecalho)
        if rotulo and any(texto(b[i]) for b in brutas)
    }
    colunas = montar_colunas(cabecalho, usadas)

    registros = []
    for bruta in brutas:
        origem = dict(zip(cabecalho, bruta))
        celulas, ordem = {}, {}

        for coluna in colunas:
            rotulo, chave, tipo = coluna["rotulo"], coluna["chave"], coluna["tipo"]
            valor = origem.get(rotulo)
            if tipo == "moeda":
                bruto = numero(valor)
                celulas[chave], ordem[chave] = moeda_br(bruto), bruto
            elif tipo == "data":
                celulas[chave], ordem[chave] = texto(valor), data_iso(valor) or ""
            elif tipo == "inteiro":
                bruto = numero(valor)
                celulas[chave] = "" if bruto is None else str(int(bruto))
                ordem[chave] = bruto
            elif tipo == "quantidade":
                bruto = numero(valor)
                celulas[chave], ordem[chave] = quantidade_br(bruto), bruto
            elif rotulo == "Linha Digitável":
                celulas[chave] = formatar_linha_digitavel(valor)
                ordem[chave] = re.sub(r"\D", "", texto(valor))
            elif rotulo == cfg["col_status"] and cfg.get("forcar_status"):
                celulas[chave] = ordem[chave] = cfg["forcar_status"]
            elif rotulo == COLUNA_FONTE:
                bruto = texto(valor)
                celulas[chave] = ordem[chave] = FONTES.get(bruto, bruto)
            else:
                celulas[chave] = ordem[chave] = texto(valor)

        registros.append({
            "c": celulas,
            "o": ordem,
            "uuid": chave_registro(origem, cfg),
            "feito": bool(texto(origem.get(COLUNA_CHECK))),
            # valor que o botao copia (o exibido tem pontuacao de leitura)
            "copia": {chave_de("Linha Digitável"): re.sub(r"\D", "", texto(origem.get("Linha Digitável")))},
            "difere": {chave_de(col): diferenca(origem, col, confrontos) for col in confrontos
                       if diferenca(origem, col, confrontos)},
            # deltas numericos: alimentam a visao Conferencia e a ordenacao
            "delta": {
                "valor": delta_bruto(origem, moeda_boleto, confrontos) if moeda_boleto else None,
                "dias": (delta_bruto(origem, "Vencimento Boleto", confrontos)
                         if "Vencimento Boleto" in confrontos else None),
            },
            # `forte` so existe nas Associacoes; em Tratar Correspondencias todo
            # titulo esta em revisao, entao o selo nunca e verde.
            "forte": "FORTE" in texto(origem.get(cfg["col_status"])).upper(),
            "selo": selo_status(origem, cfg),
            "ocr": ocr_suspeito(origem),
            "alerta": montar_alerta(origem, cfg),
            "match": montar_match(origem, cfg),
            # texto que nao cabe na celula, por coluna: abre no quadro
            "texto": {chave_de(col): texto(origem.get(col))
                      for col in cfg.get("texto_longo", ())
                      if texto(origem.get(col))},
            "busca": " ".join(texto(v) for v in origem.values() if texto(v)).lower(),
        })

    if cfg["col_score"]:
        chave_score = chave_de(cfg["col_score"])
        registros.sort(key=lambda r: -(numero(r["o"].get(chave_score)) or 0))
    else:
        # sem score: os que nao acharam par primeiro, depois por vencimento
        registros.sort(key=lambda r: (r["selo"] != "grave", r["o"].get("vencimento") or ""))
    return colunas, compacta, registros


# Colunas que NAO existem na base: o painel as desenha por conta propria. O
# valor aqui e o que o `celulaCompacta` do modelo usa para escolher o desenho --
# sem entrada, uma coluna "@nova" sairia como delta e mostraria "—" para sempre.
TIPOS_VIRTUAIS = {
    "@delta_valor": "delta",
    "@delta_dias": "delta",
    "@tratativa": "tratativa",
    "@tipo_pgto": "tipo_pgto",      # so leitura, vem da SE2
    "@tipo_real": "tipo_real",      # menu; grava no banco, por titulo
    "@alerta_tipo": "alerta_tipo",  # calculado na hora, no navegador
}


def montar_compacta(cabecalho: list[str], definicao: list, cfg: dict) -> list[dict]:
    """Colunas da visao Conferencia, respeitando a ordem em que estao na base."""
    ordem_base = {rotulo: i for i, rotulo in enumerate(cabecalho) if rotulo}
    colunas = []
    for rotulo, lado, grupo, largura in definicao:
        virtual = rotulo.startswith("@")
        # O check e recurso do painel, nao coluna da base: entra mesmo onde a
        # planilha nao tem "CHECK (FEITO)" -- e o caso de Tratar Correspondencias.
        if not virtual and rotulo != COLUNA_CHECK and rotulo not in ordem_base:
            continue  # a base mudou de nome: melhor faltar a coluna do que errar
        # Em Futuros NF a coluna "Fornecedor" e o NOME, nao o codigo: la ela nao
        # vira botao de copiar e nem se chama "Cód.". A mesma aba ganha as NFs
        # copiaveis, que nao sao copiaveis nas outras.
        rotulos = {**ROTULOS_COMPACTA, **cfg.get("rotulos", {})}
        copiaveis = {**COPIAVEIS, **cfg.get("copiar_extra", {})}
        copiavel = rotulo in copiaveis and rotulo not in cfg.get("nao_copiar", ())
        colunas.append({
            "chave": rotulo.lstrip("@") if virtual else chave_de(rotulo),
            "rotulo": rotulos.get(rotulo, rotulo),
            "tipo": TIPOS_VIRTUAIS[rotulo] if virtual else tipo_da_coluna(rotulo),
            "lado": lado,
            "grupo": grupo,
            "largura": largura,
            # o painel decide o desenho pelo PAPEL, nao pelo nome da coluna --
            # em Tratar Correspondencias o "status" chama Motivo Revisão.
            "papel": ("status" if rotulo == cfg["col_status"]
                      else "alerta" if rotulo == cfg.get("col_alerta", "Alerta Fatura")
                      # texto que nao cabe em celula (Info Comp da SEFAZ chega a
                      # 5.000 caracteres): a celula vira botao e o texto abre no
                      # quadro flutuante
                      else "texto" if rotulo in cfg.get("texto_longo", ()) else ""),
            "copiavel": copiavel,
            "copiar": copiaveis.get(rotulo, "") if copiavel else "",
            "so_botao": rotulo in SO_BOTAO,
            "check": rotulo == COLUNA_CHECK,
            # Cabecalho da EXPORTACAO. Na tela o rotulo e abreviado porque o
            # cabecalho e uma linha so; na planilha cabe o nome inteiro -- e o
            # nome da base e o que a pessoa reconhece.
            "planilha": {**ROTULOS_PLANILHA,
                         **cfg.get("rotulos_planilha", {})}.get(rotulo, rotulo),
        })
    return colunas


def tipo_alerta(valor) -> str:
    """PARCELA / FATURA / ALERTA -- rotulo curto do selo ('' quando nao ha alerta)."""
    texto_alerta = texto(valor).upper()
    if not texto_alerta:
        return ""
    for marcadores, rotulo in ALERTAS:
        if any(m in texto_alerta for m in marcadores):
            return rotulo
    return "ALERTA"


def chave_registro(origem: dict, cfg: dict) -> str:
    """Chave que amarra check e tratativa ao titulo, e sobrevive a regeracao.

    Nas duas primeiras abas e o Campo UUID da base. Em Futuros NF nao existe
    UUID, entao vale a linha digitavel (unica), com prefixo para nao colidir.
    """
    bruto = texto(origem.get(cfg["col_uuid"]))
    if not cfg["prefixo_uuid"]:
        return bruto
    return cfg["prefixo_uuid"] + re.sub(r"\D", "", bruto)


def selo_status(origem: dict, cfg: dict) -> str:
    """forte (verde) / provavel (ambar) / grave (vermelho)."""
    bruto = cfg.get("forcar_status") or texto(origem.get(cfg["col_status"]))
    # Mapa explicito da aba, quando existe. Vem antes das regras por SUBSTRING
    # porque elas nao servem para todo status: "ACHADA NA SF1" e "NÃO ACHADA"
    # sao vereditos opostos e um contem o outro.
    if bruto in cfg.get("selos", {}):
        return cfg["selos"][bruto]
    valor = bruto.upper()
    if "FORTE" in valor:
        return "forte"
    # Pendentes do Itau: "SAIU DA BASE" e boa noticia (o boleto deixou de estar
    # pendente no DDA -- na pratica, foi pago), entao entra em verde junto com o
    # match forte. Nenhum status das outras abas contem "SAIU".
    if "SAIU" in valor:
        return "forte"
    # "SEM PAR ENCONTRADO" (nao associadas) e "SEM BOLETO" (classificacao) sao a
    # mesma noticia: nada casou, e e onde a pessoa precisa agir.
    if "SEM PAR" in valor or "SEM BOLETO" in valor:
        return "grave"
    return "provavel"


# Boletos do CENTRAL cujo fornecedor saiu como "FABIO COSTA LIMA" sao leitura
# errada do OCR -- o nome nao e do fornecedor de verdade. Marcar para a pessoa
# nao sair atras do Fabio.
OCR_FORNECEDOR = "FABIO COSTA LIMA"
OCR_FONTE = "CENTRAL"


def ocr_suspeito(origem: dict) -> bool:
    fornecedor = texto(origem.get("Fornecedor")).upper()
    fonte = texto(origem.get("Fonte Boleto")).upper()
    return OCR_FORNECEDOR in fornecedor and OCR_FONTE in fonte


def montar_match(origem: dict, cfg: dict) -> dict:
    """Resumo do porque o boleto casou com o titulo -- vira o quadro que abre ao
    passar o mouse na celula de Status.

    O criterio da base e uma frase longa com as evidencias somadas por ' + ';
    aqui ela e quebrada em itens para ler de bate-pronto. A parte depois de
    'REVISAO:' (so em Tratar Correspondencias) vira o motivo em separado.
    """
    bruto = texto(origem.get(cfg["col_criterio"]))
    evidencias, revisao = bruto, ""
    if "REVISAO:" in bruto.upper():
        corte = re.split(r"\|?\s*REVIS[ÃA]O:\s*", bruto, maxsplit=1, flags=re.IGNORECASE)
        evidencias, revisao = corte[0].strip(" |"), (corte[1] if len(corte) > 1 else "").strip()

    itens = [p.strip() for p in re.split(r"\s*\+\s*", evidencias) if p.strip()]
    return {
        "status": texto(origem.get(cfg["col_status"])),
        "itens": itens,
        "revisao": revisao or texto(origem.get("Motivo Revisão")),
        "fonte": texto(origem.get("Fonte Boleto")),
        "score": texto(origem.get(cfg["col_score"])),
        "score2": texto(origem.get("2º Score")),
        "margem": texto(origem.get("Margem")),
    }


def montar_alerta(origem: dict, cfg: dict) -> dict:
    """Tipo do alerta + a NF que o drill de parcelas abre.

    A NF vem do TEXTO do alerta ("... - NF 000013303: ...") e nao da coluna
    NF/Doc Boleto: quem escreveu o alerta sabia de qual NF estava falando, e a
    coluna as vezes traz dois numeros no mesmo campo.
    """
    # A coluna do alerta muda de aba: "Alerta Fatura" nas de boleto, "Alerta" nas
    # da SEFAZ. Igual ao status, quem manda e o PAPEL declarado, nao o nome.
    bruto = texto(origem.get(cfg.get("col_alerta", "Alerta Fatura")))
    tipo = tipo_alerta(bruto)
    if not tipo:
        return {"tipo": "", "nf": ""}
    achada = re.search(r"\bNF\s+([0-9]+)", bruto, re.IGNORECASE)
    nf = nf_chave(achada.group(1)) if achada else nf_chave(origem.get("NF/Doc Boleto"))
    # O quadro que aparece ao passar o mouse: o que a base tem para justificar
    # o alerta, sem obrigar o usuario a abrir o Excel.
    return {
        "tipo": tipo,
        "nf": nf,
        "texto": bruto,
        "criterio": texto(origem.get(cfg["col_criterio"])),
        "fonte": texto(origem.get("Fonte Boleto")),
        "score": texto(origem.get(cfg["col_score"])),
        "score2": texto(origem.get("2º Score")),
        "margem": texto(origem.get("Margem")),
    }


def nf_chave(valor) -> str:
    """NF sem zeros a esquerda -- e o unico jeito de casar as duas abas."""
    return re.sub(r"\D", "", texto(valor)).lstrip("0")


def ler_parcelas(wb) -> dict[str, list[dict]]:
    """Boletos da aba 'NFs Multiplas Parcelas', agrupados por NF.

    A NF aparece em dois lugares e nem sempre no mesmo: em parte das linhas ela
    esta na coluna 'NF (9 Digitos)', em outras essa coluna vem VAZIA e a NF so
    existe dentro do 'Nosso Numero (Ref.)', no formato 0013303-01. Ignorar o
    segundo caminho deixaria a NF 13303 (3 boletos) fora do drill.
    """
    if ABA_PARCELAS not in wb.sheetnames:
        return {}

    linhas = list(wb[ABA_PARCELAS].iter_rows(values_only=True))
    if not linhas:
        return {}
    cabecalho = [texto(c) for c in linhas[0]]

    por_nf: dict[str, list[dict]] = {}
    for bruta in linhas[1:]:
        origem = dict(zip(cabecalho, bruta))
        nosso = texto(origem.get("Nosso Número (Ref.)"))
        chave = nf_chave(origem.get("NF (9 Dígitos)")) or nf_chave(nosso.split("-")[0])
        if not chave:
            continue
        digitos = re.sub(r"\D", "", texto(origem.get("Linha Digitável")))
        por_nf.setdefault(chave, []).append({
            # NF como a base mostra (9 digitos); quando a coluna vem vazia,
            # cai para a chave resolvida pelo Nosso Numero.
            "nf": texto(origem.get("NF (9 Dígitos)")) or chave,
            "parcela": texto(origem.get("Parcela")),
            "venc": texto(origem.get("Data Vencimento")),
            "emissao": texto(origem.get("Data Emissão")),
            "valor": moeda_br(numero(origem.get("Valor (R$)"))),
            "valor_n": numero(origem.get("Valor (R$)")) or 0.0,
            "fornecedor": texto(origem.get("Favorecido/Fornecedor")),
            "empresa": texto(origem.get("Empresa Pagadora")),
            "nosso": nosso,
            "linha": formatar_linha_digitavel(origem.get("Linha Digitável")),
            "copia": digitos,
        })

    for boletos in por_nf.values():
        boletos.sort(key=lambda b: (b["parcela"], b["venc"]))
    return por_nf


def ler_resumo(wb) -> dict:
    resumo = {}
    for chave, valor in wb[ABA_RESUMO].iter_rows(values_only=True):
        rotulo = texto(chave)
        if rotulo and valor is not None:
            resumo[rotulo] = valor if isinstance(valor, (int, float)) else texto(valor)
    return resumo


# Identidade do titulo sem o uuid, para alcancar as linhas da SF1 na SE2
# (14/08/2026). ⚠ Sao as chaves das CELULAS do painel, e as duas existem nos
# dois lados da aba fundida: a da SE2 traz o numero do titulo, a da SF1 traz o
# numero da nota -- que no TOTVS e o mesmo numero.
CHAVE_FORNECEDOR = chave_de("Fornecedor")
CHAVE_NO_TITULO = chave_de("No Titulo")


def conferir_na_se2(registros: list[dict], conferencia, com_tipo: bool = False) -> dict:
    """Veredito da SE2 para os titulos de uma aba, pronto para o `_meta`.

    So os titulos que SAEM da fila entram no mapa `fora` -- e a lista curta
    (dezenas), nao a aba inteira. Os outros dois numeros existem para o painel
    poder dizer em que pe esta a conferencia sem recontar nada.

    ⚠ Isto vive no `_meta` da aba, e nao dentro de cada titulo, de proposito: a
    Edge Function NAO regrava titulo que ja tem check (regra 3 da carga), entao
    um campo novo colocado na linha nunca chegaria justamente em quem ja foi
    tratado -- que sao os que precisam continuar aparecendo com o selo. O
    cabecalho `#meta:<aba>`, esse sim, e reescrito em toda carga.

    Pelo MESMO motivo o `Tipo Pgto` do titulo (coluna IK) viaja aqui, no mapa
    `tipos`, e nao em `c`: a coluna nova nasceria vazia justamente nos titulos
    que alguem ja tratou.

    ⚠ O vocabulario da SE2 (`conferencia.tipos`, 18 valores) NAO e mais enviado:
    em 14/08/2026 o menu do "tipo real" virou lista fechada, ditada pelo usuario
    (`TIPOS_INFORMAVEIS`, no painel_modelo.html). Mandar os 18 significava
    oferecer no menu os erros de digitacao que existem no TOTVS
    ("TRANFERENCIA", "TRASFERENCIA", "TRANSFERENCIA - TESTE").

    `com_tipo` vem de a aba TER a coluna, e nao de um id fixo: sem isso a aba
    Nao Associados -- que tambem confere a SE2 e nao mostra tipo nenhum --
    carregaria o mapa inteiro na carga para ninguem ler.

    ⚠ **Quem nao esta na SE2 pelo uuid e procurado pelo NUMERO do titulo**
    (14/08/2026). Isso existe pelas linhas da SF1, que entraram nesta aba com a
    fusao: o uuid delas e o da NOTA e nunca aparece na SE2, entao `olhar()`
    devolvia None para todas e as 95 escapavam do corte -- 60 pediam boleto para
    titulo que ja tinha a linha digitavel lancada e 7 ja estavam baixadas. O
    `por_nota` no `_meta` conta quantas sairam por esse caminho.
    """
    fora, tipos, abertos, sem_se2, por_nota = {}, {}, 0, 0, 0
    for registro in registros:
        situacao = conferencia.olhar(registro["uuid"])
        se2_pelo_uuid = situacao is not None
        if situacao is None:
            # Linha da SF1: o uuid dela e o da NOTA e nao existe na SE2. Acha o
            # titulo pelo NUMERO -- so assim ela passa pelo mesmo corte das
            # outras. Ver `olhar_nota`: exige a nota INTEIRA resolvida.
            situacao = conferencia.olhar_nota(registro["c"].get(CHAVE_FORNECEDOR),
                                              registro["c"].get(CHAVE_NO_TITULO))
            if situacao is not None:
                por_nota += 1
        if situacao is None:
            sem_se2 += 1          # nao esta na SE2: fica no painel, na duvida
            continue
        # o tipo vale para TODO titulo que existe na SE2, tenha ele saido da
        # fila ou nao -- ao contrario do `fora`, que so guarda os resolvidos
        if com_tipo and situacao.get("t"):
            tipos[registro["uuid"]] = situacao["t"]
        if not situacao["m"]:
            abertos += 1          # em aberto e sem boleto: e o trabalho da aba
            continue
        saida = {"m": situacao["m"], "d": situacao["d"]}
        if not se2_pelo_uuid:
            # ⚠ Achado pelo NUMERO do titulo, nao pelo uuid. Sem esta marca o
            # selo diria "OUTRO" ("lancaram um boleto diferente do indicado"),
            # que e falso: a linha da SF1 nao indicou boleto nenhum -- o `igual`
            # abaixo daria False so porque nao ha o que comparar.
            saida["nota"] = True
        if situacao["m"] == "BOLETO":
            # Lancaram o boleto que o painel indicou ou um outro? A resposta
            # muda o que a pessoa precisa conferir, entao vai junto.
            indicado = registro["copia"].get(chave_de("Linha Digitável"), "")
            saida["igual"] = bool(indicado) and situacao["dig"] == indicado
        fora[registro["uuid"]] = saida
    saida_meta = {"lido_em": conferencia.lido_em, "fora": fora,
                  "abertos": abertos, "sem_se2": sem_se2, "por_nota": por_nota}
    if com_tipo:
        saida_meta["tipos"] = tipos
    return saida_meta


# chave da coluna interna que traz o Tipo Pgto da SF1 (ver cruzamento_classificacao)
CHAVE_TIPO_SF1 = chave_de("Tipo Pgto SF1")


def meta_se2(cfg: dict, compacta: list[dict], registros: list[dict],
             conferencia) -> dict | None:
    """O `se2` do cabecalho da aba: veredito da SE2 MAIS o mapa de Tipo Pgto.

    ⚠ O Tipo Pgto das linhas da SF1 (coluna K) entra AQUI, no mesmo mapa
    `tipos` que ja carrega o da SE2 (coluna IK). Nao e atalho: a coluna da tela
    e VIRTUAL (`@tipo_pgto`) e o navegador a preenche lendo
    `aba.se2.tipos[uuid]` -- ver `processarAba` no painel_modelo.html. Por o
    valor dentro da linha nao resolveria pelo motivo de sempre: a carga nao
    regrava titulo que ja tem check, entao o campo novo nunca chegaria
    justamente em quem ja foi tratado.

    A NF da SF1 nao esta na SE2 (e nota, nao titulo), entao este e o unico
    caminho do tipo dela ate a tela.
    """
    com_tipo = any(c["tipo"] == "tipo_pgto" for c in compacta)
    saida = (conferir_na_se2(registros, conferencia, com_tipo)
             if cfg.get("conferir_se2") and conferencia else None)
    if not com_tipo:
        return saida
    da_sf1 = {r["uuid"]: r["c"][CHAVE_TIPO_SF1] for r in registros
              if r["c"].get(CHAVE_TIPO_SF1)}
    if not da_sf1:
        return saida
    if saida is None:
        # SE2 ausente ou ilegivel: o mapa de tipos ainda vale. Vai SEM `fora`
        # de proposito -- o navegador so corta linha quando esse mapa existe
        # (`if (fora)`), entao ninguem some por causa de uma leitura que falhou.
        saida = {}
    saida.setdefault("tipos", {}).update(da_sf1)
    return saida


# Teto de botoes de um grupo de filtro. Acima disso a barra vira parede -- a aba
# Nao Associados tem 40 "Motivo Revisao" distintos -- e clicar deixa de ser mais
# rapido que digitar na busca. Abaixo de 2 o botao nao filtra nada.
MIN_BOTOES_PILL = 2
MAX_BOTOES_PILL = 8


def grupos_de_pills(cfg: dict, compacta: list[dict],
                    registros: list[dict]) -> list[dict]:
    """Grupos de botoes de filtro da aba, na ordem em que aparecem na tela.

    Vem de duas origens:
    - o que a aba declara em `pills` (Fonte Boleto, Origem);
    - a coluna de STATUS (13/08/2026, pedido do usuario: "coloque um filtro de
      botao das situacao das abas"), achada pelo PAPEL `status` e nunca pelo
      nome -- ele muda de aba para aba ("Status", "Situação", "Motivo Revisão").

    ⚠ O mesmo grupo nao entra duas vezes: na Analise Base SEFAZ o papel de status
    E a coluna `Origem`, que ja e o grupo declarado. Sem esta guarda a aba ganharia
    duas barras identicas.
    """
    grupos: list[dict] = []
    candidatos = []
    # `pills` aceita uma coluna ou uma LISTA delas (a aba de pedidos declara
    # duas). A guarda de 2 a 8 valores continua valendo para cada uma: grupo que
    # nao passa simplesmente nao vira barra.
    declaradas = cfg.get("pills") or []
    if isinstance(declaradas, str):
        declaradas = [declaradas]
    for coluna in declaradas:
        candidatos.append((chave_de(coluna), coluna))
    status = next((c for c in compacta if c.get("papel") == "status"), None)
    if status:
        candidatos.append((status["chave"], status.get("rotulo") or "Situação"))

    for chave, rotulo in candidatos:
        if any(g["chave"] == chave for g in grupos):
            continue
        valores = sorted({str(r["c"].get(chave, "")) for r in registros} - {""})
        if not MIN_BOTOES_PILL <= len(valores) <= MAX_BOTOES_PILL:
            continue
        grupos.append({"chave": chave, "rotulo": rotulo, "valores": valores})
    return grupos


def gerar(base: Path, saida: Path, base_pendentes: Path | None = None,
          base_sf1: Path | None = None, base_se2: Path | None = None,
          base_sefaz: Path | None = None, base_sc7: Path | None = None,
          base_sc1: Path | None = None, base_empresas: Path | None = None,
          base_manifestacao: Path | None = None) -> dict:
    if not base.exists():
        raise FileNotFoundError(f"Base nao encontrada: {base}")
    if not MODELO.exists():
        raise FileNotFoundError(f"Modelo do painel nao encontrado: {MODELO}")

    wb, temporaria = abrir_planilha(base)
    try:
        lidas = {}
        # cabecalho+linhas cruas, por aba. A de associados fica guardada aqui
        # ate a SF1 estar pronta: as duas viram uma aba so (fundir_classificacao).
        cruas = {}
        for ident, cfg in ABAS.items():
            if cfg.get("fonte"):
                continue  # nao sai desta planilha; e lida logo abaixo
            if cfg["planilha"] not in wb.sheetnames:
                continue  # aba renomeada na base: melhor faltar do que quebrar
            cruas[ident] = ler_aba(wb, cfg)
            if ident != "associacoes":
                lidas[ident] = montar_aba(*cruas[ident], cfg)
        resumo = ler_resumo(wb)
        parcelas = ler_parcelas(wb)
        # Linhas CRUAS dos boletos que nao acharam titulo: e o acervo contra o
        # qual a aba de classificacao cruza. Tem de sair daqui, com a planilha
        # ainda aberta -- o que `lidas` guarda ja e o registro montado para a
        # tela, sem os campos crus de que o cruzamento precisa.
        boletos_nao_associados = ler_cruas(wb, ABA_FUTUROS)
    finally:
        wb.close()
        if temporaria:
            shutil.rmtree(temporaria, ignore_errors=True)

    # Aba de pendentes do Itau: outra planilha, outro arquivo de historico. Uma
    # falha aqui NAO derruba o painel inteiro -- as outras tres abas continuam
    # valendo -- mas tem de aparecer, senao a aba some em silencio e ninguem
    # descobre que ela parou de atualizar.
    pendentes_resumo = None
    caminho_pendentes = base_pendentes or pendentes_itau.BASE_PADRAO
    if caminho_pendentes.exists():
        cfg = ABAS["pendentes_itau"]
        cabecalho, brutas, pendentes_resumo = pendentes_itau.carregar(
            caminho_pendentes, HISTORICO_PENDENTES)
        if brutas:
            lidas["pendentes_itau"] = montar_aba(cabecalho, brutas, cfg)
        pendentes_resumo["base"] = str(caminho_pendentes)
        pendentes_resumo["salva_em"] = dt.datetime.fromtimestamp(
            caminho_pendentes.stat().st_mtime).strftime("%d/%m/%Y às %H:%M")
    else:
        print(f"AVISO: base de pendentes nao encontrada, aba nao atualizada:\n"
              f"       {caminho_pendentes}", file=sys.stderr)

    # Aba da classificacao: SF1 x boletos sem titulo (os do DDA lidos acima mais
    # os do e-mail do Itau). Depende do historico do Itau, entao vem DEPOIS dele.
    # Mesma regra das outras: falhar aqui nao derruba o painel, mas tem de
    # aparecer -- aba que para de atualizar em silencio e pior do que aba que
    # some.
    cruzamento_resumo = None
    caminho_sf1 = base_sf1 or cruzamento_classificacao.BASE_PADRAO
    if caminho_sf1.exists():
        cfg = ABAS["cruzamento"]
        historico_itau = {}
        if HISTORICO_PENDENTES.exists():
            historico_itau = json.loads(HISTORICO_PENDENTES.read_text(encoding="utf-8"))
        cabecalho, brutas, cruzamento_resumo = cruzamento_classificacao.carregar(
            caminho_sf1, HISTORICO_CRUZAMENTO, boletos_nao_associados, historico_itau)
        if brutas:
            # A aba continua sendo MONTADA e enviada, so que oculta: e o
            # cabecalho dela que apaga do painel a aba que ja esta no banco.
            lidas["cruzamento"] = montar_aba(cabecalho, brutas, cfg)
            cruas["cruzamento"] = (cabecalho, brutas)
        cruzamento_resumo["base"] = str(caminho_sf1)
        cruzamento_resumo["salva_em"] = dt.datetime.fromtimestamp(
            caminho_sf1.stat().st_mtime).strftime("%d/%m/%Y às %H:%M")
        cruzamento_resumo["acervo"] = len(boletos_nao_associados) + len(
            (historico_itau or {}).get("boletos", {}))
    else:
        print(f"AVISO: base SF1 nao encontrada, aba de classificacao nao atualizada:\n"
              f"       {caminho_sf1}", file=sys.stderr)

    # Fusao das duas analises numa aba so (14/08/2026). Acontece AQUI, e nao
    # junto com a leitura da planilha, porque a SF1 depende do historico do Itau
    # -- so agora as duas fontes existem. Sem a SF1 a aba sai so com os titulos
    # da SE2, que e o comportamento antigo: falta de base nunca derruba a aba.
    if "associacoes" in cruas:
        cab_base, brutas_base = cruas["associacoes"]
        cab_sf1, brutas_sf1 = cruas.get("cruzamento", ([], []))
        lidas["associacoes"] = montar_aba(
            *fundir_classificacao(cab_base, brutas_base, cab_sf1, brutas_sf1),
            ABAS["associacoes"])

    # Abas da SEFAZ: a base como ela e (so as colunas marcadas) e o cruzamento
    # dela com a SF1. Mesma regra das outras bases de fora -- falhar aqui nao
    # derruba o painel, mas tem de aparecer na tela do .cmd.
    sefaz_resumo = cruz_sefaz_resumo = pedidos_resumo = None
    caminho_sefaz = base_sefaz or sefaz.BASE_PADRAO
    # 24/08/2026: so entram na analise da SEFAZ as notas cujo TOMADOR esta na
    # LISTAGEM EMPRESAS BIOFLOR. A listagem e lida UMA vez aqui e desce para as
    # tres abas junto com as linhas -- ver sefaz.ler_cnpjs_empresas().
    # ⚠ Sem a listagem as abas da SEFAZ NAO sao atualizadas, de proposito: gerar
    # sem filtro traria de volta as notas das outras empresas do grupo, e o
    # painel diria que sao da Bioflor. Fica com o dado da rodada anterior e o
    # aviso na tela, como quando a propria SEFAZ.xlsx nao esta la.
    empresas_bioflor = None
    erro_empresas = None
    try:
        empresas_bioflor = sefaz.ler_cnpjs_empresas(base_empresas)
    except Exception as exc:  # noqa: BLE001 - listagem fora do ar/renomeada/mudada
        erro_empresas = exc
    if not caminho_sefaz.exists():
        print(f"AVISO: base SEFAZ nao encontrada, abas nao atualizadas:\n"
              f"       {caminho_sefaz}", file=sys.stderr)
    elif erro_empresas is not None:
        print(f"AVISO: abas da SEFAZ nao atualizadas (sem a listagem das empresas "
              f"BIOFLOR nao da para saber quais tomadores entram):\n"
              f"       {erro_empresas}", file=sys.stderr)
    else:
        # ⚠ A ORDEM AQUI IMPORTA: o cruzamento com a SF1 roda ANTES da aba da
        # base, porque e ele quem sabe quais notas nao foram lancadas -- e os
        # alertas de "vence em 5 dias sem lancamento" e "emitida ha mais de 3
        # dias sem lancamento" dependem disso. Sem a SF1 os dois nao aparecem em
        # lugar nenhum, e so o de prazo curto (que olha apenas a nota) continua.
        #
        # O cruzamento precisa das DUAS bases; sem a SF1 a aba simplesmente nao
        # existe (em vez de existir dizendo que nada foi achado, o que seria
        # mentira -- ninguem procurou).
        # A base e lida UMA vez aqui, e ja com o historico: `linhas_sefaz` e a
        # base de hoje MAIS as notas que sumiram dela (marcadas). As tres abas
        # que leem a SEFAZ recebem esse mesmo conjunto -- se cada uma relesse o
        # arquivo, as removidas apareceriam em nenhuma.
        linhas_sefaz, sefaz_hist = sefaz.ler_com_historico(
            caminho_sefaz, HISTORICO_SEFAZ, permitidos=empresas_bioflor)
        print(f"           tomadores BIOFLOR: {sefaz_hist['filtro_empresas']} empresas | "
              f"{sefaz_hist['filtro_notas']} de {sefaz_hist['filtro_notas_lidas']} notas "
              f"entraram ({sefaz_hist['filtro_notas_fora']} de outras "
              f"{sefaz_hist['filtro_cnpjs_fora']} empresas ficaram de fora)")
        if sefaz_hist.get("expurgadas"):
            print(f"           historico: {sefaz_hist['expurgadas']} nota(s) de fora da "
                  f"listagem sairam do historico (nao viram alerta de removida)")
        print(f"           notas: {sefaz_hist['na_base']} na base | "
              f"{sefaz_hist['novas']} NOVAS | {sefaz_hist['sairam_hoje']} sairam hoje | "
              f"{sefaz_hist['voltaram']} voltaram")
        if sefaz_hist["removidas"]:
            print(f"           ⚠ {sefaz_hist['removidas']} nota(s) REMOVIDAS da base "
                  f"continuam no painel, com alerta")
        # 01/09/2026. A manifestacao do destinatario entra AQUI, entre a leitura
        # e as abas: assim as tres abas que leem a SEFAZ enxergam a coluna, e o
        # sefaz_historico.json continua guardando a linha PURA -- manifestacao e
        # fato de hoje, relido a cada rodada. Congelada no historico, a nota
        # removida carregaria para sempre o status do dia em que sumiu.
        # ⚠ Sem a base, a coluna diz que NAO FOI LIDA, e nao fica vazia: vazio
        # seria lido como "esta nota nao tem manifestacao", afirmacao que ninguem
        # fez. O painel inteiro continua saindo -- so esta coluna fica sem dado.
        caminho_manifesto = base_manifestacao or manifestacao.BASE_PADRAO
        try:
            manifesto_resumo = manifestacao.anotar(linhas_sefaz, caminho_manifesto)
            faixa = ""
            if manifesto_resumo["de"]:
                faixa = (f" | período {manifesto_resumo['de'].strftime('%d/%m/%Y')}"
                         f" a {manifesto_resumo['ate'].strftime('%d/%m/%Y')}")
            print(f"           manifestação: {manifesto_resumo['notas']} notas no "
                  f"arquivo ({manifesto_resumo['linhas']} linhas, "
                  f"{manifesto_resumo['repetidas']} repetidas){faixa}")
            print("           " + " | ".join(
                f"{n} {s}" for s, n in manifesto_resumo["anotadas"].items()))
            if manifesto_resumo["graves"]:
                print(f"           ⚠ {manifesto_resumo['graves']} nota(s) com "
                      f"DESCONHECIMENTO/NÃO REALIZADA — há boleto ligado a nota "
                      f"que a empresa não reconhece")
        except Exception as exc:  # noqa: BLE001 - base fora do ar/renomeada/mudada
            manifestacao.marcar_sem_base(linhas_sefaz)
            print(f"AVISO: coluna de manifestação não preenchida:\n       {exc}",
                  file=sys.stderr)
        print(f"           historico: {sefaz_hist['conhecidas']} notas conhecidas "
              f"-> {HISTORICO_SEFAZ.name}")

        nao_lancadas = None
        if caminho_sf1.exists():
            # ⚠ `preparar()` le SEFAZ + SF1 UMA vez e as duas abas que dependem
            # do cruzamento reusam o resultado. Sem isso, a aba de pedidos releria
            # os 8.011 titulos da SF1 so para mostrar o mesmo veredito ao lado.
            pronto = cruzamento_sefaz_sf1.preparar(caminho_sefaz, caminho_sf1,
                                                   linhas=linhas_sefaz)
            cabecalho, brutas, cruz_sefaz_resumo = cruzamento_sefaz_sf1.carregar(
                caminho_sefaz, caminho_sf1, pronto=pronto)
            if brutas:
                lidas["sefaz_sf1"] = montar_aba(cabecalho, brutas, ABAS["sefaz_sf1"])
            # `pop`: e um set, e o resumo vai para JSON -- set nao serializa.
            nao_lancadas = cruz_sefaz_resumo.pop("nao_lancadas", None)

            # Aba dos pedidos de compra: precisa do cruzamento acima (leva o
            # veredito da SF1 e os alertas) MAIS a SC7/SC1. Mesma regra das
            # outras bases de fora: falhar aqui nao derruba o painel, mas tem de
            # aparecer -- aba que para de atualizar calada e o pior dos casos.
            # ⚠ A MESMA variavel na conferencia e na leitura: conferir
            # `sc7.BASE_SC7` e passar outro caminho adiante faria o AVISO falar
            # de um arquivo que ninguem ia ler.
            caminho_sc7 = base_sc7 or sc7.BASE_SC7
            caminho_sc1 = base_sc1 or sc7.BASE_SC1
            # ⚠ A SC1 e OPCIONAL para o sc7.carregar() (sem ela os pedidos
            # continuam valendo e so o Solicitante fica vazio) -- por isso ela
            # nao entra no `if` abaixo. Mas "opcional" nao pode virar "calada":
            # a coluna vazia e visualmente igual a um pedido sem solicitante.
            if caminho_sc7.exists() and not caminho_sc1.exists():
                print(f"AVISO: SC1 nao encontrada -- a aba de pedidos sai com a "
                      f"coluna Solicitante VAZIA:\n       {caminho_sc1}",
                      file=sys.stderr)
            if caminho_sc7.exists():
                try:
                    cabecalho, brutas, pedidos_resumo = pedidos_sefaz.carregar(
                        caminho_sefaz, caminho_sf1, base_sc7=caminho_sc7,
                        base_sc1=caminho_sc1, pronto=pronto)
                    if brutas:
                        lidas["pedidos"] = montar_aba(cabecalho, brutas, ABAS["pedidos"])
                except Exception as exc:  # noqa: BLE001 - planilha aberta, coluna movida...
                    print(f"AVISO: aba de pedidos de compra nao atualizada: {exc}",
                          file=sys.stderr)
            else:
                print(f"AVISO: SC7 nao encontrada, aba de pedidos nao criada:\n"
                      f"       {caminho_sc7}", file=sys.stderr)

        else:
            # ⚠ Sem a SF1 nao existe cruzamento, e sem cruzamento nao existem NEM
            # a aba SEFAZ x SF1 NEM a de pedidos de compra -- as duas moram
            # dentro do `if` acima. Ate 31/08/2026 isso acontecia EM SILENCIO: o
            # unico AVISO de SF1 e o da aba de classificacao, que sai bem antes e
            # fala de outra aba. Quem so olhasse a tela juraria que rodou tudo.
            print(f"AVISO: sem a SF1 as abas SEFAZ x SF1 e NF x Pedido de Compra "
                  f"NAO foram atualizadas:\n       {caminho_sf1}", file=sys.stderr)

        cabecalho, brutas, sefaz_resumo = sefaz.carregar(caminho_sefaz, nao_lancadas,
                                                         linhas=linhas_sefaz)
        sefaz_resumo.update({f"hist_{k}": v for k, v in sefaz_hist.items()})
        if brutas:
            lidas["sefaz"] = montar_aba(cabecalho, brutas, ABAS["sefaz"])
        sefaz_resumo["base"] = str(caminho_sefaz)
        sefaz_resumo["salva_em"] = dt.datetime.fromtimestamp(
            caminho_sefaz.stat().st_mtime).strftime("%d/%m/%Y às %H:%M")

    # Conferencia na SE2 (pedido do usuario, 13/08/2026): so segue na fila o
    # titulo em aberto e sem boleto lancado. Mesma regra das outras bases de
    # fora -- se ela nao estiver aqui, o painel continua funcionando inteiro,
    # mas o aviso tem de aparecer: sem a SE2 o painel volta a mostrar titulo
    # que ja foi resolvido, e ninguem descobriria sozinho.
    conferencia_se2 = None
    caminho_se2 = base_se2 or verificacao_se2.BASE_PADRAO
    if not caminho_se2.exists():
        print(f"AVISO: base SE2 nao encontrada, titulos NAO conferidos:\n"
              f"       {caminho_se2}", file=sys.stderr)
    else:
        try:
            conferencia_se2 = verificacao_se2.ler(caminho_se2)
        except Exception as exc:  # noqa: BLE001 - planilha aberta, coluna renomeada...
            # Sem conferencia o painel volta a mostrar todo mundo -- que e o
            # comportamento antigo, e o certo aqui: esconder titulo por causa de
            # uma leitura que falhou seria sumir com trabalho de verdade.
            print(f"AVISO: nao consegui ler a SE2, titulos NAO conferidos: {exc}\n"
                  f"       {caminho_se2}", file=sys.stderr)

    def do_resumo(rotulo):
        valor = resumo.get(rotulo, 0)
        return int(valor) if isinstance(valor, (int, float)) else 0

    abas = []
    for ident, cfg in ABAS.items():
        if ident not in lidas:
            continue
        # `colunas` nao vai para o HTML: o painel so desenha `compacta`. A lista
        # continua servindo para montar as celulas de cada registro.
        _colunas, compacta, registros = lidas[ident]
        # Aba oculta viaja SO com o cabecalho, sem uma linha sequer: e o
        # cabecalho que apaga do painel a aba que ja esta no banco. Zerar aqui
        # tambem tira o peso dela da carga (a SEFAZ x SF1 eram 617 linhas).
        oculta = bool(cfg.get("oculta"))
        if oculta:
            registros = []
        abas.append({
            "id": ident,
            "nome": cfg["nome"],
            "oculta": oculta,
            # aba do .xlsx exportado (o Excel corta em 31 caracteres)
            "guia": cfg.get("guia", cfg["nome"])[:31],
            "cor": cfg["cor"],
            "compacta": compacta,
            "linhas": registros,
            # quando existe, o painel abre esta aba em duas (ver ABAS)
            "particoes": cfg.get("particoes"),
            # rotulos da linha de resumo; None em um deles esconde o item
            "resumo": cfg.get("resumo"),
            # Botoes de filtro: LISTA de grupos (era um grupo so). Ver
            # grupos_de_pills -- a coluna declarada em `pills` e a de status.
            "pills": grupos_de_pills(cfg, compacta, registros),
            # Veredito da SE2 (so nas abas que declaram `conferir_se2`). None
            # quando a base nao foi lida -- o painel entende como "nao conferi"
            # e nao esconde ninguem.
            "se2": meta_se2(cfg, compacta, registros, conferencia_se2),
            "ocr": sum(1 for r in registros if r["ocr"]),
            # pelo PAPEL, nao pela chave: nas abas da SEFAZ a coluna do alerta se
            # chama "Alerta" (chave `alerta`), nao "Alerta Fatura"
            "com_alerta": sum(1 for r in registros if r["alerta"]["tipo"]),
            "divergentes": sum(1 for r in registros if r["difere"]),
            # So as NFs que algum alerta de parcela realmente abre -- nao adianta
            # levar todas as NFs da aba se o painel so tem 3 alertas.
            "parcelas": {nf: parcelas[nf] for nf in
                         {r["alerta"]["nf"] for r in registros
                          if r["alerta"]["tipo"] == "PARCELA" and r["alerta"]["nf"]}
                         if nf in parcelas},
        })

    dados = {
        "gerado_em": resumo.get("Gerado em") or dt.date.today().strftime("%d/%m/%Y"),
        "atualizado_em": dt.datetime.now().strftime("%d/%m/%Y às %H:%M"),
        # quando a PLANILHA foi salva: e o que diz se a carga esta velha ou nao
        "salva_em": dt.datetime.fromtimestamp(base.stat().st_mtime).strftime("%d/%m/%Y às %H:%M"),
        "sem_codigo": do_resumo("Títulos sem código de barras"),
        "abas": abas,
    }
    carga = json.dumps(dados, ensure_ascii=False, separators=(",", ":"))
    modelo = MODELO.read_text(encoding="utf-8")

    # O motor da planilha entra nos DOIS (dev e publicado): e so codigo, sem
    # nenhum dado, e o botao Exportar depende dele.
    if "<!--__PLANILHA__-->" not in modelo:
        raise RuntimeError("falta o marcador <!--__PLANILHA__--> no painel_modelo.html")
    modelo = modelo.replace(
        "<!--__PLANILHA__-->",
        planilha_excel_js.bloco_autonomo("PLANILHA XLSX (mesmo motor do painel SE2)"))

    # Publicacao: o HTML sai SEM a carteira (mesmo modelo dos outros paineis --
    # o Pages serve publicamente mesmo com o repo privado). A carteira vai para
    # um JSON a parte, que o supa.js entrega so depois do login.
    saida.parent.mkdir(parents=True, exist_ok=True)
    # O supa.js entra SO no publicado: no dev.html a carteira ja vem embutida e
    # uma tela de login local so atrapalharia. O ?v= e cache-bust -- sem ele o
    # navegador serve o supa.js velho depois de qualquer ajuste.
    versao = dt.datetime.now().strftime("%Y%m%d%H%M")
    saida.write_text(
        modelo.replace("<!--__SUPA__-->", f'<script src="supa.js?v={versao}"></script>'),
        encoding="utf-8")
    conferir_sem_dados(saida, dados)

    DADOS_JSON.parent.mkdir(parents=True, exist_ok=True)
    DADOS_JSON.write_text(carga, encoding="utf-8")

    # Uso local: dev.html com tudo embutido, fora do repositorio (.gitignore).
    DEV.write_text(modelo.replace("/*__DADOS__*/null", carga), encoding="utf-8")

    # So para o relatorio no terminal -- entra DEPOIS do json.dumps de proposito,
    # para nao viajar no HTML nem na carga do banco.
    dados["_pendentes_itau"] = pendentes_resumo
    dados["_cruzamento"] = cruzamento_resumo
    dados["_sefaz"] = sefaz_resumo
    dados["_sefaz_sf1"] = cruz_sefaz_resumo
    dados["_pedidos"] = pedidos_resumo
    dados["_se2"] = ({"base": str(caminho_se2), "salva_em": conferencia_se2.lido_em,
                      "total": conferencia_se2.total, "dias": conferencia_se2.dias}
                     if conferencia_se2 else None)
    return dados


# Pedacos que so podem existir no dev.html. Se aparecerem no publicado, a
# geracao para -- e a mesma trava que o painel do Geraldo tem.
def conferir_sem_dados(arquivo: Path, dados: dict) -> None:
    html = arquivo.read_text(encoding="utf-8")
    suspeitos = []
    # Nao basta olhar UUID e linha digitavel: ja escapou nome de fornecedor
    # escrito a mao dentro de um aviso do proprio painel.
    campos = ("razao_social", "fornecedor", "fornecedor_boleto", "nome_fornece",
              "cnpj_boleto", "cnpj_cpf", "historico", "empresa_origem",
              # aba NFs pendentes Itau: nome do beneficiario, empresa pagadora e
              # o CNPJ dele -- exatamente o tipo de dado que nao pode ir ao Pages
              "beneficiario", "pagador_agregado", "cpf_cnpj")
    for aba in dados["abas"]:
        for registro in aba["linhas"]:
            valores = [registro["uuid"],
                       registro["copia"].get(chave_de("Linha Digitável"), "")]
            valores += [registro["c"].get(c, "") for c in campos]
            for valor in valores:
                # nomes curtos dariam falso positivo com palavra do proprio HTML
                if valor and len(str(valor)) > 8 and str(valor) in html:
                    suspeitos.append(valor)
    if "/*__DADOS__*/null" not in html:
        suspeitos.append("marcador /*__DADOS__*/null ausente")
    if suspeitos:
        arquivo.unlink(missing_ok=True)
        raise RuntimeError(
            f"{arquivo.name} sairia com dados dentro ({len(suspeitos)} ocorrencia(s), "
            f"ex.: {suspeitos[0][:40]}). Arquivo apagado -- nada foi publicado."
        )


def conferir_motor_planilha() -> str:
    """Avisa se a copia local do motor .xlsx saiu de sincronia com a do SE2."""
    if not MOTOR_PLANILHA_FONTE.exists():
        return ""      # a pasta do SE2 nao esta aqui: segue com a copia local
    local = (PASTA / "planilha_excel_js.py").read_bytes()
    if local == MOTOR_PLANILHA_FONTE.read_bytes():
        return ""
    return ("AVISO: planilha_excel_js.py esta diferente do original do SE2\n"
            f"       ({MOTOR_PLANILHA_FONTE}).\n"
            "       A planilha exportada pode nao sair igual a do SE2.")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base", type=Path, default=BASE_PADRAO)
    parser.add_argument("--base-pendentes", type=Path, default=pendentes_itau.BASE_PADRAO,
                        help="planilha do DDA do Itau (COPIAR E COLAR BOLETOS PENDENTES)")
    parser.add_argument("--base-sf1", type=Path, default=cruzamento_classificacao.BASE_PADRAO,
                        help="base SF1 (notas classificadas no TOTVS)")
    parser.add_argument("--base-se2", type=Path, default=verificacao_se2.BASE_PADRAO,
                        help="base SE2 (posicao diaria) usada para conferir os titulos")
    parser.add_argument("--base-sefaz", type=Path, default=sefaz.BASE_PADRAO,
                        help="base SEFAZ.xlsx (abas NFes SEFAZ e NFS)")
    # As tres da aba NF x Pedido de Compra. Ate 31/08/2026 so existiam como
    # constante no modulo: numa maquina onde a biblioteca do OneDrive tem outro
    # nome, ou com a exportacao salva noutra pasta, nao havia como apontar o
    # caminho sem editar o .py.
    parser.add_argument("--base-sc7", type=Path, default=sc7.BASE_SC7,
                        help="base SC7 (pedidos de compra, por item)")
    parser.add_argument("--base-sc1", type=Path, default=sc7.BASE_SC1,
                        help="base SC1 (solicitacoes; e dela que vem o Solicitante)")
    parser.add_argument("--base-empresas", type=Path, default=sefaz.BASE_EMPRESAS,
                        help="LISTAGEM EMPRESAS BIOFLOR.xlsx (quais tomadores entram "
                             "na analise da SEFAZ)")
    parser.add_argument("--base-manifestacao", type=Path,
                        default=manifestacao.BASE_PADRAO,
                        help="MESMA PREMISSA.xlsx (manifestacao do destinatario; "
                             "so preenche a coluna, nao acrescenta nota)")
    parser.add_argument("--saida", type=Path, default=SAIDA_PADRAO)
    args = parser.parse_args()

    try:
        dados = gerar(args.base, args.saida, args.base_pendentes, args.base_sf1,
                      args.base_se2, args.base_sefaz, args.base_sc7,
                      args.base_sc1, args.base_empresas, args.base_manifestacao)
    except Exception as exc:  # noqa: BLE001 - mensagem amigavel no .cmd
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1

    # a base lida vem primeiro: e a pergunta que mais aparece ("de onde saiu?")
    salva = dt.datetime.fromtimestamp(args.base.stat().st_mtime)
    print(f"Base lida: {args.base}")
    print(f"           (planilha salva em {salva.strftime('%d/%m/%Y as %H:%M')})")

    itau = dados.pop("_pendentes_itau", None)
    if itau:
        print(f"Base do Itau: {itau['base']}")
        print(f"           (planilha salva em {itau['salva_em']})")
        print(f"           {itau['na_planilha']} na planilha de hoje | "
              f"{itau['novos']} NOVOS | {itau['sairam_hoje']} sairam | "
              f"{itau['voltaram']} voltaram")
        print(f"           historico: {itau['total']} boletos "
              f"({itau['pendentes']} ainda pendentes) -> {HISTORICO_PENDENTES.name}")

    cruz = dados.pop("_cruzamento", None)
    if cruz:
        dias = ", ".join(cruz["dias_lidos"]) or "nenhum dia novo"
        print(f"Base SF1: {cruz['base']}")
        print(f"           (planilha salva em {cruz['salva_em']})")
        print(f"           classificacao lida: {dias} (limite: {cruz['limite']})")
        print(f"           {cruz['novos']} NOVOS | cruzados contra {cruz['acervo']} boletos sem titulo")
        print(f"           {cruz['com_boleto']} com boleto | {cruz['sem_boleto']} sem boleto"
              f" | {cruz['selos']}")
        print(f"           historico: {cruz['total']} titulos -> {HISTORICO_CRUZAMENTO.name}")

    sfz = dados.pop("_sefaz", None)
    if sfz:
        print(f"Base SEFAZ: {sfz['base']}")
        print(f"           (planilha salva em {sfz['salva_em']})")
        print(f"           {sfz['nfe']} linhas de NF-e ({sfz['notas_nfe']} notas) | "
              f"{sfz['nfs']} de NFS-e ({sfz['notas_nfs']} notas) | "
              f"{sfz['com_duplicata']} com duplicata")
    cruz_sfz = dados.pop("_sefaz_sf1", None)
    if cruz_sfz:
        print(f"           SEFAZ x SF1: {cruz_sfz['notas']} notas contra {cruz_sfz['sf1']} títulos"
              f" -> {cruz_sfz['achadas']} achadas | {cruz_sfz['conferir']} conferir"
              f" | {cruz_sfz['nao_achadas']} nao achadas")

    se2 = dados.pop("_se2", None)
    if se2:
        print(f"Base SE2: {se2['base']}")
        print(f"           (planilha salva em {se2['salva_em']} | {se2['total']} titulos)")
        if se2["dias"] >= 2:
            print(f"           AVISO: a SE2 tem {se2['dias']} dias. Exporte de novo antes de "
                  f"confiar em quem saiu do painel.", file=sys.stderr)
    print()

    for aba in dados["abas"]:
        print(f"{aba['nome']}: {len(aba['linhas'])} linhas | {len(aba['compacta'])} colunas"
              f" | com alerta: {aba['com_alerta']} | boleto difere: {aba['divergentes']}"
              f" | NFs com drill: {len(aba['parcelas'])}")
        conferido = aba.get("se2")
        if conferido:
            motivos = [m["m"] for m in conferido["fora"].values()]
            print(f"           SE2: {conferido['abertos']} em aberto sem boleto (ficam)"
                  f" | {motivos.count('BOLETO')} ja com boleto lancado"
                  f" | {motivos.count('BAIXA')} baixados"
                  f" | {conferido['sem_se2']} fora da SE2 (ficam)")
    print(f"Publicavel (sem dados): {args.saida}")
    print(f"Carteira para subir:    {DADOS_JSON}")
    print(f"Painel local (com dados): {DEV}")
    aviso = conferir_motor_planilha()
    if aviso:
        print(f"\n{aviso}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
