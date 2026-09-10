#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""caminhos.py - achar as bases sem depender de ONDE esta a pasta do painel.

Criado em 10/09/2026, quando a pasta do painel saiu de

    ...\\LUCAS ABNER ARAUJO\\AUTOMACOES LUCAS\\ANALISES BOLETOS\\PAINEL ANALISE BOLETOS
para
    ...\\LUCAS ABNER ARAUJO\\PAINEIS\\ANALISE BOLETOS

Ate aqui cada modulo achava as planilhas CONTANDO niveis a partir do proprio
arquivo: `parents[1]` para a pasta ANALISES BOLETOS (SEFAZ.xlsx, MESMA
PREMISSA.xlsx) e `parents[3]` para a LUCAS ABNER ARAUJO (BASES GENERICOS). A
conta amarra o script a UM lugar da arvore -- mudou de pasta, ela aponta para o
pai errado. E o pior nao e o erro: e o silencio. Base que nao e' achada nao
derruba a rodada, so faz a aba nascer vazia com um AVISO no meio do log, que e
justamente o que ninguem le.

Aqui as pastas sao PROCURADAS por marca, e nao contadas:

  raiz_lucas()      sobe de pasta em pasta ate achar a que tem "BASES
                    GENERICOS" dentro -- essa e a LUCAS ABNER ARAUJO
  bases_genericos() a "BASES GENERICOS" de dentro dela
  pasta_analises()  a "ANALISES BOLETOS": primeiro a vizinha (quando o painel
                    ainda mora dentro dela), depois procurada a partir da raiz

Vale no lugar velho e no novo, e nas DUAS MAQUINAS: a busca sai da pasta deste
arquivo, e nao de "C:\\Users\\<fulano>". O caminho fixo continua existindo, mas
como ULTIMO recurso -- para quem tirar o script da biblioteca compartilhada.

A pasta "AUTOMACOES LUCAS" nunca e' escrita por extenso: o nome real tem
cedilha e til, e escreve-lo aqui deixaria o acerto por conta da codificacao com
que o arquivo for lido. Sempre curinga (`AUTOMA*`).
"""
from __future__ import annotations

from pathlib import Path

# Quantas pastas subir procurando a marca. 8 cobre com folga qualquer lugar
# dentro da biblioteca; sem teto, um caminho fora dela viraria uma subida ate a
# raiz do disco a cada import.
_TETO = 8

_MARCA_RAIZ = "BASES GENERICOS"
_NOME_ANALISES = "ANALISES BOLETOS"

_FIXO_RAIZ = Path(
    r"C:\Users\lucas\OneDrive - Grupo S&D\Arquivos de Gabriella Karla Oliveira Milas"
    r" - FINANCEIRO COMPARTILHADO\LUCAS ABNER ARAUJO"
)

_AQUI = Path(__file__).resolve().parent


def raiz_lucas(partida: Path | None = None) -> Path:
    """A pasta LUCAS ABNER ARAUJO -- a que tem a BASES GENERICOS dentro."""
    atual = (partida or _AQUI).resolve()
    for _ in range(_TETO):
        if (atual / _MARCA_RAIZ).is_dir():
            return atual
        if atual.parent == atual:
            break
        atual = atual.parent
    return _FIXO_RAIZ


def bases_genericos(partida: Path | None = None) -> Path:
    """...\\LUCAS ABNER ARAUJO\\BASES GENERICOS (SF1, SC7, SC1, SE2, empresas)."""
    return raiz_lucas(partida) / _MARCA_RAIZ


def pasta_analises(partida: Path | None = None) -> Path:
    """A pasta ANALISES BOLETOS -- onde ficam SEFAZ.xlsx e MESMA PREMISSA*.xlsx.

    Ordem: a pasta acima desta (vale enquanto o painel morar dentro dela), a
    propria pasta do painel (para quem trouxer as planilhas junto), e por fim a
    procura a partir da raiz, um nivel abaixo dela em qualquer pasta -- que e
    como ela e achada agora que o painel mora em PAINEIS\\ANALISE BOLETOS.
    """
    inicio = (partida or _AQUI).resolve()
    vizinhas = [inicio.parent, inicio]
    for pasta in vizinhas:
        if pasta.name.upper() == _NOME_ANALISES and pasta.is_dir():
            return pasta
    for pasta in vizinhas:
        if (pasta / "SEFAZ.xlsx").is_file():
            return pasta

    raiz = raiz_lucas(inicio)
    # "AUTOMA*/ANALISES BOLETOS" e nao "**/...": o glob fundo varreria a
    # biblioteca inteira (dezenas de milhares de arquivos) a cada import.
    achadas = sorted(raiz.glob(f"AUTOMA*/{_NOME_ANALISES}"))
    if not achadas:
        achadas = sorted(p for p in raiz.glob(f"*/{_NOME_ANALISES}") if p.is_dir())
    if achadas:
        return achadas[0]
    return raiz / "AUTOMACOES LUCAS" / _NOME_ANALISES


if __name__ == "__main__":
    print("pasta deste arquivo :", _AQUI)
    print("raiz LUCAS ABNER    :", raiz_lucas())
    print("BASES GENERICOS     :", bases_genericos(),
          "OK" if bases_genericos().is_dir() else "NAO ACHEI")
    print("ANALISES BOLETOS    :", pasta_analises(),
          "OK" if pasta_analises().is_dir() else "NAO ACHEI")
