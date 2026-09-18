# -*- coding: utf-8 -*-
"""Manda UM e-mail pelo Outlook classico. Existe para ser um PROCESSO.

Nao e' uma funcao de proposito: o `enviar()` do alertas_comuns precisa poder
DESISTIR. Uma chamada COM presa na caixa de permissao do Outlook ("Programa
tentando enviar email em seu nome") nao volta nunca, e thread travada dentro
do COM nao se mata -- processo, sim.

Le do stdin um JSON {"para","assunto","corpo"}: o endereco da pessoa nao passa
pela linha de comando (aparece no Gerenciador de Tarefas) nem por arquivo
temporario. Este repositorio e' publico -- nada de endereco aqui dentro.
"""
from __future__ import annotations

import json
import sys


def main() -> int:
    # 18/09/2026 -- o stdin do filho nasce no encoding do Windows (cp1252) e
    # o pai escreve UTF-8: json.load(sys.stdin) devolvia "tAtulo" no lugar de
    # "titulo" e o acento chegava torto no e-mail da Graziela. Ler os BYTES e
    # decodificar na mao nao depende de locale nem de PYTHONIOENCODING.
    pedido = json.loads(sys.stdin.buffer.read().decode("utf-8"))
    import win32com.client  # so aqui: em maquina sem Outlook o resto ainda roda

    outlook = win32com.client.Dispatch("Outlook.Application")
    email = outlook.CreateItem(0)  # 0 = olMailItem
    email.To = pedido["para"]
    # CC opcional: alerta que e' de um time mas precisa de testemunha manda
    # copia. Vazio ou ausente = sem copia, como sempre foi.
    if pedido.get("copia"):
        email.CC = pedido["copia"]
    email.Subject = pedido["assunto"]
    email.HTMLBody = pedido["corpo"]
    email.Send()
    # Com o Outlook fechado a mensagem fica na Caixa de Saida e o script
    # mentiria "enviado"; o SendAndReceive empurra de verdade.
    try:
        outlook.GetNamespace("MAPI").SendAndReceive(False)
    except Exception as erro:  # noqa: BLE001 - avisar e seguir
        print(f"AVISO: nao consegui forcar o envio ({erro}).")
        print("Se o Outlook estiver fechado, o e-mail sai quando ele abrir.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
