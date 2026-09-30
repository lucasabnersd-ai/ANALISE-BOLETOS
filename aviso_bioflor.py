#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""aviso_bioflor.py - tela de aviso no padrao S & D BIOFLOR.

Pedido dele em 30/09/2026: "HABILITE UMA TELA DE NOTIFICACAO NO PADRAO BIOFLOR
PARA INFORMAR QUE O PAINEL FOI ATUALIZADO". Substitui a caixa cinza do Windows
(MessageBox) que o gatilho_bases_do_dia.py mostrava no fim da rodada.

Visual = a porta dos paineis (feedback_modelo_tela_entrada): fundo #1B302F,
caixa #24403E com borda #3A5654, pastilha BRANCA com a logo oficial
(PUBLICAR\\logo-sd.png, sem recorte nem recolorir), titulo #F2F5F1, texto
#A9BDB9, botao limao #D5DF66 "Abrir painel", Calibri em tudo. Erro em #F4B6B6.

Fica por cima de tudo ate alguem fechar (Fechar, Esc ou o X) -- ele pode estar
longe do micro quando a rodada termina, e aviso que some sozinho ninguem ve.
Da para arrastar pela caixa.

Uso (o gatilho chama num processo solto, com pythonw):
    pythonw aviso_bioflor.py '{"status":"ok","titulo":"...","linhas":[["SC7","..."]],
                                "rodape":"...","url":"https://..."}'
Teste:
    python aviso_bioflor.py --exemplo
"""
from __future__ import annotations

import json
import sys
import webbrowser
from pathlib import Path

AQUI = Path(__file__).resolve().parent
LOGO = AQUI / "PUBLICAR" / "logo-sd.png"

FUNDO = "#1B302F"
CAIXA = "#24403E"
BORDA = "#3A5654"
TITULO = "#F2F5F1"
TEXTO = "#A9BDB9"
LIMAO = "#D5DF66"
LIMAO_ESCURO = "#C0D15E"
ERRO = "#F4B6B6"
FONTE = "Calibri"

EXEMPLO = {
    "status": "ok",
    "titulo": "PAINEL ANÁLISE DE BOLETOS ATUALIZADO",
    "subtitulo": "Atualizado sozinho depois da SC7, SF1 e SE2 de hoje.",
    "linhas": [
        ["SC7", "30/09 10:13 · dados novos"],
        ["SF1", "30/09 09:38 · dados novos"],
        ["SE2", "30/09 09:28 · dados novos"],
        ["ITAÚ RET", "importado antes do painel"],
        ["RODADA", "14:17 → 14:46 (28 min 51 s)"],
    ],
    "rodape": "Nada mais está rodando.",
    "url": "https://analise-boletos.vercel.app/",
}


def mostrar(dados: dict) -> None:
    import tkinter as tk

    try:  # nitidez em tela com escala (125%/150%)
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    erro = dados.get("status") == "erro"
    raiz = tk.Tk()
    raiz.title("S & D BIOFLOR")
    raiz.configure(bg=FUNDO)
    raiz.overrideredirect(True)  # sem a barra cinza do Windows: e a porta dos paineis
    raiz.attributes("-topmost", True)

    moldura = tk.Frame(raiz, bg=FUNDO, padx=14, pady=14)
    moldura.pack()
    caixa = tk.Frame(moldura, bg=CAIXA, highlightbackground=BORDA, highlightthickness=1,
                     padx=28, pady=24)
    caixa.pack()

    # Selo: pastilha branca com a logo oficial dentro.
    selo = tk.Frame(caixa, bg="white", padx=4, pady=4)
    selo.pack(pady=(0, 12))
    logo = None
    if LOGO.is_file():
        try:
            logo = tk.PhotoImage(file=str(LOGO)).subsample(4)  # 240x250 -> 60x62
        except Exception:
            logo = None
    if logo is not None:
        tk.Label(selo, image=logo, bg="white", bd=0).pack()
        selo.logo = logo  # sem referencia o Tk apaga a imagem
    else:
        tk.Label(selo, text="S & D\nBIOFLOR", bg="white", fg=FUNDO,
                 font=(FONTE, 10, "bold"), width=7, height=3).pack()

    tk.Label(caixa, text="S & D BIOFLOR", bg=CAIXA, fg=TEXTO,
             font=(FONTE, 9, "bold")).pack()
    marca = "✖  " if erro else "✔  "
    tk.Label(caixa, text=marca + dados.get("titulo", "PAINEL ATUALIZADO"), bg=CAIXA,
             fg=ERRO if erro else TITULO, font=(FONTE, 15, "bold"),
             wraplength=380, justify="center").pack(pady=(4, 2))
    if dados.get("subtitulo"):
        tk.Label(caixa, text=dados["subtitulo"], bg=CAIXA, fg=ERRO if erro else TEXTO,
                 font=(FONTE, 11), wraplength=380, justify="center").pack(pady=(0, 10))

    if dados.get("linhas"):
        tabela = tk.Frame(caixa, bg=FUNDO, highlightbackground=BORDA, highlightthickness=1,
                          padx=14, pady=8)
        tabela.pack(fill="x", pady=(2, 10))
        for i, (rotulo, valor) in enumerate(dados["linhas"]):
            tk.Label(tabela, text=rotulo, bg=FUNDO, fg=LIMAO, font=(FONTE, 10, "bold"),
                     anchor="w").grid(row=i, column=0, sticky="w", padx=(0, 16), pady=1)
            tk.Label(tabela, text=valor, bg=FUNDO, fg=TITULO, font=(FONTE, 10),
                     anchor="w").grid(row=i, column=1, sticky="w", pady=1)

    if dados.get("rodape"):
        tk.Label(caixa, text=dados["rodape"], bg=CAIXA, fg=TEXTO,
                 font=(FONTE, 10, "italic"), wraplength=380).pack(pady=(0, 12))

    botoes = tk.Frame(caixa, bg=CAIXA)
    botoes.pack(fill="x")

    def fechar(_=None):
        raiz.destroy()

    def abrir(_=None):
        if dados.get("url"):
            webbrowser.open(dados["url"])
        fechar()

    if dados.get("url"):
        b = tk.Label(botoes, text="Abrir painel", bg=LIMAO, fg=FUNDO, font=(FONTE, 12, "bold"),
                     padx=18, pady=8, cursor="hand2")
        b.pack(side="left", expand=True, fill="x", padx=(0, 6))
        b.bind("<Button-1>", abrir)
        b.bind("<Enter>", lambda e: b.configure(bg=LIMAO_ESCURO))
        b.bind("<Leave>", lambda e: b.configure(bg=LIMAO))
    f = tk.Label(botoes, text="Fechar", bg=FUNDO, fg=TITULO, font=(FONTE, 12),
                 padx=18, pady=8, cursor="hand2", highlightbackground=BORDA, highlightthickness=1)
    f.pack(side="left", expand=True, fill="x")
    f.bind("<Button-1>", fechar)
    raiz.bind("<Escape>", fechar)

    # Arrastar pela caixa (sem barra de titulo nao haveria como mover).
    pos = {}

    def pegar(e):
        pos["x"], pos["y"] = e.x_root - raiz.winfo_x(), e.y_root - raiz.winfo_y()

    def mover(e):
        raiz.geometry(f"+{e.x_root - pos.get('x', 0)}+{e.y_root - pos.get('y', 0)}")

    for w in (moldura, caixa):
        w.bind("<Button-1>", pegar)
        w.bind("<B1-Motion>", mover)

    raiz.update_idletasks()
    w, h = raiz.winfo_width(), raiz.winfo_height()
    x = (raiz.winfo_screenwidth() - w) // 2
    y = (raiz.winfo_screenheight() - h) // 3
    raiz.geometry(f"+{x}+{y}")
    raiz.lift()
    raiz.focus_force()
    raiz.mainloop()


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] == "--exemplo":
        dados = EXEMPLO
    elif len(sys.argv) > 1:
        dados = json.loads(sys.argv[1])
    else:
        dados = EXEMPLO
    mostrar(dados)
    return 0


if __name__ == "__main__":
    sys.exit(main())
