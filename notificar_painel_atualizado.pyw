"""Tela de aviso "PAINEL ANALISE DE BOLETOS ATUALIZADO" no padrao S & D BIOFLOR.

O MESMO visual da tela da SE2 (se2 - sistema\\notificar_painel_atualizado.pyw,
30/09/2026) -- pedido dele: "use o mesmo visual do da se2". Mesma paleta, mesma
fonte, mesmo canto da tela, mesmos botoes; aqui muda so o texto.

Copia do modelo do SC7 (PAINEIS_REPOS\\SC7\\notificar_painel_atualizado.pyw),
pedido dele em 30/09/2026: "COLOQUE O ALERTA NO MODELO SE2 E SC7". Muda so o
texto, o endereco e onde ficam o PID e o log de erro (pasta DADOS).

Chamada pelo gatilho_bases_do_dia.py quando a rodada automatica termina
(processo solto, a rodada nao espera): "PAINEL ANALISE DE BOLETOS ATUALIZADO"
ou "... NAO FOI ATUALIZADO" (botao VER DETALHES abre o log). Uma tela por vez.

Teste manual:  pythonw notificar_painel_atualizado.pyw --teste
Conferir o desenho sem aparecer na tela de ninguem:
               python notificar_painel_atualizado.pyw --teste --png saida.png
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tkinter as tk
import webbrowser
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PID_FILE = SCRIPT_DIR / "DADOS" / "notificar_painel_atualizado.pid"
URL_PAINEL = "https://analise-boletos.vercel.app/"
LOGO_CANDIDATAS = (
    SCRIPT_DIR / "PUBLICAR" / "logo-sd.png",
    Path(r"C:\Users\lucas\Downloads\logo_bioflor_padrao.png"),
)

# Paleta da tela de entrada S&D BIOFLOR -- a MESMA da tela da SE2
FUNDO = "#1B302F"
CARTAO = "#24403E"
BORDA = "#3A5654"
LIMA = "#C0D15E"
TEXTO = "#FFFFFF"
TEXTO_SUAVE = "#C9D6D2"
FONTE = "Calibri"


def fechar_tela_anterior() -> None:
    try:
        pid = int(PID_FILE.read_text(encoding="ascii").strip())
    except (OSError, ValueError):
        return
    if pid == os.getpid():
        return
    subprocess.run(
        ["taskkill", "/PID", str(pid), "/F"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        check=False,
    )


def salvar_png(raiz: tk.Tk, caminho: str) -> None:
    """Desenha a propria janela (PrintWindow) num PNG. A janela fica
    transparente enquanto isso: e conferencia do desenho, nao print da area
    de trabalho de ninguem."""
    import ctypes
    from ctypes import wintypes
    from PIL import Image

    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    hwnd = int(raiz.wm_frame(), 16)
    r = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    w, h = r.right - r.left, r.bottom - r.top
    dc_janela = user32.GetWindowDC(hwnd)
    dc = gdi32.CreateCompatibleDC(dc_janela)
    bmp = gdi32.CreateCompatibleBitmap(dc_janela, w, h)
    gdi32.SelectObject(dc, bmp)
    user32.PrintWindow(hwnd, dc, 2)                       # PW_RENDERFULLCONTENT

    class BIH(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                    ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD),
                    ("biCompression", wintypes.DWORD), ("biSizeImage", wintypes.DWORD),
                    ("biXPelsPerMeter", wintypes.LONG), ("biYPelsPerMeter", wintypes.LONG),
                    ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]

    bih = BIH(ctypes.sizeof(BIH), w, -h, 1, 32, 0, 0, 0, 0, 0, 0)
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(dc, bmp, 0, h, buf, ctypes.byref(bih), 0)
    Image.frombuffer("RGB", (w, h), buf, "raw", "BGRX", 0, 1).save(caminho)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(dc)
    user32.ReleaseDC(hwnd, dc_janela)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--titulo", default="PAINEL ANÁLISE DE BOLETOS ATUALIZADO")
    parser.add_argument("--linha", action="append", default=[],
                        help="linha de detalhe embaixo do titulo (pode repetir)")
    parser.add_argument("--botao", choices=("abrir", "detalhes"), default="abrir",
                        help="o botao lima: ABRIR PAINEL ou VER DETALHES")
    parser.add_argument("--detalhes", default="", help="arquivo que o VER DETALHES abre")
    parser.add_argument("--teste", action="store_true")
    parser.add_argument("--png", default="", help="desenha fora da tela, salva e sai")
    args = parser.parse_args()

    if not args.png:
        PID_FILE.parent.mkdir(parents=True, exist_ok=True)
        fechar_tela_anterior()
        try:
            PID_FILE.write_text(str(os.getpid()), encoding="ascii")
        except OSError:
            pass

    agora = datetime.now().strftime("%d/%m/%Y às %H:%M")
    detalhes = list(args.linha) or [f"Atualizado em {agora}"]

    raiz = tk.Tk()
    raiz.withdraw()
    raiz.overrideredirect(True)
    raiz.attributes("-topmost", True)
    raiz.configure(bg=BORDA)

    cartao = tk.Frame(raiz, bg=CARTAO, padx=18, pady=16)
    cartao.pack(padx=1, pady=1)

    topo = tk.Frame(cartao, bg=CARTAO)
    topo.pack(fill="x")

    logo_img = None
    for caminho in LOGO_CANDIDATAS:
        if caminho.is_file():
            try:
                logo_img = tk.PhotoImage(file=str(caminho))
                fator = max(1, logo_img.height() // 64)
                logo_img = logo_img.subsample(fator, fator)
                break
            except tk.TclError:
                logo_img = None
    if logo_img is not None:
        # a logo tem fundo branco opaco: entra numa pastilha branca
        pastilha = tk.Frame(topo, bg="#FFFFFF", padx=4, pady=4)
        pastilha.pack(side="left", padx=(0, 14))
        tk.Label(pastilha, image=logo_img, bg="#FFFFFF", bd=0).pack()

    textos = tk.Frame(topo, bg=CARTAO)
    textos.pack(side="left", fill="both", expand=True)
    tk.Label(textos, text="S & D BIOFLOR", bg=CARTAO, fg=LIMA,
             font=(FONTE, 10, "bold"), anchor="w").pack(fill="x")
    tk.Label(textos, text=args.titulo + (" (TESTE)" if args.teste else ""),
             bg=CARTAO, fg=TEXTO, font=(FONTE, 16, "bold"), anchor="w").pack(fill="x")
    for linha in detalhes:
        # wraplength so dobra linha longa (o motivo de uma base barrada);
        # as curtas ficam exatamente como na tela da SE2
        tk.Label(textos, text=linha, bg=CARTAO, fg=TEXTO_SUAVE,
                 font=(FONTE, 11), anchor="w", justify="left", wraplength=440).pack(fill="x")

    tk.Frame(cartao, bg=LIMA, height=2).pack(fill="x", pady=(14, 12))

    botoes = tk.Frame(cartao, bg=CARTAO)
    botoes.pack(fill="x")

    def abrir() -> None:
        webbrowser.open(URL_PAINEL)
        raiz.destroy()

    def ver_detalhes() -> None:
        if args.detalhes:
            try:
                subprocess.Popen(["notepad.exe", args.detalhes])
            except OSError:
                pass

    if args.botao == "detalhes":
        tk.Button(botoes, text="VER DETALHES", command=ver_detalhes, bg=LIMA, fg=FUNDO,
                  activebackground="#D4E27A", activeforeground=FUNDO, relief="flat",
                  font=(FONTE, 11, "bold"), padx=16, pady=5, cursor="hand2", bd=0
                  ).pack(side="right")
    else:
        tk.Button(botoes, text="ABRIR PAINEL", command=abrir, bg=LIMA, fg=FUNDO,
                  activebackground="#D4E27A", activeforeground=FUNDO, relief="flat",
                  font=(FONTE, 11, "bold"), padx=16, pady=5, cursor="hand2", bd=0
                  ).pack(side="right")
    tk.Button(botoes, text="FECHAR", command=raiz.destroy, bg=CARTAO, fg=TEXTO,
              activebackground=BORDA, activeforeground=TEXTO, relief="flat",
              font=(FONTE, 11), padx=12, pady=5, cursor="hand2", bd=0,
              highlightthickness=1, highlightbackground=BORDA
              ).pack(side="right", padx=(0, 8))

    # canto inferior direito, acima da barra de tarefas
    raiz.update_idletasks()
    largura = max(raiz.winfo_reqwidth(), 420)
    altura = raiz.winfo_reqheight()
    x = raiz.winfo_screenwidth() - largura - 24
    y = raiz.winfo_screenheight() - altura - 72
    if args.png:
        # no lugar de verdade, mas 100% transparente: o Tk so pinta o que esta
        # na tela (fora dela a captura sai preta), e assim ninguem ve nada
        raiz.attributes("-alpha", 0.0)
        raiz.geometry(f"{largura}x{altura}+{x}+{y}")
        raiz.deiconify()
        raiz.update()
        raiz.after(900, lambda: (salvar_png(raiz, args.png), raiz.destroy()))
        raiz.mainloop()
        return 0
    raiz.geometry(f"{largura}x{altura}+{x}+{y}")
    raiz.deiconify()
    raiz.lift()
    raiz.bell()

    try:
        raiz.mainloop()
    finally:
        try:
            if PID_FILE.read_text(encoding="ascii").strip() == str(os.getpid()):
                PID_FILE.unlink()
        except OSError:
            pass
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException:
        # pythonw nao tem onde mostrar erro: sem isto a tela falha calada e
        # a rodada acha que avisou
        import traceback
        try:
            with (SCRIPT_DIR / "DADOS" / "tela_erro.txt").open("a", encoding="utf-8") as f:
                f.write(datetime.now().strftime("[%d/%m/%Y %H:%M:%S] ") + " ".join(sys.argv[1:]) + "\n"
                        + traceback.format_exc() + "\n")
        except OSError:
            pass
        raise
