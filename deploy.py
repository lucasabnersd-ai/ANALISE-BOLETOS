# -*- coding: utf-8 -*-
"""Publica PUBLICAR/ no GitHub Pages do painel (repo ANALISE-BOLETOS).

Faz git add . / commit / push automaticos, respeitando o .gitignore. E o
passo que faltava no ATUALIZAR_PAINEL.cmd: o gerar_painel.py so escreve os
arquivos aqui na maquina -- quem coloca no ar e este script.

Trava de seguranca: o GitHub Pages serve o repositorio publicamente mesmo
sendo privado (ver README.md), entao antes de empurrar qualquer coisa este
script confere que os arquivos dentro de PUBLICAR/ continuam pequenos e sem
sequencias longas de digitos (linha digitavel, CNPJ). A carteira de verdade
vai para o Supabase pelo publicar_dados.py, nunca pelo git.

Uso:
    python deploy.py
    python deploy.py --message "Atualizacao manual"

Saidas: 0 = publicado | 2 = nada para publicar | 1 = erro / bloqueado.
"""

from __future__ import annotations

import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

PASTA = Path(__file__).resolve().parent
PUBLICAR_DIR = PASTA / "PUBLICAR"
REMOTE_URL = "https://github.com/lucasabnersd-ai/ANALISE-BOLETOS.git"
BRANCH = "main"
USER_EMAIL = "lucas.abnersd@gmail.com"
USER_NAME = "lucasabnersd-ai"
PAGES_URL = "https://lucasabnersd-ai.github.io/ANALISE-BOLETOS/PUBLICAR/"

# HTML/JS publicados sao so layout + carregador (hoje ~90 KB + ~12 KB). Com a
# carteira embutida isso passaria disso de sobra -- margem folgada, mas o
# bastante para pegar um regresso do gerar_painel.py.
LIMITE_ARQUIVO_PUBLICADO = 1024 * 1024
PADRAO_DIGITOS_LONGOS = re.compile(r"\d{20,}")

GIT_EXE = "git"


def _localizar_git() -> str:
    achado = shutil.which("git")
    if achado:
        return achado
    candidatos = [
        r"C:\Program Files\Git\cmd\git.exe",
        r"C:\Program Files\Git\bin\git.exe",
        r"C:\Program Files (x86)\Git\cmd\git.exe",
        os.path.expandvars(r"%LocalAppData%\Programs\Git\cmd\git.exe"),
    ]
    for pattern in (
        r"%LocalAppData%\GitHubDesktop\app-*\resources\app\git\cmd\git.exe",
        r"%LocalAppData%\GitHubDesktop\app-*\resources\app\git\mingw64\bin\git.exe",
    ):
        for p in sorted(glob.glob(os.path.expandvars(pattern)), reverse=True):
            candidatos.append(p)
    for c in candidatos:
        if c and os.path.isfile(c):
            return c
    raise RuntimeError(
        "git.exe nao encontrado. Instale o Git for Windows: "
        "https://git-scm.com/download/win"
    )


def _run(args: list[str], check: bool = True, quiet: bool = False) -> subprocess.CompletedProcess:
    if not quiet:
        print("  $ git " + " ".join(args))
    resultado = subprocess.run(
        [GIT_EXE] + args, cwd=PASTA, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
    )
    if not quiet:
        if resultado.stdout.strip():
            print(resultado.stdout.rstrip())
        if resultado.stderr.strip():
            print(resultado.stderr.rstrip())
    if check and resultado.returncode != 0:
        raise RuntimeError(
            "git %s: %s" % (" ".join(args), (resultado.stderr or resultado.stdout).strip())
        )
    return resultado


def _falta_protecao() -> str | None:
    """Devolve o motivo se PUBLICAR/ parecer ter dado embutido; None se ok."""
    if not PUBLICAR_DIR.is_dir():
        return "PUBLICAR/ nao existe -- rode gerar_painel.py antes."
    arquivos = [p for p in PUBLICAR_DIR.rglob("*") if p.is_file()]
    if not arquivos:
        return "PUBLICAR/ esta vazia -- rode gerar_painel.py antes."
    for caminho in arquivos:
        tamanho = caminho.stat().st_size
        nome = caminho.relative_to(PASTA)
        if tamanho > LIMITE_ARQUIVO_PUBLICADO:
            return "%s com %.1f MB -- parece ter a carteira embutida." % (nome, tamanho / 1e6)
        try:
            texto = caminho.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        achado = PADRAO_DIGITOS_LONGOS.search(texto)
        if achado:
            return "%s tem uma sequencia de %d digitos -- parece linha digitavel embutida." % (
                nome, len(achado.group())
            )
    return None


def publicar(mensagem: str | None = None) -> int:
    global GIT_EXE

    print("=" * 60)
    print("  Publicando painel Analise de Boletos em: %s" % PASTA)
    print("  Remote: %s  |  Branch: %s" % (REMOTE_URL, BRANCH))
    print("=" * 60)

    problema = _falta_protecao()
    if problema:
        print("\nPUBLICACAO BLOQUEADA: " + problema)
        return 1

    try:
        GIT_EXE = _localizar_git()
    except RuntimeError as exc:
        print("\nERRO: %s" % exc)
        return 1

    if not (PASTA / ".git").exists():
        print("\nERRO: esta pasta nao e um repositorio git (%s)." % PASTA)
        return 1

    _run(["config", "user.name", USER_NAME], check=False)
    _run(["config", "user.email", USER_EMAIL], check=False)
    r = _run(["remote"], check=False, quiet=True)
    if "origin" in (r.stdout or ""):
        _run(["remote", "set-url", "origin", REMOTE_URL], check=False)
    else:
        _run(["remote", "add", "origin", REMOTE_URL], check=False)

    r = _run(["fetch", "origin", BRANCH], check=False)
    if r.returncode == 0:
        _run(["reset", "--mixed", "origin/" + BRANCH], check=False)

    print("\n[1/3] git add .")
    _run(["add", "."])

    r = _run(["diff", "--cached", "--name-only"], check=False, quiet=True)
    preparado = [n for n in (r.stdout or "").splitlines() if n.strip()]
    if not preparado:
        print("\n  Nada mudou desde o ultimo commit. Nada a publicar.")
        return 2

    print("\n  Mudancas:")
    for nome in preparado:
        print("    - " + nome)

    msg = mensagem or ("Atualiza painel - " + datetime.now().strftime("%d/%m/%Y %H:%M"))
    print("\n[2/3] git commit -m \"%s\"" % msg)
    _run(["commit", "-m", msg])

    print("\n[3/3] git push origin %s" % BRANCH)
    try:
        _run(["push", "-u", "origin", BRANCH])
    except RuntimeError as exc:
        print("\nERRO no push: %s" % exc)
        print("  Dica: rode 'git push' manual nesta pasta uma vez para salvar as credenciais.")
        return 1

    print("\n" + "=" * 60)
    print("  PUBLICADO: " + PAGES_URL)
    print("  (o GitHub Pages leva ~1 min para refletir)")
    print("=" * 60)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Publica o painel Analise de Boletos no GitHub Pages.")
    parser.add_argument("--message", help="Mensagem do commit.")
    args = parser.parse_args()
    try:
        return publicar(mensagem=args.message)
    except Exception as exc:
        print("\nERRO FATAL: %s" % exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
