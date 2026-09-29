#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gatilho_bases_do_dia.py - atualiza o painel SOZINHO, mas so depois das 3 bases.

Pedido dele em 29/09/2026: "AUTOMATIZE O PAINEL ANALISE BOLETOS PARA ATUALIZAR
SOMENTE APOS ATUALIZAR A SC7 SF1 E SE2 O MESMO DIA".

O Agendador do Windows roda este arquivo a cada 10 minutos (tarefa
"PAINEL ANALISE BOLETOS - atualizar apos SC7 SF1 SE2"). Em cada passada ele
olha as tres bases em ...\\LUCAS ABNER ARAUJO\\BASES GENERICOS:

    SC7.xlsx
    SF1.xlsx
    SE2 - POSICAO DIARIA.xlsx   (achada por curinga: o nome tem cedilha e til)

e SO dispara o ATUALIZAR_PAINEL.cmd quando as TRES valem ao mesmo tempo:

  1. foram gravadas HOJE;
  2. foram gravadas DEPOIS da ultima rodada que este gatilho fez com sucesso
     -- ou seja, as tres de novo, e nao so uma delas (regra "somente apos");
  3. estao paradas ha pelo menos ESTAVEL_MIN minutos (a gravacao terminou e o
     OneDrive ja trouxe o arquivo inteiro).

Se faltar uma, nao faz nada e espera a proxima passada. Publicar com duas bases
de hoje e uma de ontem e justamente o que ele pediu para nao acontecer.

Antes do painel roda o CLIQUE_PARA_IMPORTAR_RET_BOLETOS_ITAU.CMD (pedido dele
29/09/2026): ele atualiza a BOLETOS ITAU.xlsx que o associador le. Se o Itau
der erro, o painel NAO roda (conta como falha e tenta de novo).

A rodada abre numa JANELA VISIVEL (regra dele: script diz sempre se esta
rodando), e no fim sobe um aviso na tela: CONCLUIDO ou TERMINOU COM ERRO, com
hora e duracao. Deu erro -> tenta de novo na passada seguinte, ate TENTATIVAS
vezes por dia; depois para e avisa.

Quem rodar o ATUALIZAR_PAINEL.cmd a mao depois das tres bases nao ganha uma
segunda rodada automatica: o dev.html (escrito pelo passo [2/10]) mais novo que
as tres bases conta como "ja atualizado".

Uso manual:
    python gatilho_bases_do_dia.py            # uma passada (o que o Agendador faz)
    python gatilho_bases_do_dia.py --status   # so diz o que faria, nao roda nada
"""
from __future__ import annotations

import argparse
import json
import re
import os
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
from caminhos import bases_genericos, raiz_lucas  # noqa: E402

CMD = AQUI / "ATUALIZAR_PAINEL.cmd"
DEV_HTML = AQUI / "dev.html"
DADOS = AQUI / "DADOS"
ESTADO = DADOS / "gatilho_bases_estado.json"
LOG = DADOS / "gatilho_bases.log"

# Minutos que cada base precisa estar parada antes de valer. Cobre a gravacao
# do Excel e o OneDrive terminando de baixar o arquivo da outra maquina.
ESTAVEL_MIN = 3
# Tentativas por dia quando a rodada termina com erro.
TENTATIVAS = 3
# Prazo da rodada inteira. Hoje ela leva ~10 min; 90 e folga para o associador.
PRAZO_RODADA_S = 90 * 60


def agora() -> datetime:
    return datetime.now()


def hhmm(d: datetime) -> str:
    return d.strftime("%d/%m %H:%M")


def registrar(msg: str) -> None:
    DADOS.mkdir(parents=True, exist_ok=True)
    linha = f"{agora():%Y-%m-%d %H:%M:%S}  {msg}"
    with LOG.open("a", encoding="utf-8") as f:
        f.write(linha + "\n")
    try:
        print(linha, flush=True)
    except Exception:  # pythonw nao tem console
        pass


def ler_estado() -> dict:
    try:
        return json.loads(ESTADO.read_text(encoding="utf-8"))
    except Exception:
        return {}


def gravar_estado(estado: dict) -> None:
    DADOS.mkdir(parents=True, exist_ok=True)
    tmp = ESTADO.with_suffix(".tmp")
    tmp.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, ESTADO)


def achar_bases() -> dict[str, Path | None]:
    pasta = bases_genericos()
    se2 = [p for p in pasta.glob("SE2 - POSI*O DIARIA.xlsx") if "copia" not in p.name.lower()]
    return {
        "SC7": pasta / "SC7.xlsx" if (pasta / "SC7.xlsx").is_file() else None,
        "SF1": pasta / "SF1.xlsx" if (pasta / "SF1.xlsx").is_file() else None,
        "SE2": se2[0] if se2 else None,
    }


def avisar_na_tela(titulo: str, texto: str, erro: bool = False) -> None:
    """Caixa de aviso por cima de tudo, num processo SOLTO: a tarefa nao fica
    presa esperando alguem clicar em OK."""
    icone = 0x10 if erro else 0x40  # MB_ICONERROR / MB_ICONINFORMATION
    flags = icone | 0x40000 | 0x10000  # MB_TOPMOST | MB_SETFOREGROUND
    codigo = (
        "import ctypes,sys;"
        "ctypes.windll.user32.MessageBoxW(0,sys.argv[2],sys.argv[1],int(sys.argv[3]))"
    )
    exe = Path(sys.executable)
    pyw = exe.with_name("pythonw.exe")
    try:
        subprocess.Popen(
            [str(pyw if pyw.is_file() else exe), "-c", codigo, titulo, texto, str(flags)],
            creationflags=0x00000008 | 0x00000200,  # DETACHED_PROCESS | NEW_PROCESS_GROUP
            close_fds=True,
        )
    except Exception as e:
        registrar(f"(nao consegui mostrar o aviso na tela: {e})")


def situacao() -> tuple[bool, str, dict]:
    """(pode_rodar, motivo, info)."""
    estado = ler_estado()
    hoje = agora().date()
    bases = achar_bases()
    faltando = [n for n, p in bases.items() if p is None]
    if faltando:
        return False, "nao achei a base: " + ", ".join(faltando), {}

    datas = {n: datetime.fromtimestamp(p.stat().st_mtime) for n, p in bases.items()}
    info = {"bases": {n: d.isoformat(timespec="seconds") for n, d in datas.items()}}
    resumo = " | ".join(f"{n} {hhmm(d)}" for n, d in datas.items())

    velhas = [n for n, d in datas.items() if d.date() != hoje]
    if velhas:
        return False, f"esperando {', '.join(velhas)} de hoje ({resumo})", info

    ultima_ok = estado.get("ultima_ok_inicio")
    ultima_ok = datetime.fromisoformat(ultima_ok) if ultima_ok else None
    if ultima_ok:
        pendentes = [n for n, d in datas.items() if d <= ultima_ok]
        if pendentes:
            if len(pendentes) == 3:
                return False, f"painel ja atualizado as {hhmm(ultima_ok)} com as 3 bases de hoje", info
            return False, (
                f"esperando {', '.join(pendentes)} de novo: as 3 precisam ser atualizadas "
                f"depois da ultima rodada ({hhmm(ultima_ok)}) -- {resumo}"
            ), info

    mais_nova = max(datas.values())
    if DEV_HTML.is_file():
        dev = datetime.fromtimestamp(DEV_HTML.stat().st_mtime)
        falhas_hoje = estado.get("falhas", {}).get(hoje.isoformat(), 0)
        if dev > mais_nova and not falhas_hoje:
            info["manual"] = dev.isoformat(timespec="seconds")
            return False, f"painel ja foi atualizado a mao as {hhmm(dev)}, depois das 3 bases", info

    recentes = [n for n, d in datas.items() if agora() - d < timedelta(minutes=ESTAVEL_MIN)]
    if recentes:
        return False, f"{', '.join(recentes)} acabou de ser gravada; espero {ESTAVEL_MIN} min parada", info

    falhas = estado.get("falhas", {}).get(hoje.isoformat(), 0)
    if falhas >= TENTATIVAS:
        return False, f"ja deu erro {falhas} vezes hoje; parei de tentar (rode o ATUALIZAR_PAINEL.cmd a mao)", info

    return True, f"as 3 bases de hoje estao prontas ({resumo})", info


def achar_cmd_itau() -> Path | None:
    """O CLIQUE_PARA_IMPORTAR_RET_BOLETOS_ITAU.CMD, procurado a partir da raiz
    (a pasta AUTOMACOES LUCAS tem cedilha e til: sempre curinga)."""
    for pasta in raiz_lucas().glob("AUTOMA*"):
        cmd = pasta / "CLIQUE_PARA_IMPORTAR_RET_BOLETOS_ITAU" / "CLIQUE_PARA_IMPORTAR_RET_BOLETOS_ITAU.CMD"
        if cmd.is_file():
            return cmd
    return None


def rodar_cmd(cmd: Path, sem_pausa_por_nul: bool = False) -> int:
    env = dict(os.environ)
    env["PAINEL_SEM_PAUSA"] = "1"  # sem "pressione qualquer tecla": a tarefa nao pode travar
    args = ["cmd.exe", "/d", "/c", "call", str(cmd)]
    if sem_pausa_por_nul:
        # O .CMD do Itau termina num `pause` sem chave para desligar; com a
        # entrada vindo do NUL ele passa direto. O caminho vai entre aspas
        # (list2cmdline), entao o & de "Grupo S&D" nao quebra o comando.
        args += ["<", "nul"]
    proc = subprocess.Popen(
        args,
        cwd=str(cmd.parent),
        env=env,
        creationflags=subprocess.CREATE_NEW_CONSOLE,  # janela visivel: ele ve rodando
    )
    try:
        return proc.wait(timeout=PRAZO_RODADA_S)
    except subprocess.TimeoutExpired:
        proc.kill()
        return 124


def rodar_itau() -> int:
    """Passo 0 (pedido dele 29/09/2026): importar os RET/DDA do Itau ANTES do
    painel. Ele grava a BOLETOS ITAU.xlsx, que o associador le no [1/10] do
    ATUALIZAR_PAINEL.cmd -- rodar depois seria publicar com o DDA de ontem."""
    cmd = achar_cmd_itau()
    if cmd is None:
        registrar("ERRO: nao achei o CLIQUE_PARA_IMPORTAR_RET_BOLETOS_ITAU.CMD")
        return 1
    registrar("[ITAU] importando RET/DDA do Itau antes do painel...")
    inicio = time.time()
    codigo = rodar_cmd(cmd, sem_pausa_por_nul=True)
    # O .CMD do Itau nao termina com `exit /b`: o codigo que chega aqui e o do
    # ultimo comando, nao o do Python. Quem diz a verdade e a linha
    # "Codigo de saida: N" que ele escreve no log desta rodada.
    logs = sorted(
        (p for p in (cmd.parent / "IMPORTACAO_ITAU" / "logs").glob("importar_txt_itau_*.log")
         if p.stat().st_mtime >= inicio - 5),
        key=lambda p: p.stat().st_mtime,
    )
    if not logs:
        registrar("[ITAU] ERRO: a rodada nao deixou log -- nao sei se importou")
        return codigo or 1

    texto = logs[-1].read_text(encoding="latin-1", errors="replace")
    achado = re.findall(r"Codigo de saida:\s*(\d+)", texto)
    codigo = int(achado[-1]) if achado else 1
    registrar(f"[ITAU] terminou com codigo {codigo} ({logs[-1].name})")
    return codigo


def rodar_painel() -> int:
    codigo = rodar_itau()
    if codigo != 0:
        # Sem o DDA do dia o painel sairia com os boletos de ontem sem ninguem
        # perceber: para aqui e tenta a rodada inteira de novo em 10 min.
        return codigo
    registrar("[PAINEL] rodando ATUALIZAR_PAINEL.cmd...")
    return rodar_cmd(CMD)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--status", action="store_true", help="so diz o que faria")
    args = ap.parse_args()

    pode, motivo, info = situacao()
    estado = ler_estado()

    if args.status:
        print(("VAI RODAR: " if pode else "NAO RODA: ") + motivo)
        return 0

    if not pode:
        # Rodada manual: marca como feita para nao voltar a perguntar o dia todo.
        if "manual" in info:
            estado["ultima_ok_inicio"] = info["manual"]
            gravar_estado(estado)
        # So escreve no log quando a situacao muda: 80 linhas iguais por dia ninguem le.
        if estado.get("ultimo_motivo") != motivo:
            estado["ultimo_motivo"] = motivo
            gravar_estado(estado)
            registrar("aguardando: " + motivo)
        return 0

    inicio = agora()
    estado["ultimo_motivo"] = "rodando"
    estado["rodando_desde"] = inicio.isoformat(timespec="seconds")
    gravar_estado(estado)
    registrar("DISPARANDO ATUALIZAR_PAINEL.cmd -- " + motivo)

    t0 = time.monotonic()
    codigo = rodar_painel()
    dur = int(time.monotonic() - t0)
    dur_txt = f"{dur // 60} min {dur % 60:02d} s"

    estado = ler_estado()
    estado.pop("rodando_desde", None)
    hoje = inicio.date().isoformat()
    if codigo == 0:
        estado["ultima_ok_inicio"] = inicio.isoformat(timespec="seconds")
        estado["ultima_ok_bases"] = info.get("bases")
        estado["falhas"] = {}
        estado["ultimo_motivo"] = "concluido"
        gravar_estado(estado)
        registrar(f"CONCLUIDO em {dur_txt}.")
        avisar_na_tela(
            "PAINEL ANALISE BOLETOS",
            f"CONCLUIDO. Nada mais esta rodando.\n\n"
            f"Atualizado sozinho depois da SC7, SF1 e SE2 de hoje.\n"
            f"Comecou {inicio:%H:%M}, terminou {agora():%H:%M} ({dur_txt}).\n\n"
            f"https://analise-boletos.vercel.app/",
        )
    else:
        falhas = estado.setdefault("falhas", {})
        falhas[hoje] = falhas.get(hoje, 0) + 1
        estado["ultimo_motivo"] = f"erro {codigo}"
        gravar_estado(estado)
        n = falhas[hoje]
        resto = (
            f"Vou tentar de novo em 10 min ({n} de {TENTATIVAS})."
            if n < TENTATIVAS
            else "Parei de tentar hoje. Rode o ATUALIZAR_PAINEL.cmd a mao para ver o erro."
        )
        registrar(f"TERMINOU COM ERRO (codigo {codigo}) em {dur_txt}. {resto}")
        avisar_na_tela(
            "PAINEL ANALISE BOLETOS - ERRO",
            f"TERMINOU COM ERRO (codigo {codigo}).\n"
            f"Comecou {inicio:%H:%M}, parou {agora():%H:%M} ({dur_txt}).\n\n{resto}",
            erro=True,
        )
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:  # nunca morrer calado dentro do Agendador
        registrar(f"FALHA DO GATILHO: {type(e).__name__}: {e}")
        raise
