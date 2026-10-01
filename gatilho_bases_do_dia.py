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
import hashlib
import json
import re
import zipfile
import os
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
from caminhos import bases_genericos, pasta_analises, raiz_lucas  # noqa: E402

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


def assinatura_dados(p: Path) -> str:
    """O que MUDA quando muda o DADO, e nao quando o arquivo so e salvo de novo.

    O .xlsx e um zip: o CRC de cada planilha (xl/worksheets/*.xml) e das
    strings (xl/sharedStrings.xml) ja esta gravado no indice do zip, entao ler
    e instantaneo mesmo na SE2 de 15 MB. Salvar sem mexer em nada muda a data
    do arquivo e o docProps (carimbo de quem salvou), mas nao esses CRCs.
    Pedido dele 30/09/2026: atualizar "quando houver dados novos"."""
    with zipfile.ZipFile(p) as z:
        partes = sorted(
            f"{i.filename}:{i.CRC}:{i.file_size}"
            for i in z.infolist()
            if i.filename.startswith(("xl/worksheets/", "xl/sharedStrings"))
        )
    return hashlib.sha1("|".join(partes).encode()).hexdigest()[:16]


def avisar_na_tela(dados: dict) -> None:
    """Tela de aviso no MODELO DA SE2 E DO SC7 (notificar_painel_atualizado.pyw,
    copia do SC7 -- pedido dele 30/09/2026), num processo SOLTO: a tarefa nao
    fica presa esperando alguem fechar. Uma tela nova fecha a anterior."""
    exe = Path(sys.executable)
    pyw = exe.with_name("pythonw.exe")
    args = [str(pyw if pyw.is_file() else exe), str(AQUI / "notificar_painel_atualizado.pyw"),
            "--titulo", dados["titulo"]]
    for rotulo, valor in dados.get("linhas", []):
        args += ["--linha", f"{rotulo}: {valor}" if rotulo else valor]
    if dados.get("status") == "erro":
        args += ["--botao", "detalhes", "--detalhes", str(LOG)]
    try:
        subprocess.Popen(
            args,
            cwd=str(AQUI),
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            # sem console + fora do job da tarefa: a tela vive depois que ela acaba
            creationflags=0x08000000 | 0x01000000,
            close_fds=True,
        )
    except OSError:
        try:
            subprocess.Popen(args, cwd=str(AQUI), creationflags=0x08000000, close_fds=True)
        except Exception as e:
            registrar(f"(nao consegui mostrar o aviso na tela: {e})")


def achar_sefaz() -> Path | None:
    """A SEFAZ.xlsx (pasta ANALISES BOLETOS), achada pelo caminhos.py."""
    p = pasta_analises() / "SEFAZ.xlsx"
    return p if p.is_file() else None


def situacao() -> tuple[bool, str, dict]:
    """(pode_rodar, motivo, info).

    Tres jeitos de rodar (o resto espera):
      A. REGRA DO DIA: SC7, SF1 e SE2 gravadas HOJE, as tres depois da ultima
         rodada boa, com pelo menos um dado novo (inclui a SEFAZ, se mudou).
      B. SEFAZ (pedido dele 01/10/2026, "COLOQUE A SEFAZ NO GATILHO TAMBEM"):
         o painel ja rodou hoje com as 3 bases do dia e a SEFAZ.xlsx trouxe
         DADO NOVO depois disso -> roda de novo. Ela sozinha nao abre o dia:
         sem as 3 de hoje continua esperando.
      C. NOVA TENTATIVA: a ultima rodada deu erro (ate TENTATIVAS por dia) ->
         repete com as mesmas bases, sem exigir que as 3 sejam salvas de novo.
         Antes de 01/10/2026 essa exigencia barrava a repeticao: o erro das
         10:10 daquele dia (janela fechada) teria deixado o painel parado."""
    estado = ler_estado()
    hoje = agora().date()
    bases = achar_bases()
    faltando = [n for n, p in bases.items() if p is None]
    if faltando:
        return False, "nao achei a base: " + ", ".join(faltando), {}
    sefaz = achar_sefaz()

    datas = {n: datetime.fromtimestamp(p.stat().st_mtime) for n, p in bases.items()}
    data_sefaz = datetime.fromtimestamp(sefaz.stat().st_mtime) if sefaz else None
    info = {"bases": {n: d.isoformat(timespec="seconds") for n, d in datas.items()}}
    if data_sefaz:
        info["bases"]["SEFAZ"] = data_sefaz.isoformat(timespec="seconds")
    resumo = " | ".join(f"{n} {hhmm(datetime.fromisoformat(d))}" for n, d in info["bases"].items())

    velhas = [n for n, d in datas.items() if d.date() != hoje]
    if velhas:
        return False, f"esperando {', '.join(velhas)} de hoje ({resumo})", info

    falhas = estado.get("falhas", {}).get(hoje.isoformat(), 0)
    if falhas >= TENTATIVAS:
        return False, f"ja deu erro {falhas} vezes hoje; parei de tentar (rode o ATUALIZAR_PAINEL.cmd a mao)", info

    envolvidas = dict(bases)
    if sefaz:
        envolvidas["SEFAZ"] = sefaz
    todas_datas = dict(datas)
    if data_sefaz:
        todas_datas["SEFAZ"] = data_sefaz

    recentes = [n for n, d in todas_datas.items() if agora() - d < timedelta(minutes=ESTAVEL_MIN)]
    if recentes:
        return False, f"{', '.join(recentes)} acabou de ser gravada; espero {ESTAVEL_MIN} min parada", info

    try:
        dados = {n: assinatura_dados(p) for n, p in envolvidas.items()}
    except Exception as e:  # zip incompleto = ainda gravando / OneDrive trazendo
        return False, f"nao consegui ler o conteudo das bases ainda ({type(e).__name__}); espero", info
    info["dados"] = dados
    anteriores = estado.get("ultima_ok_dados") or {}
    info["novas"] = [n for n in dados if dados[n] != anteriores.get(n)]

    # C. a ultima rodada deu erro hoje: repete.
    if falhas and str(estado.get("ultimo_motivo", "")).startswith("erro"):
        return True, f"NOVA TENTATIVA ({falhas + 1} de {TENTATIVAS}) depois do erro da ultima rodada ({resumo})", info

    ultima_ok = estado.get("ultima_ok_inicio")
    ultima_ok = datetime.fromisoformat(ultima_ok) if ultima_ok else None

    pendentes = [n for n, d in datas.items() if ultima_ok and d <= ultima_ok]
    if pendentes:
        # B. as 3 de hoje ja foram usadas numa rodada de HOJE; so a SEFAZ manda.
        if (ultima_ok and ultima_ok.date() == hoje and data_sefaz
                and data_sefaz > ultima_ok and "SEFAZ" in info["novas"]):
            return True, f"SEFAZ com dado novo depois da rodada de {hhmm(ultima_ok)} ({resumo})", info
        if len(pendentes) == 3:
            return False, f"painel ja atualizado as {hhmm(ultima_ok)} com as 3 bases de hoje ({resumo})", info
        return False, (
            f"esperando {', '.join(pendentes)} de novo: as 3 precisam ser atualizadas "
            f"depois da ultima rodada ({hhmm(ultima_ok)}) -- ou a SEFAZ com dado novo -- {resumo}"
        ), info

    # A. as 3 sao de hoje e mais novas que a ultima rodada boa.
    mais_nova = max(todas_datas.values())
    if DEV_HTML.is_file():
        dev = datetime.fromtimestamp(DEV_HTML.stat().st_mtime)
        if dev > mais_nova and not falhas:
            info["manual"] = dev.isoformat(timespec="seconds")
            return False, f"painel ja foi atualizado a mao as {hhmm(dev)}, depois das bases", info

    if not info["novas"]:
        return False, (
            f"as bases foram salvas de novo mas SEM dado novo desde a rodada de "
            f"{hhmm(ultima_ok) if ultima_ok else '?'} -- nao atualizo ({resumo})"
        ), info

    return True, f"as 3 bases de hoje estao prontas, dado novo em {', '.join(info['novas'])} ({resumo})", info


def achar_cmd_itau() -> Path | None:
    """O CLIQUE_PARA_IMPORTAR_RET_BOLETOS_ITAU.CMD, procurado a partir da raiz
    (a pasta AUTOMACOES LUCAS tem cedilha e til: sempre curinga)."""
    for pasta in raiz_lucas().glob("AUTOMA*"):
        cmd = pasta / "CLIQUE_PARA_IMPORTAR_RET_BOLETOS_ITAU" / "CLIQUE_PARA_IMPORTAR_RET_BOLETOS_ITAU.CMD"
        if cmd.is_file():
            return cmd
    return None


def rodar_cmd(cmd: Path, sem_pausa_por_nul: bool = False, titulo: str = "RODANDO") -> int:
    env = dict(os.environ)
    env["PAINEL_SEM_PAUSA"] = "1"  # sem "pressione qualquer tecla": a tarefa nao pode travar
    # O titulo da janela pede para nao fechar: em 01/10/2026 a rodada morreu as
    # 10:10 com 0xC000013A (janela fechada) no meio dos alertas, antes de
    # publicar. O `&` fica FORA das aspas de proposito (separa os dois comandos);
    # o caminho vai entre aspas, entao o & de "Grupo S&D" nao quebra nada.
    args = ["cmd.exe", "/d", "/c", "title", f"{titulo} - NAO FECHE ESTA JANELA", "&", "call", str(cmd)]
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
    codigo = rodar_cmd(cmd, sem_pausa_por_nul=True, titulo="ITAU RET (antes do painel ANALISE BOLETOS)")
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
    return rodar_cmd(CMD, titulo="PAINEL ANALISE BOLETOS ATUALIZANDO")


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
    linhas = [
        [n, f"{hhmm(datetime.fromisoformat(d))} · "
            + ("dados novos" if n in info.get("novas", []) else "sem alteração")]
        for n, d in (info.get("bases") or {}).items()
    ]
    if codigo == 0:
        estado["ultima_ok_inicio"] = inicio.isoformat(timespec="seconds")
        estado["ultima_ok_bases"] = info.get("bases")
        estado["ultima_ok_dados"] = info.get("dados")
        estado["falhas"] = {}
        estado["ultimo_motivo"] = "concluido"
        gravar_estado(estado)
        registrar(f"CONCLUIDO em {dur_txt}.")
        avisar_na_tela({
            "status": "ok",
            "titulo": "PAINEL ANÁLISE DE BOLETOS ATUALIZADO",
            "linhas": [
                ["Atualizado em", f"{agora():%d/%m/%Y às %H:%M} ({dur_txt})"],
            ] + linhas + [["ITAÚ RET", "importado antes do painel"]],
        })
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
        # 0xC000013A = STATUS_CONTROL_C_EXIT: alguem fechou a janela (ou Ctrl+C)
        if codigo in (3221225786, -1073741510):
            resto = "A janela foi FECHADA antes de terminar. " + resto
        registrar(f"TERMINOU COM ERRO (codigo {codigo}) em {dur_txt}. {resto}")
        avisar_na_tela({
            "status": "erro",
            "titulo": "PAINEL ANÁLISE DE BOLETOS NÃO FOI ATUALIZADO",
            "linhas": [
                ["Erro", f"código {codigo} às {agora():%H:%M} ({dur_txt})"],
                ["", resto],
            ],
        })
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:  # nunca morrer calado dentro do Agendador
        registrar(f"FALHA DO GATILHO: {type(e).__name__}: {e}")
        raise
