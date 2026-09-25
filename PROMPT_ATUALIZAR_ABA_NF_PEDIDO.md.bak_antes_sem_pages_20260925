# Prompt — atualizar SÓ a aba "NF x Pedido de Compra"

Para a **segunda máquina** (a da Gabriella, ou qualquer PC com a biblioteca do
OneDrive sincronizada). Cole a seção **O PROMPT** no Claude Code de lá.

## Antes de tudo: o que a máquina precisa ter

| | Precisa? | Por quê |
|---|---|---|
| **Conta / login / CLI do Supabase** | ❌ **Não** | A carga é um POST HTTPS comum, autenticado por um header próprio (`x-carga-token`) cujo valor está em `.analise_boletos_token`, na própria pasta — e esse arquivo chega pelo OneDrive. Só biblioteca padrão do Python, zero `pip install` |
| **Internet** | ✅ Sim | Para o POST em `*.supabase.co` (etapa 3/4) |
| **Python 3.12 + openpyxl** | ✅ Sim | O `.cmd` procura em `%PAINEL_PYTHON%`, depois `%LOCALAPPDATA%\Programs\Python\Python312\`, depois `py -3`, depois o `python` do PATH |
| **Git + credencial do GitHub** | ⚠️ **Só para a etapa 4/4** | O repo é privado, então até o `git fetch` pede login. O `.gitconfig` que fixa a conta **fica no perfil do usuário e NÃO sincroniza** pelo OneDrive |

**A etapa 4/4 não carrega dado nenhum.** O HTML publicado é só layout — a
carteira vai pelo Supabase na etapa 3/4. Numa rodada em que só as planilhas
mudaram, o único diff do `PUBLICAR/index.html` é o carimbo `supa.js?v=…`.
Se o push falhar, **o painel no ar já está atualizado**; o que fica velho é um
carimbo de cache. Ela pode ignorar o erro final da etapa 4/4 e avisar você.

⚠️ **O `.git` mora dentro do OneDrive e o `deploy.py` faz `git add .` na pasta
inteira.** Se o OneDrive dela estiver atrasado, o push pode empurrar uma versão
**antiga** dos fontes por cima da boa — e os commits saem assinados como Lucas.
Por isso: **se ela não tiver credencial do GitHub, é melhor assim mesmo.**

---

## O PROMPT (copiar daqui para baixo)

> Preciso atualizar **somente a aba `NF x Pedido de Compra`** do painel
> ANALISE-BOLETOS (https://lucasabnersd-ai.github.io/ANALISE-BOLETOS/PUBLICAR/).
> Siga os passos abaixo e **não invente atalho**.
>
> ### 1. Ache a pasta certa
> É a que tem, ao mesmo tempo, `ATUALIZAR_PAINEL.cmd`, `gerar_painel.py`,
> `sc7.py` e o arquivo oculto **`.analise_boletos_token`**. Fica em
> `...\LUCAS ABNER ARAUJO\AUTOMAÇÕES LUCAS\ANALISES BOLETOS\PAINEL ANALISE BOLETOS\`,
> dentro da biblioteca do OneDrive — cujo nome **muda em cada máquina**:
> - `C:\Users\lucas\OneDrive - Grupo S&D\Arquivos de Gabriella Karla Oliveira Milas - FINANCEIRO COMPARTILHADO\...`
> - `C:\Users\gabriella.milas\OneDrive - S & D Florestal\0 GABRIELLA\FINANCEIRO COMPARTILHADO\...`
>
> ⚠ Existe um **clone velho e enganoso** em `Documents\New project 8\analise-boletos-pages`.
> Sem o `.analise_boletos_token`, é ele — não trabalhe ali.
>
> ### 2. Confira o OneDrive ANTES de rodar
> Os arquivos têm de estar **baixados de verdade**, não como marcador "somente na
> nuvem": os caminhos padrão são resolvidos **uma vez só**, no import. Se a pasta
> `BASES GENERICOS` não estiver materializada naquele instante, o script cai num
> caminho de outra máquina pelo resto da rodada.
>
> ### 3. As 5 bases desta aba
> Confira a **data de modificação** de cada uma e me diga qual está velha:
>
> | Arquivo | Onde | Papel |
> |---|---|---|
> | `SC7.xlsx` | `LUCAS ABNER ARAUJO\BASES GENERICOS\` | pedidos de compra (por item) |
> | `SC1.xlsx` | `LUCAS ABNER ARAUJO\BASES GENERICOS\` | é daqui que vem o **Solicitante** |
> | `SF1.xlsx` | `LUCAS ABNER ARAUJO\BASES GENERICOS\` | notas lançadas no TOTVS |
> | `SEFAZ.xlsx` | `...\ANALISES BOLETOS\` | as notas da SEFAZ |
> | `LISTAGEM EMPRESAS BIOFLOR.xlsx` | `LUCAS ABNER ARAUJO\BASES GENERICOS\` | filtro de tomadores |
>
> ⚠ **A SF1 é obrigatória**, mesmo que ninguém a reexporte: sem ela a aba de
> pedidos **não é construída**. E SF1 velha é pior que SF1 ausente — a aba sai
> completa e confiante, acusando como "não lançada" toda nota lançada depois da
> última exportação. Não existe alerta de idade para ela; o `.cmd` só imprime
> `Base SF1: … (planilha salva em …)`. **Leia essa data.**
>
> ### 4. Se alguma base estiver em OUTRA pasta
> Todas as cinco aceitam caminho na linha de comando. Passe o que for preciso:
>
> ```
> ATUALIZAR_PAINEL.cmd /sembase --base-sc7 "D:\caminho\SC7.xlsx" --base-sc1 "D:\caminho\SC1.xlsx"
> ```
>
> Os argumentos são `--base-sefaz`, `--base-sf1`, `--base-sc7`, `--base-sc1` e
> `--base-empresas`. O que não for passado usa o arquivo em `BASES GENERICOS`.
>
> ### 5. Rode
> ```
> ATUALIZAR_PAINEL.cmd /sembase
> ```
> O `/sembase` pula a etapa **1/4** (o associador de boletos: demora minutos,
> refaz outra planilha e publica um segundo painel junto). Restam:
> `[2/4]` monta · `[3/4]` sobe a carteira para o Supabase · `[4/4]` publica no Pages.
> Alguns minutos (a SC7 tem ~14 MB). Pode rodar com as planilhas abertas no Excel.
>
> ### 6. Confira a saída ANTES de dizer que está pronto
> **Base faltando não derruba a rodada** — vira `AVISO` no meio do texto e o
> `.cmd` termina dizendo `CONCLUIDO` do mesmo jeito. Procure por:
>
> - ❌ `AVISO: sem a SF1 as abas SEFAZ x SF1 e NF x Pedido de Compra NAO foram atualizadas`
> - ❌ `AVISO: SC7 nao encontrada, aba de pedidos nao criada`
> - ❌ `AVISO: SC1 nao encontrada -- a aba de pedidos sai com a coluna Solicitante VAZIA`
> - ❌ `AVISO: aba de pedidos de compra nao atualizada: …`
> - ❌ `AVISO: abas da SEFAZ nao atualizadas` / `AVISO: base SEFAZ nao encontrada`
>
> Qualquer um = **a aba não foi atualizada**. Corrija e rode de novo.
>
> Tem de aparecer, no resumo final:
> ```
> NF x Pedido de Compra: <N> linhas | 40 colunas | com alerta: … | boleto difere: …
> ```
> Hoje N ≈ **692**. Muito abaixo disso = base incompleta.
>
> ### 7. A etapa 4/4 pode falhar — e tudo bem
> Se aparecer `ERRO no push` no fim: **os dados já foram publicados na etapa 3/4**
> e o painel no ar está correto. Não tente resolver credencial do GitHub sozinha,
> não rode `git push`, não force nada. Me avise e siga.
>
> ### 8. Me relate
> As datas das 5 bases, a linha de contagem da aba, todo `AVISO` que apareceu,
> e o que a etapa 4/4 disse.

---

## O que este prompt NÃO faz

**Não existe opção "gerar só uma aba"** — o `gerar_painel.py` remonta tudo, e o
`.cmd` não tem chave para pular a etapa 4/4. O `/sembase` é o recorte possível:

- ✅ **`NF x Pedido de Compra` atualiza** — é a única aba visível que muda.
- ➖ `Análise Base SEFAZ` e `SEFAZ x SF1` leem as mesmas bases, mas estão
  **ocultas** (`oculta: True`): viajam só com o cabeçalho, sem uma linha.
- ✅ As abas de boletos ficam como estão — a base delas não é regerada.
- ✅ **Nenhum check ou tratativa se perde** — moram no Supabase.
- ⚠️ Os históricos em `DADOS\` (SEFAZ, Itaú, classificação) são **reescritos** e
  ficam no OneDrive: as duas máquinas compartilham o mesmo arquivo. Não rodem as
  duas ao mesmo tempo.

## Regras da aba, para conferir o resultado

- **Só pedido EM ABERTO entra** (SC7, coluna `AU · Ped. Encerr.` sem "E"): dos
  ~3.624, ficam ~507. `PARCIAL` **fica**, com todos os itens somados.
- Nota que cita pedido já encerrado não vira "fora da base": vira o veredito
  **`PEDIDO ENCERRADO`**, com o número no Critério.
- **O número escrito na nota não é prova.** A SC7 tem PC de 1 a ~3492 quase sem
  buraco — qualquer número de 3–4 dígitos "existe". Decide o texto **+** a
  corroboração por fornecedor e valor.
- Se a SC7 mudar de layout, o `sc7.py` reencontra o bloco deslocado e avisa
  quanto andou. Se **parar com erro**, o certo é **reexportar a SC7** — não mexer
  no código.

## Se der errado

| Sintoma | O que fazer |
|---|---|
| `ERRO: nao encontrei o Python nesta maquina` | `set PAINEL_PYTHON=C:\caminho\python.exe` e rodar de novo |
| `ERRO no push` na etapa 4/4 | **Ignorar.** Os dados já subiram na 3/4. Avisar o Lucas |
| Janela do console parada, sem mensagem | Provável janela de login do GitHub escondida (o `git` roda com a saída capturada). Fechar e avisar |
| `AVISO: falta o arquivo .analise_boletos_token` | O painel **local** atualizou, o publicado não. Esperar o OneDrive |
| Etapa 4/4 diz "nada mudou" | Normal — os dados vão pelo Supabase na 3/4 |

## Nunca

- Editar `PUBLICAR\index.html` (é gerado). Layout se muda em `painel_modelo.html`.
- Rodar `git push`, `git reset` ou `git checkout` nessa pasta.
- Rodar **sem** `/sembase`: refaz a associação de boletos e republica outro painel.
- Rodar nas duas máquinas ao mesmo tempo (o `.git` e o `DADOS\` são compartilhados).

---
*31/08/2026. Desenho da aba: `PLANO_ABA_PEDIDOS_SEFAZ.md`.*
