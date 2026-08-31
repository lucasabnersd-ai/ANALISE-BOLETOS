# Prompt — atualizar SÓ a aba "NF x Pedido de Compra"

Cole o bloco abaixo no Claude Code (ou siga na mão) **na máquina que vai
atualizar**. Funciona igual na máquina do Lucas e na da Gabriella: os caminhos
são descobertos, não escritos.

---

## O PROMPT (copiar daqui para baixo)

> Preciso atualizar **somente a aba `NF x Pedido de Compra`** do painel
> ANALISE-BOLETOS (https://lucasabnersd-ai.github.io/ANALISE-BOLETOS/PUBLICAR/)
> e publicar. Siga exatamente os passos abaixo e **não invente atalho**.
>
> ### 1. Ache a pasta certa
> A pasta do painel é a que tem, ao mesmo tempo, `ATUALIZAR_PAINEL.cmd`,
> `gerar_painel.py`, `sc7.py` e o arquivo oculto **`.analise_boletos_token`**.
> Ela fica em `...\LUCAS ABNER ARAUJO\AUTOMAÇÕES LUCAS\ANALISES BOLETOS\PAINEL ANALISE BOLETOS\`,
> dentro da biblioteca do OneDrive — que **muda de nome em cada máquina**:
> - Lucas: `C:\Users\lucas\OneDrive - Grupo S&D\Arquivos de Gabriella Karla Oliveira Milas - FINANCEIRO COMPARTILHADO\...`
> - Gabriella: `C:\Users\gabriella.milas\OneDrive - S & D Florestal\0 GABRIELLA\FINANCEIRO COMPARTILHADO\...`
>
> ⚠ Existe um **clone velho e enganoso** em `Documents\New project 8\analise-boletos-pages`.
> Se a pasta não tiver o `.analise_boletos_token`, é ele — não trabalhe ali.
>
> ### 2. Confira as 4 bases que alimentam ESTA aba
> A aba cruza nota × lançamento × pedido. Antes de rodar, veja a **data de
> modificação** de cada arquivo e me diga qual está velha:
>
> | Arquivo | Onde | O que é |
> |---|---|---|
> | `SC7.xlsx` | `LUCAS ABNER ARAUJO\BASES GENERICOS\` | pedidos de compra (por item) |
> | `SC1.xlsx` | `LUCAS ABNER ARAUJO\BASES GENERICOS\` | solicitações — é daqui que vem o **Solicitante** |
> | `SF1.xlsx` | `LUCAS ABNER ARAUJO\BASES GENERICOS\` | notas lançadas no TOTVS |
> | `SEFAZ.xlsx` | `...\ANALISES BOLETOS\` | as notas da SEFAZ |
> | `LISTAGEM EMPRESAS BIOFLOR.xlsx` | `LUCAS ABNER ARAUJO\BASES GENERICOS\` | filtro de tomadores — **sem ela a aba NÃO atualiza** |
>
> O que estiver desatualizado tem de ser **reexportado do TOTVS e salvo por cima,
> com o mesmo nome**, antes de seguir. Não dá para "atualizar a aba" sem base nova.
>
> ### 3. Rode
> Na pasta do painel:
>
> ```
> ATUALIZAR_PAINEL.cmd /sembase
> ```
>
> O `/sembase` é o que segura o escopo: **pula a etapa 1/4** (o associador de
> boletos, que demora minutos, refaz a `TRATAMENTO PYTHON BOLETOS.xlsx` e ainda
> publica o painel BOLETOS-PENDENTES). Restam:
> `[2/4]` monta o painel · `[3/4]` sobe a carteira para o Supabase · `[4/4]` publica no Pages.
> Demora alguns minutos (a SC7 tem ~14 MB). Pode rodar com as planilhas abertas
> no Excel — o leitor faz cópia temporária.
>
> ### 4. Confira a saída ANTES de dizer que está pronto
> Esta é a parte que não pode ser pulada: quando uma base falta, o painel **não
> quebra** — a aba fica com o dado da rodada anterior e sai só um AVISO no meio
> do texto. Procure, no que o `.cmd` imprimiu:
>
> - ❌ `AVISO: SC7 nao encontrada, aba de pedidos nao criada`
> - ❌ `AVISO: aba de pedidos de compra nao atualizada: ...`
> - ❌ `AVISO: abas da SEFAZ nao atualizadas` (falta a LISTAGEM EMPRESAS BIOFLOR)
> - ❌ `AVISO: base SEFAZ nao encontrada` / `AVISO: base SF1 nao encontrada`
>
> Qualquer um desses = **a aba não foi atualizada**. Corrija a base e rode de novo.
>
> Tem de aparecer, no resumo final, a linha da aba com contagem plausível
> (hoje ~617 notas, ~143 com pedido amarrado):
> ```
> NF x Pedido de Compra: <N> linhas | <M> colunas | com alerta: ... | boleto difere: ...
> ```
> E, no fim: `CONCLUIDO. Painel publicado em https://...`
>
> ### 5. Me relate
> Datas das 4 bases, a linha de contagem da aba, qualquer AVISO que apareceu, e
> se o passo `[4/4]` publicou ou disse "nada mudou".

---

## O que este prompt NÃO faz (e por que está certo assim)

**Não existe opção "gerar só uma aba".** O `gerar_painel.py` remonta o painel
inteiro. O `/sembase` é o recorte possível, e ele é honesto:

- ✅ **`NF x Pedido de Compra` atualiza** — é o objetivo.
- ⚠ **`Análise Base SEFAZ` e `SEFAZ x SF1` também atualizam**, porque leem as
  mesmas SEFAZ.xlsx/SF1.xlsx. É esperado, e é o certo: deixá-las para trás faria
  duas abas afirmarem coisas diferentes sobre a mesma nota.
- ✅ **As abas de boletos ficam como estão** — a `TRATAMENTO PYTHON BOLETOS.xlsx`
  não é regerada.
- ✅ **Nenhum check/tratativa se perde.** As marcações moram no Supabase, não no
  arquivo; a carga não regrava título já tratado.

## Regras da aba que valem lembrar ao conferir o resultado

- **Só pedido EM ABERTO entra** (SC7, coluna `AU · Ped. Encerr.` sem "E"). Dos
  ~3.624 pedidos, ~3.117 ficam de fora. `PARCIAL` (parte dos itens encerrada)
  **fica**, e entra com todos os itens somados.
- Nota que cita um pedido já encerrado não vira "fora da base": vira o veredito
  **`PEDIDO ENCERRADO`**, com o número no Critério.
- **O número do pedido escrito na nota não é prova.** A SC7 tem PC de 1 a ~3492
  quase sem buraco, então qualquer número de 3–4 dígitos "existe". Quem decide é
  o texto **+** a corroboração por fornecedor e valor.
- Se a SC7 mudar de layout (coluna inserida), o `sc7.py` reencontra o bloco
  deslocado e avisa quanto andou. Se ele **parar com erro**, o certo é
  **reexportar a SC7** — não mexer no código.

## Se der errado

| Sintoma | O que fazer |
|---|---|
| `ERRO: nao encontrei o Python nesta maquina` | `set PAINEL_PYTHON=C:\caminho\python.exe` e rodar de novo (precisa de Python 3.12 com `openpyxl`) |
| `ERRO no push` na etapa 4/4 | Rodar `git push` **uma vez na mão** na pasta do painel, para o Windows guardar a credencial do repo privado `lucasabnersd-ai/ANALISE-BOLETOS` |
| `AVISO: falta o arquivo .analise_boletos_token` | O painel **local** atualizou, mas o publicado continua com a carga anterior. Esperar o OneDrive sincronizar o token |
| Etapa 4/4 diz "nada mudou" | Normal quando só os dados mudaram: o HTML publicado é só layout, os dados vão pelo Supabase na etapa 3/4 |

## Nunca

- Editar `PUBLICAR\index.html` (é gerado). Layout se muda em `painel_modelo.html`.
- Commitar planilha, `DADOS\` ou o token — o `.gitignore` barra, e é de propósito.
- Rodar **sem** `/sembase` achando que é "mais completo": isso refaz a associação
  de boletos e republica outro painel junto.

---
*Escrito em 31/08/2026. Ver `PLANO_ABA_PEDIDOS_SEFAZ.md` para o desenho da aba.*
