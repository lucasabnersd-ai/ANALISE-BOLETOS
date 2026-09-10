# PAINEL DE ANÁLISE DE BOLETOS

**No ar em dois endereços, do mesmo push:**

| onde | endereço | o que publica |
|---|---|---|
| Vercel *(preferido para mandar para fora)* | https://analise-boletos.vercel.app/ | só a pasta `PUBLICAR\` |
| GitHub Pages | https://lucasabnersd-ai.github.io/ANALISE-BOLETOS/PUBLICAR/ | o repositório inteiro |

## Para atualizar

Duplo clique em **`ATUALIZAR_PAINEL.cmd`**, aqui nesta pasta. São 6 passos:

1. refaz a base `TRATAMENTO PYTHON BOLETOS.xlsx` (chama o
   `RODAR_ASSOCIADOR_BOLETOS.CMD`; é o passo demorado, alguns minutos);
2. lê as planilhas e monta o painel;
3. avisa a Graziela de `AMARRAR ADIANTAMENTO` que está com boleto;
4. avisa dos títulos em aberto vencendo em sábado, domingo ou feriado;
5. sobe a carteira para o Supabase;
6. faz commit e push — é o que troca os dois endereços.

**Se a base não for gerada, nada é publicado.** É de propósito: painel novo em
cima de base velha é impossível de perceber depois.

Para republicar depressa, sem refazer a associação:

```
ATUALIZAR_PAINEL.cmd /sembase
```

Os dois avisos da Graziela também rodam sozinhos no fim da tarde
(`ALERTAS_DIARIOS.cmd`, pelo Agendador de Tarefas) e na rodada do SE2. A trava
diária (`DADOS\alertas_enviados.json`) garante um e-mail por dia: quem rodar
primeiro manda.

## As planilhas que ele lê

Nenhuma delas mora aqui — os programas as procuram sozinhos.

| planilha | onde |
|---|---|
| `TRATAMENTO PYTHON BOLETOS.xlsx` | `LUCAS ABNER ARAUJO\` |
| `COPIAR E COLAR BOLETOS PENDENTES.xlsx` | `LUCAS ABNER ARAUJO\` |
| `SF1`, `SC7`, `SC1`, `SE2 - POSIÇÃO DIARIA`, `LISTAGEM EMPRESAS BIOFLOR` | `LUCAS ABNER ARAUJO\BASES GENERICOS\` |
| `SEFAZ.xlsx`, `MESMA PREMISSA*.xlsx` | `AUTOMAÇÕES LUCAS\ANALISES BOLETOS\` |

Da `SEFAZ.xlsx` entram **quatro abas**, todas achadas pelo **cabeçalho** (o nome
da aba muda de exportação para exportação): a de NF-e (`sefaz.py`), a de NFS-e
(`sefaz.py`), a **NFS STATUS** (`nfs_status.py` — situação e integração ERP das
NFS-e, cruzada por CNPJ do prestador + nº da NF sem o ano grudado) e a
**PREMISSA 2 SEFAZ** (`cte_sefaz.py` — os CT-e de frete, que entram como notas
com Origem `CT-e`; o nº do CT-e faz o papel do nº da NF no cruzamento com a SF1).
Linha 100% repetida em qualquer aba é ignorada, com aviso no log.

**Como são achadas:** o `.cmd` e o `caminhos.py` **sobem de pasta em pasta até
achar a que tem a `BASES GENERICOS` dentro** — essa é a `LUCAS ABNER ARAUJO` —
e descem de lá. Não existe caminho contado por nível (`..\..\..`) nem
`C:\Users\<fulano>` no meio do caminho: por isso a pasta pôde vir para cá em
10/09/2026 sem quebrar nada, e por isso funciona igual na máquina da Gabriella,
onde a biblioteca tem outro nome.

Base que não é encontrada **não derruba a rodada**: a aba dela fica como estava
e sai um `AVISO: nao encontrei <nome>` no meio do log. Vale a pena ler os
avisos.

## O que tem em cada pasta

| pasta | o que é |
|---|---|
| `PUBLICAR\` | o que vai ao ar (`index.html`, `supa.js`, `vercel.json`) |
| `DADOS\` | a carteira em JSON, os históricos e as prévias dos e-mails |
| `BACKUPS\` | cópias do Supabase |
| `.git\` | **o repositório que publica** — é este clone, aqui mesmo |

Este é o **único** clone do `ANALISE-BOLETOS` na máquina, e é de dentro dele
que o passo 6 publica. Se um dia aparecer um segundo, o
`PAINEIS\verificar_clones_paineis.py` acusa.

## ⚠ Duas coisas que já derrubaram o painel

- **Privar o repositório APAGA o site do Pages** — e voltar a público **não** o
  recria (08/09/2026). O caminho certo é `DELETE /repos/.../pages`.
- **Planilha aberta no Excel** (a SE2, a SF1) trava a leitura ou apaga as abas
  que dependem dela. Feche antes de rodar.
