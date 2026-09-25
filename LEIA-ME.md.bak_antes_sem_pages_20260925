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

Da `SEFAZ.xlsx` entram **cinco abas**, todas achadas pelo **cabeçalho** (o nome
da aba muda de exportação para exportação; desde 18/09/2026 vêm como
`PRODUTO 1`, `PRODUTO 2`, `SERVICO`, `SERVICO 2` e `CTE`):

| aba | o que é | quem lê |
|---|---|---|
| `PRODUTO 1` | as NF-e, por item × duplicata | `sefaz.py` |
| `PRODUTO 2` | a manifestação do destinatário (DF-e, 14 colunas). **Complementa a PRODUTO 1**: preenche a coluna Manifestação das notas que estão nas duas e **acrescenta** as que só existem aqui (Origem `PRODUTO 2`) | `manifestacao.py`, `premissa2.py` |
| `SERVICO` | as NFS-e | `sefaz.py` |
| `SERVICO 2` | o relatório de situação das NFS-e (Integração ERP, situação na prefeitura). **Complementa a SERVICO**: preenche a coluna NFS STATUS e acrescenta as que só existem aqui (Origem `SERVICO 2`) | `nfs_status.py` |
| `CTE` | os CT-e de frete, que entram como notas com Origem `CT-e` | `cte_sefaz.py` |

⚠ Na `PRODUTO 2` a chave de 44 dígitos vem **como número** (o Excel guarda 15
dígitos e perde os outros 29). Por isso a nota casa com a `PRODUTO 1` por **CNPJ
do emitente + nº + série**, com o número truncado servindo de trava; a chave
inteira é recuperada dos `MESMA PREMISSA*.xlsx` da pasta quando algum a tem, e
sem ela a `Chave NF-e` sai como `PRODUTO 2 <cnpj> <nº> <série>`. Na `SERVICO 2`
o nº da NFS-e e o RPS também vêm como número e são arredondados quando passam de
15 dígitos — daí o terceiro casamento por prestador + dia + valor.
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
