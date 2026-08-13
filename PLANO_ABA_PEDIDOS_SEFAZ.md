# Plano: colunas novas na SEFAZ x SF1 + aba de Pedidos de Compra

Escrito em **13/08/2026** para ser executado por outra sessão do Claude, a pedido
do usuário. Contém o pedido dele, **a realidade medida das bases** (não suposta),
as regras propostas e as perguntas que ainda precisam de resposta.

> Tudo que está em **MEDIDO** foi conferido nos arquivos reais nesta data. Os
> números importam: são eles que dizem se uma regra funciona ou é chute.

---

## 0. Como o painel funciona hoje (leia antes de mexer)

Pasta: `…\LUCAS ABNER ARAUJO\AUTOMAÇÕES LUCAS\ANALISES BOLETOS\PAINEL ANALISE BOLETOS\`
Publicado (repo privado + Pages): <https://lucasabnersd-ai.github.io/ANALISE-BOLETOS/PUBLICAR/>

| Arquivo | Papel |
| --- | --- |
| `gerar_painel.py` | Orquestra tudo. Dicionário **`ABAS`** = uma entrada por aba; `COMPACTA_*` = colunas na tela; `montar_aba()` monta os registros |
| `sefaz.py` | Lê `ANALISES BOLETOS\SEFAZ.xlsx` (abas `NFes SEFAZ` e `NFS`) e os 3 alertas |
| `cruzamento_sefaz_sf1.py` | Cruza cada NOTA da SEFAZ com a SF1 inteira, pelo número da NF |
| `cruzamento_classificacao.py` | SF1 × boletos sem título (outra aba) |
| `verificacao_se2.py` | Tira do painel título com boleto lançado/baixado |
| `painel_modelo.html` | **A tela.** Editar SEMPRE aqui, nunca o `PUBLICAR/index.html` gerado |
| `publicar_dados.py` | Sobe a carteira para o Supabase (os dados **não** vão no HTML) |
| `deploy.py` | `git add/commit/push` do `PUBLICAR/` |
| `ATUALIZAR_PAINEL.cmd` | 4 passos: base → gerar → carteira → publicar. `/sembase` pula o passo 1 |
| `planilha_excel_js.py` | Motor do **Exportar Excel** (.xlsx), injetado no HTML pelo marcador `<!--__PLANILHA__-->` |

### Convenções que NÃO podem ser quebradas

1. **Colunas de base grande são lidas por POSIÇÃO (letra), com o nome como
   trava.** SF1 tem 220 colunas, SC7 tem 207, a SEFAZ tem 386 — e há nomes
   repetidos e grafia irregular. Ler por nome quebra calado. Ver `_conferir()` em
   `sefaz.py`: se o nome não estiver onde deveria, a leitura **para** com a
   mensagem do que saiu do lugar.
2. **O painel decide o desenho pelo PAPEL da coluna, nunca pelo nome**:
   `col_status`, `col_criterio`, `col_alerta`, `col_uuid` no `ABAS`.
3. **`pills` é uma LISTA de grupos** (barra de botões). `grupos_de_pills()` só
   aceita grupo com **2 a 8 valores** — 1 não filtra nada, 40 vira parede.
4. **A chave de marcação (`col_uuid` + `prefixo_uuid`) é sagrada**: é ela que
   amarra check/tratativa no Supabase e sobrevive à regeração. Prefixos em uso:
   `ITAU:`, `BOL:`, `SEFAZ:`, `SFZ1:`, `SF1:`. **Aba nova precisa de prefixo novo.**
5. Os dados **nunca** entram no HTML publicado (LGPD). O HTML vai vazio; a
   carteira sobe para o Supabase e é baixada após login. `deploy.py` bloqueia
   publicação se achar 20+ dígitos seguidos em `PUBLICAR/`.
6. Peso: a carteira já está em **4,0 MB / 3.274 títulos**. Cada coluna nova
   multiplica pelo número de linhas. Coluna de texto longo tem teto
   (`TETO_TEXTO = 1200` em `sefaz.py`).

---

## 1. MEDIDO — as bases novas

### SC7.xlsx (pedidos de compra) — `BASES GENERICOS\SC7.xlsx`
Aba única `SC7`, **207 colunas**, **24.401 itens de pedido**, **3.376 PCs distintos**.
⚠ É **por ITEM**: um PC tem N linhas. Toda leitura precisa agregar (ver §5).

| Coluna pedida | Letra | Observação |
| --- | --- | --- |
| `Numero PC` | **AI** | 6 dígitos com zeros à esquerda (`000007`) — a chave do cruzamento |
| `Qtd.a Classi` | **BH** | |
| `Numero da SC` | **AB** | liga com a SC1 |
| `Controle Ap.` | **BM** | ex.: `L` |
| `Usuário SC` | **GU** | ⚠ preenchido em só **83 de 24.401** itens |
| `Comprador` | **GV** | preenchido em **23.640 de 24.401** (96,9%) |
| `Ped. Encerr.` | **AU** | ex.: `E` |
| `Dt. Entrega` | **L** | data |
| `1° Venc` | **GM** | data |
| `Solicitante` | CP | ⚠ **VAZIO EM 100% DAS LINHAS (0 de 24.401)** — confirma o que o usuário disse |
| úteis: `Filial` A · `Razão Social` B · `Item` D · `Produto` E · `Descricao` F · `Quantidade` I · `Vlr.Total` V · `Fornecedor` AG · `Centro Custo` M · `DT Emissao` AH · `Cod. Usuario` BP · `Campo UUID` GO |

### SC1.xlsx (solicitações de compra) — `BASES GENERICOS\SC1.xlsx`
Aba única `SC1`, **119 colunas**.

| Coluna | Letra | Uso |
| --- | --- | --- |
| `Solicitante` | **Y** | o que falta no SC7 (ex.: `andson.oliveira`) |
| `Num. Pedido` | **Z** | **o PC** — é por aqui que se cruza |
| `Numero da SC` | B | |
| `Usuário` | DI | mesmo valor do Solicitante nos exemplos vistos — reserva |
| `Requisitante` | CZ | código (`000094`) — reserva |
| `Observacao` | P · `Descricao` E · `Centro Custo` K · `Fornecedor` W |

---

## 2. MEDIDO — como o PC/OC aparece nos textos da SEFAZ

Origem: coluna `Info Comp` (aba `NFes SEFAZ`) e `Descrição do Serviço` (aba `NFS`).
No painel as duas já chegam unificadas na coluna **`Informações`** (`sefaz.INFO`).

| | Notas |
| --- | --- |
| Notas na SEFAZ.xlsx | **617** |
| Com algum texto em Informações/Descrição | **609** |
| Com PC/OC reconhecido (regex proposta + travas) | **194** |
| …**que existem no SC7** | **146** (75%) |
| …que **não** existem no SC7 | **48** — ver §6, pergunta 3 |

**Rótulos encontrados** (o número vem depois): `PEDIDO`/`PEDIDO DE COMPRA` (163),
`OC` (23), `ORDEM DE COMPRA` (15), `PC`/`P.C.` (7).

**Formatos reais** (todos precisam casar):

```
PEDIDO DE COMPRA: 003100      PEDIDO: 2993         PC 003180
PEDIDO DE COMPRA 002985.C…    OC:1432              P.C. 001617
ORDEM DE COMPRA: 003230       OC 001412            OC 3204
REFERENTE AO PEDIDO DE COMPRA 3125                 PEDIDO DE COMPRA: 003100/1
```

⚠ **4 dígitos ou 6 (com 2 zeros na frente) são a MESMA coisa.** `003100` = `3100`.
Normalizar tirando zeros à esquerda dos dois lados.
⚠ Pode vir **`/1` no fim** (`003100/1`) = item do pedido. O `/1` **não** faz parte
do número.
⚠ **ARMADILHAS MEDIDAS — números que NÃO são o nosso PC:**
- `PEDIDO GDS: 3175071` → pedido do **fornecedor** (7 dígitos).
- `NFC-E: 32160 \ 32156` → número de nota.
- `PEDIDO: 290276` → fora da faixa dos nossos PCs.
- Qualquer coisa com **`GDS`, `VENDA(S)`, `NFC-E`, `NOTA`** perto do número.

**Regra de extração proposta** (medida: 194 achados, 146 confirmados no SC7):

```python
RUIM   = r'\b(GDS|VENDA|VENDAS|NFC-?E|NOTA)\b'          # veta a vizinhança
PADRAO = (r'\b(?:P\.?\s?C\.?|O\.?\s?C\.?'
          r'|PEDIDO(?:\s+DE\s+COMPRA)?|ORDEM(?:\s+DE\s+COMPRA)?)'
          r'\s*[:\-\.]?\s*(\d{3,6})(?!\d)')
# texto em MAIÚSCULAS; olhar ~22 chars antes até 8 depois para o veto RUIM;
# depois: numero.lstrip("0") e EXIGIR que exista no SC7.
```

⚠ **A validação contra o SC7 é parte da regra, não conferência.** Ela é o que
separa PC de número do fornecedor. Se o número não estiver no SC7, o veredito é
`PC NÃO ENCONTRADO` — **diferente** de `SEM PC NO TEXTO`. Os dois casos precisam
de selos distintos: um é lacuna de cadastro/base, o outro é a nota não citar
pedido.

---

## 3. Pedido do usuário, item por item

### 3.1 Colunas novas na aba **SEFAZ x SF1** (`ABAS["sefaz_sf1"]`, `COMPACTA_SEFAZ_SF1`)
Trazer da base SEFAZ: **`Vlr Duplicata`**, **`Venc Duplicata`**, **`Nome Fantasia`**,
**`CFOP`**, **`Tipo Operação`**, **`Saída/Entrada`**, e **`Nome Dest`** (nome do
destinatário **+ CNPJ**, copiável em 1 clique).

⚠ **Problema de nível de dado, resolver antes de codar:** a aba `sefaz_sf1` é
**uma linha por NOTA**, mas `Vlr Duplicata`/`Venc Duplicata` são **por
duplicata** (a nota pode ter 3). `notas_da_sefaz()` hoje descarta isso.
Sugestão: levar a **1ª duplicata** (menor vencimento) e, quando houver mais de
uma, marcar quantas são (o `sefaz.vencimentos_por_nota()` já agrega os
vencimentos). Confirmar com o usuário — §6, pergunta 4.

Copiável em 1 clique = `"copiar_extra"` no `ABAS` (padrão já usado: `Nº NF`,
`CNPJ Emitente`). CNPJ junto do nome segue o desenho do `Fornecedor` das outras
abas (código ao lado da razão social).

### 3.2 Remover da SEFAZ x SF1
- **`Emissão SF1`** → tirar de `COMPACTA_SEFAZ_SF1` (e do `CABECALHO` do
  `cruzamento_sefaz_sf1.py` se não for mais usada por nada).
- **Coluna `#` (contador de linha)** → é a numeração da grade no
  `painel_modelo.html`.
- **`Campo UUID`** → tirar **da tela**. ⚠ A chave continua existindo por dentro
  (`col_uuid`/marcações); só a coluna sai. Não apagar o `col_uuid`.

### 3.3 Exportar Excel
As colunas novas têm de sair na exportação **sem erro**. Pontos de atenção já
conhecidos:
- `ROTULOS_PLANILHA_SEFAZ_SF1` dá o nome por extenso de cada coluna no .xlsx.
- ⚠ No `ABAS`, **`planilha`** = aba de ORIGEM; o nome da guia do .xlsx é
  **`guia`** (já custou uma exportação com guia errada).
- Data e moeda precisam do tipo certo (`data`/`moeda`) senão saem como texto.

### 3.4 Arrastar a tabela com o mouse
Pan por clique-e-arraste no contêiner `.planilha` (que já tem `overflow-x`).
Implementar no `painel_modelo.html` com `pointerdown/move/up` alterando
`scrollLeft/scrollTop`. ⚠ Não pode roubar o clique dos botões de copiar, do
check e do selo que abre quadro: só entrar em modo arraste **depois de ~4 px de
movimento**, e ignorar o gesto se começou sobre `input`, `button` ou `[data-copiar]`.

### 3.5 Aba nova: NF × Pedido de Compra
Uma aba com o que foi pedido da `SEFAZ x SF1` **e** da `Análise Base SEFAZ`,
**mais `Nat. Operação`** (coluna da base SEFAZ ainda **não lida** — precisa
entrar em `COLUNAS_NFE`/`COLUNAS_NFS` do `sefaz.py`, por letra, com o nome como
trava), **mais** os dados do pedido:

`Numero PC` · `Numero da SC` · `Solicitante` · `Comprador` · `Usuário SC` ·
`Qtd.a Classi` · `Controle Ap.` · `Ped. Encerr.` · `Dt. Entrega` · `1° Venc`
— todas **filtráveis** (barra de filtros e/ou `pills`, respeitando o teto de 8
valores por grupo).

**NF no padrão TOTVS:** dígitos, sem zeros à esquerda, **completada com zeros à
frente até 9 dígitos** (`sefaz.nf_chave()` já normaliza; o padding de 9 é o
formato de exibição/cópia — é o que se cola no TOTVS).

### 3.6 Solicitante e Comprador
- `Comprador` sai do **SC7 (GV)** — presente em 96,9% dos itens.
- `Solicitante` **não existe no SC7** (medido: 0 de 24.401). Vem da **SC1**,
  cruzando **`Num. Pedido` (Z)** com o PC e lendo **`Solicitante` (Y)**.
  Reservas se vier vazio: `Usuário` (DI) e `Requisitante` (CZ).
- ⚠ Um PC pode ter itens de **várias SCs**, logo vários solicitantes. Definir:
  mostrar o primeiro, ou todos separados por `·`? (§6, pergunta 4)

---

## 4. Ordem de execução sugerida

1. `sefaz.py`: acrescentar `Nat. Operação` (e o que faltar) em `COLUNAS_NFE`/
   `COLUNAS_NFS`, por letra + nome-trava. Rodar e conferir que a trava não acusa.
2. `sc7.py` **(módulo novo)**: ler SC7 por posição → dict por PC já **agregado**;
   ler SC1 → `{PC: [solicitantes]}`. Expor `carregar()` no formato
   `(cabecalho, linhas, resumo)`, igual aos outros.
3. `pedidos_sefaz.py` **(módulo novo)**: extrair PC/OC do texto (regex + veto +
   validação no SC7), montar a aba nova. Prefixo de uuid novo (ex.: `PC:`).
4. `gerar_painel.py`: nova entrada em `ABAS` + `COMPACTA_*` + `ROTULOS_PLANILHA_*`;
   ajustar `COMPACTA_SEFAZ_SF1` (entram 7 colunas, saem 3).
5. `painel_modelo.html`: arraste do mouse; filtros das colunas novas.
6. Testar, **rodar `ATUALIZAR_PAINEL.cmd /sembase`**, conferir no ar.

**Testes que valem a pena** (é o que já pegou erro neste painel):
- Todo `<script>` do `PUBLICAR/index.html` tem de **compilar** — script quebrado
  deixa a aba em branco e nada avisa.
- Contagem de notas usa **`r.c.chave_nf_e`** (não `copia.chave_nf_e`, que não
  existe e cai no uuid, por linha).
- Regra com data precisa de teste de **data controlada**: na base de hoje várias
  condições não têm nenhum caso, e a regra iria para produção sem nunca disparar.

---

## 5. Agregação: o ponto mais delicado

`SC7` é por item. Ao trazer "o pedido" para uma linha de nota:

| Campo | Sugestão |
| --- | --- |
| `Dt. Entrega`, `1° Venc` | a **menor** data do PC (a que chega antes) |
| `Comprador`, `Controle Ap.`, `Ped. Encerr.` | 1º valor não vazio do PC |
| `Qtd.a Classi` | **soma** dos itens |
| `Numero da SC` | distintos, juntos por `·` (um PC pode nascer de várias SCs) |
| `Vlr.Total` | soma (serve para confrontar com o valor da nota) |
| `Solicitante` | distintos da SC1, juntos por `·` |

⚠ Se o usuário preferir **uma linha por item do pedido**, a aba muda de natureza
(deixa de ser "uma linha por nota") e o desenho é outro. **Perguntar** (§6, p. 4).

---

## 5-B. PEDIDO NOVO (13/08, fazer ANTES do resto): Tipo de Pagamento na aba Títulos Associados

Três coisas: tirar a coluna `#`, trazer o **`Tipo Pgto` da SE2**, e ter ao lado um
campo para o **tipo real** com um **alerta quando divergir**.

**MEDIDO na `SE2 - POSIÇÃO DIARIA.xlsx`** (259 colunas): `Tipo Pgto` está em
**IK** (índice 244) e o nome é **único** — dá para ler por posição com o nome
como trava, o mais seguro. Valores reais (amostra de 4.000 títulos):

| Valor | Qtd |
| --- | --- |
| NÃO DEFINIDO | 1.333 |
| BOLETO | 1.299 |
| TRANSFERENCIA | 577 |
| BOLETO S/C | 250 |
| IMPOSTO | 182 |
| CONVENIO | 153 |
| AMARRAR ADIANTAMENTO | 47 |
| (vazio) | 135 |

⚠ **O campo do tipo real tem de ser LISTA FECHADA com esses valores, não texto
livre.** Com digitação livre, "TRANSFERÊNCIA" vs "TRANSFERENCIA" viraria alerta
falso e o recurso perderia sentido. Mesmo desenho do dropdown de classificação do
painel FLUXO. São ~8 valores, então também caberia como grupo de `pills`.

**O que muda em cada arquivo:**
1. `verificacao_se2.py` — hoje lê 4 colunas **por nome** (`_indices()`). Acrescentar
   `Tipo Pgto`, de preferência **por posição IK conferindo o nome**. O módulo já
   devolve um veredito por UUID; passa a devolver também o tipo.
2. `gerar_painel.py` — a coluna entra em `COMPACTA` (aba `associacoes`) vinda do
   veredito da SE2, como já acontece com o selo `BOLETO`/`BAIXADO`.
3. `painel_modelo.html` — o campo do tipo real e o alerta são **colunas virtuais**,
   calculadas no navegador. Há precedente exato: `@delta_valor`, `@delta_dias` e
   `@tratativa` já são assim. Sugestão: `@tipo_real` (o dropdown) e `@tipo_diverge`
   (o alerta, aceso só quando `tipo_real` existe **e** difere do da SE2).
   ⚠ Só alertar quando o campo estiver preenchido — vazio é "ninguém conferiu
   ainda", não divergência.
4. Coluna `#`: é `<th class="canto">#</th>` no `desenharConferencia()`, com o `<td>`
   por linha. É **global** (todas as abas), não da aba. Tirar de lá tira de todas —
   se tiver de sair só nas Títulos Associados, precisa virar opção do `ABAS`.

### Os 5 pontos da coluna `#` (já localizados, no `painel_modelo.html`)

Ela **não** é uma coluna do `COMPACTA`: é a "gutter" desenhada à mão em
`desenharConferencia()`. Sair tem de sair dos cinco lugares, senão a grade
desalinha:

| Onde | O que está lá |
| --- | --- |
| `<colgroup>` (perto da linha 1366) | o `<col>` da gutter — **se ficar, todas as larguras deslocam uma coluna** |
| linha ~1371 | `<tr class="grupos"><th class="canto"></th>` |
| linha ~1376 | `<tr class="titulos"><th class="canto">#</th>` |
| linha ~1385 | `<td class="gutter">${i + 1}</td>` |
| linha ~1386 | `colspan="${cols.length + 1}"` do "Nenhuma linha encontrada" → passa a `cols.length` |
| CSS linhas 311-312 + as 5 regras `body[data-aba=…] th.canto` (104, 117, 134, 151, 168) | a cor do canto por aba |

⚠ `th.canto` é `position:sticky; left:0` — é ele que segura a primeira coluna
parada na rolagem horizontal. Sem ele, a coluna que ficar em primeiro lugar
**perde a fixação**; se isso for indesejado, a classe sticky tem de migrar para a
primeira coluna real.

### ⛔ O bloqueio: onde gravar o tipo real

O tipo real é **dado digitado**, então precisa de lugar para morar. Hoje as
marcações vivem em `analise_boletos_marcacoes` (Supabase), **uma linha por
título**, com `check` e `tratativa` — gravadas por `upsert … onConflict:"uuid"`
no `PUBLICAR/supa.js`. **Não existe coluna para o tipo real**, e criá-la é DDL no
banco.

Duas saídas, e a escolha é do usuário:

- **(A) Compartilhado (recomendado)** — criar a coluna e gravar junto da
  marcação, do mesmo jeito que a tratativa (inclusive entrando na fila de
  reenvio, que é o que garante o "não perder nada"):

  ```sql
  alter table public.analise_boletos_marcacoes
    add column if not exists tipo_pgto_real text;
  ```

  Depois: incluir `tipo_pgto_real` no `select` do `lerTudo(...)` e no objeto do
  `upsert` (`supa.js`), e no `backup_supabase.py`. A policy de escrita já existe
  para a linha; conferir se ela não lista colunas.
- **(B) Só neste navegador** — `localStorage`, como o `painel_boletos_feitos`.
  Sai hoje, sem tocar no banco, mas **cada pessoa vê o seu** — e como quem usa o
  painel é a assistente, isso provavelmente não serve.

⚠ Não dá para reaproveitar o campo `tratativa`: misturar o tipo real com a nota
escrita corromperia as duas informações.

---

## 6. Perguntas em aberto (responder antes de codar)

1. **"Ajuste o script para quando atualizar não perder essas modificações"** —
   não ficou claro o que se perde hoje. É (a) as colunas que você pinta de verde
   na SEFAZ.xlsx, que o script precisa passar a ler; (b) as marcações
   (check/tratativa) do painel; ou (c) outra coisa? Nada no painel apaga
   marcação hoje — elas moram no Supabase, por UUID.
2. **A aba nova SUBSTITUI a `SEFAZ x SF1` ou as duas convivem?** O pedido tem as
   duas coisas: ajustar a SEFAZ x SF1 **e** criar a aba nova com um superconjunto
   das colunas.
3. **Cobertura do SC7:** 48 PCs citados nas notas não existem no SC7 — e entre
   eles há uma sequência (`003251, 003253, 003258, 003261, 003266, 003273`), o
   que tem cara de **exportação incompleta** (filial ou período de fora), não de
   erro de leitura. O SC7.xlsx exportado cobre **todas as filiais e todo o
   período**? Se não, o painel vai marcar `PC NÃO ENCONTRADO` em nota boa.
4. **Agregação (§5):** uma linha por **nota** (pedido resumido) ou uma linha por
   **item do pedido**? E `Solicitante`/`Numero da SC` múltiplos: mostrar todos ou
   o primeiro?
5. **`Nome Dest`:** nome + CNPJ na mesma coluna (com o CNPJ copiável ao lado,
   como o `Fornecedor`), ou duas colunas?
6. Peso: a aba nova soma ~600 linhas à carteira, que já está em 4,0 MB. Segue
   valendo carregar tudo de uma vez, ou é hora de carregar **por aba, sob
   demanda**?
