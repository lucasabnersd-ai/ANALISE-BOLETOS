# Painel Análise de Boletos

Versão visual da planilha `TRATAMENTO PYTHON BOLETOS.xlsx`. Hoje traz a aba
**Associações Encontradas** — os títulos do TOTVS que estavam *sem código de
boleto* e foram associados a um boleto do DDA.

O painel tem **uma única visão, a Conferência**: um recorte das colunas da aba,
na mesma ordem da planilha, com título e boleto lado a lado e a diferença entre
os dois logo depois de cada par. Cabeçalho e numeração de linha ficam fixos e a
ordenação é por clique no título da coluna. Nada de barra de filtros — só uma
busca.

O painel acrescenta o que a planilha não faz:

- **Check por título** — a coluna `CHECK (FEITO)` vira caixa de seleção; a linha
  marcada fica riscada. A marcação é gravada **neste navegador**, pela chave
  UUID do título, e sobrevive a atualizações do painel. O registro definitivo
  continua sendo a coluna `CHECK (FEITO)` na planilha — o que estiver preenchido
  lá entra marcado.
- **Diferença calculada** — as colunas `Δ valor` e `Δ dias` comparam o boleto
  com o título. Iguais mostram `ok`; diferentes saem em âmbar, em moeda
  (`−R$ 0,85`) ou em dias (`+6d`).
- **Status em selo** — verde para `MATCH FORTE`, âmbar para `MATCH PROVÁVEL`.
- **Alerta em selo roxo** — diz o tipo (`PARCELA` ou `FATURA`); o texto completo
  do alerta fica no tooltip.
- **Campos de copiar** — `UUID` (para buscar no SE2) e `Linha dig.` (os 47
  dígitos, sem a pontuação de leitura) são só botões, sem gastar largura com o
  valor. O código do **Fornecedor**, ao lado da razão social, também copia.
- **Abas da planilha** no rodapé, mostrando o que ainda falta levar.

## ⚠️ Os dados estão dentro do `index.html`

O `index.html` contém **linhas digitáveis, CNPJs, valores e fornecedores** em
texto puro, no próprio arquivo.

O GitHub Pages serve o site publicamente **mesmo com o repositório privado** —
controle de acesso só existe no plano Enterprise. Ou seja: **com o Pages ligado,
qualquer pessoa com a URL baixa esses dados sem login**, e a página pode ser
indexada por buscadores.

O Pages foi ligado em **04/08/2026, por decisão do usuário, ciente disso**. Para
voltar a fechar: desligar o Pages em *Settings › Pages* e manter o repositório
privado. Para ter URL sem expor os dados, o caminho é o dos outros painéis —
dados fora do HTML, no Supabase, carregados só depois do login.

Vale lembrar que desligar o Pages **não apaga o histórico**: os commits antigos
seguem com os dados para quem tiver acesso ao repositório.

## Como atualizar

Duplo clique em `ATUALIZAR_PAINEL.cmd`, ou:

```
python gerar_painel.py
```

O script lê a planilha (mesmo aberta no Excel, via cópia temporária) e reescreve
`PUBLICAR/index.html`. Para ver o painel, é só abrir esse arquivo no navegador.

## Arquivos

| Arquivo | O que é |
|---|---|
| `gerar_painel.py` | Lê a planilha e injeta os dados no modelo |
| `painel_modelo.html` | Layout do painel (marcador `/*__DADOS__*/null`) |
| `PUBLICAR/index.html` | Painel gerado — é o arquivo que se abre |
| `ATUALIZAR_PAINEL.cmd` | Atalho para rodar o gerador |

A planilha de origem **não** entra no repositório (ver `.gitignore`).

## Próximas abas

Ainda faltam levar: `Sem Match` (1.112), `Tratar Correspondências` (122),
`NFs Múltiplas Parcelas` e `Futuros NF Não Associada` (817).
