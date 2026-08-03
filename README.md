# Painel Análise de Boletos

Versão visual da planilha `TRATAMENTO PYTHON BOLETOS.xlsx`. Hoje traz a aba
**Associações Encontradas** — os títulos do TOTVS que estavam *sem código de
boleto* e foram associados a um boleto do DDA.

É a **mesma planilha**: as colunas da aba, na mesma ordem, com cabeçalho e
numeração de linha fixos e ordenação ao clicar no título da coluna. Nada de
modos nem de barra de filtros — só uma busca.

O painel acrescenta as quatro coisas que a planilha não faz:

- **Check por título** — a coluna `CHECK (FEITO)` vira caixa de seleção; a linha
  marcada fica riscada. A marcação é gravada **neste navegador**, pela chave
  UUID do título, e sobrevive a atualizações do painel. O registro definitivo
  continua sendo a coluna `CHECK (FEITO)` na planilha — o que estiver preenchido
  lá entra marcado.
- **Divergência destacada** — quando `Valor Boleto` ou `Vencimento Boleto`
  diferem do título, a célula fica em âmbar com a diferença ao lado
  (`− R$ 0,85`, `+6 dias`).
- **Copiar** — o UUID (para buscar no SE2) e a linha digitável só com os 47
  dígitos, sem a pontuação de leitura.
- **Abas da planilha** no rodapé, mostrando o que ainda falta levar.

Colunas 100% vazias nas linhas associadas não são exibidas (exceto o check).

## ⚠️ Este repositório é privado e deve continuar assim

O `index.html` contém **linhas digitáveis, CNPJs, valores e fornecedores**.

- **Não tornar o repositório público.**
- **Não ativar o GitHub Pages.** O Pages serve o site publicamente *mesmo com o
  repositório privado* (controle de acesso só existe no plano Enterprise) —
  qualquer pessoa com a URL baixaria os dados sem login.
- Se um dia o painel precisar de URL para outras pessoas, o caminho é o mesmo
  dos outros painéis: dados fora do HTML, num bucket privado do Supabase,
  carregados só depois do login.

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
