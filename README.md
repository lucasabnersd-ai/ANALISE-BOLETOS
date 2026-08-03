# Painel Análise de Boletos

Versão visual da planilha `TRATAMENTO PYTHON BOLETOS.xlsx`. Hoje traz a aba
**Associações Encontradas** — os títulos do TOTVS que estavam *sem código de
boleto* e foram associados a um boleto do DDA.

O painel mostra, para cada título: o boleto correspondente (linha digitável),
os indicadores do cruzamento, os **critérios** que sustentam a associação e os
**alertas** de conferência (parcelas de fatura, NFs com várias parcelas).

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
