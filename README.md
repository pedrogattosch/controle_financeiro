# Automação de Faturas para Controle Financeiro

Este projeto importa uma fatura em PDF, identifica as compras pela data real da transação e lança os gastos nas abas mensais da sua planilha `.xlsm`.

## O que esta primeira versão faz

- lê PDFs de fatura do Banco do Brasil em formato texto;
- extrai `data`, `descrição` e `valor` de cada compra;
- classifica a compra em uma categoria da sua aba `Categorias`;
- decide se o gasto vai para `Gastos essenciais` ou `Gastos não essenciais`;
- grava na aba correta do mês da compra;
- cria uma aba oculta `_Importacoes` para evitar importações duplicadas;
- carrega as regras de categorização a partir de `categories.json`.

## Instalação

```bash
pip install -r requirements.txt
```

## Categorias

As regras de categorização ficam em [categories.json](C:/Users/pedro/dev/controle_financeiro/categories.json).

- `heading_to_category`: fallback com base na categoria que o banco já trouxe na fatura;
- `merchant_rules`: regras por estabelecimento, avaliadas de cima para baixo.

Quanto mais específica a regra, mais no topo ela deve ficar.

## Teste manual

Use primeiro em modo de teste:

```bash
python finance_automation.py import-pdf --pdf "C:/caminho/fatura.pdf" --workbook "C:/caminho/Controle financeiro (2026).xlsm" --dry-run
```

Para gravar em uma cópia:

```bash
python finance_automation.py import-pdf --pdf "C:/caminho/fatura.pdf" --workbook "C:/caminho/Controle financeiro (2026).xlsm" --output-workbook "C:/caminho/Controle financeiro (2026) - teste.xlsm"
```

## Monitoramento da pasta do Drive

1. Copie `config.example.json` para `config.json`.
2. Ajuste os caminhos da sua pasta do Drive, da planilha e do `categories.json`.
3. Rode:

```bash
python finance_automation.py watch --config config.json
```

Quando um novo PDF aparecer na pasta monitorada:

- a fatura será processada;
- os lançamentos serão escritos na planilha;
- o PDF será movido para a pasta `processadas`.

## Observações

- Esta versão foi preparada para o layout da sua fatura atual do BB.
- A categorização usa regras por estabelecimento e um fallback pela categoria da própria fatura.
- Se aparecer um estabelecimento novo, basta adicionar uma nova regra em `categories.json`.
