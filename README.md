# Controle Financeiro

Este projeto importa uma fatura em PDF, identifica as compras pela data real da transação e lança os gastos nas abas mensais da sua planilha `.xlsx` ou `.xlsm`.

## O que esta primeira versão faz?

- lê PDFs de fatura do Banco do Brasil em formato texto;
- extrai `data`, `descrição` e `valor` de cada compra;
- classifica a compra em uma categoria da sua aba `Categorias`;
- decide se o gasto vai para `Gastos essenciais` ou `Gastos não essenciais`;
- grava na aba correta do mês da compra;
- cria uma aba oculta `_Importacoes` para evitar importações duplicadas;
- carrega as regras de categorização a partir de `categories.json`;
- reaproveita categorias aprendidas em `learned_categories.json`;
- aprende correções manuais feitas por você em lançamentos já importados;
- preserva os resumos e gráficos que já existirem na lateral da sua planilha.

## Instalação

```bash
pip install -r requirements.txt
```

## Categorias

As regras de categorização ficam em [categories.json](C:/Users/pedro/dev/controle_financeiro/categories.json).

- `heading_to_category`: fallback com base na categoria que o banco já trouxe na fatura;
- `merchant_rules`: regras por estabelecimento, avaliadas de cima para baixo.

Quanto mais específica a regra, mais no topo ela deve ficar.

## Aprendizado

As categorias aprendidas ficam em [learned_categories.json](C:/Users/pedro/dev/controle_financeiro/learned_categories.json).

Fluxo:

- primeiro o script tenta uma categoria já aprendida para a mesma descrição;
- se não encontrar, tenta as regras do `categories.json`;
- se você corrigir manualmente a célula de categoria de um lançamento importado, essa correção passa a valer para as próximas vezes.

## Resumos e gráficos da planilha

Se a sua planilha já tiver fórmulas, totais por categoria e gráficos na lateral, o script não sobrescreve essa estrutura.

Ele apenas preenche os lançamentos nas colunas principais, deixando a própria planilha recalcular:

- totais por categoria;
- gráficos de pizza;
- resumos mensais.

## Processo mensal

Este é o fluxo recomendado para o seu uso mensal:

1. Adicione a nova fatura em PDF em uma pasta do seu Google Drive.
2. Aguarde o Google Drive sincronizar esse arquivo no seu computador.
3. Abra o terminal na pasta do projeto.
4. Rode o comando que usa os caminhos já configurados no `config.json`.
5. Aguarde o script atualizar a planilha.
6. Aguarde o Google Drive sincronizar a planilha atualizada.
7. Abra a planilha pelo Google Drive no celular ou no navegador para visualizar os totais por categoria e os gráficos.

Comando base:

```bash
python finance_automation.py process-drive --config config.json
```

Se quiser conferir antes de gravar:

```bash
python finance_automation.py process-drive --config config.json --dry-run
```

O `config.json` deve apontar para:

- a pasta do Drive onde você coloca as faturas novas;
- a pasta `processadas`;
- a planilha financeira;
- `categories.json`;
- `learned_categories.json`.

Se alguma categoria vier errada:

1. Corrija manualmente a célula de categoria na planilha.
2. Salve a planilha.
3. Na próxima execução, o script reaproveita essa correção automaticamente.

Observações importantes:

- a planilha e a fatura precisam estar em uma pasta do Google Drive sincronizada localmente no computador;
- o script atualiza o arquivo local sincronizado, e o Google Drive sobe essa alteração para a nuvem;
- o comando `process-drive` processa todos os PDFs pendentes da pasta configurada em `invoice_dir` e move os processados para `processed_dir`;
- a visualização no celular acontece depois que a sincronização do Drive terminar;
- a correção manual de categoria funciona melhor quando você altera a célula de categoria de um lançamento que foi importado pelo próprio script.

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
2. Ajuste os caminhos da sua pasta do Drive, da planilha, do `categories.json` e do `learned_categories.json`.
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
- Se você corrigir manualmente a categoria de um lançamento importado, o script aprende essa decisão automaticamente.
