# Controle Financeiro

Importa faturas Banco do Brasil em PDF para a planilha de controle financeiro e preserva os resumos e gráficos já existentes.

## Instalação

```bash
pip install -r requirements.txt
```

## Arquivos

- `config.json`: caminhos da pasta de faturas, planilha e arquivos de categorias.
  Sugestão: deixe a planilha e as faturas em uma pasta sincronizada do Google Drive.
- `categories.json`: regras fixas de categorização.
  Sugestão: use para regras conhecidas, como estabelecimentos que sempre caem na mesma categoria.
- `learned_categories.json`: categorias aprendidas a partir de correções manuais.
  Explicação: esse arquivo é atualizado pelo script ao longo do uso.

## Configuração inicial

Use como referência a planilha padrão `Controle Financeiro.xlsm`.

1. Copie `config.example.json` para `config.json`.
2. Copie `categories.example.json` para `categories.json`.

## Uso mensal

1. Coloque a fatura PDF na pasta do Google Drive configurada em `invoice_dir`.
2. Abra o terminal na pasta do projeto.
3. Rode:

```bash
python finance_automation.py process-drive --config config.json
```

4. Aguarde a sincronização do Google Drive e visualize a planilha.

## Correção de categorias

Se uma categoria vier errada:

1. Corrija a célula de categoria na planilha.
2. Salve a planilha.
3. Na próxima execução, o script reaproveita essa correção.
