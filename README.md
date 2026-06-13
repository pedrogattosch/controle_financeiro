# Controle Financeiro

Importa faturas Banco do Brasil em PDF para a planilha de controle financeiro e preserva os resumos e gráficos já existentes.

## Instalação

```bash
pip install -r requirements.txt
```

## Arquivos

- `config.json`: caminhos da pasta de faturas, planilha e arquivos de categorias.
- `categories.json`: regras fixas de categorização.
- `learned_categories.json`: categorias aprendidas a partir de correções manuais.

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
