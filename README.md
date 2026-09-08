# 📊 Automação de Relatórios de Vendas com Python

Projeto desenvolvido em Python para automatizar a análise de dados de vendas e a geração de relatórios profissionais em Excel.

## 🚀 Funcionalidades

- Leitura automática de dados de uma planilha Excel
- Cálculo do faturamento total
- Cálculo da quantidade de vendas
- Cálculo do ticket médio
- Geração automática de relatório em Excel
- Formatação de valores em R$
- Cabeçalho e título estilizados
- Ajuste da largura das colunas
- Registro automático da data e hora de geração

## 🛠️ Tecnologias utilizadas

- Python
- Pandas
- OpenPyXL
- Microsoft Excel

## 📁 Estrutura do projeto

```text
automacao-relatorios-excel/
│
├── data/
│   └── vendas.xlsx
│
├── reports/
│   └── relatorio_vendas.xlsx
│
├── src/
│   └── main.py
│
└── README.md
```

## ⚙️ Como executar

Instale as dependências:

```bash
python -m pip install pandas openpyxl
```

Execute o programa:

```bash
python src/main.py
```

O relatório será gerado automaticamente em:

```text
reports/relatorio_vendas.xlsx
```

## 📊 Indicadores gerados

O programa calcula automaticamente:

- Faturamento total
- Quantidade de vendas
- Ticket médio

## 🎯 Objetivo

Este projeto foi criado para praticar automação de tarefas com Python e demonstrar como processos manuais realizados em planilhas podem ser automatizados.

O projeto pode ser adaptado para diferentes tipos de relatórios e necessidades empresariais.