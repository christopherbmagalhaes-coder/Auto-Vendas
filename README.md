# 🚗 Auto Vendas — Automação de Relatórios com Python

Projeto desenvolvido em Python para automatizar a consolidação de dados de clientes, produtos e pedidos em planilhas Excel, com validação de dados, formatação profissional e geração de um dashboard para análise de vendas.

## 🎯 Objetivo

Simplificar a geração de relatórios de vendas, reduzindo tarefas manuais e facilitando a análise de informações como faturamento por produto, desempenho ao longo do tempo e quantidade de itens vendidos.

## Funcionalidades

- Leitura de planilhas Excel com dados de clientes, estoque e pedidos.
- Validação de dados e identificação de inconsistências.
- Cruzamento de informações por meio dos identificadores de clientes e produtos.
- Consolidação dos pedidos em uma única planilha de vendas.
- Geração automática do arquivo `Vendas.xlsx`.
- Formatação de células, colunas e valores monetários.
- Criação de um dashboard com quatro gráficos:
  - Faturamento por produto.
  - Participação de cada produto no faturamento.
  - Faturamento por data.
  - Quantidade vendida por produto.

## 🛠️ Tecnologias utilizadas

- **Python** — desenvolvimento da automação.
- **Pandas** — leitura, manipulação e processamento dos dados.
- **OpenPyXL** — criação e formatação de planilhas e gráficos no Excel.

## Estrutura do projeto

```text
Auto Vendas/
├── main.py
├── leitura.py
├── validacao.py
├── processamento.py
├── exportacao.py
├── formatacao.py
├── requirements.txt
├── Clientes.xlsx
├── Estoque.xlsx
├── Pedidos.xlsx
└── Vendas.xlsx
```

O projeto utiliza uma arquitetura modular, separando as responsabilidades de leitura, validação, processamento, exportação e formatação dos dados.

## Resultado

A automação gera o arquivo `Vendas.xlsx`, contendo duas planilhas:

- **Vendas:** reúne os dados consolidados dos pedidos.
- **Dashboard:** apresenta gráficos para facilitar a visualização e a análise dos resultados.

## Aprendizados e práticas aplicadas

- Modularização e separação de responsabilidades.
- Manipulação e validação de dados com Python.
- Leitura e escrita de arquivos Excel.
- Tratamento de inconsistências e erros.
- Automação de relatórios e criação de dashboards.

## 👨‍💻 Autor

**Christopher Bruno Magalhães**

[GitHub](https://github.com/christopherbmagalhaes-coder)
