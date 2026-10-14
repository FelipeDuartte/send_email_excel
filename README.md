# 📊 Automação de Relatórios de Vendas com Python

Projeto de estudo desenvolvido para automatizar a análise de dados de vendas e o envio de relatórios por e-mail utilizando Python.

A aplicação lê dados de uma planilha Excel, processa as informações por loja, calcula indicadores de desempenho e envia um relatório formatado em HTML pelo Microsoft Outlook no macOS.

## 🚀 Funcionalidades

- 📄 Leitura de dados de planilhas Excel.
- 📈 Cálculo do faturamento total por loja.
- 📦 Soma da quantidade de produtos vendidos por loja.
- 💰 Cálculo do ticket médio por loja.
- 🎨 Geração de tabelas HTML para apresentação dos resultados.
- 📧 Envio automatizado de relatórios por e-mail através do Microsoft Outlook.

## 🛠️ Tecnologias utilizadas

- **Python** — linguagem principal.
- **Pandas** — manipulação e análise de dados.
- **OpenPyXL** — leitura de arquivos Excel pelo Pandas.
- **Appscript** — integração com o Microsoft Outlook no macOS.
- **HTML** — formatação das tabelas no corpo do e-mail.

## 📂 Estrutura do projeto

```text
send_email_excel/
├── send_email.py
├── Vendas.xlsx
└── README.md
```

## ⚙️ Como executar

### 1. Clone o repositório

```bash
git clone URL_DO_SEU_REPOSITORIO
cd send_email_excel
```

### 2. Crie um ambiente virtual

```bash
python3 -m venv venv
```

Ative o ambiente virtual:

**macOS/Linux**
```bash
source venv/bin/activate
```

### 3. Instale as dependências

```bash
pip install pandas openpyxl appscript
```

### 4. Configure a planilha

Coloque o arquivo `Vendas.xlsx` na pasta do projeto.

A planilha deve conter as colunas necessárias para os cálculos:

- `ID Loja`
- `Valor Final`
- `Quantidade`

### 5. Configure o destinatário

No arquivo `send_email.py`, altere o endereço de e-mail para o destinatário desejado.

### 6. Execute o projeto

```bash
python3 send_email.py
```

O Microsoft Outlook deverá estar instalado e configurado no macOS, com as permissões necessárias para a automação via Apple Events.

## 📊 Indicadores calculados

| Indicador | Descrição |
|---|---|
| Faturamento | Soma dos valores finais das vendas por loja. |
| Quantidade vendida | Total de unidades vendidas por loja. |
| Ticket médio | Faturamento dividido pela quantidade de unidades vendidas. |

## 🎯 Objetivo do projeto

Praticar conceitos de automação de processos com Python, análise de dados, manipulação de arquivos Excel, geração de relatórios HTML e integração com aplicações externas para reduzir tarefas manuais e aumentar a produtividade.

## 👨‍💻 Autor

**Felipe Duarte**

Estudante de Análise e Desenvolvimento de Sistemas, com foco em Python, automação de processos (RPA) e desenvolvimento de software.
