# Workshop de Automação com Python

Projeto base para uma atividade prática de automação de tarefas com Python.

A proposta deste projeto é simples:

- ler uma planilha Excel
- calcular o valor total de cada venda
- criar resumos por vendedor e por produto
- salvar um novo arquivo Excel com o relatório final

## Estrutura do projeto

```text
projeto/
├── dados/
│   └── base_simples_workshop.xlsx
├── saida/
├── main.py
├── requirements.txt
├── README.md
└── .gitignore
```

## Objetivo da atividade

O objetivo é fazer os alunos praticarem:

- leitura de planilhas com `pandas`
- manipulação de colunas
- criação de novas informações a partir dos dados
- agrupamento de dados com `groupby`
- geração de relatório em Excel

## Arquivo de entrada

O arquivo de entrada deve estar dentro da pasta `dados` com o nome:

```text
base_simples_workshop.xlsx
```

A aba usada no script é:

```text
Vendas
```

## Colunas esperadas na planilha

A planilha deve conter estas colunas:

- `Data`
- `Vendedor`
- `Produto`
- `Quantidade`
- `Preço Unitário`

## O que o script faz

O script:

1. lê a planilha da pasta `dados`
2. cria uma nova coluna chamada `Total`
3. gera um resumo por vendedor
4. gera um resumo por produto
5. salva um novo arquivo Excel na pasta `saida`

## Arquivo de saída

O relatório gerado será salvo em:

```text
saida/relatorio_simples.xlsx
```

Esse arquivo terá 3 abas:

- `Base_Com_Total`
- `Resumo_Vendedor`
- `Resumo_Produto`

## Instalação

Crie e ative seu ambiente virtual, depois instale as dependências:

```bash
pip install -r requirements.txt
```

## Como executar

Rode o script com:

```bash
python main.py
```

## Bibliotecas usadas

- `pandas`
- `openpyxl`

## Sugestões de exercícios para os alunos

Algumas ideias de atividades para trabalhar em sala:

- alterar o nome das abas do relatório
- adicionar um resumo por data
- ordenar os resumos do maior para o menor valor
- formatar cabeçalhos em negrito
- adicionar novas colunas na planilha
- adaptar o projeto para outro contexto, como estoque ou notas

## Observação

Este projeto foi mantido propositalmente simples para facilitar o entendimento de alunos com pouca experiência em programação.
