# Data Analytics Script

## Descrição
Este repositório contém um script em Python para análise de dados mensais a partir de um arquivo Excel. O script permite ao usuário interagir com os dados, selecionando um mês específico para ordenar e visualizar as informações, além de fornecer estatísticas e gráficos de forma interativa. O usuário tem a opção de encerrar o programa a qualquer momento.

## Funcionalidades
- **Leitura de Arquivo Excel**: O script lê um arquivo Excel chamado `analyticsDataFacepass.xlsx` e carrega os dados em um DataFrame usando a biblioteca `pandas`.
- **Interatividade com o Usuário**: O usuário pode selecionar um mês para análise (por exemplo, `JAN`, `FEV`, `MAR`, etc.), e o script ordenará os dados com base no mês escolhido, exibirá estatísticas básicas ou gráficos conforme a escolha.
- **Verificação de Colunas**: O script verifica se o mês inserido pelo usuário corresponde a uma das colunas do DataFrame `excel`. Se não existir, informa o usuário.
- **Estatísticas Básicas**: O script pode calcular a média, mediana e soma dos valores de um mês específico.
- **Visualização Gráfica**: O usuário pode visualizar os dados do mês selecionado em forma de gráfico.
- **Opção de Saída**: O usuário pode optar por encerrar o programa a qualquer momento, digitando "sair" em qualquer etapa de interação.

## Pré-requisitos
Antes de executar o script, você precisará ter o Python instalado em sua máquina, bem como as bibliotecas necessárias:

- Python 3.x
- pandas
- matplotlib

Você pode instalar as dependências usando o seguinte comando:

```bash
pip install pandas matplotlib
