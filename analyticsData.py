import pandas as pd
import matplotlib.pyplot as plt

# Lê o arquivo Excel
excel = pd.read_excel("analyticsDataFacepass.xlsx")

# Função para exibir o menu principal
def show_menu():
    print("\nMenu de Análise:")
    print("1. Ordenar e exibir dados por mês")
    print("2. Calcular estatísticas básicas (Média, Mediana, Soma)")
    print("3. Visualizar dados em gráfico")
    print("4. Sair")

# Função para analisar o mês selecionado
def analyze_month(month):
    if month in excel.columns:
        analytics = excel.sort_values(by=month, ascending=False)
        print("\nDados ordenados por", month, ":")
        print(analytics.head())  # Mostra apenas os primeiros 5 registros
    else:
        print(f"O mês '{month}' não existe nas colunas do arquivo.")

# Função para calcular estatísticas básicas
def calculate_statistics(month):
    if month in excel.columns:
        print(f"\nEstatísticas para o mês de {month}:")
        print(f"Média: {excel[month].mean()}")
        print(f"Mediana: {excel[month].median()}")
        print(f"Soma: {excel[month].sum()}")
    else:
        print(f"O mês '{month}' não existe nas colunas do arquivo.")

# Função para visualizar os dados em gráfico
def plot_data(month):
    if month in excel.columns:
        plt.figure(figsize=(10, 6))
        excel[month].plot(kind='bar')
        plt.title(f"Dados do mês de {month}")
        plt.xlabel("Índice")
        plt.ylabel("Valor")
        plt.show()
    else:
        print(f"O mês '{month}' não existe nas colunas do arquivo.")

# Loop principal do programa
while True:
    show_menu()
    choice = input("\nEscolha uma opção (ou digite 'sair' para encerrar): ").strip().lower()

    if choice == '1':
        month = input("Qual mês você deseja analisar? (Use JAN, FEV, MAR, etc. ou digite 'sair' para encerrar)").upper()
        if month == 'SAIR':
            print("Obrigada! Saindo do programa.")
            break
        analyze_month(month)
    elif choice == '2':
        month = input("Para qual mês deseja calcular as estatísticas? (Use JAN, FEV, MAR, etc. ou digite 'sair' para encerrar)").upper()
        if month == 'SAIR':
            print("Obrigada! Saindo do programa.")
            break
        calculate_statistics(month)
    elif choice == '3':
        month = input("Para qual mês deseja visualizar os dados em gráfico? (Use JAN, FEV, MAR, etc. ou digite 'sair' para encerrar)").upper()
        if month == 'SAIR':
            print("Obrigada! Saindo do programa.")
            break
        plot_data(month)
    elif choice == '4' or choice == 'sair':
        print("Obrigada! Saindo do programa.")
        break
    else:
        print("Opção inválida! Por favor, escolha uma opção do menu.")
