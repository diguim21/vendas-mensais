import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
from openpyxl.styles import Font, Alignment, PatternFill

# Função para validar entrada numérica
def validar_entrada(mensagem):
    while True:
        entrada = input(mensagem)
        # Substituir vírgulas por pontos e remover múltiplos pontos
        entrada = entrada.replace(',', '.')
        if entrada.count('.') > 1:
            print("Formato inválido. Por favor, insira um número válido.")
            continue
        try:
            return float(entrada)
        except ValueError:
            print("Entrada inválida. Digite um número válido (ex: 16046.78).")

# Função principal
def registrar_vendas_mes():
    """Registra as vendas diárias de um mês, gera uma planilha e exibe um gráfico."""
    # Pedir o mês e o ano
    ano = int(input("Digite o ano (ex: 2025): "))
    mes = int(input("Digite o mês (1-12): "))

    # Determinar o número de dias no mês
    try:
        dias_no_mes = (datetime(ano, mes + 1, 1) - datetime(ano, mes, 1)).days
    except ValueError:
        dias_no_mes = 31 if mes == 12 else (datetime(ano + 1, 1, 1) - datetime(ano, mes, 1)).days

    # Coletar as vendas diárias
    vendas = []
    for dia in range(1, dias_no_mes + 1):
        valor = validar_entrada(f"Digite o valor vendido no dia {dia}/{mes}/{ano} (R$): ")
        vendas.append({"Data": f"{ano}-{mes:02d}-{dia:02d}", "Valor": valor})

    # Criar DataFrame e ordenar
    df = pd.DataFrame(vendas)
    df = df[["Data", "Valor"]]  # Garantir ordem das colunas
    total_vendas = df["Valor"].sum()

    # Salvar a planilha com valores formatados
    nome_arquivo = f"vendas_{ano}_{mes:02d}.xlsx"
    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=f"Vendas_{mes:02d}_{ano}")
        workbook = writer.book
        sheet = writer.sheets[f"Vendas_{mes:02d}_{ano}"]

        # Título da planilha
        sheet.merge_cells("A1:B1")
        title_cell = sheet["A1"]
        title_cell.value = f"Vendas do Mês - {mes:02d}/{ano}"
        title_cell.font = Font(size=14, bold=True)
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        title_cell.fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")

        # Formatar os valores na coluna B
        for cell in sheet["B"]:
            if cell.value and isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"

        # Adicionar linha do total
        total_row = len(df) + 2
        sheet[f"A{total_row}"] = "VALOR TOTAL MÊS"
        sheet[f"A{total_row}"].font = Font(bold=True)
        sheet[f"A{total_row}"].alignment = Alignment(horizontal="right")
        sheet[f"B{total_row}"] = total_vendas
        sheet[f"B{total_row}"].font = Font(color="0000FF", bold=True)
        sheet[f"B{total_row}"].number_format = "#,##0.00"

    # Exibir resultados
    print(f"\nPlanilha '{nome_arquivo}' gerada com sucesso!")
    print(f"Total de vendas no mês: R$ {total_vendas:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    # Gerar gráfico
    plt.figure(figsize=(10, 6))
    plt.plot(df["Data"], df["Valor"], marker='o', color='b', label="Vendas Diárias")
    plt.axhline(y=0, color='r', linestyle='--', label="Comércio fechado")
    plt.xticks(rotation=45)
    plt.title(f"Vendas Diárias - {mes:02d}/{ano}", fontsize=16)
    plt.xlabel("Data", fontsize=12)
    plt.ylabel("Valor (R$)", fontsize=12)
    plt.legend()
    plt.tight_layout()
    plt.show()

# Fluxo principal
if __name__ == "__main__":
    registrar_vendas_mes()
