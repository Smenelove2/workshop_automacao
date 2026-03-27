from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

PASTA_DADOS = Path("dados")
PASTA_SAIDA = Path("saida")

ARQUIVO_ENTRADA = PASTA_DADOS / "base_simples_workshop.xlsx"
ARQUIVO_SAIDA = PASTA_SAIDA / "relatorio_simples.xlsx"


def ler_planilha():
    df = pd.read_excel(ARQUIVO_ENTRADA, sheet_name="Vendas")
    return df


def processar_dados(df):
    df["Total"] = df["Quantidade"] * df["Preço Unitário"]
    return df


def criar_resumos(df):
    resumo_vendedor = df.groupby("Vendedor")["Total"].sum().reset_index()
    resumo_produto = df.groupby("Produto")["Total"].sum().reset_index()
    return resumo_vendedor, resumo_produto


def salvar_relatorio(df, resumo_vendedor, resumo_produto):
    PASTA_SAIDA.mkdir(exist_ok=True)

    with pd.ExcelWriter(ARQUIVO_SAIDA, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Base_Com_Total", index=False)
        resumo_vendedor.to_excel(writer, sheet_name="Resumo_Vendedor", index=False)
        resumo_produto.to_excel(writer, sheet_name="Resumo_Produto", index=False)

    formatar_moeda()


def formatar_moeda():
    wb = load_workbook(ARQUIVO_SAIDA)

    # Aba Base_Com_Total
    ws_base = wb["Base_Com_Total"]
    for cell in ws_base["F"][1:]:  # coluna Total
        cell.number_format = 'R$ #,##0.00'

    for cell in ws_base["E"][1:]:  # coluna Preço Unitário
        cell.number_format = 'R$ #,##0.00'

    # Aba Resumo_Vendedor
    ws_vendedor = wb["Resumo_Vendedor"]
    for cell in ws_vendedor["B"][1:]:  # coluna Total
        cell.number_format = 'R$ #,##0.00'

    # Aba Resumo_Produto
    ws_produto = wb["Resumo_Produto"]
    for cell in ws_produto["B"][1:]:  # coluna Total
        cell.number_format = 'R$ #,##0.00'

    wb.save(ARQUIVO_SAIDA)


print("Lendo a planilha...")
df = ler_planilha()

print("Processando os dados...")
df = processar_dados(df)

print("Criando os resumos...")
resumo_vendedor, resumo_produto = criar_resumos(df)

print("Salvando o relatório...")
salvar_relatorio(df, resumo_vendedor, resumo_produto)

print("Concluído! Arquivo gerado em:", ARQUIVO_SAIDA)
