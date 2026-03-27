from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

PASTA_DADOS = Path("dados")
PASTA_SAIDA = Path("saida")

ARQUIVO_ENTRADA = PASTA_DADOS / "base_simples_workshop.xlsx"
ARQUIVO_SAIDA = PASTA_SAIDA / "relatorio_simples.xlsx"


