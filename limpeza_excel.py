import pandas as pd
from datetime import datetime


df = pd.read_excel("input.xlsx", engine="openpyxl", dtype=str)


df = df.dropna(how="all")


df.columns = df.columns.str.strip().str.lower()


if "nome" in df.columns:
    df["nome"] = df["nome"].str.strip().str.lower().str.title()

df = df.fillna("Não informado")


data = datetime.now().strftime("%Y-%m-%d")
output_file = f"output_{data}.xlsx"


df.to_excel(output_file, index=False)


print("RELATÓRIO")
print("---------")
print(f"Linhas finais: {len(df)}")
print(f"Colunas: {len(df.columns)}")

print(f"Arquivo gerado: {output_file}")
