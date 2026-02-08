import pandas as pd
from datetime import datetime

# Ler o Excel
df = pd.read_excel("input.xlsx", engine="openpyxl", dtype=str)

# Remover linhas vazias
df = df.dropna(how="all")

# Padronizar nomes das colunas
df.columns = df.columns.str.strip().str.lower()

# Padronizar coluna nome
if "nome" in df.columns:
    df["nome"] = df["nome"].str.strip().str.lower().str.title()

# Preencher valores vazios
df = df.fillna("Não informado")

# Criar nome do arquivo com data
data = datetime.now().strftime("%Y-%m-%d")
output_file = f"output_{data}.xlsx"

# Salvar arquivo final
df.to_excel(output_file, index=False)

# Relatório simples
print("RELATÓRIO")
print("---------")
print(f"Linhas finais: {len(df)}")
print(f"Colunas: {len(df.columns)}")
print(f"Arquivo gerado: {output_file}")