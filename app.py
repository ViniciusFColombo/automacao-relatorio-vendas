import pandas as pd

df = pd.read_excel("planilha_vendas_baguncada.xlsx")

print(df.isnull().sum()) 
print("ANTES:", len(df))
df = df.dropna(how="all") 
print("DEPOIS:", len(df))

print("ANTES:", len(df))
df = df.drop_duplicates()
print("DEPOIS:", len(df))

df = df[df["Quantidade"].notna()]
df = df[df["Preço"].notna()]

df = df[df["Quantidade"] > 0]
df = df[df["Preço"] > 0]

print("Linhas restantes: ", len(df))
print(df.head())

df = df.dropna(subset=["Produto", "Categoria"])

df["Vendedor"] = df["Vendedor"].fillna("Desconhecido")

df = df.dropna(subset=["Data"])

print(df.isnull().sum())

df["Total"] = df["Quantidade"] * df["Preço"]
print(df.head())

faturamento_total = df["Total"].sum()
print("Faturamento total: $", faturamento_total)

produto_mais_vendido = df.groupby("Produto")["Quantidade"].sum().sort_values(ascending=False)
print(produto_mais_vendido.to_string())

vendedor_top = df.groupby("Vendedor")["Total"].sum().sort_values(ascending=False)
print(vendedor_top.to_string())

with pd.ExcelWriter("Relatorio_Vendas_Limpo.xlsx", engine="openpyxl") as writer:


    df.to_excel(writer, index=False, sheet_name="Sheet1")
    worksheet = writer.sheets["Sheet1"]

    
    linha_inicial = len(df) + 3

    worksheet.cell(row=linha_inicial, column=1, value="Faturamento Total")
    worksheet.cell(row=linha_inicial, column=2, value=faturamento_total)

    produto_top_nome = produto_mais_vendido.index[0]
    produto_top_valor = produto_mais_vendido.iloc[0]
    worksheet.cell(row=linha_inicial + 2, column=1, value="Produto mais Vendido")
    worksheet.cell(row=linha_inicial + 2, column=2, value=produto_top_nome)
    worksheet.cell(row=linha_inicial + 2, column=3, value=produto_top_valor)

    vendedor_nome = vendedor_top.index[0]
    vendedor_valor = vendedor_top.iloc[0]
    worksheet.cell(row=linha_inicial + 4, column=1, value="Vendedor Top")
    worksheet.cell(row=linha_inicial + 4, column=2, value=vendedor_nome)
    worksheet.cell(row=linha_inicial + 4, column=3, value=vendedor_valor)

print("Relatório gerado com sucesso!")