# salvar_cotacao.py
import os
import pandas as pd
from teste import extrair_cotacao

CAMINHO_ARQUIVO = "cotacoes_diarias.xlsx"

df = extrair_cotacao()

if df.empty:
    print("Nenhuma cotação extraída.")
    exit()

df = df.dropna()
df["Data"] = df["Data"].astype(str)

if os.path.exists(CAMINHO_ARQUIVO):
    df_existente = pd.read_excel(CAMINHO_ARQUIVO)
    df_existente["Data"] = df_existente["Data"].astype(str)

    data_hoje = df.iloc[0]["Data"]

    if data_hoje not in df_existente["Data"].values:
        df_final = pd.concat([df_existente, df], ignore_index=True)
        df_final.to_excel(CAMINHO_ARQUIVO, index=False)
        print(f"Cotação do dia {data_hoje} adicionada com sucesso!")
    else:
        print(f"Cotação do dia {data_hoje} já está no arquivo.")
else:
    df.to_excel(CAMINHO_ARQUIVO, index=False)
    print("Arquivo criado com a cotação do dia.")
