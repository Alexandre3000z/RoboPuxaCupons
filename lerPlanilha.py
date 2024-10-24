import pandas as pd
from fuzzywuzzy import fuzz

# Carregando a planilha
df = pd.read_excel('PLANILHA QSA.xlsx')

def nomeEmpresas():
    empresas_series = df['Empresa']
    empresas_array = empresas_series.to_numpy()
    return empresas_array


def puxarCnpj():
    metodo_series = df['CNPJ']
    metodo_array = metodo_series.to_numpy()
    return metodo_array


