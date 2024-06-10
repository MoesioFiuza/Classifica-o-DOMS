import pandas as pd

arquivo_original = r'C:\Users\moesios\Desktop\EXPANSÃO\DESLOCAMENTOS EM POA_.xlsx'
arquivo_saida = r'C:\Users\moesios\Desktop\EXPANSÃO\CLASSIFICAÇÃO DOMICILIOS.xlsx'

df_domicilio = pd.read_excel(arquivo_original, sheet_name='DOMICÍLIO')
df_morador = pd.read_excel(arquivo_original, sheet_name='MORADOR')

regras_domicilio = {
    'Banheiros': [0, 3, 7, 10, 14],
    'Empregados domésticos': [0, 3, 7, 10, 13],
    'Automóveis': [0, 3, 5, 8, 11],
    'Microocomputador': [0, 3, 6, 8, 11],
    'Lava louça': [0, 3, 6, 6, 6],
    'Geladeira': [0, 2, 3, 5, 5],
    'Freezer': [0, 2, 4, 6, 6],
    'Lava roupa': [0, 2, 4, 6, 6],
    'Microo-ondas': [0, 2, 4, 4, 4],
    'Motocicleta': [0, 1, 3, 3, 3],
    'Secadora Roupa': [0, 2, 2, 2, 2],
    'Água encanada': {'SIM': 4, 'NÃO': 0},
    'TRECHO DA RUA DO DOMICÍLIO TEM PAVIMENTAÇÃO?': {'SIM': 2, 'NÃO': 0}
}

regras_morador = {
    'Grau de Instrução': {
        'Analfabeto / Fundamental I incompleto': 0,
        'Fundamental I completo / Fundamental II incompleto': 1,
        'Fundamental II completo / Médio incompleto': 2,
        'Médio completo / Superior incompleto': 4,
        'Superior completo': 7
    }
}

def calcular_pontuacao_domicilio(row):
    responsavel = df_morador.loc[(df_morador['ID_DOMICILIO'] == row['ID_DOMICILIO']) & (df_morador['SITUAÇÃO DO MORADOR NO DOMICÍLIO'] == 'RESPONSÁVEL PELO DOMICÍLIO')]
    if not responsavel.empty:
        pontuacao = 0
        for coluna, pontos in regras_domicilio.items():
            if isinstance(pontos, list):
                try:
                    quantidade = int(row[coluna])
                except ValueError:
                    quantidade = 0
                if quantidade >= len(pontos):
                    quantidade = len(pontos) - 1
                pontuacao += pontos[quantidade]
            elif isinstance(pontos, dict):
                pontuacao += pontos.get(row[coluna], 0)
        return pontuacao
    else:
        return 0

def calcular_pontuacao_morador(row):
    return regras_morador['Grau de Instrução'].get(row['Grau de Instrução'], 0)

df_domicilio['Pontuacao'] = df_domicilio.apply(calcular_pontuacao_domicilio, axis=1)

df_morador['Pontuacao'] = df_morador.apply(calcular_pontuacao_morador, axis=1)
pontuacao_moradores = df_morador.groupby('ID_DOMICILIO')['Pontuacao'].sum()
df_domicilio = df_domicilio.join(pontuacao_moradores, on='ID_DOMICILIO', rsuffix='_moradores')

df_domicilio['Pontuacao_total'] = df_domicilio['Pontuacao'] + df_domicilio['Pontuacao_moradores']

def classificar_renda(pontuacao):
    if 45 <= pontuacao <= 100:
        return 'A'
    elif 38 <= pontuacao <= 44:
        return 'B1'
    elif 29 <= pontuacao <= 37:
        return 'B2'
    elif 23 <= pontuacao <= 28:
        return 'C1'
    elif 17 <= pontuacao <= 22:
        return 'C2'
    else:
        return 'D-E'

df_domicilio['Classificacao'] = df_domicilio['Pontuacao_total'].apply(classificar_renda)

df_domicilio.to_excel(arquivo_saida, index=False)

df_domicilio.head()
