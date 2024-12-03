import pandas as pd
import numpy as np

# Carrega o CSV, especificando o delimitador
# tabela_Movel = pd.read_csv('C:\\Users\\Math\\Documents\\fechamento\\movel.csv', delimiter=';', encoding='utf-8-sig')
tabela_Movel = pd.read_csv('/home/math/Documents/fechamento/vada.csv', delimiter=';', encoding='utf-8-sig')

# Substituir vírgulas por pontos em todas as células do DataFrame
tabela_Movel = tabela_Movel.replace(',', '.', regex=True)

def remover_acentos(palavra):
    mapeamento = {
        'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a',
        'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
        'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
        'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
        'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u'
    }
    palavra_sem_acentos = ''.join(mapeamento.get(letra, letra) for letra in palavra)
    return palavra_sem_acentos

def separarExcel(executivo):
    vendas_Executivo = tabela_Movel.loc[tabela_Movel['CONSULTOR'] == executivo]
    nomeSemAcento = remover_acentos(executivo.lower())
    # vendas_Executivo.to_excel(f'C:\\Users\\Math\\Documents\\fechamento\\vendas_Movel_{nomeSemAcento}.xlsx', index=False, engine='openpyxl')
    vendas_Executivo.to_excel(f'/home/math/Documents/fechamento/convertidos/vendas_Vada_{nomeSemAcento}.xlsx', index=False, engine='openpyxl')

# Lista de nomes únicos dos executivos
nomes_unicos = ['']
print(nomes_unicos)

# Iteração para salvar arquivos separados
for nome in nomes_unicos:
    separarExcel(nome)
