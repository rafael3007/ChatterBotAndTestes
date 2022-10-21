import pandas as pd
from IPython.display import display

DATA = pd.read_excel('testando.xlsx')
TIPOS_DE_DADOS = pd.DataFrame(DATA.dtypes, columns = ["Tipos de Dados"])

TIPOS_DE_DADOS.columns.name = 'Variaveis'
# numLinhas 

DATA.shape[0]

#numColumns
DATA.shape[1]


vetor = DATA.values


#print('A base de dados apresenta {} registros (imóveis) e {} variávies'.format(DATA.shape[0], DATA.shape[1]))
print(vetor)