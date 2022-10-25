from os import chdir, getcwd, listdir
import zipfile
import pandas as pd
#from IPython.display import display
from datetime import date, timedelta

td = timedelta(-1)
USUARIO = "RAFAEL.SAMPAIO"
data_atual = date.today() + td

#cam = input('Digite o caminho: ')
CAMINHO = 'C:/Users/ECOELETRICA/Downloads'

# nome do arquivo vem no formato consulta _ servicos _ USUARIO _ DD-MM-YYYY _ HH-mm-ss

# movimentacao_fat_obra_gpm_USUARIO_DD-MM-YYYY_HH-mm-ss

data_em_texto = '{}-{}-{}'.format(data_atual.day, data_atual.month,
data_atual.year)

chdir(CAMINHO)
print("Caminho 1>>"+str(getcwd()))
EXCEPTION = ''

RELATORIOS = [
    "consulta_turno_"+str(USUARIO)+"_"+str(data_em_texto),
    "consulta_servicos_"+str(USUARIO)+"_"+str(data_em_texto),
    "movimentacao_fat_obra_gpm_"+str(USUARIO)+"_"+str(data_atual)
]

for c in listdir():
    try:
        LISTA_ARQUIVO = c.split("_")
        if len(LISTA_ARQUIVO)==5:
            NOME_ARQUIVO = str(c.split("_")[0])+"_"+str(c.split("_")[1])+"_"+str(c.split("_")[2])+"_"+str(c.split("_")[3])
            if NOME_ARQUIVO in RELATORIOS :
                z = zipfile.ZipFile(c,'r')
                z.extractall('C:/Users/ECOELETRICA/Desktop/BANCO_DE_RELATORIOS')
                z.close()
        else:
            if len(LISTA_ARQUIVO) == 6:
                DATA_TEMP = str(c.split("_")[5].split("-")[0])+"-"+str(c.split("_")[5].split("-")[1])+"-"+str(c.split("_")[5].split("-")[2])
                NOME_ARQUIVO = str(c.split("_")[0])+"_"+str(c.split("_")[1])+"_"+str(c.split("_")[2])+"_"+str(c.split("_")[3])+"_"+str(c.split("_")[4])+"_"+str(DATA_TEMP)
                #print(NOME_ARQUIVO)
                if NOME_ARQUIVO in RELATORIOS:
                    z = zipfile.ZipFile(c,'r')
                    z.extractall('C:/Users/ECOELETRICA/Desktop/BANCO_DE_RELATORIOS')
                    z.close()    
    except Exception:
        EXCEPTION = ''

CAMINHO = 'C:/Users/ECOELETRICA/Desktop/BANCO_DE_RELATORIOS'
chdir(CAMINHO)
print("Caminho 2>>"+str(getcwd()))

DATABASE_SERVICOS = []
DATABASE_TURNO = []
DATABASE_FATURAMENTO = []

ARQUIVO = ''
for arquivo in listdir():
    try:
        LISTA_ARQUIVO_EXTRAIDOS = arquivo.split("_")
        if len(LISTA_ARQUIVO_EXTRAIDOS)==5:
            NOME_ARQUIVO_EXTRAIDO = str(arquivo.split("_")[0])+"_"+str(arquivo.split("_")[1])+"_"+str(arquivo.split("_")[2])+"_"+str(arquivo.split("_")[3])
            if NOME_ARQUIVO_EXTRAIDO in RELATORIOS : 
                if str(arquivo.split("_")[0])+"_"+str(arquivo.split("_")[1]) == 'consulta_turno':
                    DATABASE_TURNO = pd.read_cvs(arquivo)
                    pass
                if str(arquivo.split("_")[0])+"_"+str(arquivo.split("_")[1]) == 'consulta_servicos':
                    DATABASE_SERVICOS = pd.read_cvs(arquivo)
                    pass
        else:
            if len(LISTA_ARQUIVO_EXTRAIDOS) == 6:
                DATA_TEMP = str(arquivo.split("_")[5].split("-")[0])+"-"+str(arquivo.split("_")[5].split("-")[1])+"-"+str(arquivo.split("_")[5].split("-")[2])
                NOME_ARQUIVO_EXTRAIDO = str(arquivo.split("_")[0])+"_"+str(arquivo.split("_")[1])+"_"+str(arquivo.split("_")[2])+"_"+str(arquivo.split("_")[3])+"_"+str(arquivo.split("_")[4])+"_"+str(DATA_TEMP)
                if NOME_ARQUIVO_EXTRAIDO in RELATORIOS:
                    DATABASE_FATURAMENTO = pd.read_cvs(arquivo)
                    #Ler os dados de movimentação faturamento obra gpm
    except:
        EXCEPTION = ''

print("-------------------------------------------")
print("DATABASE: FATURAMENTO")
print(DATABASE_FATURAMENTO)
print("-------------------------------------------")

print("-------------------------------------------")
print("DATABASE: SERVICO")
print(DATABASE_SERVICOS)
print("-------------------------------------------")

print("-------------------------------------------")
print("DATABASE: TURNO")
print(DATABASE_TURNO)
print("-------------------------------------------")
