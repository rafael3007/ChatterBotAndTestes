from os import chdir, getcwd, listdir
import zipfile
import pandas as pd
from IPython.display import display

#cam = input('Digite o caminho: ')
CAMINHO = 'C:/Users/ECOELETRICA/Downloads'
chdir(CAMINHO)
print(getcwd())
EXCEPTION = ''

for c in listdir():
    try:
        if c.split(".")[1] == 'zip':
            if c.split("_")[0] == "TESTANDO":
                z = zipfile.ZipFile(c,'r')
                z.extractall('C:/Users/ECOELETRICA/Desktop')
                z.close()
                break
    except Exception:
        EXCEPTION = ''

CAMINHO = 'C:/Users/ECOELETRICA/Desktop'
chdir(CAMINHO)
print(getcwd())

ARQUIVO = ''
for arquivo in listdir():
    try:
        if arquivo == 'Testando.xlsx':
            ARQUIVO = arquivo
            
    except Exception:
        EXCEPTION = ''


DATA = pd.read_excel(ARQUIVO)

print(DATA.values)
