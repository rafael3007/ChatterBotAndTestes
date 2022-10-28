from os import chdir, getcwd, listdir
import zipfile
import pandas as pd
#from IPython.display import display
from datetime import date, timedelta


def inserirDados(spreadsheet_id, range_name,_values):
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    try:
        service = build('sheets', 'v4', credentials=creds)

        values = [
            
               _values,
            # Additional rows
        ]
        data = [
            {
                'range': range_name,
                'values': values
            },
            # Additional ranges to update ...
        ]
        body = {
            'valueInputOption': "USER_ENTERED",
            'data': data
        }
        result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id, body=body).execute()
        #print(f"{(result.get('totalUpdatedCells'))} cells updated.")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


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
                    print(arquivo)
                    DATABASE_TURNO = pd.read_csv(arquivo,sep=";").values
                    
                if str(arquivo.split("_")[0])+"_"+str(arquivo.split("_")[1]) == 'consulta_servicos':
                    print(arquivo)
                    DATABASE_SERVICOS = pd.read_csv(arquivo,sep=";").values
                
        else:
            if len(LISTA_ARQUIVO_EXTRAIDOS) == 6:
                DATA_TEMP = str(arquivo.split("_")[5].split("-")[0])+"-"+str(arquivo.split("_")[5].split("-")[1])+"-"+str(arquivo.split("_")[5].split("-")[2])
                NOME_ARQUIVO_EXTRAIDO = str(arquivo.split("_")[0])+"_"+str(arquivo.split("_")[1])+"_"+str(arquivo.split("_")[2])+"_"+str(arquivo.split("_")[3])+"_"+str(arquivo.split("_")[4])+"_"+str(DATA_TEMP)
                if NOME_ARQUIVO_EXTRAIDO in RELATORIOS:
                    print(arquivo)
                    DATABASE_FATURAMENTO = pd.read_csv(arquivo,sep=";").values
                            #Ler os dados de movimentação faturamento obra gpm
    except:
        EXCEPTION = ''

print("-------------------------------------------")
print("DATABASE: FATURAMENTO")
print(len(DATABASE_FATURAMENTO))
print("-------------------------------------------")

print("-------------------------------------------")
print("DATABASE: SERVICO")
print(len(DATABASE_SERVICOS))
print("-------------------------------------------")

print("-------------------------------------------")
print("DATABASE: TURNO")
print(len(DATABASE_TURNO))
print("-------------------------------------------")
