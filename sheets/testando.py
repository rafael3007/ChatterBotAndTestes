from __future__ import print_function
from os import chdir, getcwd, listdir
import zipfile
import pandas as pd
#from IPython.display import display
from datetime import date, timedelta


import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError



import time
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SAMPLE_SPREADSHEET_ID_CONTROLE_GPM_TEST = '1fYCntz2rXMj4782Dy805XENPk4VKlgZu2ioE8XO27oI'



def pegarDados(spreadsheet_id, range_name):
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

        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=range_name).execute()
        #rows = result.get('values', [])
        #print(f"{len(rows)} rows retrieved")
        return result["values"]
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error



def update_values(spreadsheet_id, range_name, value_input_option,
                  _values):
    """
    Creates the batch_update the user has access to.
    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
        """
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # pylint: disable=maybe-no-member
    try:

        service = build('sheets', 'v4', credentials=creds)
        values = [
                # Cell values ...
                _values,
            # Additional rows ...
        ]
        body = {
            'values': values
        }
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption=value_input_option, body=body).execute()
        print(f"{result.get('updatedCells')} cells updated.")
        return result
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error



def trata_dados():

    CONT_REQ = 0
    online = True

    
    index = 0


    td = timedelta(0)
    USUARIO = "RAFAEL.SAMPAIO"
    data_atual = date.today()
    init_timer = time.time()

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


    #pegar a ultima linha da aba de turno e concatenar no range
    chdir('C:/Users/ECOELETRICA/Desktop/python/Relatório/python_with_Gsheets')
    
    LAST_ROW_TURNO = len(pegarDados(SAMPLE_SPREADSHEET_ID_CONTROLE_GPM_TEST,"BASE DE TURNO!E2:E"))+1

    LAST_ROW_FATURAMENTO = len(pegarDados(SAMPLE_SPREADSHEET_ID_CONTROLE_GPM_TEST,"BASE FATURAMENTO!E2:E"))+1

    LAST_ROW_SERVICOS = len(pegarDados(SAMPLE_SPREADSHEET_ID_CONTROLE_GPM_TEST,"BASE SERVIÇOS!A2:A"))+1

    tempoInicial = time.time()
    for n,row in enumerate(DATABASE_FATURAMENTO):

        numRow = LAST_ROW_FATURAMENTO+1+n
        temp = LAST_ROW_FATURAMENTO + len(DATABASE_FATURAMENTO)
        print(f'qtd: {temp}')
        RANGE_BASE_FATURAMENTO = "BASE FATURAMENTO!E"+str(numRow)+":W"+str(numRow)
        
        lista_temporaria = []


        time_after = time.time()
        PARADO = round(time_after - tempoInicial,2) > 50.00

        if CONT_REQ == 24 or PARADO:
            print("parada iniciada"+str(round(time_after - tempoInicial,2)))
            PARADO = True 
            while PARADO:
                PARADO = round(time.time() - time_after,2) < 21.00
            CONT_REQ = 0
            print("parada acabou"+str(round(time_after - tempoInicial,2)))
            tempoInicial = time.time()
        else:
            CONT_REQ = CONT_REQ + 1

        
        for coluna in row:
            coluna = str(coluna)
            
            if len(coluna) == 3:
                if coluna == 'nan':
                    lista_temporaria.append('')
                elif coluna == '0,00':
                    lista_temporaria.append(0.00)
                else:
                  lista_temporaria.append(coluna)  
            else:
                lista_temporaria.append(coluna)
            
        print(f'Linha original: {row}')
        print(f'Linha Convertida: {lista_temporaria}')
        #pause if 25 or 1min
        update_values(SAMPLE_SPREADSHEET_ID_CONTROLE_GPM_TEST,RANGE_BASE_FATURAMENTO,"USER_ENTERED",lista_temporaria)
    print(RANGE_BASE_FATURAMENTO)
    print("-------------------------------------------")



    print("-------------------------------------------")
    print("DATABASE: SERVICO")
    temp = LAST_ROW_SERVICOS + len(DATABASE_SERVICOS)
    RANGE_BASE_SERVICO = "BASE SERVICO!E"+str(LAST_ROW_SERVICOS+1)+":BX"+str(temp)
    print(RANGE_BASE_SERVICO)
    print("-------------------------------------------")

    print("-------------------------------------------")
    print("DATABASE: TURNO")
    temp = LAST_ROW_TURNO + len(DATABASE_TURNO)
    RANGE_BASE_DE_TURNO = "BASE DE TURNO!E"+str(LAST_ROW_TURNO+1)+":BQ"+str(temp)
    print(RANGE_BASE_DE_TURNO)
    print("-------------------------------------------")



if __name__ == "__main__":
    #update_values("1fYCntz2rXMj4782Dy805XENPk4VKlgZu2ioE8XO27oI","TEST!A2:B","USER_ENTERED",['Rafael12412','Br125125ito'])
    trata_dados()
   
    