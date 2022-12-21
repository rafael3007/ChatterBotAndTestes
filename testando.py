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
TEMPO_DE_ESPERA = 0.25


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
    inicio = time.time()
    CONT_REQ = 0
    duracao = round(time.time() - inicio,2)
    duracao = 0.50

    if  duracao >= 0.45:
            print("time>> "+str(duracao))
            inicio = time.time()
            CONT_REQ = 0
        
    else:
        if CONT_REQ >= 20:
            print("req>>"+str(CONT_REQ))
            inicio = time.time()
            parada = True
            while parada:
                duracao = round(time.time() - inicio,2)
                print("Pause iniciado aos :"+str(duracao))
                parada = duracao < TEMPO_DE_ESPERA
            duracao = round(time.time() - inicio,2)
            print("Pause Finalizado aos :"+str(duracao))
            inicio = time.time()
            CONT_REQ = 0
        
        


    LAST_ROW_TURNO = len(pegarDados(SAMPLE_SPREADSHEET_ID_CONTROLE_GPM_TEST,"BASE DE TURNO!E2:E"))+1
    CONT_REQ = CONT_REQ + 1
    LAST_ROW_FATURAMENTO = len(pegarDados(SAMPLE_SPREADSHEET_ID_CONTROLE_GPM_TEST,"BASE FATURAMENTO!E2:E"))+1
    CONT_REQ = CONT_REQ + 1 

    LAST_ROW_SERVICOS = len(pegarDados(SAMPLE_SPREADSHEET_ID_CONTROLE_GPM_TEST,"BASE SERVIÇOS!A2:A"))+1
    CONT_REQ = CONT_REQ + 1
    
    lista_de_dados_tratados = []
    numRow = LAST_ROW_FATURAMENTO+1
    lista_de_dados_tratados = []

    for n,row in enumerate(DATABASE_FATURAMENTO):
        linha = []
        controleColunas = 0

        for colunas in row:
            linha.append(str(colunas))
            controleColunas = controleColunas +1
        
        duracao = round(time.time() - inicio,2)
        '''
        if  duracao >= 0.45:
            print("time>> "+str(duracao))
            inicio = time.time()
            CONT_REQ = 0
        else:
            if CONT_REQ >= 200:
                print("req>>"+str(CONT_REQ))
                inicio = time.time()
                parada = True
                while parada:
                    duracao = round(time.time() - inicio,2)
                    print("Pause iniciado aos :"+str(duracao))
                    parada = duracao < TEMPO_DE_ESPERA
                duracao = round(time.time() - inicio,2)
                print("Pause Finalizado aos :"+str(duracao))
                inicio = time.time()
                CONT_REQ = 0
        '''

        try:
            RANGE_BASE_FATURAMENTO = "BASE FATURAMENTO!E"+str(LAST_ROW_FATURAMENTO+n+1)+":W"+str(LAST_ROW_FATURAMENTO+n+1)
            response = update_values(SAMPLE_SPREADSHEET_ID_CONTROLE_GPM_TEST,RANGE_BASE_FATURAMENTO,"USER_ENTERED",linha)
            response_flag = response["spreadsheetId"]
            print("ok")
        except:
            flag_controle = True
            contador_de_erros = 0
            while flag_controle:       
                try:
                    contador_de_erros = contador_de_erros + 1
                    response = update_values(SAMPLE_SPREADSHEET_ID_CONTROLE_GPM_TEST,RANGE_BASE_FATURAMENTO,"USER_ENTERED",linha)
                    response_flag = response["spreadsheetId"]
                    print(response_flag)
                    if response_flag == '1fYCntz2rXMj4782Dy805XENPk4VKlgZu2ioE8XO27oI':
                        flag_controle = False
                except:
                    print(" Deu erro pula 1")
                



        CONT_REQ = CONT_REQ + 1 
        
        
        print("--------------")
        print(len(linha))
        print("--------------")

        lista_de_dados_tratados.append(linha)
        

    #update_values(SAMPLE_SPREADSHEET_ID_CONTROLE_GPM_TEST,RANGE_BASE_FATURAMENTO,"USER_ENTERED",lista_de_dados_tratados)    
    #print(RANGE_BASE_FATURAMENTO)
    #print("-------------------------------------------")
    


    #print("-------------------------------------------")
    #print("DATABASE: SERVICO")
    #temp = LAST_ROW_SERVICOS + len(DATABASE_SERVICOS)
    #RANGE_BASE_SERVICO = "BASE SERVICO!E"+str(LAST_ROW_SERVICOS+1)+":BX"+str(temp)
    #print(RANGE_BASE_SERVICO)
    #print("-------------------------------------------")

    #print("-------------------------------------------")
    #print("DATABASE: TURNO")
    #temp = LAST_ROW_TURNO + len(DATABASE_TURNO)
    #RANGE_BASE_DE_TURNO = "BASE DE TURNO!E"+str(LAST_ROW_TURNO+1)+":BQ"+str(temp)
    #print(RANGE_BASE_DE_TURNO)


if __name__ == "__main__":
    #update_values("1fYCntz2rXMj4782Dy805XENPk4VKlgZu2ioE8XO27oI","TEST!A2:B","USER_ENTERED",['Rafael12412','Br125125ito'])
    trata_dados()
   
    