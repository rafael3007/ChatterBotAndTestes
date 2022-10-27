from __future__ import print_function
from ast import If

import time
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID_PE = '1KJaz1qfn4HZDpGwZEIrD1EXdPPgXSg6esjbaprjSlkM'
SAMPLE_RANGE_NAME = 'Colar A6!A2:G'


# Variables of ambient
RANGE_COLAR_A6 = "Colar A6!A2:G"
RANGE_COLAR_C9 = "Colar C9!A2:G"


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


def inserirDados(spreadsheet_id, range_name, _values):
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


def getDadosRelatorioC9(DIA, PLACA):

    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

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
        data = []

        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_PE,
                                    range=RANGE_COLAR_C9).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return
        cont = 1
        for row in values:
            cont = cont + 1
            DATA_TEMP = row[3].split(" ")[0]
            HORA_TEMP = row[3].split(" ")[1]

            if (DATA_TEMP == DIA and HORA_TEMP != "00:00:00" and row[1] == PLACA):
                data.append(row)
        return data
    except HttpError as err:
        print(err)


def getDadosRelatorioA6(DIA, PLACA):

    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

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
        data = []

        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_PE,
                                    range=RANGE_COLAR_A6).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return
        cont = 1
        for row in values:
            cont = cont + 1

            if (row[0] == DIA and row[1] == PLACA):
                data.append(row)
        return data
    except HttpError as err:
        print(err)


if __name__ == '__main__':

    DATA_DE_ANALISE = "25/10/2022"

    CONT_REQ = 0
    online = True

    index = 0
    lista_de_placas = pegarDados(SAMPLE_SPREADSHEET_ID_PE, "Listas!A2:A")

    while online:

        for placa in lista_de_placas:

            if CONT_REQ == 25:
                time.sleep(10)
                CONT_REQ = 0
            else:
                CONT_REQ = CONT_REQ + 1
            IGNICAO_LIGADA = ""
            SAIDA_DA_BASE = ""
            IGNICAO_DESLIGADA = ""
            OBS = ""
            LOCAL = ""
            MOTORISTA = ""
            LISTA_DE_DATAS = []

            print(str(placa[0]))

            # pegar dados dos relatóriso, tratar e inserir nas variaveis
            DATA_A6 = getDadosRelatorioA6(DATA_DE_ANALISE, str(placa[0]))
            DATA_C9 = getDadosRelatorioC9(DATA_DE_ANALISE, str(placa[0]))

            # print(DATA_A6)
            if DATA_A6 != [] and DATA_C9 != []:
                try:
                    IGNICAO_LIGADA = DATA_A6[0][2]
                    SAIDA_DA_BASE = DATA_C9[0][3].split(" ")[1]
                    IGNICAO_DESLIGADA = DATA_A6[len(DATA_A6)-1][4]
                    LOCAL = DATA_C9[0][0].split("_")[1]

                    if LOCAL == "RAULLINS":
                        LOCAL = "PETROLINA"
                    elif LOCAL == "NORTE":
                        OBS = "Evasão Norte"
                        LOCAL = "OURICURI"
                    elif LOCAL == "SUL":
                        OBS = "Evasão Sul"
                        LOCAL = "Petrolina"
                    elif LOCAL == "BOMNOME":
                        LOCAL == "SERRA TALHADA"
                    elif LOCAL == "SANTAMARIA":
                        LOCAL == "PETROLINA"
                        
                    DRIVER = []
                    contD = 0

                    for data in DATA_A6:
                        if data[5] not in DRIVER:
                            DRIVER.append(data[5])

                    for motorista in DRIVER:
                        if len(DRIVER)-1 == contD:
                            MOTORISTA = MOTORISTA+str(motorista)
                        else:
                            MOTORISTA = MOTORISTA+str(motorista)+"/"
                        contD = contD + 1

                    # getDadosRelatorioC9(DATA_DE_ANALISE,"QPX9I72")#str(placa[0])

                    # abrir aba
                    DATAS = pegarDados(
                        SAMPLE_SPREADSHEET_ID_PE, str(placa[0])+"!A1:A")

                    #LISTA_DE_DATAS = pegarDados(SAMPLE_SPREADSHEET_ID_PE,str(placa[0])+"!A2:A")
                    for data in DATAS:
                        LISTA_DE_DATAS = LISTA_DE_DATAS + data

                    # pesquisar o index da data analisada
                    index = LISTA_DE_DATAS.index(DATA_DE_ANALISE)+2

                    qtdLinhas = len(LISTA_DE_DATAS)

                    NEW_RANGE = str(placa[0])+"!B"+str(index)+":G"+str(index)

                    print(
                        f"------------------------------------PLACA {placa[0]}--------------------------------------")
                    print(f'Ignição ligada: {IGNICAO_LIGADA}')
                    print(f'Saída da base: {SAIDA_DA_BASE}')
                    print(f'Ignição desligada: {IGNICAO_DESLIGADA}')
                    print(f'OBS: {OBS}')
                    print(f'Local: {LOCAL}')
                    print(f'Motorista: {MOTORISTA}')

                except:

                    print("ERRO404")

            else:
                print(
                    f"------------------------------------PLACA {placa[0]}--------------------------------------")
                print("Dados vazios")

            # inserirDados(SAMPLE_SPREADSHEET_ID_PE,NEW_RANGE,[IGNICAO_LIGADA,SAIDA_DA_BASE,IGNICAO_DESLIGADA,OBS,LOCAL,MOTORISTA])

        online = False
