## <h1 align='center'>Python_with_Gsheets</h1>
<br>
<p>Este projeto tem como principal intuito melhorar o nosso relatório diário de telemetria, por meio de uma automação em Python.</p>
<br>
<a  name="ancora"></a>

## :dvd: Principais Funcionalidades

- [Pegar Dados](#PegarDados)
- [Inserir Dados](#InserirDados)
- [Dados Relatorio A6](#getDadosRelatorioA6)
- [Dados Relatorio C9](#getDadosRelatorioC9)

<a id="PegarDados"></a>
<br/>
### :bulb: Pegar Dados
```php
def pegarDados(spreadsheet_id,range_name):
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
```

<a id="InserirDados"></a>
<br/>
### :bulb: Inserir Dados

```php
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
```

<a id="getDadosRelatorioA6"></a>
<br/>
### :bulb: Pegar Dados Relatorio A6

```php
def getDadosRelatorioC9(DIA,PLACA):
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
            if(DATA_TEMP == DIA and HORA_TEMP != "00:00:00" and row[1] == PLACA):
                data.append(row)
        return data
    except HttpError as err:
        print(err)
```

<a id="getDadosRelatorioC9"></a>
<br/>
### :bulb: Pegar Dados Relatorio C9

```php
def getDadosRelatorioA6(DIA,PLACA):
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
            #apenas test
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
            if(row[0] == DIA and row[1] == PLACA):
                data.append(row)
        return data
    except HttpError as err:
        print(err)
```

## :computer: Autores

* **Rafael** - [Github](https://github.com/rafael3007)
* **João Vittor** - [Github](https://github.com/JoaoVittorL)