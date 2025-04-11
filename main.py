from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pandas as pd
import re
import os
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()

# Obter valores sensíveis das variáveis de ambiente
spreadsheet_id = os.getenv('SPREADSHEET_ID')

# Configurar a autenticação OAuth com escopos adequados
flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', scopes=[
    'https://www.googleapis.com/auth/spreadsheets.readonly', 
    'https://www.googleapis.com/auth/drive.file', 
    'https://www.googleapis.com/auth/drive.readonly'
])
creds = flow.run_local_server(port=0)

# Conectar à API do Google Drive e Google Sheets
drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)

# Intervalo que contém os links
range_links = 'Respostas ao formulário 1!C2:C'

# Ler os links
result_links = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_links).execute()
links = result_links.get('values', [])

# Lista para acumular todos os dados extraídos
all_data = []

# Função para extrair ID do Google Drive ou Sheets
def extract_sheet_id(link):
    match_drive = re.search(r'/d/([a-zA-Z0-9-_]+)', link)
    match_open = re.search(r'id=([a-zA-Z0-9-_]+)', link)
    if match_drive:
        return match_drive.group(1)
    elif match_open:
        return match_open.group(1)
    return None

# Função para verificar se o arquivo é uma planilha do Google Sheets
def is_google_sheet(file_id):
    try:
        file = drive_service.files().get(fileId=file_id, fields='mimeType').execute()
        return file['mimeType'] == 'application/vnd.google-apps.spreadsheet'
    except Exception as error:
        print(f"Erro ao verificar o arquivo: {error}")
        return False

# Função para verificar se o arquivo está acessível
def is_accessible(file_id):
    try:
        drive_service.files().get(fileId=file_id).execute()
        return True
    except Exception as error:
        print(f"Erro ao verificar o acesso ao arquivo: {error}")
        return False

# Função para obter o nome da primeira aba da planilha
def get_first_sheet_name(file_id):
    try:
        sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=file_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        if sheets:
            return sheets[0].get("properties", {}).get("title", "Sheet1")
        else:
            return "Sheet1"
    except Exception as error:
        print(f"Erro ao obter o nome da aba: {error}")
        return "Sheet1"

# Função para converter arquivo Excel para Google Sheets
def convert_to_google_sheets(file_id):
    try:
        file_metadata = {
            'name': 'ConvertedSheet',
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        }
        converted_file = drive_service.files().copy(fileId=file_id, body=file_metadata).execute()
        return converted_file['id']
    except Exception as error:
        print(f"Erro ao converter o arquivo: {error}")
        return None

# Iterar sobre cada link e extrair dados
for link in links:
    link = link[0]
    file_id = extract_sheet_id(link)
    if file_id and is_accessible(file_id):
        if not is_google_sheet(file_id):
            file_id = convert_to_google_sheets(file_id)
        if file_id and is_google_sheet(file_id):
            try:
                sheet_name = get_first_sheet_name(file_id)
                range_data = f'{sheet_name}!A1:Z'
                result_data = sheets_service.spreadsheets().values().get(spreadsheetId=file_id, range=range_data).execute()
                data = result_data.get('values', [])
                all_data.extend(data)
                print(f"Dados extraídos da planilha: {data}")
            except Exception as e:
                print(f"Erro ao processar o link: {link}\n{e}")
        else:
            print(f"Formato de link inválido ou não é uma planilha do Google Sheets: {link}")
    else:
        print(f"Arquivo não acessível ou não encontrado: {link}")

# Verificar os dados extraídos
if all_data:
    print(f"Dados acumulados: {all_data}")
    df = pd.DataFrame(all_data)
    # caminho_arquivo_excel = 'L:\\Umadeb\\Uploads\\dados_extraidos.xlsx'
    caminho_arquivo_excel = os.getenv('caminho_arquivo_excel')
    if not os.path.exists(os.path.dirname(caminho_arquivo_excel)):
        os.makedirs(os.path.dirname(caminho_arquivo_excel))
    df.to_excel(caminho_arquivo_excel, index=False)
    print(f"Dados extraídos e salvos em: {caminho_arquivo_excel}")
else:
    print("Nenhum dado foi extraído para salvar.")
