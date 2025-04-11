from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pandas as pd
import re
import os
from dotenv import load_dotenv

def extract_sheet_id(link):
    match_drive = re.search(r'/d/([a-zA-Z0-9-_]+)', link)
    match_open = re.search(r'id=([a-zA-Z0-9-_]+)', link)
    if match_drive:
        return match_drive.group(1)
    elif match_open:
        return match_open.group(1)
    return None


def is_google_sheet(file_id, drive_service):
    try:
        file = drive_service.files().get(fileId=file_id, fields='mimeType').execute()
        return file['mimeType'] == 'application/vnd.google-apps.spreadsheet'
    except Exception as error:
        print(f"Erro ao verificar o tipo do arquivo: {error}")
        return False


def is_accessible(file_id, drive_service):
    try:
        drive_service.files().get(fileId=file_id).execute()
        return True
    except Exception as error:
        print(f"Erro ao verificar o acesso: {error}")
        return False


def get_first_sheet_name(file_id, sheets_service):
    try:
        sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=file_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        if sheets:
            return sheets[0].get("properties", {}).get("title", "Sheet1")
    except Exception as error:
        print(f"Erro ao obter o nome da aba: {error}")
    return "Sheet1"


def convert_to_google_sheets(file_id, drive_service):
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


def main():
    load_dotenv()

    spreadsheet_id = os.getenv('SPREADSHEET_ID')
    output_path = os.getenv('CAMINHO_ARQUIVO_EXCEL')

    if not spreadsheet_id or not output_path:
        print("Erro: Verifique se SPREADSHEET_ID e CAMINHO_ARQUIVO_EXCEL estão definidos no .env")
        return

    # Autenticação
    flow = InstalledAppFlow.from_client_secrets_file(
        'client_secret.json',
        scopes=[
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/drive.file',
            'https://www.googleapis.com/auth/drive.readonly'
        ]
    )
    creds = flow.run_local_server(port=0)

    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)

    # Buscar os links
    range_links = 'Respostas ao formulário 1!C2:C'
    result_links = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_links).execute()
    links = result_links.get('values', [])

    all_data = []

    print(f"{len(links)} links encontrados. Iniciando processamento...\n")

    for row in links:
        link = row[0]
        file_id = extract_sheet_id(link)

        if not file_id:
            print(f"Link inválido: {link}")
            continue

        if not is_accessible(file_id, drive_service):
            print(f"Arquivo inacessível: {link}")
            continue

        if not is_google_sheet(file_id, drive_service):
            file_id = convert_to_google_sheets(file_id, drive_service)

        if file_id and is_google_sheet(file_id, drive_service):
            try:
                sheet_name = get_first_sheet_name(file_id, sheets_service)
                range_data = f'{sheet_name}!A1:Z'
                result_data = sheets_service.spreadsheets().values().get(
                    spreadsheetId=file_id, range=range_data).execute()
                data = result_data.get('values', [])
                all_data.extend(data)
                print(f"Dados extraídos com sucesso do arquivo {file_id}")
            except Exception as e:
                print(f"Erro ao extrair dados do link: {link}\n{e}")
        else:
            print(f"Link não aponta para planilha válida: {link}")

    if all_data:
        df = pd.DataFrame(all_data)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        df.to_excel(output_path, index=False)
        print(f" Dados salvos com sucesso em: {output_path}")
    else:
        print("Nenhum dado foi extraído.")


if __name__ == "__main__":
    main()
