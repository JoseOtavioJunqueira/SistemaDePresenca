import cv2 as cv
from PIL import Image
from pyzbar.pyzbar import decode
import datetime
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from defs import *

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    spreadsheet_id = '17HqMNAT7WVbt9FsrtNz-7SfFIejBsLL9yZq7ysHFUhQ'
    range_name = 'Sheet1!E2'
    sheet = service.spreadsheets()

    # Verificar a data atual
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    current_date = datetime.datetime.now().strftime("%d/%B/%Y")

    # Se a data da planilha for diferente da data atual, criar uma nova aba
    if not values or values[0][0] != current_date:
        new_sheet_title = f"Sheet_{current_date.replace('/', '_')}"
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', '')

        # Criar nova aba
        new_sheet = {
            'properties': {
                'title': new_sheet_title
            }
        }
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                'requests': [
                    {'addSheet': new_sheet}
                ]
            }
        ).execute()

        new_sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']
        original_sheet_id = sheet_metadata['sheets'][0]['properties']['sheetId']

        # Copiar dados da aba original para a nova aba
        source_range = 'Sheet1!A:E'
        source_values = sheet.values().get(spreadsheetId=spreadsheet_id, range=source_range).execute().get('values', [])

        if source_values:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{new_sheet_title}!A1",
                valueInputOption="RAW",
                body={"values": source_values}
            ).execute()

        # Marcar todos como ausentes na nova aba
        row_count = len(source_values)
        absence_values = [["Ausente"] for _ in range(1, row_count)]
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{new_sheet_title}!C2:C{row_count}",
            valueInputOption="RAW",
            body={"values": absence_values}
        ).execute()

        # Atualizar a data na nova aba
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{new_sheet_title}!E2",
            valueInputOption="RAW",
            body={"values": [[current_date]]}
        ).execute()

        # Copiar formatação condicional
        copiarformatacaodepresenca(service, spreadsheet_id, original_sheet_id, new_sheet_id)

        # Copiar formatação de células
        copiarformatacao(service, spreadsheet_id, original_sheet_id, new_sheet_id)

        # Copiar tamanhos das colunas
        copiarlarguradecoluna(service, spreadsheet_id, original_sheet_id, new_sheet_id)

        range_name = f"{new_sheet_title}!B:B"
    else:
        range_name = 'Sheet1!B:B'

    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    data = current_date
    camera = cv.VideoCapture(0)
    rodando = True

    try:
        while rodando:
            status, frame = camera.read()
            if not status:
                print("Erro ao capturar imagem da câmera!")
                break
            frame_height, frame_width = frame.shape[:2]
            codigos = converter(frame)
            texto = ""
            if codigos:
                for codigo in codigos:
                    codigo_decodificado = codigo.data.decode("utf-8")
                    try:
                        codigo_int = int(codigo_decodificado)
                    except ValueError:
                        print(f"Erro ao converter '{codigo_decodificado}'.")
                        continue
                    for i, row in enumerate(values):
                        if row and row[0] != 'NUsp' and int(row[0]) == codigo_int:
                            cell_range = f'{new_sheet_title}!C{i + 1}'  
                            receber = [["Presente"]]
                            update_result = sheet.values().update(
                                spreadsheetId=spreadsheet_id,
                                range=cell_range,
                                valueInputOption="RAW",
                                body={"values": receber}).execute()
                            nome_range = f'{new_sheet_title}!A{i + 1}'
                            nome_result = sheet.values().get(spreadsheetId=spreadsheet_id, range=nome_range).execute()
                            if 'values' in nome_result:
                                nome = nome_result['values'][0][0]
                                texto = f'{nome} presente!'
                                print(texto)
                                exibir_texto(frame, texto, tempo_exibicao=1500)
            
            cv.imshow("Camera", frame)
            key = cv.waitKey(1) & 0xFF
            if key == ord('q'):
                rodando = False
    finally:
        camera.release()
        cv.destroyAllWindows()

if __name__ == '__main__':
    main()
