from __future__ import print_function
import os
from PIL import Image
from pyzbar.pyzbar import decode
import datetime
import cv2 as cv
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

def converter(frame):
    try:
        img = Image.fromarray(frame)  
        img_gray = img.convert("L")  
        return decode(img_gray)
    except Exception as e:
        print("Erro ao processar a imagem:", e)
        return []


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

    # ID da planilha e nome da aba
    spreadsheet_id = '17HqMNAT7WVbt9FsrtNz-7SfFIejBsLL9yZq7ysHFUhQ'
    range_name = 'Sheet1!B:B'  
    
    # Chama a API do Sheets
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    #colocar data
    data = datetime.datetime.now().strftime("%d/%B/%Y")
    result = sheet.values().update(spreadsheetId=spreadsheet_id,
                                range='Sheet1!E2', valueInputOption="RAW",
                                   body={"values": [[data]]}).execute()
    #camera
    camera = cv.VideoCapture(0)
    rodando = True

    while rodando:
        status, frame = camera.read()

        if not status:
            print("Erro ao capturar imagem da câmera!")
            break
        cv.imshow("Camera", frame)
        key = cv.waitKey(1) & 0xFF

        if key == ord('s'):
            codigos = converter(frame)
            if codigos:
                for codigo in codigos:
                    codigo_decodificado = codigo.data.decode("utf-8")
                    try:
                        codigo_int = int(codigo_decodificado)
                    except ValueError:
                        print(f"Erro ao converter '{codigo_decodificado}'.")
                    else:
                        for i, row in enumerate(values):
                            try:
                                if row and row[0] != 'NUsp' and int(row[0]) == codigo_int:
                                    cell_range = f'Sheet1!C{i + 1}'
                                    receber = [["Presente"]]
                                    update_result = sheet.values().update(
                                        spreadsheetId=spreadsheet_id,
                                        range=cell_range,
                                        valueInputOption="RAW",
                                        body={"values": receber}).execute()
                                    nome_range = 'Sheet1!A{}'.format(i + 1)
                                    nome_result = sheet.values().get(spreadsheetId=spreadsheet_id,range=nome_range).execute()
                                    if 'values' in nome_result:
                                        nome = nome_result['values'][0][0]
                                        print(f'{nome} presente!')
                            except:
                                print(f'Código {codigo_int} não encontrado na lista')                       
        if key == ord('q'):
            rodando = False

if __name__ == '__main__':
    main()