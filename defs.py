from PIL import Image
from pyzbar.pyzbar import decode
import cv2 as cv

def converter(frame):
    try:
        img = Image.fromarray(frame)
        img_gray = img.convert("L")
        return decode(img_gray)
    except Exception as e:
        print("Erro ao processar a imagem:", e)
        return []

def exibir_texto(frame, texto, tempo_exibicao):
    frame_height, frame_width = frame.shape[:2]
    fundo = (0, 0, 0)  #Cor do fundo
    texto_cor = (255, 255, 255)  #Cor do texto
    fonte = cv.FONT_HERSHEY_SIMPLEX
    tamanho_fonte = 0.7
    espessura = 2
    linha_tipo = cv.LINE_AA

    (largura_texto, altura_texto), _ = cv.getTextSize(texto, fonte, tamanho_fonte, espessura)
    posicao_texto = ((frame_width - largura_texto) // 2, 50) 
    cv.rectangle(frame, (0, 0), (frame_width, 70), fundo, -1)  
    cv.putText(frame, texto, posicao_texto, fonte, tamanho_fonte, texto_cor, espessura, linha_tipo)
    cv.imshow("Camera", frame)
    cv.waitKey(tempo_exibicao)

def copiarformatacaodepresenca(service, spreadsheet_id, source_sheet_id, target_sheet_id):
    source_sheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[],
        fields="sheets.conditionalFormats"
    ).execute()['sheets'][source_sheet_id].get('conditionalFormats', [])

    for rule in source_sheet:
        for range_ in rule['ranges']:
            range_['sheetId'] = target_sheet_id

    requests = [{
        'addConditionalFormatRule': {
            'rule': rule,
            'index': 0
        }
    } for rule in source_sheet]

    if requests:
        body = {'requests': requests}
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

def copiarformatacao(service, spreadsheet_id, source_sheet_id, target_sheet_id):
    source_sheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[],
        fields="sheets.data.rowData.values.userEnteredFormat,sheets.data.rowData.values.effectiveFormat"
    ).execute()['sheets'][source_sheet_id].get('data', [])[0].get('rowData', [])

    requests = []
    for row_index, row in enumerate(source_sheet):
        for col_index, cell in enumerate(row.get('values', [])):
            if 'userEnteredFormat' in cell:
                requests.append({
                    'repeatCell': {
                        'range': {
                            'sheetId': target_sheet_id,
                            'startRowIndex': row_index,
                            'endRowIndex': row_index + 1,
                            'startColumnIndex': col_index,
                            'endColumnIndex': col_index + 1
                        },
                        'cell': {
                            'userEnteredFormat': cell['userEnteredFormat']
                        },
                        'fields': 'userEnteredFormat'
                    }
                })

    if requests:
        body = {'requests': requests}
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

def copiarlarguradecoluna(service, spreadsheet_id, source_sheet_id, target_sheet_id):
    source_sheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[],
        fields="sheets.data.columnMetadata"
    ).execute()['sheets'][source_sheet_id].get('data', [])[0].get('columnMetadata', [])

    requests = []
    for col_index, col in enumerate(source_sheet):
        if 'pixelSize' in col:
            requests.append({
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': target_sheet_id,
                        'dimension': 'COLUMNS',
                        'startIndex': col_index,
                        'endIndex': col_index + 1
                    },
                    'properties': {
                        'pixelSize': col['pixelSize']
                    },
                    'fields': 'pixelSize'
                }
            })

    if requests:
        body = {'requests': requests}
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
