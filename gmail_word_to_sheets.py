import os
import pickle
import base64
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from docx import Document

# SCOPES necesarios
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

# ID de tu hoja de cálculo de Google Sheets
SPREADSHEET_ID = '16Jz0JVAcuCk2qfjuG2xeJzTystSdbnyfF2L8B2dvYX0'  # <-- Cambia esto por el ID de tu hoja
RANGE_NAME = 'A1'  # Puedes cambiar el rango si lo deseas

def authenticate_google():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def buscar_ultimo_correo_con_word(service):
    results = service.users().messages().list(userId='me', q="has:attachment filename:docx", maxResults=1).execute()
    messages = results.get('messages', [])
    if not messages:
        print("No se encontraron correos con adjuntos Word.")
        return None
    return messages[0]['id']

def descargar_adjunto_word(service, msg_id):
    message = service.users().messages().get(userId='me', id=msg_id).execute()
    for part in message['payload'].get('parts', []):
        filename = part.get('filename')
        if filename and filename.endswith('.docx'):
            att_id = part['body']['attachmentId']
            att = service.users().messages().attachments().get(userId='me', messageId=msg_id, id=att_id).execute()
            data = att['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            with open(filename, 'wb') as f:
                f.write(file_data)
            print(f"Adjunto guardado como: {filename}")
            return filename
    print("No se encontró adjunto Word en el correo.")
    return None

def extraer_texto_word(filename):
    doc = Document(filename)
    texto = []
    for para in doc.paragraphs:
        texto.append(para.text)
    return '\n'.join(texto)

def escribir_en_sheets(creds, texto, spreadsheet_id, range_name):
    service = build('sheets', 'v4', credentials=creds)
    # Divide el texto en filas (puedes ajustar esto según tu formato)
    values = [[line] for line in texto.split('\n') if line.strip()]
    body = {'values': values}
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id, range=range_name,
        valueInputOption="RAW", body=body).execute()
    print(f"{result.get('updates').get('updatedCells')} celdas actualizadas en Google Sheets.")

def main():
    creds = authenticate_google()
    gmail_service = build('gmail', 'v1', credentials=creds)
    msg_id = buscar_ultimo_correo_con_word(gmail_service)
    if not msg_id:
        return
    filename = descargar_adjunto_word(gmail_service, msg_id)
    if not filename:
        return
    texto = extraer_texto_word(filename)
    escribir_en_sheets(creds, texto, SPREADSHEET_ID, RANGE_NAME)
    # Limpia el archivo descargado
    os.remove(filename)

if __name__ == '__main__':
    main() 