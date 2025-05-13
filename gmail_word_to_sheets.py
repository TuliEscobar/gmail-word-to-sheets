import os
import pickle
import base64
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from docx import Document
from dotenv import load_dotenv
import google.generativeai as genai
import json

# Cargar variables de entorno
load_dotenv()
GEMINI_API_KEY = os.getenv('API_KEY_GOOGLE')
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
RANGE_NAME = os.getenv('RANGE_NAME', 'A1')

# SCOPES necesarios
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]



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

def extraer_parametros_con_gemini(texto):
    if not GEMINI_API_KEY:
        raise Exception("No se encontró la API Key de Gemini. Asegúrate de tener API_KEY_GOOGLE en tu .env.")
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel('gemini-1.5-flash')
    prompt = (
        """
        Extrae los siguientes parámetros del texto de un documento oficial. Devuelve solo un JSON con las claves exactas:
        REF, CONCURSO, FECHA Y HORA, DESCRIPCION, CANTIDAD.
        Si algún dato no está, deja el valor vacío.
        Ejemplo de respuesta:
        {"REF": "Bordon Ruben Anibal", "CONCURSO": "2171/2025", "FECHA Y HORA": "12/05/2025 09:05:00", "DESCRIPCION": "Pembrolizumab 100 mg. Fco. Amp. x 1 x 4 ml.", "CANTIDAD": "2"}
        Texto:
        """ + texto
    )
    response = model.generate_content(prompt)
    # Buscar el primer bloque JSON en la respuesta
    try:
        json_str = response.text[response.text.index('{'):response.text.rindex('}')+1]
        datos = json.loads(json_str)
    except Exception as e:
        print("Error al parsear la respuesta de Gemini:", e)
        print("Respuesta completa:", response.text)
        datos = {"REF": "", "CONCURSO": "", "FECHA Y HORA": "", "DESCRIPCION": "", "CANTIDAD": ""}
    return datos

def escribir_en_sheets_parametros(creds, parametros, spreadsheet_id):
    service = build('sheets', 'v4', credentials=creds)
    # Orden: A: REF, B: CONCURSO, C: FECHA Y HORA, D: DESCRIPCION, E: CANTIDAD
    values = [[
        parametros.get("REF", ""),
        parametros.get("CONCURSO", ""),
        parametros.get("FECHA Y HORA", ""),
        parametros.get("DESCRIPCION", ""),
        parametros.get("CANTIDAD", "")
    ]]
    body = {'values': values}
    result = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id, range='A1',
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
    parametros = extraer_parametros_con_gemini(texto)
    escribir_en_sheets_parametros(creds, parametros, SPREADSHEET_ID)
    # Limpia el archivo descargado
    os.remove(filename)

if __name__ == '__main__':
    main() 