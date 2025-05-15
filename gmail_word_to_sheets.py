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
import re
import win32com.client as win32
from tempfile import mkstemp
import time
import subprocess
import schedule
from PyPDF2 import PdfReader

# Cargar variables de entorno
load_dotenv()
GEMINI_API_KEY = os.getenv('API_KEY_GOOGLE')
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
RANGE_NAME = os.getenv('RANGE_NAME', 'A1')

# SCOPES necesarios
SCOPES = [
    'https://www.googleapis.com/auth/gmail.modify',
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

def buscar_ultimo_correo_con_adjunto_no_leido(service):
    results = service.users().messages().list(
        userId='me', 
        q="is:unread has:attachment (filename:docx OR filename:doc OR filename:pdf)", 
        maxResults=1
    ).execute()
    messages = results.get('messages', [])
    if not messages:
        print("No se encontraron correos con adjuntos Word o PDF.")
        return None
    return messages[0]['id']

def descargar_adjunto(service, msg_id):
    message = service.users().messages().get(userId='me', id=msg_id).execute()
    for part in message['payload'].get('parts', []):
        filename = part.get('filename')
        if filename and filename.lower().endswith(('.docx', '.doc', '.pdf')):
            # Limpiar nombre de archivo y usar ruta temporal segura
            safe_filename = ''.join(c for c in filename if c.isalnum() or c in (' ', '.', '_')).rstrip()
            temp_path = os.path.join(os.getenv('TEMP', '.'), safe_filename)
            
            att_id = part['body']['attachmentId']
            att = service.users().messages().attachments().get(userId='me', messageId=msg_id, id=att_id).execute()
            data = att['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            
            try:
                with open(temp_path, 'wb') as f:
                    f.write(file_data)
                print(f"Adjunto guardado como: {temp_path}")
                return temp_path
            except PermissionError:
                # Segundo intento con nombre alternativo
                alt_path = os.path.join(os.getenv('TEMP', '.'), f"temp_{int(time.time())}.pdf")
                with open(alt_path, 'wb') as f:
                    f.write(file_data)
                print(f"Adjunto guardado como: {alt_path}")
                return alt_path
    
    print("No se encontró adjunto Word o PDF en el correo.")
    return None

def convert_doc_to_docx(doc_path):
    """Convierte archivo .doc a .docx usando Word"""
    try:
        # Primero cerramos cualquier instancia previa de Word
        os.system('taskkill /f /im winword.exe')
        
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False  # Deshabilitar alertas
        
        # Usar ruta absoluta y esperar entre operaciones
        doc = word.Documents.Open(os.path.abspath(doc_path))
        time.sleep(1)  # Pequeña pausa
        
        new_path = mkstemp(suffix='.docx')[1]
        doc.SaveAs(new_path, FileFormat=16)
        time.sleep(1)
        
        doc.Close(False)
        word.Quit()
        time.sleep(1)
        
        # Forzar liberación de recursos
        del doc
        del word
        
        os.remove(doc_path)
        return new_path
    except Exception as e:
        print(f"Error al convertir {doc_path} a .docx: {str(e)}")
        try:
            word.Quit()
        except:
            pass
        return None

def extraer_texto_archivo(filename):
    """Extrae texto de archivos Word (.doc, .docx) o PDF"""
    try:
        if filename.lower().endswith('.pdf'):
            # Extraer texto de PDF
            with open(filename, 'rb') as f:
                reader = PdfReader(f)
                text = ''
                for page in reader.pages:
                    text += page.extract_text() + '\n'
            return text
        elif filename.lower().endswith('.doc'):
            # Método alternativo para .doc sin antiword
            try:
                with open(filename, 'rb') as f:
                    content = f.read()
                    # Extraer texto entre secuencias de texto comunes en .doc
                    text_parts = re.findall(b'[\x20-\x7E\x0A\x0D]{20,}', content)
                    return b' '.join(text_parts).decode('latin-1', errors='ignore')
            except Exception as e:
                print(f"Error al leer .doc: {str(e)}")
                return ""
        else:
            # Método para .docx
            try:
                doc = Document(filename)
                return '\n'.join(para.text for para in doc.paragraphs if para.text)
            except Exception as e:
                print(f"Error al leer .docx: {str(e)}")
                return ""
    except Exception as e:
        print(f"Error al leer archivo {filename}: {str(e)}")
        return ""
    finally:
        if filename and os.path.exists(filename):
            try:
                os.remove(filename)
            except:
                pass

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
        El dato de REF siempre tiene que ser un nombre completo o dejarlo vacio.
        IMPORTANTE: Siempre hay nombres asi que busca y agregalo.
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

def marcar_como_leido(service, msg_id):
    service.users().messages().modify(
        userId='me',
        id=msg_id,
        body={'removeLabelIds': ['UNREAD']}
    ).execute()

def main():
    while True:
        try:
            print("\n--- Iniciando ciclo de procesamiento ---")
            creds = authenticate_google()
            gmail_service = build('gmail', 'v1', credentials=creds)
            msg_id = buscar_ultimo_correo_con_adjunto_no_leido(gmail_service)
            if msg_id:
                print(f"Procesando correo con ID: {msg_id}")
                filename = descargar_adjunto(gmail_service, msg_id)
                if filename:
                    print(f"Archivo descargado: {filename}")
                    if filename.lower().endswith('.doc'):
                        filename = convert_doc_to_docx(filename)
                    texto = extraer_texto_archivo(filename)
                    parametros = extraer_parametros_con_gemini(texto)
                    escribir_en_sheets_parametros(creds, parametros, SPREADSHEET_ID)
                    marcar_como_leido(gmail_service, msg_id)
            else:
                print("No hay correos nuevos con adjuntos Word o PDF")
        except Exception as e:
            print(f"Error en ejecución: {str(e)}")
        
        print("Esperando 1 minuto para el próximo ciclo...")
        time.sleep(60)  # Espera 1 minuto (60 segundos)

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\nDeteniendo el servicio...")