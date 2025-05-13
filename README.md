# Gmail Word a Google Sheets

Este proyecto en Python detecta el correo más reciente en Gmail con un archivo Word adjunto (.docx), descarga el adjunto, extrae el texto y lo pega en una hoja de Google Sheets.

## Requisitos previos

1. **Habilita las APIs de Gmail y Google Sheets** en Google Cloud Console.
2. **Descarga el archivo `credentials.json`** y colócalo en la carpeta del proyecto.
3. **Crea una hoja de cálculo en Google Sheets** y copia su ID (de la URL).

## Instalación

Instala las dependencias con:

```bash
pip install -r requirements.txt
```

## Configuración

1. Coloca tu archivo `credentials.json` en la raíz del proyecto.
2. Abre `gmail_word_to_sheets.py` y reemplaza `TU_ID_DE_SHEET_AQUI` por el ID de tu hoja de Google Sheets.

## Ejecución

```bash
python gmail_word_to_sheets.py
```

La primera vez te pedirá iniciar sesión con tu cuenta de Google y autorizar permisos.

## Funcionamiento

- Busca el correo más reciente con un adjunto Word.
- Descarga el adjunto.
- Extrae el texto del Word.
- Pega cada párrafo en una fila de la hoja de Google Sheets.

## Notas

- El script solo procesa el correo más reciente con adjunto Word.
- El archivo Word descargado se elimina automáticamente después de procesarse.

---

¿Dudas o mejoras? ¡No dudes en preguntar! 