# Gmail Word a Google Sheets con Gemini 2.5

Este proyecto en Python detecta el correo NO LEÍDO más reciente en Gmail con un archivo Word adjunto (.docx), descarga el adjunto, extrae el texto y utiliza Gemini 2.5 para identificar parámetros clave, que luego pega automáticamente en una hoja de Google Sheets.

## Requisitos previos

1. **Habilita las APIs de Gmail y Google Sheets** en Google Cloud Console.
2. **Descarga el archivo `credentials.json`** y colócalo en la carpeta del proyecto.
3. **Crea una hoja de cálculo en Google Sheets** y copia su ID (de la URL).
4. **Obtén una API Key de Gemini 2.5** desde [Google AI Studio](https://aistudio.google.com/app/apikey).

## Instalación

Instala las dependencias con:

```bash
pip install -r requirements.txt
```

## Configuración

1. Coloca tu archivo `credentials.json` en la raíz del proyecto.
2. Crea un archivo `.env` en la raíz del proyecto con el siguiente contenido:
   ```
   API_KEY_GOOGLE=tu_api_key_aqui
   SPREADSHEET_ID=tu_id_de_sheet
   RANGE_NAME=A1
   ```
3. No es necesario modificar el script para cambiar el ID de la hoja o el rango, solo actualiza el `.env`.

## Ejecución

```bash
python gmail_word_to_sheets.py
```

La primera vez te pedirá iniciar sesión con tu cuenta de Google y autorizar permisos.

## Funcionamiento

- Busca el correo NO LEÍDO más reciente con un adjunto Word.
- Descarga el adjunto.
- Extrae el texto del Word.
- Envía el texto a Gemini 2.5, que devuelve los siguientes parámetros:
  - REF
  - CONCURSO
  - FECHA Y HORA
  - DESCRIPCION
  - CANTIDAD
- Pega cada parámetro en una columna específica de la hoja de Google Sheets (A: REF, B: CONCURSO, C: FECHA Y HORA, D: DESCRIPCION, E: CANTIDAD).
- Marca el correo como leído para evitar reprocesarlo.

## Automatización recomendada

Para que el script funcione automáticamente cada vez que llegue un correo nuevo:

- Usa el **Programador de tareas de Windows** para ejecutar el script cada 1, 5 o 10 minutos.
- El script solo actuará si hay un correo NO LEÍDO con adjunto Word.

## Notas

- El script solo procesa el correo NO LEÍDO más reciente con adjunto Word.
- El archivo Word descargado se elimina automáticamente después de procesarse.
- La API Key de Gemini se gestiona de forma segura mediante un archivo `.env`.

---

¿Dudas o mejoras? ¡No dudes en preguntar! 