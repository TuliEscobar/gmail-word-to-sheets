# Gmail a Google Sheets Automático

Este programa revisa automáticamente tu bandeja de Gmail en busca de correos no leídos con archivos adjuntos (Word o PDF), extrae la información importante y la guarda en una hoja de cálculo de Google.

## 📋 Requisitos Previos

Antes de comenzar, necesitarás:

1. Una cuenta de Google (Gmail)
2. Python instalado en tu computadora
3. Conexión a Internet

## 🚀 Instalación Paso a Paso

### 1. Instalar Python
Si no tienes Python instalado:

1. Ve a [python.org](https://www.python.org/downloads/)
2. Descarga la última versión de Python
3. **IMPORTANTE:** Durante la instalación, marca la casilla que dice "Add Python to PATH"
4. Haz clic en "Install Now"

### 2. Descargar los archivos del proyecto

1. Descarga el código del proyecto (deberías tener ya estos archivos si estás leyendo esto)
2. Guarda todos los archivos en una carpeta de tu computadora

### 3. Instalar los programas necesarios

1. Abre el menú Inicio y escribe "CMD"
2. Haz clic derecho en "Símbolo del sistema" y selecciona "Ejecutar como administrador"
3. Copia y pega el siguiente comando y presiona Enter:
   ```
   pip install --upgrade pip
   ```
4. Luego, copia y pega este otro comando (asegúrate de estar en la carpeta del proyecto):
   ```
   cd "ruta\completa\a\tu\carpeta\del\proyecto"
   pip install -r requirements.txt
   ```
   (Reemplaza la ruta con la ubicación real de tus archivos)

### 4. Configurar acceso a Google

1. Sigue las instrucciones para habilitar las APIs de Gmail y Google Sheets en [Google Cloud Console](https://console.cloud.google.com/)
2. Descarga el archivo `credentials.json` y guárdalo en la carpeta del proyecto

### 5. Configurar el archivo .env

1. En la carpeta del proyecto, crea un archivo llamado `.env`
2. Abre el archivo con el Bloc de notas y pega lo siguiente:
   ```
   API_KEY_GOOGLE=tu_api_key_aquí
   SPREADSHEET_ID=el_id_de_tu_hoja_de_cálculo
   RANGE_NAME=A1
   ```
3. Reemplaza `tu_api_key_aquí` con tu clave de API de Google AI Studio
4. Reemplaza `el_id_de_tu_hoja_de_cálculo` con el ID de tu hoja de Google Sheets

## ▶️ Cómo Usar

### Método 1: Usando el archivo por lotes (recomendado)

1. Navega hasta la carpeta del proyecto en el Explorador de archivos
2. Haz doble clic en `run_project.bat`
3. La primera vez, se abrirá una ventana del navegador para que inicies sesión con Google
4. ¡Listo! El programa se ejecutará automáticamente

### Método 2: Usando la línea de comandos

1. Abre el símbolo del sistema (CMD)
2. Navega hasta la carpeta del proyecto:
   ```
   cd "ruta\completa\a\tu\carpeta\del\proyecto"
   ```
3. Ejecuta el programa:
   ```
   python gmail_word_to_sheets.py
   ```

## 🔧 Solución de Problemas

### Error: "invalid_grant: Bad Request"

Si ves este error:
```
Error en ejecución: ('invalid_grant: Bad Request', {'error': 'invalid_grant', 'error_description': 'Bad Request'})
```

Sigue estos pasos:

1. Cierra todas las ventanas del programa si están abiertas
2. Busca y elimina el archivo `token.pickle` en la carpeta del proyecto
3. Vuelve a ejecutar el programa
4. Se abrirá una ventana del navegador para que vuelvas a iniciar sesión con Google

### El programa no encuentra Python

Si al hacer doble clic en `run_project.bat` la ventana se cierra inmediatamente:

1. Abre el Bloc de notas
2. Copia y pega lo siguiente:
   ```
   @echo off
   cd /d "%~dp0"
   python gmail_word_to_sheets.py
   pause
   ```
3. Guarda el archivo como `run_project.bat` (asegúrate de seleccionar "Todos los archivos" en el tipo de archivo)
4. Intenta ejecutarlo de nuevo

## 📞 Soporte

Si tienes problemas o preguntas, por favor contacta al soporte técnico proporcionando:

1. Una descripción del problema
2. Una captura de pantalla del error (si aplica)
3. Los pasos que seguiste antes de que ocurriera el error

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