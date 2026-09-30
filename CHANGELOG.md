# Registro de Cambios

## [1.0.0] - 2026-09-29

### Agregado
- Conversión de Word a PDF usando `reportlab` con soporte para texto, tablas e imágenes
- Conversión de PDF a Word usando `pdf2docx` (local) y `LibreOffice` (producción)
- Interfaz web con Flask para subir archivos y descargar resultados
- Soporte para caracteres Unicode (ñ, á, é, í, ó, ú, etc.)
- Detección automática de entorno (local vs producción) para elegir el mejor método de conversión
- Archivo `README.md` con instrucciones de instalación y uso
- Archivo `Procfile` para despliegue en Render
- Archivo `CHANGELOG.md` con el registro de cambios

### Corregido
- Error con fpdf2 y fuentes personalizadas → Se cambió a `reportlab`
- Error con caracteres Unicode en PDF → Se implementó soporte con `xml.sax.saxutils.escape`
- Error con tablas no apareciendo en PDF → Se agregó procesamiento de tablas con `reportlab.Table`
- Error con imágenes no apareciendo en PDF → Se agregó extracción de imágenes usando `python-docx` y `BytesIO`
- Error con directorio temporal `temp_images` → Se eliminó la necesidad de carpetas temporales
- Error con `pdf2docx` distorsionando PDFs complejos → Se agregó `LibreOffice` como alternativa en producción

### Cambios técnicos
- Se reemplazó `docx2pdf` por `reportlab` para evitar dependencia de Microsoft Word
- Se reemplazó `fpdf2` por `reportlab` para mejor soporte Unicode
- Se implementó detección de entorno con variable `RENDER` para usar `LibreOffice` en producción
- Se agregó `gunicorn` para el servidor de producción en Render
