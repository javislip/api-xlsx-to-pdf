import logging
import os
import subprocess
import tempfile
import shutil
import uuid
import time
import re
import zipfile
import xml.etree.ElementTree as ET
from fastapi import FastAPI, File, UploadFile, HTTPException, Depends, status, BackgroundTasks, Request
from fastapi.security import APIKeyHeader
from fastapi.responses import FileResponse

# Configurar el registro de logs
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("xlsx-to-pdf-api")

app = FastAPI(
    title="Microservicio Conversor XLSX a PDF",
    description="Convierte archivos XLSX a PDF de forma segura usando LibreOffice headless y enlaces de descarga temporal.",
    version="1.1.0"
)

# Directorio donde se guardarán las descargas temporales públicas
DOWNLOADS_DIR = "/app/downloads"
os.makedirs(DOWNLOADS_DIR, exist_ok=True)

# Definir cabecera para autenticación
API_KEY_NAME = "X-API-Key"
api_key_header = APIKeyHeader(name=API_KEY_NAME, auto_error=False)

# Verificar la variable de entorno en el inicio
API_KEY = os.environ.get("API_KEY")
if not API_KEY:
    logger.warning("ALERTA: La variable de entorno API_KEY no está configurada. Las solicitudes de conversión fallarán con un error 500.")

async def validate_api_key(api_key: str = Depends(api_key_header)):
    """
    Valida la API Key provista en la cabecera X-API-Key.
    Si la variable de entorno API_KEY no está configurada, devuelve un error 500.
    Si la API Key es incorrecta o falta, devuelve un error 401.
    """
    server_api_key = os.environ.get("API_KEY")
    if not server_api_key:
        logger.error("Error de configuración: API_KEY no definida en el servidor.")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Error de configuración del servidor: API_KEY no configurada."
        )
    if not api_key or api_key != server_api_key:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="No autorizado: API Key faltante o inválida."
        )
    return api_key

def cleanup_directory(path: str):
    """
    Elimina un directorio temporal de manera recursiva.
    """
    if os.path.exists(path):
        try:
            shutil.rmtree(path)
            logger.info(f"Directorio temporal de conversión eliminado: {path}")
        except Exception as e:
            logger.error(f"Error al eliminar el directorio temporal {path}: {e}")

def cleanup_expired_files(directory: str, max_age_seconds: int = 900):
    """
    Escanea el directorio de descargas y elimina los archivos que superen
    el tiempo de vida configurado (por defecto, 15 minutos / 900 segundos).
    """
    now = time.time()
    try:
        if os.path.exists(directory):
            deleted_count = 0
            for filename in os.listdir(directory):
                file_path = os.path.join(directory, filename)
                if os.path.isfile(file_path):
                    file_age = now - os.path.getmtime(file_path)
                    if file_age > max_age_seconds:
                        os.remove(file_path)
                        deleted_count += 1
                        logger.info(f"Archivo de descarga expirado eliminado: {file_path}")
            if deleted_count > 0:
                logger.info(f"Limpieza completada. Se eliminaron {deleted_count} archivos expirados.")
    except Exception as e:
        logger.error(f"Error al limpiar descargas expiradas en {directory}: {e}")

def optimize_xlsx_layout_zip(file_path: str):
    """
    Modifica el archivo .xlsx directamente manipulando el ZIP y los archivos XML
    para aplicar configuraciones de impresión sin perder imágenes o gráficos:
    - PageOrientation: portrait (vertical)
    - PageSize: letter (carta, ID 1)
    - AutoPageFit: ajustar columnas a 1 página (fitToWidth=1, fitToHeight=0)
    - ClearPrintArea: borrar áreas de impresión del workbook
    """
    _, ext = os.path.splitext(file_path)
    if ext.lower() != ".xlsx":
        logger.info(f"El archivo {file_path} no es .xlsx, se omitirá el formateo de página XML.")
        return

    # Registrar namespaces estándar de OpenXML para mantener consistencia en etiquetas al guardar
    namespaces = {
        "": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
        "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
        "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
        "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    }
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)

    temp_fd, temp_path = tempfile.mkstemp()
    os.close(temp_fd)

    try:
        logger.info(f"Ajustando XML de impresión en el ZIP para {file_path}")
        with zipfile.ZipFile(file_path, 'r') as yin:
            with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as yout:
                for item in yin.infolist():
                    data = yin.read(item.filename)
                    
                    # 1. Modificar hojas de cálculo (worksheets)
                    if item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                        try:
                            root = ET.fromstring(data)
                            ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
                            
                            # A) Habilitar fitToPage en <sheetPr> -> <pageSetUpPr fitToPage="1"/>
                            sheetPr = root.find(f"{ns}sheetPr")
                            if sheetPr is None:
                                sheetPr = ET.Element(f"{ns}sheetPr")
                                root.insert(0, sheetPr)
                            
                            pageSetUpPr = sheetPr.find(f"{ns}pageSetUpPr")
                            if pageSetUpPr is None:
                                pageSetUpPr = ET.Element(f"{ns}pageSetUpPr")
                                sheetPr.append(pageSetUpPr)
                            pageSetUpPr.set("fitToPage", "1")
                            
                            # B) Configurar orientación, tamaño carta y escala de página en <pageSetup>
                            pageSetup = root.find(f"{ns}pageSetup")
                            if pageSetup is None:
                                pageSetup = ET.Element(f"{ns}pageSetup")
                                root.append(pageSetup)
                            
                            pageSetup.set("orientation", "portrait")
                            pageSetup.set("paperSize", "1")  # Letter
                            pageSetup.set("fitToWidth", "1")
                            pageSetup.set("fitToHeight", "0") # Alto dinámico (fluye hacia abajo)
                            
                            # Eliminar atributo 'scale' si existe para evitar que anule fitToWidth/fitToHeight
                            if "scale" in pageSetup.attrib:
                                del pageSetup.attrib["scale"]
                                
                            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                        except Exception as ex:
                            logger.error(f"Error al modificar worksheet XML {item.filename}: {ex}")
                    
                    # 2. Modificar workbook para limpiar áreas de impresión (ClearPrintArea)
                    elif item.filename == "xl/workbook.xml":
                        try:
                            root = ET.fromstring(data)
                            ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
                            
                            # Buscar <definedNames>
                            definedNames = root.find(f"{ns}definedNames")
                            if definedNames is not None:
                                print_areas = [
                                    dn for dn in definedNames 
                                    if dn.get("name") == "_xlnm.Print_Area"
                                ]
                                for pa in print_areas:
                                    definedNames.remove(pa)
                                    logger.info("Área de impresión eliminada en workbook.xml")
                                
                                if len(definedNames) == 0:
                                    root.remove(definedNames)
                                    
                            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                        except Exception as ex:
                            logger.error(f"Error al modificar workbook.xml: {ex}")
                            
                    yout.writestr(item.filename, data)
                    
        # Reemplazar el archivo original con el modificado
        shutil.move(temp_path, file_path)
        logger.info(f"Modificación del ZIP completada exitosamente para {file_path}")
    except Exception as e:
        logger.error(f"Error general al modificar el archivo ZIP {file_path}: {e}")
        if os.path.exists(temp_path):
            os.remove(temp_path)

@app.get("/health", status_code=status.HTTP_200_OK)
def health_check():
    """
    Endpoint de salud simple para verificar que la API está respondiendo.
    """
    return {
        "status": "healthy",
        "api_key_configured": os.environ.get("API_KEY") is not None
    }

@app.post("/convert")
async def convert_xlsx_to_pdf(
    request: Request,
    background_tasks: BackgroundTasks,
    _api_key: str = Depends(validate_api_key)
):
    """
    Endpoint para convertir un archivo XLSX a PDF.
    - Soporta tanto multipart/form-data (campos 'file' o 'data') como envíos binarios en bruto (raw body).
    - Requiere la cabecera 'X-API-Key'.
    - Guarda el PDF resultante en una ruta pública temporal.
    - Retorna un JSON con la URL de descarga.
    - Ejecuta la limpieza de archivos de descarga antiguos en segundo plano.
    """
    content_type = request.headers.get("content-type", "")
    
    uploaded_bytes = b""
    filename = "document.xlsx"

    # Caso A: Carga estándar multipart/form-data (ej: curl -F "file=@archivo.xlsx")
    if "multipart/form-data" in content_type:
        try:
            form = await request.form()
            form_file = form.get("file") or form.get("data")
            if not form_file or not hasattr(form_file, "filename") or not form_file.filename:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="No se ha proporcionado ningún archivo en los campos 'file' o 'data'."
                )
            filename = os.path.basename(form_file.filename or "document.xlsx")
            uploaded_bytes = await form_file.read()
            await form_file.close()
        except Exception as e:
            if isinstance(e, HTTPException):
                raise e
            logger.error(f"Error al procesar petición multipart: {e}")
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Error al procesar los datos de formulario: {str(e)}"
            )
            
    # Caso B: Envío binario directo en el cuerpo de la petición (ej: n8n Binary File)
    else:
        try:
            uploaded_bytes = await request.body()
            if not uploaded_bytes:
                raise HTTPException(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    detail="El cuerpo de la petición está vacío y no se envió ningún archivo."
                )
            
            # Determinar extensión por content-type o por defecto xlsx
            ext = ".xlsx"
            if "excel" in content_type or "ms-excel" in content_type:
                ext = ".xls"
            filename = f"document{ext}"
            
            # Intentar obtener el nombre original de la cabecera Content-Disposition si existe
            content_disposition = request.headers.get("content-disposition", "")
            if "filename=" in content_disposition:
                match = re.search(r'filename="?([^";]+)"?', content_disposition)
                if match:
                    filename = os.path.basename(match.group(1))
        except Exception as e:
            if isinstance(e, HTTPException):
                raise e
            logger.error(f"Error al leer el cuerpo binario de la petición: {e}")
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="Error interno al leer el cuerpo binario de la petición."
            )

    # Validar extensión final del archivo
    base_name, extension = os.path.splitext(filename)
    if extension.lower() not in [".xlsx", ".xls"]:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Extensión no permitida: {extension}. Solo se permiten archivos .xlsx o .xls."
        )

    # 2. Crear un directorio temporal seguro para la conversión
    temp_dir = tempfile.mkdtemp()
    input_path = os.path.join(temp_dir, filename)

    # 3. Guardar el archivo cargado en el directorio temporal
    try:
        with open(input_path, "wb") as buffer:
            buffer.write(uploaded_bytes)
        logger.info(f"Archivo cargado ({len(uploaded_bytes)} bytes) y guardado temporalmente en: {input_path}")
    except Exception as e:
        cleanup_directory(temp_dir)
        logger.error(f"Error al guardar el archivo cargado: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Error al procesar el archivo subido en el servidor."
        )

    # 3.5. Formatear y optimizar el diseño del archivo XLSX antes de la conversión sin perder recursos
    optimize_xlsx_layout_zip(input_path)

    # 4. Ejecutar la conversión mediante LibreOffice
    cmd = [
        "libreoffice",
        "--headless",
        "--invisible",
        "--nologo",
        "--convert-to",
        "pdf",
        "--outdir",
        temp_dir,
        input_path
    ]

    try:
        logger.info(f"Iniciando conversión de LibreOffice: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60, check=True)
        logger.info("Conversión de LibreOffice ejecutada exitosamente.")
    except subprocess.TimeoutExpired as e:
        cleanup_directory(temp_dir)
        logger.error("La conversión de LibreOffice superó el límite de tiempo.")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="La conversión del archivo tardó demasiado tiempo y fue cancelada."
        )
    except subprocess.CalledProcessError as e:
        cleanup_directory(temp_dir)
        logger.error(f"LibreOffice retornó código de error {e.returncode}. Stderr: {e.stderr}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error al realizar la conversión del archivo: {e.stderr or e.stdout}"
        )
    except Exception as e:
        cleanup_directory(temp_dir)
        logger.error(f"Error inesperado al ejecutar LibreOffice: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Error inesperado durante la ejecución del conversor."
        )

    # 5. Comprobar si el archivo PDF se generó
    pdf_filename = f"{base_name}.pdf"
    pdf_path = os.path.join(temp_dir, pdf_filename)

    if not os.path.exists(pdf_path):
        cleanup_directory(temp_dir)
        logger.error(f"El archivo PDF esperado no fue encontrado en: {pdf_path}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="La conversión se ejecutó pero el archivo PDF no se pudo generar correctamente."
        )

    # 6. Mover el archivo PDF a la ruta de descargas públicas temporales con un UUID para evitar colisiones
    unique_file_id = f"{uuid.uuid4()}_{pdf_filename}"
    public_pdf_path = os.path.join(DOWNLOADS_DIR, unique_file_id)
    
    try:
        shutil.move(pdf_path, public_pdf_path)
        logger.info(f"Archivo PDF movido a descargas públicas: {public_pdf_path}")
    except Exception as e:
        logger.error(f"Error al mover el archivo PDF a la carpeta de descargas: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Error interno al guardar el archivo convertido en el servidor."
        )
    finally:
        # Limpiar el directorio temporal de conversión inmediatamente
        cleanup_directory(temp_dir)

    # 7. Programar tarea de limpieza de descargas expiradas (más de 15 minutos de antigüedad)
    background_tasks.add_task(cleanup_expired_files, DOWNLOADS_DIR)

    # 8. Retornar JSON con la URL de descarga
    base_url = os.environ.get("BASE_URL", "http://localhost:8000").rstrip("/")
    download_url = f"{base_url}/download/{unique_file_id}"
    
    logger.info(f"Proceso completado. Enlace de descarga generado: {download_url}")
    return {"download_url": download_url}

@app.get("/download/{file_id}", response_class=FileResponse)
def download_pdf(file_id: str):
    """
    Endpoint para descargar un archivo PDF previamente convertido.
    - Recibe el 'file_id' que incluye el UUID.
    - Valida que el archivo exista.
    - Retorna el archivo con su nombre original (sin el prefijo UUID).
    """
    # Prevenir ataques de directory traversal sanitizando el nombre
    sanitized_file_id = os.path.basename(file_id)
    file_path = os.path.join(DOWNLOADS_DIR, sanitized_file_id)
    
    if not os.path.exists(file_path):
        logger.warning(f"Intento de descarga fallido. Archivo no encontrado: {file_path}")
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="El archivo no existe, ha expirado o el enlace es incorrecto."
        )
    
    # Extraer el nombre original del archivo quitando el UUID y el guion bajo
    # El formato del nombre es: <UUID-36-chars>_<original_name>
    original_filename = sanitized_file_id[37:] if len(sanitized_file_id) > 37 else sanitized_file_id
    
    logger.info(f"Sirviendo descarga del archivo: {file_path} con nombre original: {original_filename}")
    return FileResponse(
        path=file_path,
        media_type="application/pdf",
        filename=original_filename
    )
