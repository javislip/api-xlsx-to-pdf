# Especificación del Proyecto: Microservicio Conversor XLSX a PDF Seguro

## 1. Contexto y Objetivo
Se requiere construir un microservicio backend ligero para reemplazar una dependencia externa (ConvertAPI). El objetivo principal del servicio es recibir un archivo de Excel (`.xlsx`) a través de una petición HTTP, convertirlo a `.pdf` utilizando LibreOffice en modo *headless*, y devolver el archivo PDF resultante. El proyecto debe estar protegido por una API Key y se desplegará en un servidor propio utilizando Coolify mediante Docker Compose.

## 2. Stack Tecnológico
* **Lenguaje:** Python 3.11+
* **Framework Web:** FastAPI
* **Servidor ASGI:** Uvicorn
* **Motor de Conversión:** LibreOffice (ejecutado mediante subprocesos en modo headless)
* **Contenedorización y Orquestación:** Docker y Docker Compose
* **Seguridad:** Autenticación mediante API Key en el Header de la petición.

## 3. Estructura de Archivos Esperada
El agente debe generar los siguientes cuatro archivos en la raíz del proyecto:
1.  `main.py` (Lógica de la aplicación FastAPI y seguridad)
2.  `requirements.txt` (Dependencias de Python)
3.  `Dockerfile` (Instrucciones de construcción de la imagen)
4.  `docker-compose.yml` (Instrucciones de orquestación y variables de entorno)

## 4. Requisitos Funcionales y Lógica (`main.py`)

### 4.1. Seguridad (API Key)
* Leer una variable de entorno llamada `API_KEY`.
* Implementar una dependencia en FastAPI utilizando `APIKeyHeader` (por ejemplo, esperando el header `X-API-Key`).
* Si el header no coincide con la variable de entorno `API_KEY`, el endpoint debe rechazar la petición con un error `401 Unauthorized`.

### 4.2. Endpoint Principal
* **Método:** `POST`
* **Ruta:** `/convert`
* **Dependencias:** Debe requerir la validación de la API Key.
* **Entrada:** Debe recibir un archivo binario mediante `multipart/form-data`. El nombre del campo en el formulario DEBE ser exactamente `file`.
* **Salida:** Un objeto `FileResponse` de FastAPI con el `media_type="application/pdf"` y el nombre original del archivo con la extensión cambiada a `.pdf`.

### 4.3. Flujo de Procesamiento
1.  **Recepción:** Recibir el archivo y crear un directorio temporal seguro (usando `tempfile`).
2.  **Escritura:** Guardar el archivo `.xlsx` subido dentro de ese directorio temporal.
3.  **Conversión:** Invocar LibreOffice a través de `subprocess.run()`. Argumentos: `libreoffice --headless --invisible --nologo --convert-to pdf --outdir <directorio_temporal> <ruta_archivo_xlsx>`.
4.  **Validación:** Comprobar si el archivo `.pdf` se generó. Si falla, devolver un error JSON (`500 Internal Server Error`) con los detalles y limpiar la carpeta temporal inmediatamente.
5.  **Limpieza Diferida:** Utilizar `BackgroundTasks` de FastAPI para programar una función que elimine recursivamente todo el directorio temporal **después** de que FastAPI haya terminado de transmitir el archivo al cliente.

## 5. Requisitos de Dependencias (`requirements.txt`)
Las dependencias estrictamente necesarias son:
* `fastapi`
* `uvicorn`
* `python-multipart`

No incluir bases de datos ni ORMs. El microservicio debe ser *stateless*.

## 6. Requisitos de Contenedorización (`Dockerfile`)
* **Imagen Base:** Usar `python:3.11-slim` (o similar basada en Debian).
* **Instalación del Sistema:** Ejecutar `apt-get update` e instalar `libreoffice` con `--no-install-recommends`. Limpiar la caché de apt (`rm -rf /var/lib/apt/lists/*`).
* **Entorno Python:** Copiar `requirements.txt`, instalar dependencias mediante `pip install --no-cache-dir`.
* **Ejecución:** Exponer el puerto `8000` y definir el comando CMD para ejecutar Uvicorn en `0.0.0.0:8000`.

## 7. Requisitos de Orquestación (`docker-compose.yml`)
* Definir un servicio llamado `api` o `pdf-converter`.
* Configurar el `build: .` para que construya la imagen desde el Dockerfile local.
* Mapear el puerto `8000:8000`.
* Inyectar la variable de entorno `API_KEY` (puede leerse de un `.env` local o definirse directamente, dejando un valor de marcador de posición como `tu_super_secreto_aqui` para que el usuario lo cambie en producción).
* Configurar la política de reinicio a `unless-stopped` o `always`.

## 8. Instrucciones Finales para el Agente
Genera el código completo para los 4 archivos respetando estas directrices. El código debe ser robusto, manejar errores de subprocesos y de falta de variables de entorno de forma elegante, e incluir comentarios descriptivos.