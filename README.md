# Microservicio Conversor XLSX a PDF Seguro (Optimizado para n8n)

Este es un microservicio backend ligero y *stateless* desarrollado con **FastAPI** y **LibreOffice Headless** que permite convertir archivos de Excel (`.xlsx` o `.xls`) a formato PDF de forma segura. 

El servicio ha sido especialmente diseñado para su integración con plataformas de automatización como **n8n**, resolviendo los problemas comunes de parseo de flujos binarios y formateo de hojas de cálculo.

---

## 🚀 Características Clave

1.  **Seguridad por API Key:** Todas las peticiones de conversión están protegidas mediante la validación de una clave en la cabecera HTTP `X-API-Key`.
2.  **Preservación de Recursos y Formato (Edición Directa de XML):** 
    Para evitar que librerías como `openpyxl` destruyan los logotipos, firmas de firmas y gráficos embebidos al guardar, este microservicio procesa el archivo `.xlsx` directamente como un archivo comprimido (`ZIP`) usando la biblioteca estándar de Python. Modifica de forma quirúrgica los ficheros XML internos de configuración de impresión sin alterar las relaciones de imágenes ni los recursos multimedia (`/xl/media`).
3.  **Ajuste de Diseño de Página Automático:**
    *   **PageOrientation:** Configura de forma automática la orientación vertical (`portrait`).
    *   **PageSize:** Ajusta el tamaño del papel a tamaño Carta (`Letter` / ID 1).
    *   **AutoPageFit:** Escala el contenido para que todas las columnas quepan perfectamente a lo ancho de **una sola página**, permitiendo que las filas fluyan verticalmente de forma dinámica.
    *   **ClearPrintArea:** Limpia cualquier área de impresión previamente guardada en el Excel para garantizar que se exporten todas las hojas por completo.
4.  **Flujo Desacoplado para n8n:**
    *   Soporta la subida de archivos como formularios standard (`multipart/form-data`) y cargas directas de binarios en bruto (`raw request body`) que utiliza n8n por defecto al activar `"Body Content Type: n8n Binary File"`.
    *   En lugar de retornar los bytes del PDF directamente en la petición `/convert` (lo cual genera un error de serialización JSON circular en n8n), la API almacena el PDF en un directorio de descargas temporales y responde con un JSON que contiene un enlace único:
        ```json
        {
          "download_url": "http://<BASE_URL>/download/<uuid>_<nombre_archivo>.pdf"
        }
        ```
5.  **Limpieza Automática Diferida:**
    Cada petición de conversión ejecuta una tarea en segundo plano (`BackgroundTasks`) que escanea el directorio de descargas públicas y elimina automáticamente los archivos cuya antigüedad sea mayor a **15 minutos** (900 segundos).

---

## 🛠️ Stack Tecnológico

*   **Python 3.11+**
*   **FastAPI** & **Uvicorn**
*   **LibreOffice** (modo headless en el contenedor)
*   **Docker** & **Docker Compose**

---

## 📋 Variables de Entorno

Puedes configurar estas variables en el host, en un archivo `.env` o en la interfaz de Coolify:

| Variable | Descripción | Valor por defecto |
| :--- | :--- | :--- |
| `API_KEY` | Clave secreta necesaria para autorizar las solicitudes en la cabecera `X-API-Key`. | `tu_super_secreto_aqui` |
| `BASE_URL` | Dominio o dirección IP pública del servidor que se usará para construir los enlaces de descarga. | `http://localhost:8000` |
| `HOST_PORT` | Puerto físico expuesto en el servidor host. | `8000` |

---

## 💻 Instalación y Despliegue Local

### Requisitos Previos
*   Tener **Docker** y **Docker Compose** instalados y ejecutándose en el sistema.

### Levantar el servicio
Ejecuta el siguiente comando en la raíz del proyecto para construir la imagen y levantar el contenedor en segundo plano:
```bash
docker compose up --build -d
```

Una vez iniciado, puedes verificar el estado de salud accediendo a:
[http://localhost:8000/health](http://localhost:8000/health)

---

## 🧪 Ejemplos de Uso (curl)

### 1. Conversión de archivo (Multipart / Form-Data)
```bash
curl -X POST http://localhost:8000/convert \
  -H "X-API-Key: tu_super_secreto_aqui" \
  -F "data=@mi_archivo.xlsx"
```

### 2. Conversión de archivo (Binary Raw Body - n8n default)
```bash
curl -X POST http://localhost:8000/convert \
  -H "X-API-Key: tu_super_secreto_aqui" \
  -H "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" \
  --data-binary "@mi_archivo.xlsx"
```

**Respuesta exitosa (HTTP 200):**
```json
{
  "download_url": "http://localhost:8000/download/1fe14709-bf51-46b7-b040-2ae6715b081e_mi_archivo.pdf"
}
```

### 3. Descarga del PDF
Haciendo una petición GET simple a la URL del JSON:
```bash
curl http://localhost:8000/download/1fe14709-bf51-46b7-b040-2ae6715b081e_mi_archivo.pdf --output mi_archivo.pdf
```

---

## ☁️ Despliegue en Producción (Coolify)

Este microservicio está optimizado para Coolify:
1. Crea un recurso de tipo **Application** en Coolify.
2. Apúntalo al repositorio Git del proyecto.
3. Elige **Docker Compose** como el tipo de build.
4. En el panel de variables de entorno, configura `API_KEY` y `BASE_URL` (este último con el dominio público asignado por Coolify, ej. `https://excel-pdf.tudominio.com`).
5. Coolify detectará el archivo expuesto en el `Dockerfile` y enrutará el tráfico de forma automática.
