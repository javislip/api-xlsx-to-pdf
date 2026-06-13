# Dockerfile
FROM python:3.11-slim

# Evitar que Python escriba archivos .pyc en el disco
ENV PYTHONDONTWRITEBYTECODE=1
# Evitar que Python almacene en búfer stdout y stderr
ENV PYTHONUNBUFFERED=1

# Instalar dependencias del sistema (LibreOffice headless)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-nogui \
    && rm -rf /var/lib/apt/lists/*

# Establecer el directorio de trabajo en el contenedor
WORKDIR /app

# Copiar el archivo de requerimientos
COPY requirements.txt .

# Instalar dependencias de Python
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el código de la aplicación
COPY main.py .

# Exponer el puerto en el que corre Uvicorn
EXPOSE 8000

# Comando para iniciar la aplicación mediante Uvicorn
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
