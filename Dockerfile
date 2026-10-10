# ==============================================================================
# [INICIO SECCIÓN 1 - IMAGEN BASE DE PYTHON]
# (Propósito: Utiliza la versión oficial de Python 3.11 en Linux Debian slim)
# ==============================================================================
FROM python:3.11-slim
# ==============================================================================
# [FIN SECCIÓN 1 - IMAGEN BASE DE PYTHON]
# ==============================================================================

# ==============================================================================
# [INICIO SECCIÓN 2 - CONFIGURACIÓN DEL DIRECTORIO DE TRABAJO]
# (Propósito: Define la carpeta /app como la raíz de nuestro proyecto en el servidor)
# ==============================================================================
WORKDIR /app
# ==============================================================================
# [FIN SECCIÓN 2 - CONFIGURACIÓN DEL DIRECTORIO DE TRABAJO]
# ==============================================================================

# ==============================================================================
# [INICIO SECCIÓN 3 - INSTALACIÓN DE DEPENDENCIAS DEL SISTEMA OPERATIVO]
# (Propósito: Instala librerías nativas de Linux necesarias para gráficos y navegador)
# ==============================================================================
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget \
    gnupg \
    curl \
    ca-certificates \
    libglib2.0-0 \
    libnss3 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libcups2 \
    libdrm2 \
    libxkbcommon0 \
    libxcomposite1 \
    libxdamage1 \
    libxfixes3 \
    libxrandr2 \
    libgbm1 \
    libpango-1.0-0 \
    libcairo2 \
    libasound2 \
    && rm -rf /var/lib/apt/lists/*
# ==============================================================================
# [FIN SECCIÓN 3 - INSTALACIÓN DE DEPENDENCIAS DEL SISTEMA OPERATIVO]
# ==============================================================================

# ==============================================================================
# [INICIO SECCIÓN 4 - INSTALACIÓN DE LIBRERÍAS PYTHON Y PLAYWRIGHT]
# (Propósito: Instala las librerías de requirements.txt y el navegador Chromium para capturas)
# ==============================================================================
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN playwright install chromium
RUN playwright install-deps chromium
# ==============================================================================
# [FIN SECCIÓN 4 - INSTALACIÓN DE LIBRERÍAS PYTHON Y PLAYWRIGHT]
# ==============================================================================

# ==============================================================================
# [INICIO SECCIÓN 5 - COPIA DE CÓDIGO Y COMANDO DE ARRANQUE]
# (Propósito: Copia los archivos del proyecto y ejecuta el servidor FastAPI en el puerto 10000)
# ==============================================================================
COPY . .

EXPOSE 10000

CMD ["uvicorn", "main.py:app", "--host", "0.0.0.0", "--port", "10000"]
# ==============================================================================
# [FIN SECCIÓN 5 - COPIA DE CÓDIGO Y COMANDO DE ARRANQUE]
# ==============================================================================