FROM python:3.11-slim

# Instala LibreOffice para conversão Excel → PDF
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-calc \
    fonts-liberation \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Instala dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia a aplicação
COPY app.py .

# Cria pasta para templates
RUN mkdir -p /app/templates

# Expõe a porta
EXPOSE 5000

CMD ["python", "app.py"]
