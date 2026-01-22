# Imagem base Python enxuta
FROM python:3.11-slim

# Instala LibreOffice para o docx2pdf funcionar
RUN apt-get update && \
    apt-get install -y libreoffice && \
    rm -rf /var/lib/apt/lists/*

# Diretório de trabalho
WORKDIR /app

# Copia tudo do projeto para dentro do container
COPY . .

# Instala as dependências Python
RUN pip install --no-cache-dir -r requirements.txt

# Comando para iniciar o app no Render
# Render passa a porta na variável de ambiente $PORT
CMD gunicorn app:app --bind 0.0.0.0:$PORT
