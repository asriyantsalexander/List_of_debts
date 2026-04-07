FROM python:3.11-slim

# Install system dependencies (unrar, unzip, 7z)
RUN apt-get update && apt-get install -y \
    unrar-free \
    unzip \
    p7zip-full \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE $PORT

CMD gunicorn --bind 0.0.0.0:$PORT --timeout 300 --workers 2 app:app
