# Używamy standardowego obrazu Pythona
FROM python:3.10-slim

# Instalujemy LibreOffice (krytyczne dla konwersji PDF)
RUN apt-get update && apt-get install -y libreoffice && apt-get clean

# Ustawiamy katalog roboczy
WORKDIR /app

# Kopiujemy pliki projektu
COPY . /app

# Instalujemy zależności Pythonowe (z requirements.txt)
RUN pip install --no-cache-dir -r requirements.txt

# Definiujemy komendę uruchamiającą Gunicorn
CMD gunicorn --bind 0.0.0.0:$PORT app:app