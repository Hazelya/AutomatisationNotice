FROM python:3.11-slim

# Installer les dépendances système nécessaires à WeasyPrint
RUN apt-get update && apt-get install -y \
    build-essential \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libcairo2 \
    libgdk-pixbuf2.0-0 \
    libffi-dev \
    libgobject-2.0-0 \
    curl \
    && apt-get clean

# Définir le répertoire de travail
WORKDIR /app

# Copier tous les fichiers dans l'image Docker
COPY . .

# Installer les dépendances Python
RUN pip install --upgrade pip && pip install -r requirements.txt

# Démarrer l'application Streamlit
CMD ["streamlit", "run", "app.py", "--server.port=8000", "--server.address=0.0.0.0"]
