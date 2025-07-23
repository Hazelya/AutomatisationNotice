# Étape 1 : image Python avec libs nécessaires à WeasyPrint
FROM python:3.11-slim

# Installation des dépendances système nécessaires pour WeasyPrint et les fonts
RUN apt-get update && apt-get install -y \
    build-essential \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libcairo2 \
    libgdk-pixbuf2.0-0 \
    libffi-dev \
    libgobject-2.0-0 \
    fonts-liberation \
    curl \
    && apt-get clean

# Créer le dossier de travail
WORKDIR /app

# Copier les fichiers
COPY . /app

# Installer les dépendances Python
RUN pip install --no-cache-dir -r requirements.txt

# Port utilisé par Streamlit
EXPOSE 8501

# Lancer Streamlit
CMD ["streamlit", "run", "main.py", "--server.port=8501", "--server.address=0.0.0.0"]
