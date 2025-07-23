# Utilise une image de base Debian avec Python
FROM python:3.11-slim

# Installer les librairies système requises pour WeasyPrint
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

# Définir le dossier de travail
WORKDIR /app

# Copier les fichiers du projet
COPY . .

# Installer les dépendances Python
RUN pip install --no-cache-dir -r requirements.txt

# Exposer le port (pour Streamlit : 8501 ; pour FastAPI via uvicorn : 8000)
EXPOSE 8501

# Lancer l'application (ajuste si ce n'est pas Streamlit)
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]