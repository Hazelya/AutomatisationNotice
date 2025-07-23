FROM python:3.11-slim

# Installer les dépendances système pour WeasyPrint
RUN apt-get update && apt-get install -y \
    build-essential \
    libpango1.0-0 \
    libpangocairo-1.0-0 \
    libcairo2 \
    libgdk-pixbuf2.0-0 \
    libffi-dev \
    libgobject-2.0-0 \
    libglib2.0-0 \
    curl \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . .

RUN pip install --upgrade pip && pip install -r requirements.txt

EXPOSE 8000
CMD ["streamlit", "run", "app.py", "--server.port=8000", "--server.address=0.0.0.0"]
