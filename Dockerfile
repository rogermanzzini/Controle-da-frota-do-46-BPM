FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=5000 \
    FLASK_ENV=production

WORKDIR /app

# Instalar dependências de sistema para compilação se necessário
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Instalar dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar código do projeto
COPY . .

# Criar pastas persistentes de dados e uploads
RUN mkdir -p uploads/manutencao/viaturas relatorios_gerados exemplos

EXPOSE 5000

CMD ["gunicorn", "--workers=4", "--bind=0.0.0.0:5000", "wsgi:application"]
