FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# Dependencias del extractor + watcher
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

# Watcher
COPY watcher.py /app/watcher.py
COPY watcher_import_partida.py /app/watcher_import_partida.py


CMD ["python3", "/app/watcher.py"]
