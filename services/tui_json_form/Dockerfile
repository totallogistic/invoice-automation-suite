FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# deps mínimas
RUN pip install --no-cache-dir textual-serve

# Copiamos server.py (lo demás lo montamos en runtime para iterar rápido)
COPY server.py /app/server.py

EXPOSE 8000
CMD ["python3", "/app/server.py"]