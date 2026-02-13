FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY processor.py tool_registry.py /app/

CMD ["python3", "processor.py"]
