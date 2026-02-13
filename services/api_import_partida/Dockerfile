FROM python:3.12-alpine
WORKDIR /app
RUN apk add --no-cache tzdata
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app /app/app
ENV PYTHONUNBUFFERED=1
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
