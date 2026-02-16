FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy all application files
COPY tui_json_form.py .
COPY serve_packing_list.py .
COPY serve_customs_clearance.py .

# Schemas will be mounted from apps/forms
VOLUME /schemas

# Default command
CMD ["python3", "--version"]
