# Usage — Lear Cable (CLI / SFTP / API / UI)

> **Note:** This guide assumes a single-stack deployment. For multi-stack deployments (pre, prod, dev on the same host), see [MULTI_STACK_DEPLOYMENT.md](MULTI_STACK_DEPLOYMENT.md) for details on using environment-specific configuration files.
>
> All `docker compose` commands in this guide should include the environment files:
> ```bash
> docker compose --env-file env/common.env --env-file env/prod.env [command]
> ```

Este stack soporta 4 modos de entrada:
1) UI Web (upload ZIP)
2) API (upload ZIP)
3) SFTP (drop PDFs)
4) CLI local (ejecutar extractor manual)

**Salida**
- `invoices_extracted.json`
- `invoices_extracted.csv`
- `invoices_extracted.xlsx`
- Email (modo batch) a los recipients configurados

---

## 0) Pre-requisitos

Servicios levantados:
```bash
docker ps --format "table {{.Names}}\t{{.Status}}\t{{.Ports}}"
```

Health:
```bash
curl -sS http://localhost:8081/api/lear_cable/health
```

---

## 1) UI Web (recomendado para cliente)

URL:
- `http://<HOST>:8081/tools/lear_cable/`

Flujo:
1) Subir un `ZIP` con PDFs dentro
2) La UI obtiene un `batch_id`
3) La UI hace polling del `status`
4) Al finalizar, se muestra:
   - estado final (DONE/ERROR)
   - progress (0..100%)
   - recipients a los que se envió email

Notas:
- El ZIP debe contener PDFs.
- Si hay `CMR.pdf`, se ignora (según la regla actual del extractor).
- Si sale `413 Request Entity Too Large`, sube `client_max_body_size` en nginx.

---

## 2) API (upload ZIP)

### 2.1 Subir ZIP
```bash
curl -sS \
  -F "file=@lote_demo.zip;type=application/zip" \
  http://<HOST>:8081/api/lear_cable/batches | jq
```

Respuesta típica:
```json
{
  "batch_id": "20260126_132826_sxtn",
  "extracted_pdfs": 5,
  "status_url": "/api/lear_cable/batches/20260126_132826_sxtn/status"
}
```

### 2.2 Consultar status
```bash
B=20260126_132826_sxtn
curl -sS http://<HOST>:8081/api/lear_cable/batches/$B/status | jq
```

Campos relevantes:
- `state`: `UPLOADED | PROCESSING | DONE | ERROR`
- `processed_files / total_files`
- `recipients` (cuando termina)
- `message` (texto humano)

---

## 3) SFTP (drop PDFs)

### 3.1 Conectar
```bash
sftp -P 2222 lear_cable@<HOST>
```

Si tienes que forzar password:
```bash
sftp \
  -o PreferredAuthentications=password \
  -o PubkeyAuthentication=no \
  -o IdentitiesOnly=yes \
  -P 2222 \
  lear_cable@<HOST>
```

### 3.2 Layout remoto
Normalmente verás:
- `inbox/`
- `processing/`
- `out/`
- `processed/`
- `error/`
- `status/`

### 3.3 Subir PDFs como lote
Recomendado: crea un folder de lote y sube ahí los PDFs:
```sftp
mkdir inbox/mi_lote_001
put DM*.pdf inbox/mi_lote_001/
```

El processor cerrará el lote por “quiet time” (inactividad) y lo procesará.

---

## 4) CLI (extractor manual)

### 4.1 Ejecutar extractor
```bash
python3 apps/lear_cable/extractor/extract_lear_fields.py <inputs...> -o <output_dir>
```

Ejemplo (folder):
```bash
python3 apps/lear_cable/extractor/extract_lear_fields.py ./samples/lote1 -o ./out/lote1
```

Salida:
- `invoices_extracted.json`
- `invoices_extracted.csv`
- `invoices_extracted.xlsx`

### 4.2 Validación rápida
```bash
jq '.[0]' ./out/lote1/invoices_extracted.json
head -n 5 ./out/lote1/invoices_extracted.csv
ls -lah ./out/lote1/invoices_extracted.xlsx
```

---

## 5) Email (batch)

Variables relevantes:
- `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`
- `MAIL_FROM`
- `MAIL_TO` (lista separada por comas)
- `EMAIL_MODE=BATCH_ONLY`

El watcher envía los 3 ficheros como adjuntos al finalizar el lote.

---

## 6) Dónde quedan los ficheros en host

En el host (máquina):
- `/srv/sftpgo/lear_cable/data/out/<batch_id>/invoices_extracted.*`
- `/srv/sftpgo/lear_cable/data/status/<batch_id>/status.json`
- `/srv/sftpgo/lear_cable/data/error/<batch_id>/...` si falla
