# invoice-automation-suite (staging-lab)

Suite Docker para automatizar la extracción de datos de facturas (Lear Cable) y entregar resultados por email.
Incluye:
- **API** (FastAPI) para subir ZIPs y consultar estado.
- **Processor** (watcher) que procesa lotes (extractor) y envía emails.
- **Web/UI** (nginx estático) + reverse proxy a la API.
- **SFTPGo** (opcional) para entradas por SFTP / administración web.

> Stack pensado para LAN primero. Luego se expondrá externamente desde un servidor dedicado.

---

## Arquitectura (alto nivel)

Cliente:
- Subida por **API** (ZIP) → `POST /api/lear_cable/batches`
- (Opcional) Subida por **SFTP** a `inbox/` + marker `_DONE`

Servidor:
- `api_lear_cable` escribe lote en `/data/inbox/...`
- `processor_lear_cable` detecta lote listo, lo mueve a `processing/`, ejecuta extractor, genera outputs en `out/`,
  actualiza status y envía email con adjuntos.

Rutas en disco (host):
- Data root: `/srv/sftpgo/lear_cable/data`
  - `inbox/`, `processing/`, `out/`, `processed/`, `error/`, `status/`

---

## Servicios y puertos (LAN)

- **UI + reverse proxy (nginx)**: `http://<host>:8081/`
  - UI Lear Cable: `http://<host>:8081/tools/lear_cable/`
  - API: `http://<host>:8081/api/lear_cable/...`
- **SFTPGo Admin UI**: `http://<host>:8080/`
- **SFTP**: `sftp://<host>:2222` (SFTPGo expone 2022 en el contenedor)

---

## Requisitos

- Docker + Docker Compose
- (Recomendado) `jq` para ver JSON bonito en CLI

---

## Quickstart (LAN)

1) Crea tu `.env` (NO se commitea):
```bash
cp env.example .env
# edita .env con tus credenciales SMTP y settings
