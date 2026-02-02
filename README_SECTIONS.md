# README Sections — Quick Start + Operator Guide (paste-ready)

Copia/pega estas secciones en tu `README.md` principal.

---

## Quick Start

### 1) Requisitos
- Docker + Docker Compose
- Un directorio persistente en host para datos: `/srv/sftpgo/lear_cable/data`

### 2) Config
1. Copia `env.example` → `.env` (NO se comitea)
2. Edita `.env` con tus valores:
   - SMTP (si quieres email)
   - `MAIL_TO` (lista separada por comas)
   - `BATCH_QUIET_SECONDS` (cierre automático de lote)
   - (opcional) puertos publicados

### 3) Levantar stack

> **Multi-stack deployment:** For running multiple environments (pre, prod, etc.) on the same host, see [MULTI_STACK_DEPLOYMENT.md](MULTI_STACK_DEPLOYMENT.md)

Production deployment:
```bash
docker compose --env-file env/common.env --env-file env/prod.env up -d --build
docker ps --format "table {{.Names}}\t{{.Status}}\t{{.Ports}}"
```

### 4) Verificar
```bash
curl -sS http://localhost:8081/api/lear_cable/health
curl -sS http://localhost:8081/tools/lear_cable/ | head
```

### 5) Usar
- UI: `http://<HOST>:8081/tools/lear_cable/`
- API: `POST http://<HOST>:8081/api/lear_cable/batches`
- SFTP: `sftp -P 2222 lear_cable@<HOST>`

---

## Operator Guide

### Logs

> **Note:** For multi-stack deployments, add `--env-file env/common.env --env-file env/{environment}.env` to all docker compose commands.

```bash
docker compose --env-file env/common.env --env-file env/prod.env logs --tail 200 tools_web
docker compose --env-file env/common.env --env-file env/prod.env logs --tail 200 api_lear_cable
docker compose --env-file env/common.env --env-file env/prod.env logs --tail 200 processor_lear_cable
docker compose --env-file env/common.env --env-file env/prod.env logs --tail 200 sftp_lear_cable
```

### Restart / Recreate
- Si cambiaste código (Dockerfile + COPY):
```bash
docker compose --env-file env/common.env --env-file env/prod.env up -d --build --force-recreate api_lear_cable processor_lear_cable
```
- Si cambiaste solo nginx/html (bind-mounted):
```bash
docker compose --env-file env/common.env --env-file env/prod.env up -d --force-recreate tools_web
```

### Datos persistentes (backup recomendado)
Backups mínimos:
- `/srv/sftpgo/lear_cable/data` (inbox/out/status/processed/error)
- `/srv/sftpgo/lear_cable_var` (si guardas provider/db/config de SFTPGo ahí)

Ejemplo tar:
```bash
sudo tar -czf backup_lear_cable_$(date +%Y%m%d).tgz /srv/sftpgo/lear_cable /srv/sftpgo/lear_cable_var
```

### Tamaño máximo de ZIP (Nginx 413)
Si subes ZIPs grandes, en nginx:
```nginx
client_max_body_size 200m;
```
Luego:
```bash
docker compose --env-file env/common.env --env-file env/prod.env up -d --build --force-recreate tools_web
```

### Verificación de mounts
```bash
docker exec -it api_lear_cable sh -lc 'ls -la /data; ls -la /data/inbox /data/out /data/status || true'
docker exec -it processor_lear_cable sh -lc 'ls -la /data; ls -la /data/inbox /data/out /data/status || true'
```
