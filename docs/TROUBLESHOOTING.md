# Troubleshooting — Invoice Automation Suite (Lear Cable)

Este documento cubre troubleshooting del stack Docker:
- `tools_web` (nginx static + reverse proxy `/api/`)
- `api_lear_cable` (FastAPI/Uvicorn)
- `processor_lear_cable` (watcher/processor)
- `sftp_lear_cable` (SFTPGo)

> Regla de oro: identifica si el fallo es **red/proxy**, **API**, **processor**, o **filesystem/volúmenes**.

---

## 0) Comandos básicos

### Ver estado y puertos
```bash
docker ps --format "table {{.Names}}\t{{.Status}}\t{{.Ports}}"
sudo ss -lntup | sed -n '1,200p'
```

### Ver logs por servicio
```bash
docker compose logs --tail 200 tools_web
docker compose logs --tail 200 api_lear_cable
docker compose logs --tail 200 processor_lear_cable
docker compose logs --tail 200 sftp_lear_cable
```

### Logs “en vivo”
```bash
docker compose logs -f api_lear_cable
docker compose logs -f processor_lear_cable
```

### Entrar a un contenedor
```bash
docker exec -it processor_lear_cable sh
docker exec -it tools_web sh
docker exec -it api_lear_cable sh
docker exec -it sftp_lear_cable sh
```

---

## 1) Health checks rápidos

### Web y proxy
```bash
curl -i http://localhost:8081/ | sed -n '1,25p'
curl -i http://localhost:8081/tools/lear_cable/ | sed -n '1,25p'
curl -i http://localhost:8081/api/lear_cable/health | sed -n '1,25p'
```

### API directo (desde dentro de nginx)
```bash
docker exec -it tools_web sh -lc 'wget -qO- http://api_lear_cable:8000/api/lear_cable/health && echo'
```

Si eso funciona pero desde fuera falla → el problema es Nginx/routing.

---

## 2) Errores típicos y fixes

### A) `502 Bad Gateway` en `/api/...`
**Síntoma**
- `curl http://localhost:8081/api/lear_cable/health` → 502

**Diagnóstico**
1) ¿La API está up?
```bash
docker ps --filter name=api_lear_cable
docker compose logs --tail 80 api_lear_cable
```

2) ¿Nginx resuelve el upstream?
```bash
docker exec -it tools_web sh -lc 'getent hosts api_lear_cable || true'
```

3) ¿Nginx puede llamar al upstream?
```bash
docker exec -it tools_web sh -lc 'wget -S -O- http://api_lear_cable:8000/api/lear_cable/health 2>&1 | sed -n "1,60p"'
```

4) Logs de Nginx:
```bash
docker exec -it tools_web sh -lc 'tail -n 120 /var/log/nginx/error.log'
```

**Causas comunes**
- `api_lear_cable` caído o reiniciando
- puerto interno cambiado
- contenedores en redes distintas (no deberían si todo está en el mismo compose)

**Fix**
```bash
docker compose up -d --force-recreate api_lear_cable tools_web
```
Si cambiaste Dockerfile/código, añade `--build`.

---

### B) `413 Request Entity Too Large` al subir ZIP
**Síntoma**
- UI o `curl -F file=@...` devuelve 413

**Causa**
- Nginx limita tamaño de request.

**Fix (nginx)**
En el `server { ... }` de nginx añade:
```nginx
client_max_body_size 200m;
```
(ajusta el tamaño a tu realidad)

**Aplicar**
```bash
docker compose up -d --build --force-recreate tools_web
```

**Verificar**
```bash
docker exec -it tools_web sh -lc 'nginx -T 2>/dev/null | grep -n "client_max_body_size"'
```

---

### C) Upload OK, pero status “se queda en WAITING”
**Síntoma**
- API devuelve `WAITING` aunque el processor ya procesó y mandó mail.

**Causa típica**
- La API lee status desde `/data/status/<batch>/status.json` pero el processor no lo actualiza (o escribe en otra ruta).

**Diagnóstico**
```bash
docker exec -it api_lear_cable sh -lc 'ls -la /data/status | tail -n 50'
docker exec -it processor_lear_cable sh -lc 'find /data/status -maxdepth 3 -type f -name status.json | tail -n 20'
```

**Fix**
- Asegurar que el **processor** actualiza `status.json` en cada transición (`UPLOADED → PROCESSING → DONE/ERROR`).
- Verifica que ambos contenedores montan el mismo `/data` (mismo bind mount).

---

### D) Email “enviado”, pero adjuntos vacíos
**Síntoma**
- Logs: “Email enviado correctamente”
- Mail llega con adjuntos 0 bytes

**Checklist**
1) ¿Los ficheros existen y tienen tamaño en `/data/out/<batch>/...`?
```bash
B=<batch_id>
docker exec -it processor_lear_cable sh -lc "ls -lah /data/out/$B; wc -c /data/out/$B/invoices_extracted.*"
```

2) ¿El código adjunta bytes reales (`read_bytes`)?
- Debe hacer `data = f.read_bytes()` y `msg.add_attachment(...)`.

3) Dump debug (si lo mantienes):
```bash
docker exec -it processor_lear_cable sh -lc 'ls -lah /tmp/last_email.eml; grep -nE "filename=|Content-Transfer-Encoding|Content-Type" /tmp/last_email.eml | sed -n "1,200p"'
```

**Causa típica**
- Se adjunta un handler/stream en vez de bytes, o se reabre en modo texto y queda vacío.

---

### E) `Permission denied (publickey)` con GitHub SSH
**Síntoma**
- `ssh -T git@github.com` → denied

**Diagnóstico**
```bash
ls -la ~/.ssh
ssh -vvv -i ~/.ssh/id_ed25519_github -T git@github.com
```

**Fix rápido**
Añade `~/.ssh/config`:
```sshconfig
Host github.com
  HostName github.com
  User git
  IdentityFile ~/.ssh/id_ed25519_github
  IdentitiesOnly yes
```
Luego:
```bash
ssh -T git@github.com
git clone git@github.com:totallogistic/invoice-automation-suite.git
```

---

### F) Processor no procesa lotes (quiet time / marker)
**Síntoma**
- Se suben ficheros pero nunca aparece “Lote listo”.

**Diagnóstico**
1) Variables runtime:
```bash
docker exec -it processor_lear_cable sh -lc 'env | egrep "INBOX_DIR|DONE_MARKER|BATCH_QUIET_SECONDS|POLL_SECONDS" | sort'
```

2) ¿Está entrando el lote donde el processor mira?
```bash
docker exec -it processor_lear_cable sh -lc 'find /data/inbox -maxdepth 4 -type f | head'
```

3) Logs:
```bash
docker compose logs -f processor_lear_cable
```

**Causa común**
- La UI/API sube a `/data/inbox/<batch>/...` pero el watcher espera otro layout.
- Quiet time demasiado alto y el folder no “se estabiliza”.
- El lote se está cerrando por marker diferente al esperado.

---

## 3) Verificar mounts / rutas (el 80% de los bugs)

En cada contenedor:
```bash
docker exec -it api_lear_cable sh -lc 'ls -la /data; ls -la /data/inbox /data/out /data/status || true'
docker exec -it processor_lear_cable sh -lc 'ls -la /data; ls -la /data/inbox /data/out /data/status || true'
```

Si el contenido no coincide → estás montando rutas distintas.

---

## 4) Rebuild/recreate “bien hecho”

### Cuando tocaste código copiado en imagen (`Dockerfile` + `COPY`)
```bash
docker compose up -d --build --force-recreate <service>
```

### Cuando tocaste solo HTML/config y está bind-mounted
Normalmente basta:
```bash
docker compose up -d --force-recreate <service>
```

### Reinicio completo
```bash
docker compose down
docker compose up -d --build
```
