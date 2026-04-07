# Banking Dashboard — Guía de despliegue a producción

## Estrategia recomendada

Conecta todos los bancos desde el servidor DEV (con ngrok).
Cuando todo esté validado, copia el repo + la DB + la clave privada al servidor de producción.
Las sesiones bancarias son válidas en cualquier servidor — no hace falta re-autorizar en cada banco.

---

## Paso 1 — En DEV: conectar todos los bancos

1. Edita `banks.yaml` con los 9 bancos de producción
2. Para cada banco, ve a `/setup` → "Conectar →"
3. Autoriza en la web del banco
4. Verifica que aparece "✓ conectado" para cada uno
5. Pulsa "Sync ahora" y confirma que el dashboard muestra saldos

Una vez conectados los 9 bancos en DEV, la `data/balances.db` contiene todas las sesiones.

---

## Paso 2 — Preparar el servidor de producción

### Requisitos
- Ubuntu 20.04+ con Python 3.10+
- Dominio con HTTPS (o nginx + certbot)
- Puerto 443/80 abierto

### Clonar el repo
```bash
git clone tu-repo /opt/banking-dashboard
cd /opt/banking-dashboard/standalone_apps/banking_dashboard
```

### Copiar la clave privada desde DEV
```bash
# Ejecuta esto desde DEV
scp standalone_apps/banking_dashboard/keys/private.key \
    usuario@servidor-prod:/opt/banking-dashboard/standalone_apps/banking_dashboard/keys/private.key

# En prod, ajusta permisos
ssh usuario@servidor-prod
chmod 600 /opt/banking-dashboard/standalone_apps/banking_dashboard/keys/private.key
```

### Copiar la base de datos con las sesiones bancarias
```bash
# Ejecuta esto desde DEV (con los 9 bancos ya conectados)
scp standalone_apps/banking_dashboard/data/balances.db \
    usuario@servidor-prod:/opt/banking-dashboard/standalone_apps/banking_dashboard/data/balances.db
```

---

## Paso 3 — Configurar nginx + HTTPS en producción

```nginx
server {
    listen 443 ssl;
    server_name banking.totallogistic.com;  # tu dominio

    ssl_certificate     /etc/letsencrypt/live/banking.totallogistic.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/banking.totallogistic.com/privkey.pem;

    location / {
        proxy_pass http://127.0.0.1:8085;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

```bash
# Instala certbot si no lo tienes
sudo apt install certbot python3-certbot-nginx
sudo certbot --nginx -d banking.totallogistic.com
```

---

## Paso 4 — Instalar y arrancar en producción

```bash
cd /opt/banking-dashboard/standalone_apps/banking_dashboard

# Crea el .env de producción
cp .env.example .env
nano .env
```

Contenido del `.env` de producción:
```bash
ENABLE_BANKING_APP_ID=84b28f71-7413-41b9-bd3d-3a8e0ef7fe46
ENABLE_BANKING_KEY_PATH=/opt/banking-dashboard/standalone_apps/banking_dashboard/keys/private.key
APP_BASE_URL=https://banking.totallogistic.com
SYNC_HOUR=7
REPORT_EMAIL=josepino@totallogistic.es
PORT=8085
```

```bash
# Instala el servicio (sin --dev, sin ngrok)
sudo bash install/install.sh

# Arranca
sudo systemctl start banking-dashboard
sudo systemctl status banking-dashboard
```

---

## Paso 5 — Actualizar Enable Banking con la URL de producción

En el dashboard de enablebanking.com → tu app → Redirect URLs:
- Añade: `https://banking.totallogistic.com/callback`
- Puedes mantener también la de ngrok para DEV

---

## Paso 6 — Verificar

```bash
# Logs en tiempo real
journalctl -u banking-dashboard -f

# Forzar sync manual
curl -X POST https://banking.totallogistic.com/sync
```

Abre `https://banking.totallogistic.com` — debería mostrar los 9 bancos con saldos.

---

## Mantenimiento recurrente

| Tarea | Frecuencia | Acción |
|---|---|---|
| Renovar sesiones bancarias | Cada 90 días | Ve a `/setup` → "Reconectar" en los que expiren |
| Logs de errores | Semanal | `journalctl -u banking-dashboard --since "7 days ago"` |
| Backup de la DB | Mensual | `cp data/balances.db data/balances.db.bak` |

---

## Troubleshooting rápido

```bash
# El servicio no arranca
journalctl -u banking-dashboard -n 50

# El sync falla para un banco concreto
# → Ve a /setup y reconecta ese banco (sesión expirada)

# Ver saldos en la DB directamente
sqlite3 data/balances.db "SELECT b.name, bl.amount, bl.recorded_at FROM balances bl JOIN banks b ON b.id=bl.bank_id ORDER BY bl.recorded_at DESC LIMIT 20;"
```