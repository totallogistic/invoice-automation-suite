
# Banking Dashboard — Resumen Técnico

## 1. Descripción general

Sistema interno de monitorización de saldos bancarios para Total Logistic Services SL.
Conecta con múltiples entidades bancarias vía API PSD2 (Open Banking), centraliza los
saldos en un dashboard web y genera un informe Excel diario agrupado por empresa y
tipo de producto financiero.

**Stack:** Python 3.12 / FastAPI / uvicorn / SQLite / openpyxl
**Despliegue:** systemd service (standalone, fuera de Docker)
**Patrón:** igual que `form_generator` y `bl_inbox_watcher` del stack principal

---

## 2. Arquitectura

```
Red local Totallogistic
  PC usuarios (navegador)
       ↓
  https://banking.totallogistic  (nginx puerto 443, SSL autofirmado)
       ↓
  localhost:8085  (uvicorn — banking-dashboard.service)
       ↓
  Enable Banking API  (PSD2/AIS, HTTPS saliente)
       ↓
  Bancos: BBVA, Santander, Banca March, Bankinter,
          CaixaBank, Caja Rural del Sur, Deutsche Bank
          (Unicaja: sin soporte PSD2)
```

---

## 3. Estructura de ficheros

```
standalone_apps/banking_dashboard/
│
├── app.py                  FastAPI principal — rutas y endpoints
├── bank_fetcher.py         Cliente Enable Banking API (JWT + OAuth)
├── db.py                   Capa SQLite (banks, accounts, balances)
├── scheduler.py            Sync diario 07:00 + exports post-sync
├── bancos_export.py        Generador Excel por empresa/titular
├── banks.yaml              Bancos a conectar (aspsp_name para PSD2)
├── requirements.txt        Dependencias Python
├── .env                    Credenciales (NO en git)
│
├── keys/
│   └── private.key         Clave RSA 4096 para JWT Enable Banking (NO en git)
│
├── config/                 [Samba: banking-config — editable]
│   └── bancos_config.yaml  Mapeo completo cuentas → titular/tipo/params
│
├── data/                   [Samba: banking-dashboard — solo lectura]
│   ├── balances.db         SQLite: historial de saldos
│   ├── saldos_actuales.yaml  Último sync (puente para Excel)
│   └── situacion_YYYYMMDD.xlsx  Excel diario generado
│
├── templates_excel/        [Sin Samba — protegido]
│   └── (no se usa en flujo actual)
│
├── templates/
│   ├── dashboard.html      UI web del dashboard
│   └── setup.html          UI de conexión de bancos
│
└── install/
    ├── install.sh                              Instala servicio systemd
    ├── banking-dashboard.service.template      Plantilla servicio uvicorn
    └── banking-dashboard-ngrok.service.template  Solo DEV
```

---

## 4. Dependencias Python

```
fastapi==0.115.0
uvicorn[standard]==0.30.6
aiosqlite==0.20.0
httpx==0.27.2
PyJWT[crypto]==2.9.0
APScheduler==3.10.4
Jinja2==3.1.4
PyYAML==6.0.2
python-multipart==0.0.9
openpyxl==3.1.5
gspread==6.1.2        (Google Sheets — opcional)
google-auth==2.29.0   (Google Sheets — opcional)
```

Instalación:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

---

## 5. Variables de entorno (.env)

```bash
# Enable Banking — credenciales API
ENABLE_BANKING_APP_ID=84b28f71-7413-41b9-bd3d-3a8e0ef7fe46
ENABLE_BANKING_KEY_PATH=/home/mpino/invoice-automation-suite/standalone_apps/banking_dashboard/keys/private.key

# URL pública para callback OAuth (debe coincidir con Enable Banking dashboard)
APP_BASE_URL=https://banking.totallogistic          # PROD
# APP_BASE_URL=https://xxxx.ngrok-free.dev          # DEV

# Sync automático (hora en formato 24h)
SYNC_HOUR=7

# Email de reporte diario
REPORT_EMAIL=josepino@totallogistic.es

# Puerto uvicorn
PORT=8085

# Google Sheets (opcional — dejar vacío si no se usa)
# GOOGLE_SHEETS_CREDENTIALS=/ruta/service-account.json
# GOOGLE_SHEETS_ID=spreadsheet-id
```

---

## 6. Enable Banking

**Proveedor:** [enablebanking.com](https://enablebanking.com)
**Protocolo:** PSD2 / Open Banking AIS (Account Information Services)
**Modo:** Restricted (vinculación manual de cuentas, sin contrato)

**Credenciales:**

| Elemento             | Valor / Ubicación                               |
| -------------------- | ------------------------------------------------ |
| Application ID       | `84b28f71-7413-41b9-bd3d-3a8e0ef7fe46`         |
| Clave privada RSA    | `keys/private.key`                             |
| Certificado público | `keys/public.crt` (subido a enablebanking.com) |
| Vence certificado    | **3 abril 2027**                           |
| Redirect URL PROD    | `https://banking.totallogistic/callback`       |
| Redirect URL DEV     | `https://xxxx.ngrok-free.dev/callback`         |

**Autenticación JWT:**

```python
payload = {
    "iss": "enablebanking.com",       # fijo
    "aud": "api.enablebanking.com",   # fijo
    "iat": now,
    "exp": now + 3600,
}
headers = {"kid": APP_ID}             # app_id en el header, no en payload
algoritmo = "RS256"
```

**Bancos configurados en `banks.yaml`:**

| ID                 | Nombre display        | ASPSP name (Enable Banking)      |
| ------------------ | --------------------- | -------------------------------- |
| bbva               | BBVA                  | BBVA                             |
| banca_march        | Banca March           | Banca March                      |
| santander          | Banco Santander       | Banco Santander                  |
| bankinter          | Bankinter             | Bankinter                        |
| caixabank          | CaixaBank             | CaixaBank                        |
| caja_rural_sur     | Caja Rural del Sur    | Caja Rural del Sur               |
| caja_rural_granada | Caja Rural de Granada | Caja Rural de Granada            |
| unicaja            | Unicaja Banco         | Unicaja Banco (sin soporte PSD2) |
| deutsche_bank      | Deutsche Bank España | Deutsche Bank                    |

**Sesiones bancarias:**

- Válidas ~90 días
- Renovar desde `https://banking.totallogistic/setup` → "Reconectar"
- Se almacenan en `data/balances.db` (tabla `banks.session_id`)

---

## 7. Base de datos (SQLite)

Fichero: `data/balances.db`

```sql
-- Bancos configurados y estado de conexión
banks (id, name, aspsp_name, country,
       session_id, session_expires_at,
       connected_at, last_sync_at, status)

-- Cuentas descubiertas tras autorización OAuth
accounts (id, bank_id, iban, name, currency)

-- Historial de saldos (una entrada por sync por cuenta)
balances (id, account_id, bank_id, amount,
          currency, balance_type, recorded_at)
```

---

## 8. Flujo de sync

```
07:00 (APScheduler cron)  o  botón "Sync ahora"
    ↓
scheduler.sync_all_banks()
    ↓
Para cada banco conectado (status='connected'):
    GET /accounts/{id}/balances  (Enable Banking API)
    → guarda en balances (SQLite)
    ↓
_post_sync_exports()
    ↓
    1. Genera data/saldos_actuales.yaml
       (IBAN → dispuesto, banco, cuenta, moneda)
    ↓
    2. Genera data/situacion_YYYYMMDD.xlsx
       (bancos_export.generate_bancos)
    ↓
    3. Envía email resumen a REPORT_EMAIL
       (vía iasuite_common.email si disponible)
```

---

## 9. bancos_config.yaml

Mapeo manual de todas las cuentas bancarias. Es la pieza clave que conecta
los saldos de la API con el informe Excel estructurado por empresa.

**Campos por cuenta:**

| Campo                | Descripción                                                        |
| -------------------- | ------------------------------------------------------------------- |
| `iban`             | IBAN completo tal como devuelve la API (null para líneas sin IBAN) |
| `titular`          | Empresa propietaria (una de las 5 definidas en`orden_titulares`)  |
| `banco`            | Nombre de la entidad                                                |
| `tipo`             | `POLIZA` / `CTA_CTE` / `CTA_CTE_USD` / `LINEA` / `ANTIC`  |
| `cta_auxiliar`     | Número de cuenta contable                                          |
| `no_traer`         | `true` = excluir del total y del análisis                        |
| `euribor_tipo`     | `"Euribor 3"` / `"Euribor 6"` / `"Euribor 12"`                |
| `diferencial`      | Spread sobre Euribor (decimal, ej:`0.0095`)                       |
| `nd`               | Comisión ND (decimal)                                              |
| `importe`          | Límite del crédito (pólizas y líneas)                           |
| `dispuesto_manual` | Saldo manual para cuentas sin API (Unicaja, Caja Rural GR)          |

**Lógica de dispuesto (prioridad):**

1. `saldos_actuales.yaml` por IBAN (sync automático)
2. `dispuesto_manual` del config (manual, para bancos sin API)

**Sección euribor** (actualizar mensualmente):

```yaml
euribor:
  "Euribor 3":  0.02079
  "Euribor 6":  0.02475
  "Euribor 12": 0.02870
```

**Orden de titulares en el Excel:**

```yaml
orden_titulares:
  - "TOTAL LOGISTIC SERVICES SL"
  - "TOTAL ENGINEERING S.L."
  - "ASOCIACION MELILLA INTEGRA"
  - "COMERCIAL ANDALUZA DE SALDOS S.L."
  - "CARMELO MARTINEZ RODRIGUEZ S.L."
```

---

## 10. Excel generado (situacion_YYYYMMDD.xlsx)

**Hoja "Saldos actuales":**

Bloque 1 — Lista plana de todas las cuentas:

```
Banco | IBAN | Cuenta | Saldo | Fecha | [vacío] | NO TRAER | TITULAR | ENTIDAD
```

- Cuentas marcadas `no_traer: true` → fondo amarillo, etiqueta "NO TRAER"
- Total al final (solo cuentas válidas en EUR)

Bloque 2 — Secciones por empresa/titular:

```
TOTAL LOGISTIC SERVICES SL
  PÓLIZAS:  Entidad | IBAN | Cta.Aux. | Euribor tipo | Euribor val | Diferencial | ND | Coste | Límite | Dispuesto | % | Disponible | %
  CTA CTE €: ...
  CTA CTE $: ...
  LÍNEAS:   igual que PÓLIZAS

TOTAL ENGINEERING S.L.
  PÓLIZAS: ...
  CTA CTE: ...
...
```

**Cálculos en Excel:**

- Coste = Euribor + Diferencial
- % Dispuesto = Dispuesto / Límite
- Disponible = Límite − Dispuesto
- % Disponible = Disponible / Límite
- Alerta "OJO" si % disponible < 40%
- Alerta "OJO NEGATIVO" si saldo es negativo (póliza sobredispuesta)

---

## 11. Nginx (servidor de producción)

Fichero: `/etc/nginx/sites-enabled/toolbox-tools`

```nginx
server {
  listen 443 ssl;
  server_name banking.totallogistic;
  client_max_body_size 10m;

  ssl_certificate     /etc/ssl/certs/banking.crt;
  ssl_certificate_key /etc/ssl/private/banking.key;

  location / {
    proxy_pass http://127.0.0.1:8085;
    proxy_http_version 1.1;
    proxy_set_header Host $host;
    proxy_set_header X-Forwarded-Proto $scheme;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
  }
}

server {
  listen 80;
  server_name banking.totallogistic;
  return 301 https://$host$request_uri;
}
```

**Certificado SSL autofirmado:**

- Ubicación: `/etc/ssl/certs/banking.crt` y `/etc/ssl/private/banking.key`
- Generado con: `openssl req -x509 -nodes -days 3650 -newkey rsa:2048 -subj "/CN=banking.totallogistic"`
- Vence: 3650 días desde generación (~10 años)

**Hosts file** (en cada PC de la red):

```
192.168.0.42    banking.totallogistic
```

---

## 12. Samba shares

En `/etc/samba/smb.conf` del servidor Totallogistic:

```ini
[banking-dashboard]
   comment = Banking Dashboard — Saldos generados (solo lectura)
   path = /home/mpino/invoice-automation-suite/standalone_apps/banking_dashboard/data
   browseable = yes
   read only = yes
   valid users = josepino, mpino

[banking-config]
   comment = Banking Dashboard — Configuración editable
   path = /home/mpino/invoice-automation-suite/standalone_apps/banking_dashboard/config
   browseable = yes
   read only = no
   create mask = 0664
   directory mask = 0775
   valid users = josepino, mpino
```

---

## 13. Servicio systemd

Fichero: `/etc/systemd/system/banking-dashboard.service`

```ini
[Unit]
Description=Banking Dashboard (uvicorn)
After=network.target

[Service]
Type=simple
User=mpino
WorkingDirectory=/home/mpino/invoice-automation-suite/standalone_apps/banking_dashboard
EnvironmentFile=/home/mpino/invoice-automation-suite/standalone_apps/banking_dashboard/.env
ExecStart=/home/mpino/.../banking_dashboard/.venv/bin/uvicorn app:app --host 0.0.0.0 --port 8085
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

**Comandos útiles:**

```bash
sudo systemctl start banking-dashboard
sudo systemctl stop banking-dashboard
sudo systemctl restart banking-dashboard
sudo systemctl status banking-dashboard
journalctl -u banking-dashboard -f        # logs en tiempo real
journalctl -u banking-dashboard -n 50     # últimas 50 líneas
```

---

## 14. Entornos DEV vs PROD

|              | DEV (ubuntu24)                   | PROD (toolbox 192.168.0.42)   |
| ------------ | -------------------------------- | ----------------------------- |
| URL          | ngrok HTTPS                      | https://banking.totallogistic |
| Servicios    | uvicorn + ngrok (systemd)        | uvicorn (systemd)             |
| Bancos       | ING personal (pruebas)           | 7-8 bancos empresa            |
| Samba        | `banking-dashboard` solo mpino | + josepino                    |
| APP_BASE_URL | https://xxxx.ngrok-free.dev      | https://banking.totallogistic |

---

## 15. Mantenimiento recurrente

| Tarea                              | Frecuencia           | Acción                                                           |
| ---------------------------------- | -------------------- | ----------------------------------------------------------------- |
| Renovar sesiones bancarias         | Cada 90 días        | `/setup` → Reconectar bancos que expiren                       |
| Actualizar Euribor                 | Mensual              | Editar`config/bancos_config.yaml` → sección `euribor`       |
| Actualizar dispuesto líneas       | Cuando cambie        | Editar`config/bancos_config.yaml` → campo `dispuesto_manual` |
| Actualizar condiciones contratos   | Al renegociar        | Editar`bancos_config.yaml` → diferencial, nd, importe          |
| Renovar certificado Enable Banking | **Abril 2027** | Regenerar`keys/public.crt` y subir a enablebanking.com          |
| Limpiar Excel antiguos en`data/` | Manual/periódico    | Borrar`situacion_*.xlsx` anteriores                             |

---

## 16. Renovación del certificado Enable Banking (abril 2027)

```bash
cd standalone_apps/banking_dashboard/keys

# Regenera el par de claves
openssl genrsa -out private.key 4096
openssl req -new -x509 -days 365 \
  -key private.key \
  -out public.crt \
  -subj "/C=ES/ST=Andalucia/L=Almeria/O=TotalLogistic/CN=totallogistic.com"

# Sube public.crt a enablebanking.com:
# API applications → banking-dashboard → editar → Public certificate → reemplazar
# (private.key nunca sale del servidor)

sudo systemctl restart banking-dashboard
```

---

## 17. Troubleshooting

```bash
# Banking Dashboard — Resumen Técnico

## 1. Descripción general

Sistema interno de monitorización de saldos bancarios para Total Logistic Services SL.
Conecta con múltiples entidades bancarias vía API PSD2 (Open Banking), centraliza los
saldos en un dashboard web y genera un informe Excel diario agrupado por empresa y
tipo de producto financiero.

**Stack:** Python 3.12 / FastAPI / uvicorn / SQLite / openpyxl  
**Despliegue:** systemd service (standalone, fuera de Docker)  
**Patrón:** igual que `form_generator` y `bl_inbox_watcher` del stack principal

---

## 2. Arquitectura

```
Red local Totallogistic
  PC usuarios (navegador)
       ↓
  https://banking.totallogistic  (nginx puerto 443, SSL autofirmado)
       ↓
  localhost:8085  (uvicorn — banking-dashboard.service)
       ↓
  Enable Banking API  (PSD2/AIS, HTTPS saliente)
       ↓
  Bancos: BBVA, Santander, Banca March, Bankinter,
          CaixaBank, Caja Rural del Sur, Deutsche Bank
          (Unicaja: sin soporte PSD2)
```

---

## 3. Estructura de ficheros

```
standalone_apps/banking_dashboard/
│
├── app.py                  FastAPI principal — rutas y endpoints
├── bank_fetcher.py         Cliente Enable Banking API (JWT + OAuth)
├── db.py                   Capa SQLite (banks, accounts, balances)
├── scheduler.py            Sync diario 07:00 + exports post-sync
├── bancos_export.py        Generador Excel por empresa/titular
├── banks.yaml              Bancos a conectar (aspsp_name para PSD2)
├── requirements.txt        Dependencias Python
├── .env                    Credenciales (NO en git)
│
├── keys/
│   └── private.key         Clave RSA 4096 para JWT Enable Banking (NO en git)
│
├── config/                 [Samba: banking-config — editable]
│   └── bancos_config.yaml  Mapeo completo cuentas → titular/tipo/params
│
├── data/                   [Samba: banking-dashboard — solo lectura]
│   ├── balances.db         SQLite: historial de saldos
│   ├── saldos_actuales.yaml  Último sync (puente para Excel)
│   └── situacion_YYYYMMDD.xlsx  Excel diario generado
│
├── templates_excel/        [Sin Samba — protegido]
│   └── (no se usa en flujo actual)
│
├── templates/
│   ├── dashboard.html      UI web del dashboard
│   └── setup.html          UI de conexión de bancos
│
└── install/
    ├── install.sh                              Instala servicio systemd
    ├── banking-dashboard.service.template      Plantilla servicio uvicorn
    └── banking-dashboard-ngrok.service.template  Solo DEV
```

---

## 4. Dependencias Python

```
fastapi==0.115.0
uvicorn[standard]==0.30.6
aiosqlite==0.20.0
httpx==0.27.2
PyJWT[crypto]==2.9.0
APScheduler==3.10.4
Jinja2==3.1.4
PyYAML==6.0.2
python-multipart==0.0.9
openpyxl==3.1.5
gspread==6.1.2        (Google Sheets — opcional)
google-auth==2.29.0   (Google Sheets — opcional)
```

Instalación:
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

---

## 5. Variables de entorno (.env)

```bash
# Enable Banking — credenciales API
ENABLE_BANKING_APP_ID=84b28f71-7413-41b9-bd3d-3a8e0ef7fe46
ENABLE_BANKING_KEY_PATH=/home/mpino/invoice-automation-suite/standalone_apps/banking_dashboard/keys/private.key

# URL pública para callback OAuth (debe coincidir con Enable Banking dashboard)
APP_BASE_URL=https://banking.totallogistic          # PROD
# APP_BASE_URL=https://xxxx.ngrok-free.dev          # DEV

# Sync automático (hora en formato 24h)
SYNC_HOUR=7

# Email de reporte diario
REPORT_EMAIL=josepino@totallogistic.es

# Puerto uvicorn
PORT=8085

# Google Sheets (opcional — dejar vacío si no se usa)
# GOOGLE_SHEETS_CREDENTIALS=/ruta/service-account.json
# GOOGLE_SHEETS_ID=spreadsheet-id
```

---

## 6. Enable Banking

**Proveedor:** [enablebanking.com](https://enablebanking.com)  
**Protocolo:** PSD2 / Open Banking AIS (Account Information Services)  
**Modo:** Restricted (vinculación manual de cuentas, sin contrato)

**Credenciales:**

| Elemento | Valor / Ubicación |
|---|---|
| Application ID | `84b28f71-7413-41b9-bd3d-3a8e0ef7fe46` |
| Clave privada RSA | `keys/private.key` |
| Certificado público | `keys/public.crt` (subido a enablebanking.com) |
| Vence certificado | **3 abril 2027** |
| Redirect URL PROD | `https://banking.totallogistic/callback` |
| Redirect URL DEV | `https://xxxx.ngrok-free.dev/callback` |

**Autenticación JWT:**
```python
payload = {
    "iss": "enablebanking.com",       # fijo
    "aud": "api.enablebanking.com",   # fijo
    "iat": now,
    "exp": now + 3600,
}
headers = {"kid": APP_ID}             # app_id en el header, no en payload
algoritmo = "RS256"
```

**Bancos configurados en `banks.yaml`:**

| ID | Nombre display | ASPSP name (Enable Banking) |
|---|---|---|
| bbva | BBVA | BBVA |
| banca_march | Banca March | Banca March |
| santander | Banco Santander | Banco Santander |
| bankinter | Bankinter | Bankinter |
| caixabank | CaixaBank | CaixaBank |
| caja_rural_sur | Caja Rural del Sur | Caja Rural del Sur |
| caja_rural_granada | Caja Rural de Granada | Caja Rural de Granada |
| unicaja | Unicaja Banco | Unicaja Banco (sin soporte PSD2) |
| deutsche_bank | Deutsche Bank España | Deutsche Bank |

**Sesiones bancarias:**
- Válidas ~90 días
- Renovar desde `https://banking.totallogistic/setup` → "Reconectar"
- Se almacenan en `data/balances.db` (tabla `banks.session_id`)

---

## 7. Base de datos (SQLite)

Fichero: `data/balances.db`

```sql
-- Bancos configurados y estado de conexión
banks (id, name, aspsp_name, country,
       session_id, session_expires_at,
       connected_at, last_sync_at, status)

-- Cuentas descubiertas tras autorización OAuth
accounts (id, bank_id, iban, name, currency)

-- Historial de saldos (una entrada por sync por cuenta)
balances (id, account_id, bank_id, amount,
          currency, balance_type, recorded_at)
```

---

## 8. Flujo de sync

```
07:00 (APScheduler cron)  o  botón "Sync ahora"
    ↓
scheduler.sync_all_banks()
    ↓
Para cada banco conectado (status='connected'):
    GET /accounts/{id}/balances  (Enable Banking API)
    → guarda en balances (SQLite)
    ↓
_post_sync_exports()
    ↓
    1. Genera data/saldos_actuales.yaml
       (IBAN → dispuesto, banco, cuenta, moneda)
    ↓
    2. Genera data/situacion_YYYYMMDD.xlsx
       (bancos_export.generate_bancos)
    ↓
    3. Envía email resumen a REPORT_EMAIL
       (vía iasuite_common.email si disponible)
```

---

## 9. bancos_config.yaml

Mapeo manual de todas las cuentas bancarias. Es la pieza clave que conecta
los saldos de la API con el informe Excel estructurado por empresa.

**Campos por cuenta:**

| Campo | Descripción |
|---|---|
| `iban` | IBAN completo tal como devuelve la API (null para líneas sin IBAN) |
| `titular` | Empresa propietaria (una de las 5 definidas en `orden_titulares`) |
| `banco` | Nombre de la entidad |
| `tipo` | `POLIZA` / `CTA_CTE` / `CTA_CTE_USD` / `LINEA` / `ANTIC` |
| `cta_auxiliar` | Número de cuenta contable |
| `no_traer` | `true` = excluir del total y del análisis |
| `euribor_tipo` | `"Euribor 3"` / `"Euribor 6"` / `"Euribor 12"` |
| `diferencial` | Spread sobre Euribor (decimal, ej: `0.0095`) |
| `nd` | Comisión ND (decimal) |
| `importe` | Límite del crédito (pólizas y líneas) |
| `dispuesto_manual` | Saldo manual para cuentas sin API (Unicaja, Caja Rural GR) |

**Lógica de dispuesto (prioridad):**
1. `saldos_actuales.yaml` por IBAN (sync automático)
2. `dispuesto_manual` del config (manual, para bancos sin API)

**Sección euribor** (actualizar mensualmente):
```yaml
euribor:
  "Euribor 3":  0.02079
  "Euribor 6":  0.02475
  "Euribor 12": 0.02870
```

**Orden de titulares en el Excel:**
```yaml
orden_titulares:
  - "TOTAL LOGISTIC SERVICES SL"
  - "TOTAL ENGINEERING S.L."
  - "ASOCIACION MELILLA INTEGRA"
  - "COMERCIAL ANDALUZA DE SALDOS S.L."
  - "CARMELO MARTINEZ RODRIGUEZ S.L."
```

---

## 10. Excel generado (situacion_YYYYMMDD.xlsx)

**Hoja "Saldos actuales":**

Bloque 1 — Lista plana de todas las cuentas:
```
Banco | IBAN | Cuenta | Saldo | Fecha | [vacío] | NO TRAER | TITULAR | ENTIDAD
```
- Cuentas marcadas `no_traer: true` → fondo amarillo, etiqueta "NO TRAER"
- Total al final (solo cuentas válidas en EUR)

Bloque 2 — Secciones por empresa/titular:
```
TOTAL LOGISTIC SERVICES SL
  PÓLIZAS:  Entidad | IBAN | Cta.Aux. | Euribor tipo | Euribor val | Diferencial | ND | Coste | Límite | Dispuesto | % | Disponible | %
  CTA CTE €: ...
  CTA CTE $: ...
  LÍNEAS:   igual que PÓLIZAS

TOTAL ENGINEERING S.L.
  PÓLIZAS: ...
  CTA CTE: ...
...
```

**Cálculos en Excel:**
- Coste = Euribor + Diferencial
- % Dispuesto = Dispuesto / Límite
- Disponible = Límite − Dispuesto
- % Disponible = Disponible / Límite
- Alerta "OJO" si % disponible < 40%
- Alerta "OJO NEGATIVO" si saldo es negativo (póliza sobredispuesta)

---

## 11. Nginx (servidor de producción)

Fichero: `/etc/nginx/sites-enabled/toolbox-tools`

```nginx
server {
  listen 443 ssl;
  server_name banking.totallogistic;
  client_max_body_size 10m;

  ssl_certificate     /etc/ssl/certs/banking.crt;
  ssl_certificate_key /etc/ssl/private/banking.key;

  location / {
    proxy_pass http://127.0.0.1:8085;
    proxy_http_version 1.1;
    proxy_set_header Host $host;
    proxy_set_header X-Forwarded-Proto $scheme;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
  }
}

server {
  listen 80;
  server_name banking.totallogistic;
  return 301 https://$host$request_uri;
}
```

**Certificado SSL autofirmado:**
- Ubicación: `/etc/ssl/certs/banking.crt` y `/etc/ssl/private/banking.key`
- Generado con: `openssl req -x509 -nodes -days 3650 -newkey rsa:2048 -subj "/CN=banking.totallogistic"`
- Vence: 3650 días desde generación (~10 años)

**Hosts file** (en cada PC de la red):
```
192.168.0.42    banking.totallogistic
```

---

## 12. Samba shares

En `/etc/samba/smb.conf` del servidor Totallogistic:

```ini
[banking-dashboard]
   comment = Banking Dashboard — Saldos generados (solo lectura)
   path = /home/mpino/invoice-automation-suite/standalone_apps/banking_dashboard/data
   browseable = yes
   read only = yes
   valid users = josepino, mpino

[banking-config]
   comment = Banking Dashboard — Configuración editable
   path = /home/mpino/invoice-automation-suite/standalone_apps/banking_dashboard/config
   browseable = yes
   read only = no
   create mask = 0664
   directory mask = 0775
   valid users = josepino, mpino
```

---

## 13. Servicio systemd

Fichero: `/etc/systemd/system/banking-dashboard.service`

```ini
[Unit]
Description=Banking Dashboard (uvicorn)
After=network.target

[Service]
Type=simple
User=mpino
WorkingDirectory=/home/mpino/invoice-automation-suite/standalone_apps/banking_dashboard
EnvironmentFile=/home/mpino/invoice-automation-suite/standalone_apps/banking_dashboard/.env
ExecStart=/home/mpino/.../banking_dashboard/.venv/bin/uvicorn app:app --host 0.0.0.0 --port 8085
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

**Comandos útiles:**
```bash
sudo systemctl start banking-dashboard
sudo systemctl stop banking-dashboard
sudo systemctl restart banking-dashboard
sudo systemctl status banking-dashboard
journalctl -u banking-dashboard -f        # logs en tiempo real
journalctl -u banking-dashboard -n 50     # últimas 50 líneas
```

---

## 14. ngrok (solo DEV)

ngrok es el túnel HTTPS que permite a Enable Banking hacer el callback OAuth hacia
el servidor DEV, que no tiene IP pública ni dominio real.

**Instalación:**
```bash
snap install ngrok
ngrok config add-authtoken TU_TOKEN  # cuenta en ngrok.com — gratuita
```

**Servicio systemd** (`banking-dashboard-ngrok.service`):
```ini
[Unit]
Description=ngrok tunnel for Banking Dashboard (DEV only)
After=network.target banking-dashboard.service
Requires=banking-dashboard.service

[Service]
Type=simple
User=mpino
ExecStart=/snap/bin/ngrok http 8085 --log=stdout
Restart=on-failure
RestartSec=15

[Install]
WantedBy=multi-user.target
```

**Comandos:**
```bash
sudo systemctl start banking-dashboard-ngrok
sudo systemctl status banking-dashboard-ngrok
journalctl -u banking-dashboard-ngrok -f

# Ver URL asignada
curl -s http://127.0.0.1:4040/api/tunnels | \
  python3 -c "import sys,json; print(json.load(sys.stdin)['tunnels'][0]['public_url'])"
```

**URL actual DEV:** `https://dustin-microclimatologic-cordelia.ngrok-free.dev`

⚠️ **Importante:** Con cuenta gratuita la URL se mantiene entre reinicios del servicio
pero puede cambiar si se reinstala ngrok o se crea un nuevo túnel. Si cambia:
1. Actualizar `APP_BASE_URL` en `.env` de DEV
2. Actualizar Redirect URL en enablebanking.com → tu app → editar
3. `sudo systemctl restart banking-dashboard`

**En PROD no se usa ngrok** — nginx hace de proxy con SSL propio.

---

## 18. Entornos DEV vs PROD

| | DEV (ubuntu24) | PROD (toolbox 192.168.0.42) |
|---|---|---|
| URL | ngrok HTTPS | https://banking.totallogistic |
| Servicios | uvicorn + ngrok (systemd) | uvicorn (systemd) |
| Bancos | ING personal (pruebas) | 7-8 bancos empresa |
| Samba | `banking-dashboard` solo mpino | + josepino |
| APP_BASE_URL | https://xxxx.ngrok-free.dev | https://banking.totallogistic |

---

## 18. Mantenimiento recurrente

| Tarea | Frecuencia | Acción |
|---|---|---|
| Renovar sesiones bancarias | Cada 90 días | `/setup` → Reconectar bancos que expiren |
| Actualizar Euribor | Mensual | Editar `config/bancos_config.yaml` → sección `euribor` |
| Actualizar dispuesto líneas | Cuando cambie | Editar `config/bancos_config.yaml` → campo `dispuesto_manual` |
| Actualizar condiciones contratos | Al renegociar | Editar `bancos_config.yaml` → diferencial, nd, importe |
| Renovar certificado Enable Banking | **Abril 2027** | Regenerar `keys/public.crt` y subir a enablebanking.com |
| Limpiar Excel antiguos en `data/` | Manual/periódico | Borrar `situacion_*.xlsx` anteriores |

---

## 18. Renovación del certificado Enable Banking (abril 2027)

```bash
cd standalone_apps/banking_dashboard/keys

# Regenera el par de claves
openssl genrsa -out private.key 4096
openssl req -new -x509 -days 365 \
  -key private.key \
  -out public.crt \
  -subj "/C=ES/ST=Andalucia/L=Almeria/O=TotalLogistic/CN=totallogistic.com"

# Sube public.crt a enablebanking.com:
# API applications → banking-dashboard → editar → Public certificate → reemplazar
# (private.key nunca sale del servidor)

sudo systemctl restart banking-dashboard
```

---

## 18. Troubleshooting

```bash
# Servicio no arranca
journalctl -u banking-dashboard -n 50

# 401 Unauthorized → Enable Banking
# → Verificar ENABLE_BANKING_APP_ID en .env
# → Verificar que private.key corresponde al public.crt subido
# → Verificar que el certificado no ha expirado (ver §16)

# Banco falla al conectar (400 Bad Request)
# → Verificar que la cuenta está vinculada en enablebanking.com → Link accounts
# → Algunos bancos (Unicaja) no soportan PSD2 empresarial

# Sesión expirada (banco aparece como pendiente)
# → https://banking.totallogistic/setup → Reconectar

# Excel no genera / botón da error
journalctl -u banking-dashboard -f
# → Verificar que config/bancos_config.yaml existe
# → Verificar que data/saldos_actuales.yaml tiene datos

# Ver saldos directamente en DB
sqlite3 data/balances.db \
  "SELECT b.name, a.iban, bl.amount, bl.recorded_at
   FROM balances bl
   JOIN accounts a ON a.id = bl.account_id
   JOIN banks b ON b.id = bl.bank_id
   ORDER BY bl.recorded_at DESC LIMIT 20;"
```
```
