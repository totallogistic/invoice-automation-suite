# Despliegue — invoice-automation-suite

Guía de todas las formas de desplegar el stack, por servicio y global, más el
procedimiento para levantar **prod env en un host DEV** desde la rama
`feature/versionado-procesos-forms`. Incluye una sección de **problemas
detectados + fixes propuestos** (§7) pensada para host nuevo (Ubuntu DEV).

> Los valores reales (puertos, credenciales) viven en `.env.prod`, que **no está
> en git** (gitignorado). Hay que copiarlo desde el host de prod a cada host.

---

## 1. Arquitectura

Dos bloques independientes:

**A. Stack Docker** (proyecto compose `ias_<env>`, p.ej. `ias_prod`)

| Servicio              | Imagen                                           | Puerto                                                                                                                   | Notas                                                                                                                                                                    |
| --------------------- | ------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `unified_api`       | build`services/unified_api` (python:3.12-slim) | `${UNIFIED_API_PORT:-8000}`                                                                                            | FastAPI.`main.py` va **horneado** en la imagen → cambiarlo exige rebuild. `apps/`, `config/`, `libs/` van **montados** (ro) → cambios sin rebuild. |
| `unified_processor` | build`services/unified_processor`              | — (worker)                                                                                                              | Poll de inboxes, corre extractores, envía emails. Igual:`apps/`, `config/` montados.                                                                                |
| `tools_web`         | `nginx:alpine`                                 | `${TOOLS_WEB_PORT:-8081}:80` | Sirve `services/web/sites/${ENVIRONMENT:-dev}` y proxy `/api/ → unified_api:8000`. |                                                                                                                                                                          |

**B. Servicios standalone (systemd, en el host, fuera de Docker)**

| Servicio          | Unidad systemd                                               | Puerto                                | Env source                                                     |
| ----------------- | ------------------------------------------------------------ | ------------------------------------- | -------------------------------------------------------------- |
| Form Generator    | `form-generator-<env>` (p.ej. `form-generator-prod`)     | `FORM_UI_PORT` (env.json: `8201`) | `.env.prod` de la raíz del repo                             |
| Banking Dashboard | `banking-dashboard` (+ `banking-dashboard-ngrok` en dev) | `PORT` (default `8085`)           | **su propio** `standalone_apps/banking_dashboard/.env` |

La landing (`services/web/sites/prod/index.html`) y `env.json` (`{"env":"PROD","formPort":"8201"}`)
enlazan los forms al `formPort`; el badge PROD/DEV lo pinta el JS leyendo `env.json`.

---

## 2. Requisitos del host DEV (Ubuntu)

1. **Docker Engine + plugin compose** (`docker compose version`).
2. **Python 3 + pip** (para los standalone). En Ubuntu con PEP 668 usar
   `--break-system-packages` o venv (banking ya usa venv propio).
3. **`.env.prod`** copiado desde prod a la raíz del repo (no está en git):
   ```bash
   scp usuario@toolbox:/ruta/repo/.env.prod  ~/repos/.../invoice-automation-suite/.env.prod
   ```
4. **Banking** (si se despliega): su propio `.env`, `keys/private.key` y
   `data/balances.db` con las sesiones bancarias (ver §4.3 y §7-E).
5. Rutas de host que hoy usa `compose.yaml` (ver §7-A) — el punto más importante
   a resolver antes del primer `up` en un host nuevo.

---

## 3. Desplegar ESTA rama a DEV (camino corto)

```bash
# 1. En el host DEV: traer la rama
git fetch origin
git checkout feature/versionado-procesos-forms      # o git clone + checkout

# 2. .env.prod presente en la raíz (copiado de prod)  → ver §2.3
ls .env.prod

# 3. Stack docker (rebuild del API con nuestros cambios de main.py)
./scripts/deploy.sh --env prod

# 4. Form generator (recoge app.py + los nuevos 'version' de los schemas)
sudo systemctl restart form-generator-prod

# 5. Banking: NO lo tocan nuestros cambios → sólo si es host nuevo (ver §4.3)
```

### Qué cambio de esta rama necesita qué acción

| Cambio (esta rama)                                                   | Vía                                  | Acción para que aplique                                                       |
| -------------------------------------------------------------------- | ------------------------------------- | ------------------------------------------------------------------------------ |
| `services/unified_api/app/main.py` (endpoints + `/api/versions`) | **horneado** en imagen          | `docker compose build` → lo hace `deploy.sh` (`--no-cache` por defecto) |
| `config/tools.yaml` (bloque retirados)                             | montado`:ro`                        | `up -d` recrea / reinicia contenedores                                       |
| `apps/*/extractor/*.py` (SCRIPT_VERSION/CHANGELOG)                 | montado`:ro`, leído por subprocess | inmediato, sin rebuild                                                         |
| `services/web/sites/prod/*` (landing, tools)                       | montado en nginx                      | inmediato (fichero estático)                                                  |
| `standalone_apps/form_generator/app.py` + `schemas/*.json`       | systemd                               | `systemctl restart form-generator-prod`                                      |

> Nuestros cambios **no tocan banking**. Si el host DEV ya tenía el stack, con
> `deploy.sh --env prod` + `restart form-generator-prod` es suficiente.

---

## 4. Desplegar cada servicio (referencia por comandos)

### 4.1 Stack Docker

Manual, pieza a pieza:

```bash
docker compose --env-file .env.prod -p ias_prod down
docker compose --env-file .env.prod -p ias_prod build --no-cache
docker compose --env-file .env.prod -p ias_prod up -d
docker compose --env-file .env.prod -p ias_prod restart tools_web   # CRÍTICO: refresca DNS upstream
```

Con script (hace build → up → restart tools_web → smoke test):

```bash
./scripts/deploy.sh --env prod                # build --no-cache
./scripts/deploy.sh --env prod --no-cache=false  # build con cache (iteración rápida)
```

> El `restart tools_web` es obligatorio tras rebuild del API: nginx resuelve el
> upstream al arrancar; si el API cambia de IP, da 502 hasta el restart
> (incidente 2026-04-30, ya cubierto por `deploy.sh`).

### 4.2 Form Generator (systemd, multi-entorno)

Host nuevo (instala servicio(s) `form-generator-<env>` por cada `.env.*`):

```bash
# ⚠ Instalar dependencias primero (install-all.sh NO lo hace — ver §7-D)
pip install --user --break-system-packages -r standalone_apps/form_generator/requirements.txt
sudo standalone_apps/form_generator/install/install-all.sh
```

Actualizar tras cambios de código/schemas:

```bash
sudo systemctl restart form-generator-prod
sudo journalctl -u form-generator-prod -f      # logs
```

Reinstalar limpio:

```bash
sudo standalone_apps/form_generator/install/uninstall-all.sh
sudo standalone_apps/form_generator/install/install-all.sh
```

### 4.3 Banking Dashboard (systemd, env propio)

**No usa `.env.prod`** — usa `standalone_apps/banking_dashboard/.env` propio, un
venv, y necesita la clave privada de Enable Banking + `data/balances.db` con las
sesiones bancarias.

```bash
cd standalone_apps/banking_dashboard
sudo bash install/install.sh            # crea venv, instala deps, crea servicio
# luego:
cp .env.example .env && nano .env        # ENABLE_BANKING_APP_ID, KEY_PATH, PORT, ...
mkdir -p keys data
scp usuario@prod:.../keys/private.key keys/    # copiar clave de prod/DEV origen
scp usuario@prod:.../data/balances.db  data/   # copiar sesiones (o reconectar bancos)
sudo systemctl start banking-dashboard
sudo systemctl restart banking-dashboard
```

Ver `standalone_apps/banking_dashboard/install/deploy.md` para el flujo completo
(conectar bancos en DEV con ngrok → copiar DB+clave). En DEV con ngrok:
`sudo bash install/install.sh --dev`.

---

## 5. Deploy global (docker + standalone) — secuencia recomendada

Versión limpia de la secuencia manual (sin duplicados ni el typo
`uninstallall-all.sh`):

```bash
cd ~/repos/totallogistic/invoice-automation-suite

# 1) Docker
./scripts/deploy.sh --env prod

# 2) Form generator (host nuevo: pip install primero, ver §4.2)
sudo systemctl restart form-generator-prod
#   host nuevo:  sudo standalone_apps/form_generator/install/install-all.sh

# 3) Banking (sólo si aplica / host nuevo)
sudo systemctl restart banking-dashboard
```

> **Propuesta:** unificar esto en un único `scripts/deploy_all.sh --env prod`
> que haga docker + standalone en orden (ver §7-C).

---

## 6. Verificación post-deploy

```bash
# API agregada (nuestra) — debe listar 11 procesos activos, up=11
curl -s http://localhost:${TOOLS_WEB_PORT}/api/versions | python3 -m json.tool | head -40

# Un proceso concreto
curl -s http://localhost:${TOOLS_WEB_PORT}/api/merge_pdf/version

# Forms (en el puerto del form-generator, NO tras nginx)
curl -s http://localhost:8201/api/forms/status | python3 -m json.tool | head -40

# Estado
docker compose --env-file .env.prod -p ias_prod ps
sudo systemctl status form-generator-prod
```

---

## 7. Problemas detectados + fixes propuestos

Ordenados por impacto en un **host DEV nuevo**. Nada de esto está aplicado aún:
son propuestas para revisar.

### A. 🔴 Rutas de host hardcodeadas en `compose.yaml` (bloqueante en host nuevo)

`unified_api` y `unified_processor` montan paths absolutos que **no existen en
un host nuevo**; Docker los crea como **directorios vacíos** silenciosamente,
lo que puede romper o dejar el servicio a medias:

```yaml
# unified_api.volumes
- /opt/bl_sync.sh:/opt/bl_sync.sh:ro          # BL retirado
- /usr/bin/rclone:/usr/bin/rclone:ro          # BL retirado
- /home/mpino/.config/rclone:/root/.config/rclone   # usuario 'mpino' específico
# unified_processor.volumes
- /mnt/canon-hotfolder:/mnt/canon-hotfolder   # scanner de camion (host-specific)
```

**Propuesta:**

- Quitar los 3 montajes de BL/rclone (BL está retirado) y el env `BL_SYNC_SCRIPT`.
- Parametrizar el hotfolder: `- ${CANON_HOTFOLDER:-./data/canon-hotfolder}:/mnt/canon-hotfolder`.
  En prod se fija `CANON_HOTFOLDER=/mnt/canon-hotfolder` en su `.env.prod`; en
  DEV se usa el default relativo al repo (que siempre existe).
- Alternativa "sin tocar el env compartido": mover lo host-specific a un
  `compose.override.yml` por host (gitignorado, como el `.env`).

### B. 🟠 `ENVIRONMENT=prod` es obligatorio en `.env.prod`

`tools_web` monta `services/web/sites/${ENVIRONMENT:-dev}`, pero **sólo existe
`sites/prod`** (no `sites/dev`). Si `ENVIRONMENT` no está definido, nginx sirve
un directorio vacío (404). El comentario del compose dice `WEB_SITE_DIR` pero el
código usa `ENVIRONMENT` (desalineado). **Propuesta:** confirmar
`ENVIRONMENT=prod` en `.env.prod` y corregir el comentario del compose.

### C. 🟠 `deploy.sh` sólo hace Docker, no los standalone

El "deploy global" real es docker + form_generator + banking, pero está en la
cabeza del operador. **Propuesta:** añadir `scripts/deploy_all.sh --env <env>`
que orqueste: `deploy.sh` → `systemctl restart form-generator-<env>` (o install
si no existe) → banking opcional. (Borrador listo si lo quieres.)

### D. 🟠 `form_generator/install/install-all.sh` no instala dependencias

Crea el servicio systemd pero nunca hace `pip install`. En host nuevo el servicio
arranca y **falla** (falta fastapi/uvicorn/jsonschema/…). **Propuesta:** añadir
un paso de `pip install --break-system-packages -r requirements.txt` al principio
de `install-all.sh` (una vez para la app).

### E. 🔴 `form_generator/requirements.txt` no incluye `pyyaml`

`app.py` hace `import yaml` (para leer `tools.yaml`), pero `pyyaml` **no está** en
`requirements.txt`. En prod funciona por suerte (está instalado por otra vía); en
un host nuevo con sólo `pip install -r requirements.txt`, el servicio **no
arranca** (`ModuleNotFoundError: yaml`). **Propuesta:** añadir `pyyaml==6.0.1` a
`standalone_apps/form_generator/requirements.txt`.

### F. 🟡 Banking usa su propio `.env`, no `.env.prod`

"Usar el mismo `.env.prod`" no cubre banking: necesita su `.env` propio +
`keys/private.key` + `data/balances.db`. Es un track aparte (§4.3). Además el
`.env.example` de banking usa `PORT=8080` mientras `install.sh`/`deploy.md`
asumen `8085` (inconsistencia menor a alinear).

### G. 🟡 `services/web/apply_env_tag.sh` usa `sed -i ''` (sintaxis macOS)

`sed -i ''` es BSD/macOS; en Linux (host DEV) **falla**. No lo llama `deploy.sh`
(el badge lo pinta el JS con `env.json`), así que es opcional/legacy, pero si se
ejecuta en la Ubuntu, rompe. **Propuesta:** hacerlo portable (detectar GNU vs
BSD sed) o marcarlo como "sólo macOS".

### H. 🟢 Cosméticos en `deploy.sh`

El primer smoke test se etiqueta `/health` pero pega a `/api/camion/version`
(funciona, camion sigue activo). **Propuesta:** renombrar la etiqueta y añadir un
check de `/api/versions` (nuestro agregador) al smoke test.

---

## 8. Resumen de cambios propuestos (para decidir cuáles aplico)

| # | Fichero                                   | Cambio                                               | Severidad |
| - | ----------------------------------------- | ---------------------------------------------------- | --------- |
| A | `compose.yaml`                          | Quitar montajes BL/rclone; parametrizar hotfolder    | 🔴 alta   |
| E | `form_generator/requirements.txt`       | Añadir`pyyaml`                                    | 🔴 alta   |
| D | `form_generator/install/install-all.sh` | Añadir`pip install` de deps                       | 🟠 media  |
| C | `scripts/deploy_all.sh` (nuevo)         | Orquestador docker + standalone                      | 🟠 media  |
| B | `compose.yaml`                          | Corregir comentario`WEB_SITE_DIR`→`ENVIRONMENT` | 🟠 media  |
| G | `services/web/apply_env_tag.sh`         | `sed` portable                                     | 🟡 baja   |
| H | `scripts/deploy.sh`                     | Etiqueta smoke + check`/api/versions`              | 🟢 baja   |
