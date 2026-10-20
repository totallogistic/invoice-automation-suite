# dua_watcher

Watcher standalone que automatiza el envío de **DUAs de aduanas a un SharePoint**
del cliente, renombrados con su número de DUA. Patrón watcher del stack
(como `bl` / `merge_pdf`).

## Flujo

```
Escáner → Samba (dua-inbox) → dua_watcher (Linux)
   → OCR: extrae el nº DUA de la 1ª pág, valida contra el resto de págs
   → ALTA confianza  → renombra a <DUA>.pdf → Samba (dua-outbox)
   → DUDOSO          → carpeta review/ + email interno con recorte del DUA
   → conserva SIEMPRE el original en processed/ · historial.json
         ↓ (dua-outbox)
   Windows Server con OneDrive sincronizando la biblioteca de Teams/SharePoint
   → sube solo a SharePoint
```

## Por qué OCR + revisión (importante)

Los DUA llegan como **PDF escaneado sin capa de texto**, así que hay que usar OCR
(tesseract). El OCR sobre estos escaneos tiene un **sesgo sistemático** conocido
(confunde `I`↔`1`, `O`↔`0`). Como el nombre del fichero es dato de **compliance**,
el sistema NO sube a ciegas:

- **Alta confianza** (formato exacto, confirmado en varias páginas, sin caracteres
  ambiguos) → renombra y entrega automático.
- **Dudoso** (caracteres ambiguos, discrepancia entre páginas, formato no exacto) →
  `review/` + email interno con el PDF, el DUA propuesto y un **recorte PNG de la zona
  del DUA** para que el operador confirme de un vistazo y coloque el fichero en
  `outbox/`.

> En la práctica, con los escaneos actuales la mayoría caerá en "dudoso" (el sesgo
> del OCR es consistente y la validación cruzada no siempre lo detecta). Es
> deliberado: mejor revisión humana rápida que un nombre incorrecto en compliance.
> La vía ideal sería que el origen entregue el DUA en el nombre/un CSV, o PDF con
> capa de texto — pendiente de pedir al cliente (ver doc de proyecto).

## Carpetas (`/data/ias_prod/data/dua/`)

| Carpeta | Qué contiene |
|---|---|
| `inbox/` | PDFs que deja el escáner (vía Samba `dua-inbox`) |
| `outbox/` | `<DUA>.pdf` renombrados (vía Samba `dua-outbox` → OneDrive → SharePoint) |
| `review/` | Dudosos: PDF + `.propuesto.txt` + `.dua_zona.png` |
| `processed/` | Copia del original (traza, siempre) |
| `error/` | PDFs que fallaron al procesar |
| `historial.json` | Registro: fecha, original, DUA, confianza, motivo, nº páginas |

## Instalación

```bash
# 1. Watcher (instala tesseract+poppler, deps Python, crea carpetas y servicio)
sudo install/install-dua-watcher.sh

# 2. Revisa el entorno (SMTP + a quién van los emails de revisión)
nano dua.env           # MAIL_INTERNO_DUA, SMTP_*

# 3. Shares Samba de entrada y salida
sudo install/install-samba-dua.sh duascan
```

## Configuración (`dua.env` o entorno)

| Variable | Default | Descripción |
|---|---|---|
| `DUA_INBOX` | `/data/ias_prod/data/dua/inbox` | Carpeta vigilada |
| `DUA_OUTBOX` | `.../dua/outbox` | Salida (OneDrive sincroniza) |
| `DUA_QUIET_SECONDS` | `30` | Quietud antes de procesar |
| `MAIL_INTERNO_DUA` | — | Destinatario(s) de los emails de revisión |
| `SMTP_HOST/PORT/USER/PASS`, `MAIL_FROM` | del stack | Envío de email (mismo relay) |

## Lado Windows (OneDrive → SharePoint)

1. El cliente añade vuestra cuenta de correo a su grupo de Teams y os pasa el link.
2. En un Windows Server interno: sesión iniciada con esa cuenta, OneDrive instalado
   y sincronizando esa biblioteca; OneDrive al arranque (desatendido).
3. Monta `\\<toolbox>\dua-outbox` y haz que su contenido llegue a la carpeta de
   OneDrive (montar la carpeta de OneDrive como destino, o un robocopy programado).
4. OneDrive sube a SharePoint automáticamente.

> Punto de mantenimiento: OneDrive puede pedir reautenticación cada cierto tiempo.
> Alternativa más robusta a futuro: Microsoft Graph API con app en Azure.

## Gestión

```bash
journalctl -u dua-watcher -f      # logs
systemctl restart dua-watcher
```

## Dependencias de sistema

`tesseract-ocr` y `poppler-utils` (los instala el script). Opcional:
`tesseract-ocr-spa` para mejor OCR de texto español (el DUA en sí no lo necesita).
