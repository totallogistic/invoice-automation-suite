# merge_pdf_watcher

Watcher standalone que vigila el inbox de `merge_pdf` y monta batches
automaticamente para que el `unified_processor` los consolide, sin pasar por la
web UI. Pensado para alimentar por **Samba** (escaner / ERP via unidad de red).

## Flujo

```
ERP/escaner --(Samba)--> inbox/  --(watcher, tras 30s de quietud)-->
   inbox/<YYYYmmdd_HHMMSS_merge_pdf>/ + _DONE  --> unified_processor consolida
```

No copia `rules.json`: el extractor usa `default_rules.json` (reglas por defecto
del servidor).

## Robustez (escrituras SMB de Windows)

Antes de empaquetar, el watcher exige TRES condiciones:
1. **Quietud**: 30s sin cambios de `mtime` (configurable).
2. **Estabilidad de tamano**: ningun fichero crecio entre ciclos.
3. **PDF completo**: cada PDF abre con pypdf y tiene >=1 pagina.

Ademas ignora ocultos (`._*`, `.*`) y temporales de Windows/ERP
(`.tmp`, `.crdownload`, `.part`, `~$*`).

## Instalacion

```bash
# 1. Dependencias
pip install --user -r requirements.txt          # o --break-system-packages en Ubuntu 24

# 2. Servicio systemd
sudo install/install-merge-pdf-watcher.sh

# 3. Share Samba (opcional, si se alimenta por red)
sudo install/install-samba-merge-pdf.sh mergescan
```

## Gestion

```bash
journalctl -u merge-pdf-watcher -f      # logs
systemctl restart merge-pdf-watcher     # reiniciar
systemctl stop merge-pdf-watcher        # parar
```

## Configuracion (env o flags)

| Variable | Flag | Default | Descripcion |
|---|---|---|---|
| `MERGE_PDF_INBOX` | `--inbox` | `/data/ias_prod/data/merge_pdf/inbox` | Carpeta vigilada |
| `MERGE_PDF_QUIET_SECONDS` | `--quiet` | `30` | Segundos de quietud antes de empaquetar |
| `MERGE_PDF_POLL_SECONDS` | `--poll` | `5` | Intervalo de comprobacion |
