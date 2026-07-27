#!/usr/bin/env bash
#
# Modo TEST del DeCA con un bucket MinIO local (S3-compatible, cero coste).
# Levanta MinIO, crea el bucket y le pone lectura anónima SOLO en el prefijo d/*,
# e imprime las líneas de .env.prod para pegar. Idempotente: se puede relanzar.
#
# Uso:  bash deca-minio-test.sh
#
set -euo pipefail

BUCKET="${DECA_BUCKET:-deca-prod}"
MINIO_USER="${DECA_MINIO_USER:-decatest}"
MINIO_PASS="${DECA_MINIO_PASS:-decatest123}"
S3_PORT="${DECA_S3_PORT:-9000}"
CONSOLE_PORT="${DECA_CONSOLE_PORT:-9001}"
HOST_IP="$(hostname -I | awk '{print $1}')"

command -v docker >/dev/null || { echo "❌ Falta docker"; exit 1; }

echo "▶ Levantando MinIO (si no está ya)..."
if ! docker ps --format '{{.Names}}' | grep -qx 'deca-minio'; then
  docker rm -f deca-minio >/dev/null 2>&1 || true
  docker run -d --name deca-minio \
    -p "${S3_PORT}:9000" -p "${CONSOLE_PORT}:9001" \
    -e MINIO_ROOT_USER="${MINIO_USER}" \
    -e MINIO_ROOT_PASSWORD="${MINIO_PASS}" \
    -v deca-minio-data:/data \
    quay.io/minio/minio server /data --console-address ":9001" >/dev/null
fi

echo "▶ Esperando a que MinIO responda..."
for _ in $(seq 1 30); do
  curl -sf "http://localhost:${S3_PORT}/minio/health/ready" >/dev/null && break
  sleep 1
done

echo "▶ Creando bucket '${BUCKET}' + lectura anónima en d/* ..."
docker run --rm --network host --entrypoint /bin/sh quay.io/minio/mc -c "
  mc alias set local 'http://localhost:${S3_PORT}' '${MINIO_USER}' '${MINIO_PASS}' >/dev/null &&
  mc mb --ignore-existing local/${BUCKET} >/dev/null &&
  mc anonymous set download local/${BUCKET}/d/ >/dev/null &&
  echo '  ✓ bucket y política listos'
"

cat <<EOF

✅ MinIO de test listo.
   Consola web:  http://${HOST_IP}:${CONSOLE_PORT}   (user: ${MINIO_USER} / pass: ${MINIO_PASS})

1) Añade estas líneas a  .env.prod  (raíz del repo):

DECA_S3_ENDPOINT=http://${HOST_IP}:${S3_PORT}
DECA_S3_BUCKET=${BUCKET}
DECA_S3_KEY=${MINIO_USER}
DECA_S3_SECRET=${MINIO_PASS}
DECA_PUBLIC_BASEURL=http://${HOST_IP}:${S3_PORT}/${BUCKET}
DECA_URL_TTL_DAYS=7
DECA_RETENTION_DIR=/home/mpino/repos/invoice-automation-suite/standalone_apps/form_generator/_deca_retencion

2) Reinicia el servicio y prueba:

   sudo systemctl restart form-generator-prod
   # abre http://${HOST_IP}:8201/form/deca , genera un DeCA
   # y escanea el QR con el móvil (misma red) → debe descargar el PDF

Para parar/borrar el MinIO de test:
   docker rm -f deca-minio && docker volume rm deca-minio-data
EOF