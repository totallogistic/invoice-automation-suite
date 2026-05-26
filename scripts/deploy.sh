#!/bin/bash
# deploy.sh — script de deploy para el stack invoice-automation-suite
#
# Uso:
#   ./scripts/deploy.sh --env prod
#   ./scripts/deploy.sh --env dev
#   ./scripts/deploy.sh --env staging --no-cache=false
#
# Lo que hace:
#   1. Verifica que el .env existe
#   2. Build de imágenes (--no-cache por defecto, override con --no-cache=false)
#   3. Up -d (recrea containers con imágenes nuevas)
#   4. Restart tools_web (CRÍTICO: refresca DNS upstream tras rebuild del API)
#   5. Smoke test de endpoints clave
#   6. Reporta estado final

set -euo pipefail

# ── Parsing de argumentos ──────────────────────────────────────────────────

ENV=""
NO_CACHE="true"

while [[ $# -gt 0 ]]; do
    case "$1" in
        --env)
            ENV="$2"
            shift 2
            ;;
        --env=*)
            ENV="${1#*=}"
            shift
            ;;
        --no-cache)
            NO_CACHE="true"
            shift
            ;;
        --no-cache=*)
            NO_CACHE="${1#*=}"
            shift
            ;;
        -h|--help)
            cat << EOF
Uso: $0 --env <prod|dev|staging|...> [--no-cache=true|false]

Argumentos:
  --env ENV         Entorno a deployar. Busca .env.\$ENV en la raíz del repo.
                    El proyecto docker compose se nombra ias_\$ENV.

  --no-cache BOOL   true (default) → build --no-cache (lento pero limpio)
                    false → build con cache (rápido para iteración)

Ejemplos:
  $0 --env prod
  $0 --env dev --no-cache=false

EOF
            exit 0
            ;;
        *)
            echo "❌ Argumento desconocido: $1"
            echo "Usa --help para ver opciones."
            exit 1
            ;;
    esac
done

# ── Validación ─────────────────────────────────────────────────────────────

if [ -z "$ENV" ]; then
    echo "❌ Falta --env. Usa --help para ver opciones."
    exit 1
fi

# Localizar raíz del repo (asume script en scripts/ del repo)
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$REPO_ROOT"

ENV_FILE=".env.${ENV}"
PROJECT="ias_${ENV}"

if [ ! -f "$ENV_FILE" ]; then
    echo "❌ Error: $ENV_FILE no encontrado en $REPO_ROOT"
    echo "   ¿Estás seguro de que '$ENV' es un entorno válido?"
    exit 1
fi

# Determinar puerto del API según .env (default 8081 que es el de nginx tools_web)
TOOLS_WEB_PORT=$(grep -E '^TOOLS_WEB_PORT=' "$ENV_FILE" | head -1 | cut -d= -f2- || echo "8081")
TOOLS_WEB_PORT="${TOOLS_WEB_PORT:-8081}"

echo ""
echo "═══════════════════════════════════════════════════════════"
echo " 🚀 Deploy → entorno: $ENV"
echo "═══════════════════════════════════════════════════════════"
echo "  Repo:      $REPO_ROOT"
echo "  Env file:  $ENV_FILE"
echo "  Project:   $PROJECT"
echo "  Web port:  $TOOLS_WEB_PORT"
echo "  No-cache:  $NO_CACHE"
echo ""

# ── 1. Build ───────────────────────────────────────────────────────────────

echo "[1/4] 🔨 Build de imágenes..."
if [ "$NO_CACHE" = "true" ]; then
    docker compose --env-file "$ENV_FILE" -p "$PROJECT" build --no-cache
else
    docker compose --env-file "$ENV_FILE" -p "$PROJECT" build
fi

# ── 2. Up ─────────────────────────────────────────────────────────────────

echo ""
echo "[2/4] 🟢 Up containers..."
docker compose --env-file "$ENV_FILE" -p "$PROJECT" up -d

# ── 3. Restart tools_web (CRÍTICO) ────────────────────────────────────────
#
# Razón: nginx resuelve hostnames del proxy_pass al ARRANCAR, no en cada
# request. Cuando unified_api se rebuildea cambia de IP en la red docker.
# tools_web (que usa nginx:alpine, no rebuildeado) sigue apuntando a la IP
# vieja → connection refused → 502 Bad Gateway.
# 
# El restart fuerza re-resolución DNS y arregla el upstream.
# Ver incidente del 2026-04-30 para más detalles.

echo ""
echo "[3/4] 🔁 Restart tools_web (refresca DNS upstream)..."
docker compose --env-file "$ENV_FILE" -p "$PROJECT" restart tools_web

# ── 4. Smoke test ─────────────────────────────────────────────────────────

echo ""
echo "[4/4] ⏳ Esperando estabilización..."
sleep 5

echo ""
echo "═══════════════════════════════════════════════════════════"
echo " 🩺 Smoke test"
echo "═══════════════════════════════════════════════════════════"

failed=0

# Test 1: API health endpoint
echo -n "  /health           → "
code=$(curl -s -o /dev/null -w "%{http_code}" "http://localhost:${TOOLS_WEB_PORT}/api/camion/version" 2>/dev/null || echo "000")
if [ "$code" = "200" ]; then
    echo "✅ HTTP $code"
else
    echo "❌ HTTP $code (esperado 200)"
    failed=$((failed + 1))
fi

# Test 2: Upload endpoint rechaza archivo inválido
echo -n "  /batches (invalid)→ "
code=$(curl -s -o /dev/null -w "%{http_code}" -F "files=@/etc/hostname" "http://localhost:${TOOLS_WEB_PORT}/api/camion/batches" 2>/dev/null || echo "000")
if [ "$code" = "400" ]; then
    echo "✅ HTTP $code (rechazo correcto)"
elif [ "$code" = "200" ]; then
    echo "⚠️  HTTP $code (aceptó hostname como input — revisar validaciones)"
else
    echo "❌ HTTP $code (esperado 400)"
    failed=$((failed + 1))
fi

echo ""
echo "═══════════════════════════════════════════════════════════"
echo " 📋 Estado containers"
echo "═══════════════════════════════════════════════════════════"
docker compose --env-file "$ENV_FILE" -p "$PROJECT" ps

# ── Resultado final ───────────────────────────────────────────────────────

echo ""
if [ $failed -eq 0 ]; then
    echo "═══════════════════════════════════════════════════════════"
    echo " ✅ Deploy de '$ENV' completado correctamente"
    echo "═══════════════════════════════════════════════════════════"
    exit 0
else
    echo "═══════════════════════════════════════════════════════════"
    echo " ⚠️  Deploy completado pero $failed smoke test(s) falló"
    echo "═══════════════════════════════════════════════════════════"
    echo ""
    echo "Revisa logs:"
    echo "  docker compose --env-file $ENV_FILE -p $PROJECT logs --tail=80 unified_api"
    echo "  docker compose --env-file $ENV_FILE -p $PROJECT logs --tail=80 unified_processor"
    echo "  docker compose --env-file $ENV_FILE -p $PROJECT logs --tail=80 tools_web"
    exit 1
fi
