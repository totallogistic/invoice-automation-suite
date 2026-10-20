#!/usr/bin/env bash
# install-samba-dua.sh
# Crea DOS shares Samba para el flujo DUA:
#   - dua-inbox  : el escaner deja aqui los PDFs de DUAs (ENTRADA)
#   - dua-outbox : el <DUA>.pdf renombrado (SALIDA) -> el Windows con OneDrive
#                  sincroniza esta carpeta con SharePoint
# Uso: sudo ./install-samba-dua.sh [usuario_samba]

set -euo pipefail

BASE="/data/ias_prod/data/dua"
IN_DIR="$BASE/inbox"
OUT_DIR="$BASE/outbox"
SAMBA_USER="${1:-duascan}"
SMB_CONF="/etc/samba/smb.conf"

echo "=== Samba DUA Setup (inbox + outbox) ==="
echo "  Inbox  : $IN_DIR    -> share \\\\$(hostname -I | awk '{print $1}')\\dua-inbox"
echo "  Outbox : $OUT_DIR   -> share \\\\$(hostname -I | awk '{print $1}')\\dua-outbox"
echo "  Usuario: $SAMBA_USER"
echo ""

# 1. Samba
if ! command -v smbd &>/dev/null; then
  echo "[1/5] Instalando Samba..."; apt-get update -qq && apt-get install -y samba
else
  echo "[1/5] Samba ya instalado — OK"
fi

# 2. Carpetas + permisos
echo "[2/5] Preparando carpetas..."
mkdir -p "$IN_DIR" "$OUT_DIR"
groupadd -f samba-share
usermod -aG samba-share mpino
for d in "$IN_DIR" "$OUT_DIR"; do
  chmod 2775 "$d"          # 2 = setgid: lo nuevo hereda el grupo samba-share
  chgrp samba-share "$d"
done
echo "      OK"

# 3. Usuario del sistema
echo "[3/5] Usuario del sistema '$SAMBA_USER'..."
if ! id "$SAMBA_USER" &>/dev/null; then
  useradd -M -s /sbin/nologin -G samba-share "$SAMBA_USER"; echo "      Creado"
else
  usermod -aG samba-share "$SAMBA_USER"; echo "      Ya existe"
fi

# 4. Contraseña Samba
echo "[4/5] Contrasena Samba para '$SAMBA_USER'..."
smbpasswd -a "$SAMBA_USER"
smbpasswd -e "$SAMBA_USER"

# 5. smb.conf — dos secciones
echo "[5/5] Configurando smb.conf..."
cp "$SMB_CONF" "${SMB_CONF}.bak.$(date +%Y%m%d_%H%M%S)"

add_share() {
  local name="$1" path="$2" comment="$3"
  if grep -q "\[$name\]" "$SMB_CONF"; then
    echo "      [$name] ya existe — no se modifica"
  else
    cat >> "$SMB_CONF" <<EOF

[$name]
   comment         = $comment
   path            = $path
   browseable      = yes
   read only       = no
   writable        = yes
   create mask     = 0664
   directory mask  = 0775
   force group     = samba-share
   valid users     = $SAMBA_USER
EOF
    echo "      [$name] anadido"
  fi
}
add_share "dua-inbox"  "$IN_DIR"  "DUA Inbox — el escaner deja aqui los PDF"
add_share "dua-outbox" "$OUT_DIR" "DUA Outbox — <DUA>.pdf para SharePoint (OneDrive sync)"

systemctl restart smbd nmbd
systemctl enable smbd nmbd

HOST_IP=$(hostname -I | awk '{print $1}')
echo ""
echo "=== Instalacion completada ==="
echo ""
echo "  ENTRADA (escaner deja PDFs):  \\\\${HOST_IP}\\dua-inbox"
echo "  SALIDA  (OneDrive sincroniza): \\\\${HOST_IP}\\dua-outbox"
echo "  Usuario Samba: $SAMBA_USER"
echo ""
echo "  En el Windows Server con OneDrive:"
echo "    1. Monta  \\\\${HOST_IP}\\dua-outbox  como unidad de red."
echo "    2. OneDrive sincroniza la biblioteca de Teams/SharePoint del cliente."
echo "    3. Un robocopy/script (o la propia unidad) vuelca dua-outbox -> carpeta OneDrive,"
echo "       o monta la carpeta de OneDrive directamente como destino del outbox."
echo ""
echo "  FIREWALL (si ufw activo): sudo ufw allow samba"
echo ""
