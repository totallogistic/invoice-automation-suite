#!/usr/bin/env bash
# install-samba-merge-pdf.sh
# Crea un share Samba de escritura para el inbox de merge_pdf.
# Uso: sudo ./install-samba-merge-pdf.sh [usuario_samba]
#
# Resultado: \\<IP>\merge-pdf-inbox  — el escaner/PC deja PDFs aqui
#            → el watcher los empaqueta → el processor los consolida.

set -euo pipefail

MERGE_INBOX="/data/ias_prod/data/merge_pdf/inbox"
SAMBA_USER="${1:-mergescan}"
SHARE_NAME="merge-pdf-inbox"
SMB_CONF="/etc/samba/smb.conf"

echo "=== Samba merge_pdf Inbox Setup ==="
echo "  Carpeta : $MERGE_INBOX"
echo "  Share   : \\\\$(hostname -I | awk '{print $1}')\\$SHARE_NAME"
echo "  Usuario : $SAMBA_USER"
echo ""

# 1. Instalar Samba si no esta
if ! command -v smbd &>/dev/null; then
  echo "[1/5] Instalando Samba..."
  apt-get update -qq && apt-get install -y samba
else
  echo "[1/5] Samba ya instalado — OK"
fi

# 2. Carpeta y permisos
echo "[2/5] Preparando carpeta $MERGE_INBOX..."
mkdir -p "$MERGE_INBOX"
groupadd -f samba-share
usermod -aG samba-share mpino          # usuario que corre watcher/processor
chmod 775 "$MERGE_INBOX"
chgrp samba-share "$MERGE_INBOX"
chmod g+s "$MERGE_INBOX"
echo "      OK"

# 3. Usuario del sistema para Samba
echo "[3/5] Creando usuario del sistema '$SAMBA_USER'..."
if ! id "$SAMBA_USER" &>/dev/null; then
  useradd -M -s /sbin/nologin -G samba-share "$SAMBA_USER"
  echo "      Creado"
else
  usermod -aG samba-share "$SAMBA_USER"
  echo "      Ya existe — anadido al grupo samba-share"
fi

# 4. Contrasena Samba
echo "[4/5] Configurando contrasena Samba para '$SAMBA_USER'..."
echo "      Introduce la contrasena que usara el escaner/PC:"
smbpasswd -a "$SAMBA_USER"
smbpasswd -e "$SAMBA_USER"

# 5. smb.conf
echo "[5/5] Configurando smb.conf..."
cp "$SMB_CONF" "${SMB_CONF}.bak.$(date +%Y%m%d_%H%M%S)"

if grep -q "\[$SHARE_NAME\]" "$SMB_CONF"; then
  echo "      Share '$SHARE_NAME' ya existe en smb.conf — no se modifica"
else
  cat >> "$SMB_CONF" << EOF

[$SHARE_NAME]
   comment         = merge_pdf Inbox — deposita PDFs a consolidar aqui
   path            = $MERGE_INBOX
   browseable      = yes
   read only       = no
   writable        = yes
   create mask     = 0664
   directory mask  = 0775
   force group     = samba-share
   valid users     = $SAMBA_USER
EOF
  echo "      Seccion [$SHARE_NAME] anadida"
fi

testparm -s 2>/dev/null | grep -A8 "\[$SHARE_NAME\]" && echo "" || true

systemctl restart smbd nmbd
systemctl enable smbd nmbd

HOST_IP=$(hostname -I | awk '{print $1}')

echo ""
echo "=== Instalacion completada ==="
echo ""
echo "  Ruta de red  : \\\\${HOST_IP}\\${SHARE_NAME}"
echo "  Usuario      : $SAMBA_USER"
echo "  Carpeta local: $MERGE_INBOX"
echo ""
echo "  Windows: Ejecutar → \\\\${HOST_IP}\\${SHARE_NAME}"
echo "  Mac    : Finder → Ir → Conectar al servidor → smb://${HOST_IP}/${SHARE_NAME}"
echo ""
echo "  Escaner (SCAN TO FOLDER): SMB · Host ${HOST_IP} · Share ${SHARE_NAME} · Usuario ${SAMBA_USER}"
echo ""
echo "  IMPORTANTE: deja los PDFs sueltos en la RAIZ del share. El watcher espera"
echo "  ${QUIET:-15}s de quietud y los empaqueta en un batch automaticamente."
echo ""
echo "  FIREWALL (si ufw activo): sudo ufw allow samba"
echo ""
