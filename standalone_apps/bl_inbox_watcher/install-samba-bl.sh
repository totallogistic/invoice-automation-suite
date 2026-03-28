#!/usr/bin/env bash
# install-samba-bl.sh
# Crea un share Samba de solo escritura para el inbox de BL
# Uso: sudo ./install-samba-bl.sh
#
# Resultado: \\<IP>\bl-inbox  — accesible desde Windows, Mac, escáner
# El scanner/PC copia PDFs aquí → el watcher los procesa automáticamente

set -euo pipefail

BL_INBOX="/data/ias_prod/data/bl/inbox"
SAMBA_USER="${1:-blscan}"          # usuario Samba (pasar como arg o usar default)
SHARE_NAME="bl-inbox"
SMB_CONF="/etc/samba/smb.conf"

echo "=== Samba BL Inbox Setup ==="
echo "  Carpeta : $BL_INBOX"
echo "  Share   : \\\\$(hostname -I | awk '{print $1}')\\$SHARE_NAME"
echo "  Usuario : $SAMBA_USER"
echo ""

# 1. Instalar Samba si no está
if ! command -v smbd &>/dev/null; then
  echo "[1/5] Instalando Samba..."
  apt-get update -qq && apt-get install -y samba
else
  echo "[1/5] Samba ya instalado — OK"
fi

# 2. Asegurar que la carpeta existe y tiene permisos correctos
echo "[2/5] Preparando carpeta $BL_INBOX..."
mkdir -p "$BL_INBOX"
# El grupo samba-share tendrá acceso; mpino (watcher) también debe estar en él
groupadd -f samba-share
usermod -aG samba-share mpino          # el usuario que corre el processor
chmod 775 "$BL_INBOX"
chgrp samba-share "$BL_INBOX"
# Sticky bit para que solo el owner pueda borrar sus propios ficheros
chmod g+s "$BL_INBOX"
echo "      OK"

# 3. Crear usuario del sistema para Samba (si no existe)
echo "[3/5] Creando usuario del sistema '$SAMBA_USER'..."
if ! id "$SAMBA_USER" &>/dev/null; then
  # Sin shell, sin home — solo para autenticación Samba
  useradd -M -s /sbin/nologin -G samba-share "$SAMBA_USER"
  echo "      Creado"
else
  # Asegurar que está en el grupo
  usermod -aG samba-share "$SAMBA_USER"
  echo "      Ya existe — añadido al grupo samba-share"
fi

# 4. Establecer contraseña Samba
echo "[4/5] Configurando contraseña Samba para '$SAMBA_USER'..."
echo "      Introduce la contraseña que usará el escáner/PC:"
smbpasswd -a "$SAMBA_USER"
smbpasswd -e "$SAMBA_USER"   # habilitar usuario

# 5. Añadir la sección al smb.conf si no existe ya
echo "[5/5] Configurando smb.conf..."

# Backup
cp "$SMB_CONF" "${SMB_CONF}.bak.$(date +%Y%m%d_%H%M%S)"

if grep -q "\[$SHARE_NAME\]" "$SMB_CONF"; then
  echo "      Share '$SHARE_NAME' ya existe en smb.conf — no se modifica"
else
  cat >> "$SMB_CONF" << EOF

[$SHARE_NAME]
   comment         = BL Inbox — deposita PDFs aqui
   path            = $BL_INBOX
   browseable      = yes
   read only       = no
   writable        = yes
   create mask     = 0664
   directory mask  = 0775
   force group     = samba-share
   valid users     = $SAMBA_USER
   ; Solo PDFs — el watcher ignora otros tipos pero mejor filtrar
   ; veto files    = /*.exe/*.bat/*.sh/
   ; delete veto files = no
EOF
  echo "      Sección [$SHARE_NAME] añadida"
fi

# Verificar configuración
testparm -s 2>/dev/null | grep -A8 "\[$SHARE_NAME\]" && echo "" || true

# Reiniciar Samba
systemctl restart smbd nmbd
systemctl enable smbd nmbd

HOST_IP=$(hostname -I | awk '{print $1}')

echo ""
echo "=== Instalación completada ==="
echo ""
echo "  Ruta de red  : \\\\${HOST_IP}\\${SHARE_NAME}"
echo "  Usuario      : $SAMBA_USER"
echo "  Carpeta local: $BL_INBOX"
echo ""
echo "  Conexión desde Windows:"
echo "    Ejecutar → \\\\${HOST_IP}\\${SHARE_NAME}"
echo ""
echo "  Conexión desde Mac:"
echo "    Finder → Ir → Conectar al servidor → smb://${HOST_IP}/${SHARE_NAME}"
echo ""
echo "  Config escáner (SCAN TO FOLDER):"
echo "    Protocolo : SMB"
echo "    Host/IP   : ${HOST_IP}"
echo "    Share     : ${SHARE_NAME}"
echo "    Usuario   : ${SAMBA_USER}"
echo "    Carpeta   : (dejar vacío o /)"
echo ""
echo "  Gestión:"
echo "    Ver usuarios  : sudo pdbedit -L"
echo "    Cambiar pass  : sudo smbpasswd $SAMBA_USER"
echo "    Ver conexiones: sudo smbstatus"
echo "    Logs          : sudo journalctl -u smbd -f"
echo ""
echo "  FIREWALL — si ufw está activo:"
echo "    sudo ufw allow samba"
echo ""