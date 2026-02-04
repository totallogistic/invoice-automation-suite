# Guía de Despliegue - Invoice Automation Suite

Esta guía te ayudará a desplegar el stack completo de Invoice Automation Suite en una nueva máquina Ubuntu.

## Tabla de Contenidos
- [Pre-requisitos](#pre-requisitos)
- [Instalación Automática](#instalación-automática)
- [Instalación Manual](#instalación-manual)
- [Configuración](#configuración)
- [Verificación](#verificación)
- [Troubleshooting](#troubleshooting)

---

## Pre-requisitos

### Sistema Operativo
- Ubuntu 20.04 LTS o superior
- Usuario con permisos sudo
- Acceso a Internet

### Requisitos Mínimos de Hardware
- CPU: 2 cores
- RAM: 4 GB
- Disco: 20 GB libres (más espacio según volumen de facturas)

---

## Instalación Automática

### Opción 1: Script de Despliegue Rápido

```bash
# 1. Clonar el repositorio (si no lo tienes ya)
git clone https://github.com/totallogistic/invoice-automation-suite.git
cd invoice-automation-suite

# 2. Ejecutar el script de instalación
chmod +x deploy.sh
sudo ./deploy.sh
```

El script realizará automáticamente:
- Instalación de Docker y Docker Compose
- Creación de directorios necesarios
- Configuración de permisos
- Setup inicial del entorno

---

## Instalación Manual

### Paso 1: Instalar Docker

```bash
# Actualizar paquetes del sistema
sudo apt update
sudo apt upgrade -y

# Instalar dependencias necesarias
sudo apt install -y apt-transport-https ca-certificates curl software-properties-common

# Añadir clave GPG oficial de Docker
curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /usr/share/keyrings/docker-archive-keyring.gpg

# Añadir repositorio de Docker
echo "deb [arch=$(dpkg --print-architecture) signed-by=/usr/share/keyrings/docker-archive-keyring.gpg] https://download.docker.com/linux/ubuntu $(lsb_release -cs) stable" | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null

# Instalar Docker
sudo apt update
sudo apt install -y docker-ce docker-ce-cli containerd.io

# Verificar instalación
sudo docker --version
```

### Paso 2: Instalar Docker Compose

```bash
# Instalar Docker Compose (v2 viene con Docker Desktop, pero en servidor lo instalamos por separado)
sudo apt install -y docker-compose-plugin

# O si prefieres la versión standalone:
sudo curl -L "https://github.com/docker/compose/releases/latest/download/docker-compose-$(uname -s)-$(uname -m)" -o /usr/local/bin/docker-compose
sudo chmod +x /usr/local/bin/docker-compose

# Verificar instalación
docker compose version
```

### Paso 3: Configurar Permisos de Docker

```bash
# Añadir tu usuario al grupo docker (para no usar sudo)
sudo usermod -aG docker $USER

# Aplicar cambios (o cerrar sesión y volver a entrar)
newgrp docker

# Verificar que funciona sin sudo
docker ps
```

### Paso 4: Crear Estructura de Directorios

```bash
# Crear directorio raíz para datos persistentes
sudo mkdir -p /data/ias_prod

# Crear subdirectorios para producción
sudo mkdir -p /data/ias_prod/data/{inbox,processing,out,processed,error,status}
sudo mkdir -p /data/ias_prod/sftpgo
sudo mkdir -p /data/ias_prod/sftpgo_var

# (Opcional) Para entorno de pre-producción
sudo mkdir -p /data/ias_pre/data/{inbox,processing,out,processed,error,status}
sudo mkdir -p /data/ias_pre/sftpgo
sudo mkdir -p /data/ias_pre/sftpgo_var

# Configurar permisos adecuados
sudo chown -R $USER:$USER /data/ias_prod
sudo chown -R $USER:$USER /data/ias_pre  # si creaste pre

# Dar permisos de escritura
sudo chmod -R 755 /data/ias_prod
sudo chmod -R 755 /data/ias_pre  # si creaste pre
```

### Paso 5: Instalar Herramientas Adicionales (Opcional pero Recomendado)

```bash
# jq para formatear JSON en CLI
sudo apt install -y jq

# tree para visualizar estructura de directorios
sudo apt install -y tree

# htop para monitorear recursos
sudo apt install -y htop
```

---

## Configuración

### Paso 1: Clonar el Repositorio (si no lo has hecho)

```bash
cd ~
git clone https://github.com/totallogistic/invoice-automation-suite.git
cd invoice-automation-suite
```

### Paso 2: Configurar Variables de Entorno

```bash
# Copiar archivo de ejemplo
cp .env.example .env

# Editar con tu editor favorito
nano .env
# o
vim .env
```

**Variables Importantes a Configurar:**

```bash
# --- SMTP / Email ---
# Configuración para envío de emails con resultados
EMAIL_MODE=BATCH_ONLY
SMTP_HOST=smtp.gmail.com              # Tu servidor SMTP
SMTP_PORT=587                         # Puerto (587 para TLS, 465 para SSL)
SMTP_USER=tu_usuario@gmail.com        # Usuario SMTP
SMTP_PASS=tu_contraseña_app           # Contraseña o App Password
MAIL_FROM=sender@domain.com           # Email del remitente
MAIL_TO=recipient@domain.com          # Email(s) destinatarios (separados por comas)

# --- Watcher / Processor ---
DONE_MARKER=_DONE                     # Marker para indicar lote completo
POLL_SECONDS=3                        # Frecuencia de chequeo (segundos)
BATCH_QUIET_SECONDS=300               # Tiempo de inactividad para cerrar lote (segundos)

# --- SFTPGo Admin (primera vez) ---
SFTPGO_DEFAULT_ADMIN_USERNAME=admin   # Usuario administrador
SFTPGO_DEFAULT_ADMIN_PASSWORD=CHANGE_ME_STRONG_PASSWORD  # ⚠️ CAMBIAR!

# --- Puertos (opcional, ajustar si hay conflictos) ---
SFTP_PORT=2222                        # Puerto SFTP
SFTPGO_WEB_PORT=8080                  # Puerto web admin SFTPGo
TOOLS_WEB_PORT=8081                   # Puerto web principal

# --- Paths de datos ---
STACK_ROOT=/data/ias_prod             # Raíz de datos
DATA_ROOT=/data/ias_prod/data         # Datos de procesamiento
```

**⚠️ Importante para Gmail:**
- Si usas Gmail, necesitas crear una "Contraseña de Aplicación" en vez de tu contraseña normal
- Activa la verificación en dos pasos: https://myaccount.google.com/security
- Genera una contraseña de app: https://myaccount.google.com/apppasswords

### Paso 3: Construir e Iniciar los Servicios

```bash
# Asegúrate de estar en el directorio del proyecto
cd ~/invoice-automation-suite

# Construir e iniciar todos los servicios
docker compose up -d --build

# Ver el estado de los contenedores
docker compose ps

# Ver logs en tiempo real (Ctrl+C para salir)
docker compose logs -f
```

### Paso 4: Configurar SFTPGo (Primera Vez)

```bash
# 1. Acceder a la interfaz web de SFTPGo
# Abre en tu navegador: http://<IP_DE_TU_SERVIDOR>:8080

# 2. Iniciar sesión con las credenciales del .env
# Usuario: admin (o el que configuraste en SFTPGO_DEFAULT_ADMIN_USERNAME)
# Password: (el que pusiste en SFTPGO_DEFAULT_ADMIN_PASSWORD)

# 3. Crear usuario SFTP para Lear Cable
# - Ir a "Users" > "Add"
# - Username: lear_cable
# - Password: <contraseña_segura>
# - Home directory: /srv/sftpgo/lear_cable/data
# - Permissions: Leer, Escribir, Listar, Crear directorios
# - Save

# 4. (Opcional) Cambiar contraseña del admin por seguridad
```

---

## Verificación

### Verificación Básica

```bash
# 1. Verificar que todos los contenedores están corriendo
docker compose ps
# Todos deben mostrar estado "Up"

# 2. Verificar salud de la API
curl -sS http://localhost:8081/api/lear_cable/health | jq

# Respuesta esperada:
# {
#   "status": "healthy",
#   "service": "api_lear_cable",
#   "version": "1.0.0"
# }

# 3. Verificar UI
curl -sS http://localhost:8081/tools/lear_cable/ | head -n 20
# Debe devolver HTML

# 4. Verificar directorios de datos
ls -la /data/ias_prod/data/
# Debe mostrar: inbox, processing, out, processed, error, status
```

### Verificación de Conectividad

```bash
# Verificar puertos abiertos
sudo ss -lntup | grep -E ':(2222|8080|8081)'

# Debe mostrar:
# *:2222  (SFTP)
# *:8080  (SFTPGo Web)
# *:8081  (Tools Web)
```

### Prueba de Subida de Archivo

```bash
# Opción 1: Via API (crea un ZIP de prueba primero)
curl -sS \
  -F "file=@test.zip;type=application/zip" \
  http://localhost:8081/api/lear_cable/batches | jq

# Opción 2: Via UI
# Abre: http://<IP_SERVIDOR>:8081/tools/lear_cable/
# Sube un ZIP con PDFs de facturas
```

### Verificación de Logs

```bash
# Ver logs de cada servicio
docker compose logs --tail 50 tools_web
docker compose logs --tail 50 api_lear_cable
docker compose logs --tail 50 processor_lear_cable
docker compose logs --tail 50 sftp_lear_cable

# Ver logs en tiempo real de un servicio específico
docker compose logs -f processor_lear_cable
```

---

## Acceso a los Servicios

Una vez desplegado, tendrás acceso a:

| Servicio | URL | Descripción |
|----------|-----|-------------|
| **Web UI** | `http://<IP>:8081/tools/lear_cable/` | Interfaz para subir ZIPs |
| **API** | `http://<IP>:8081/api/lear_cable/` | API REST para integración |
| **SFTPGo Admin** | `http://<IP>:8080/` | Administración de usuarios SFTP |
| **SFTP** | `sftp://<IP>:2222` | Acceso SFTP para subir archivos |

---

## Mantenimiento

### Actualizar el Stack

```bash
# 1. Obtener últimos cambios
cd ~/invoice-automation-suite
git pull

# 2. Reconstruir e reiniciar servicios
docker compose up -d --build --force-recreate

# 3. Verificar que todo funciona
docker compose ps
curl -sS http://localhost:8081/api/lear_cable/health | jq
```

### Backups

```bash
# Backup de datos
sudo tar -czf backup_ias_prod_$(date +%Y%m%d_%H%M%S).tar.gz /data/ias_prod

# Backup solo de outputs procesados
sudo tar -czf backup_outputs_$(date +%Y%m%d_%H%M%S).tar.gz /data/ias_prod/data/out /data/ias_prod/data/processed

# Restaurar backup
sudo tar -xzf backup_ias_prod_YYYYMMDD_HHMMSS.tar.gz -C /
```

### Limpieza de Datos Antiguos

```bash
# Eliminar lotes procesados hace más de 30 días
find /data/ias_prod/data/processed -type d -mtime +30 -exec rm -rf {} +

# Eliminar logs antiguos de Docker
docker system prune -a --volumes
```

### Monitoreo

```bash
# Ver uso de recursos
docker stats

# Ver espacio en disco
df -h /data/ias_prod

# Ver logs de errores recientes
docker compose logs --tail 100 | grep -i error
```

### Reiniciar Servicios

```bash
# Reiniciar todos los servicios
docker compose restart

# Reiniciar un servicio específico
docker compose restart processor_lear_cable

# Detener todos los servicios
docker compose down

# Iniciar servicios
docker compose up -d
```

---

## Troubleshooting

### Problema: Los contenedores no inician

```bash
# Verificar logs de error
docker compose logs

# Verificar puertos en uso
sudo ss -lntup | grep -E ':(2222|8080|8081)'

# Si hay conflicto de puertos, edita .env y cambia los puertos
nano .env
# Luego reinicia
docker compose down
docker compose up -d --build
```

### Problema: Error 502 en la API

```bash
# Verificar que la API está corriendo
docker compose ps api_lear_cable

# Ver logs de la API
docker compose logs --tail 100 api_lear_cable

# Reiniciar API
docker compose restart api_lear_cable
```

### Problema: No se procesan los lotes

```bash
# Verificar el processor
docker compose logs --tail 100 processor_lear_cable

# Verificar permisos de directorios
ls -la /data/ias_prod/data/

# Verificar que existen los directorios necesarios
docker exec -it processor_lear_cable ls -la /data
```

### Problema: No se envían emails

```bash
# Verificar configuración SMTP en .env
cat .env | grep SMTP

# Ver logs del processor (quien envía los emails)
docker compose logs --tail 200 processor_lear_cable | grep -i email

# Probar conexión SMTP manualmente
telnet <SMTP_HOST> <SMTP_PORT>
```

Para más detalles, consulta [TROUBLESHOOTING.md](./TROUBLESHOOTING.md)

---

## Seguridad

### Recomendaciones de Seguridad

1. **Cambiar contraseñas por defecto**
   ```bash
   # Edita .env y cambia SFTPGO_DEFAULT_ADMIN_PASSWORD
   nano .env
   ```

2. **Configurar firewall**
   ```bash
   # Permitir solo puertos necesarios
   sudo ufw allow 22/tcp     # SSH
   sudo ufw allow 2222/tcp   # SFTP
   sudo ufw allow 8080/tcp   # SFTPGo (solo si necesitas acceso externo)
   sudo ufw allow 8081/tcp   # Web UI
   sudo ufw enable
   ```

3. **Usar HTTPS en producción**
   - Considera usar un reverse proxy como Nginx o Traefik con certificados SSL/TLS
   - Herramientas recomendadas: Let's Encrypt + Certbot

4. **Restringir acceso a SFTPGo Admin**
   - Solo permitir acceso desde IPs específicas
   - O usar VPN para acceso seguro

5. **Rotación de credenciales**
   - Cambiar contraseñas SMTP y SFTPGo periódicamente
   - Mantener el .env seguro (nunca lo commitees a git)

---

## Entornos Múltiples (Pre-producción y Producción)

Si quieres tener dos entornos separados:

```bash
# Producción (puerto 8081)
docker compose --env-file env/prod.env up -d

# Pre-producción (puerto 8082 o diferente)
docker compose --env-file env/pre.env -p ias_pre up -d

# Ver todos los contenedores
docker ps
```

**Nota:** Asegúrate de que cada entorno tenga:
- Puertos diferentes (SFTP_PORT, SFTPGO_WEB_PORT, TOOLS_WEB_PORT)
- Directorios de datos diferentes (STACK_ROOT, DATA_ROOT)
- Nombres de proyecto diferentes (COMPOSE_PROJECT_NAME)

---

## Migración desde Otro Servidor

Si ya tienes el stack corriendo en otra máquina (por ejemplo "toolboox") y quieres migrar:

### 1. En el servidor antiguo (toolboox):

```bash
# Hacer backup de datos
cd /data/ias_prod
sudo tar -czf ~/ias_prod_backup_$(date +%Y%m%d).tar.gz .

# Exportar configuración
cd ~/invoice-automation-suite
cp .env ~/env_backup
```

### 2. En el servidor nuevo:

```bash
# Seguir esta guía de instalación completa
# Luego restaurar datos:

# Copiar backup desde servidor antiguo
scp user@toolboox:~/ias_prod_backup_YYYYMMDD.tar.gz ~/
scp user@toolboox:~/env_backup ~/.env_from_toolboox

# Restaurar datos
sudo tar -xzf ~/ias_prod_backup_YYYYMMDD.tar.gz -C /data/ias_prod/

# Copiar configuración (revisar antes de aplicar)
cp ~/.env_from_toolboox ~/invoice-automation-suite/.env

# Ajustar permisos
sudo chown -R $USER:$USER /data/ias_prod

# Iniciar servicios
cd ~/invoice-automation-suite
docker compose up -d --build
```

---

## Soporte

Para problemas o preguntas:
- Revisa [TROUBLESHOOTING.md](./TROUBLESHOOTING.md)
- Revisa [USAGE.md](./USAGE.md) para instrucciones de uso
- Contacta al equipo de desarrollo

---

## Apéndice: Comandos Útiles

```bash
# Ver versión de Docker
docker --version
docker compose version

# Ver todos los contenedores (incluso parados)
docker ps -a

# Ver uso de espacio de Docker
docker system df

# Limpiar recursos no usados
docker system prune -a

# Ver redes de Docker
docker network ls

# Ver volúmenes de Docker
docker volume ls

# Ejecutar comando en contenedor
docker exec -it processor_lear_cable sh

# Copiar archivos desde/hacia contenedor
docker cp processor_lear_cable:/data/out/batch_123/invoices_extracted.json ./
docker cp ./file.txt processor_lear_cable:/data/inbox/

# Exportar logs a archivo
docker compose logs > logs_full_$(date +%Y%m%d).log
docker compose logs processor_lear_cable > logs_processor_$(date +%Y%m%d).log
```
