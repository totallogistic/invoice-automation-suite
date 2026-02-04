# Quick Start - Despliegue en Nueva Ubuntu

Esta es una guía rápida para desplegar el stack en una máquina Ubuntu nueva. 

Para instrucciones detalladas, consulta [DEPLOYMENT.md](./DEPLOYMENT.md).

---

## Opción 1: Despliegue Automático (Recomendado)

```bash
# 1. Clonar repositorio
git clone https://github.com/totallogistic/invoice-automation-suite.git
cd invoice-automation-suite

# 2. Ejecutar script de despliegue
chmod +x deploy.sh
sudo ./deploy.sh
```

El script instalará Docker, creará los directorios necesarios y te guiará en la configuración.

---

## Opción 2: Despliegue Manual Rápido

```bash
# 1. Instalar Docker (si no lo tienes)
curl -fsSL https://get.docker.com -o get-docker.sh
sudo sh get-docker.sh
sudo usermod -aG docker $USER
newgrp docker

# 2. Instalar Docker Compose
sudo apt install docker-compose-plugin -y

# 3. Clonar repositorio
git clone https://github.com/totallogistic/invoice-automation-suite.git
cd invoice-automation-suite

# 4. Crear directorios de datos
sudo mkdir -p /data/ias_prod/data/{inbox,processing,out,processed,error,status}
sudo mkdir -p /data/ias_prod/{sftpgo,sftpgo_var}
sudo chown -R $USER:$USER /data/ias_prod
sudo chmod -R 755 /data/ias_prod

# 5. Configurar variables de entorno
cp .env.example .env
nano .env  # Editar con tus valores (SMTP, contraseñas, etc.)

# 6. Iniciar servicios
docker compose up -d --build

# 7. Verificar
docker compose ps
curl http://localhost:8081/api/lear_cable/health
```

---

## Configuración Mínima Requerida

Antes de iniciar los servicios, debes editar `.env` y configurar:

### 1. Configuración SMTP (para envío de emails)
```bash
SMTP_HOST=smtp.gmail.com              # Tu servidor SMTP
SMTP_PORT=587                         # Puerto (587 para Gmail con TLS)
SMTP_USER=tu_email@gmail.com          # Tu email
SMTP_PASS=tu_app_password             # App Password de Gmail
MAIL_FROM=tu_email@gmail.com          # Remitente
MAIL_TO=destinatario@domain.com       # Destinatario(s)
```

**Para Gmail:** Necesitas crear una "App Password":
1. Activa verificación en 2 pasos: https://myaccount.google.com/security
2. Genera App Password: https://myaccount.google.com/apppasswords
3. Usa esa contraseña en `SMTP_PASS`

### 2. Contraseña de Administrador SFTPGo
```bash
SFTPGO_DEFAULT_ADMIN_PASSWORD=una_contraseña_segura_aqui
```

⚠️ **IMPORTANTE:** Cambia `CHANGE_ME` por una contraseña segura.

---

## Acceso a los Servicios

Una vez desplegado:

| Servicio | URL | Credenciales |
|----------|-----|--------------|
| **Web UI** | http://localhost:8081/tools/lear_cable/ | No requiere |
| **API** | http://localhost:8081/api/lear_cable/ | No requiere |
| **SFTPGo Admin** | http://localhost:8080/ | admin / (tu SFTPGO_DEFAULT_ADMIN_PASSWORD) |
| **SFTP** | sftp://localhost:2222 | (configurar usuario en SFTPGo primero) |

---

## Próximos Pasos

1. **Configurar usuario SFTP en SFTPGo**
   - Accede a http://localhost:8080
   - Login con admin / tu_contraseña
   - Crea usuario `lear_cable` con permisos de lectura/escritura

2. **Probar el sistema**
   - Sube un ZIP con PDFs via UI: http://localhost:8081/tools/lear_cable/
   - Verifica que procesa correctamente
   - Revisa que llegue el email con resultados

3. **Revisar logs si hay problemas**
   ```bash
   docker compose logs -f
   ```

---

## Migración desde Otro Servidor

Si ya tienes el stack corriendo en otro servidor (ej: "toolboox"):

### En el servidor antiguo:
```bash
# Backup de datos
cd /data/ias_prod
sudo tar -czf ~/ias_backup_$(date +%Y%m%d).tar.gz .

# Backup de configuración
cd ~/invoice-automation-suite
cp .env ~/.env.backup
```

### En el servidor nuevo:
```bash
# 1. Desplegar usando una de las opciones de arriba
# 2. Copiar backup desde servidor antiguo
scp user@servidor_antiguo:~/ias_backup_*.tar.gz ~/

# 3. Restaurar datos
sudo tar -xzf ~/ias_backup_*.tar.gz -C /data/ias_prod/
sudo chown -R $USER:$USER /data/ias_prod

# 4. Copiar configuración .env
scp user@servidor_antiguo:~/.env.backup ~/invoice-automation-suite/.env

# 5. Reiniciar servicios
cd ~/invoice-automation-suite
docker compose up -d --build
```

---

## Comandos Útiles

```bash
# Ver estado de servicios
docker compose ps

# Ver logs en tiempo real
docker compose logs -f

# Reiniciar servicios
docker compose restart

# Detener servicios
docker compose down

# Ver logs de un servicio específico
docker compose logs processor_lear_cable

# Verificar salud de la API
curl http://localhost:8081/api/lear_cable/health | jq

# Backup rápido
sudo tar -czf backup_$(date +%Y%m%d).tar.gz /data/ias_prod
```

---

## Documentación Completa

- **[DEPLOYMENT.md](./DEPLOYMENT.md)** - Guía completa de despliegue e instalación
- **[USAGE.md](./USAGE.md)** - Instrucciones de uso del sistema
- **[TROUBLESHOOTING.md](./TROUBLESHOOTING.md)** - Solución de problemas
- **[README.md](./README.md)** - Visión general del proyecto

---

## Soporte

Si tienes problemas:
1. Revisa [TROUBLESHOOTING.md](./TROUBLESHOOTING.md)
2. Verifica logs: `docker compose logs -f`
3. Contacta al equipo de desarrollo
