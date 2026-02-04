# Guía de Entornos Múltiples

Esta guía explica cómo gestionar múltiples entornos (producción, pre-producción, desarrollo) del Invoice Automation Suite en la misma máquina.

## Conceptos Clave

Cada entorno debe tener:
1. **Puertos diferentes** - Para evitar conflictos
2. **Directorios de datos diferentes** - Para evitar corrupción de datos
3. **Nombre de proyecto diferente** - Para que Docker Compose los mantenga separados

## Estructura de Entornos

```
/data/
├── ias_prod/          # Entorno de Producción
│   ├── data/
│   ├── sftpgo/
│   └── sftpgo_var/
└── ias_pre/           # Entorno de Pre-producción
    ├── data/
    ├── sftpgo/
    └── sftpgo_var/
```

## Archivos de Configuración

El repositorio incluye archivos de ejemplo en `env/`:
- `env/prod.env` - Configuración para producción
- `env/pre.env` - Configuración para pre-producción

### Crear tus archivos de entorno:

```bash
# Producción
cp env/prod.env .env.prod

# Pre-producción  
cp env/pre.env .env.pre

# Editar cada uno con sus valores específicos
nano .env.prod
nano .env.pre
```

## Configuración por Entorno

### Producción (.env.prod)

```bash
# Identificación
COMPOSE_PROJECT_NAME=ias_prod

# Directorios
STACK_ROOT=/data/ias_prod
DATA_ROOT=/data/ias_prod/data
SFTPGO_ROOT=/data/ias_prod/sftpgo
SFTPGO_VAR=/data/ias_prod/sftpgo_var

# Puertos
SFTP_PORT=2222
SFTPGO_WEB_PORT=8080
TOOLS_WEB_PORT=8081

# Email
MAIL_TO=produccion@domain.com
BATCH_QUIET_SECONDS=300  # 5 minutos

# SMTP, contraseñas, etc.
# ...
```

### Pre-producción (.env.pre)

```bash
# Identificación
COMPOSE_PROJECT_NAME=ias_pre

# Directorios (diferentes a producción)
STACK_ROOT=/data/ias_pre
DATA_ROOT=/data/ias_pre/data
SFTPGO_ROOT=/data/ias_pre/sftpgo
SFTPGO_VAR=/data/ias_pre/sftpgo_var

# Puertos (diferentes a producción)
SFTP_PORT=2223          # Diferente!
SFTPGO_WEB_PORT=8083    # Diferente!
TOOLS_WEB_PORT=8082     # Diferente!

# Email (puede ser diferente)
MAIL_TO=testing@domain.com
BATCH_QUIET_SECONDS=60  # 1 minuto para testing más rápido

# SMTP, contraseñas, etc.
# ...
```

## Crear Directorios para Cada Entorno

```bash
# Producción
sudo mkdir -p /data/ias_prod/data/{inbox,processing,out,processed,error,status}
sudo mkdir -p /data/ias_prod/{sftpgo,sftpgo_var}
sudo chown -R $USER:$USER /data/ias_prod
sudo chmod -R 755 /data/ias_prod

# Pre-producción
sudo mkdir -p /data/ias_pre/data/{inbox,processing,out,processed,error,status}
sudo mkdir -p /data/ias_pre/{sftpgo,sftpgo_var}
sudo chown -R $USER:$USER /data/ias_pre
sudo chmod -R 755 /data/ias_pre
```

## Gestión de Servicios

### Iniciar Entornos

```bash
# Iniciar SOLO producción
docker compose --env-file .env.prod up -d --build

# Iniciar SOLO pre-producción
docker compose --env-file .env.pre -p ias_pre up -d --build

# Iniciar AMBOS
docker compose --env-file .env.prod up -d --build
docker compose --env-file .env.pre -p ias_pre up -d --build
```

### Ver Estado

```bash
# Ver todos los contenedores
docker ps --format "table {{.Names}}\t{{.Status}}\t{{.Ports}}"

# Filtrar por entorno
docker ps --filter "name=ias_prod"
docker ps --filter "name=ias_pre"

# Estado completo de un entorno
docker compose --env-file .env.prod ps
docker compose --env-file .env.pre -p ias_pre ps
```

### Ver Logs

```bash
# Logs de producción
docker compose --env-file .env.prod logs -f

# Logs de pre-producción
docker compose --env-file .env.pre -p ias_pre logs -f

# Logs de un servicio específico en un entorno
docker compose --env-file .env.prod logs -f processor_lear_cable
docker compose --env-file .env.pre -p ias_pre logs -f processor_lear_cable
```

### Detener Entornos

```bash
# Detener producción
docker compose --env-file .env.prod down

# Detener pre-producción
docker compose --env-file .env.pre -p ias_pre down

# Detener TODO
docker compose --env-file .env.prod down
docker compose --env-file .env.pre -p ias_pre down
```

### Reiniciar Servicios

```bash
# Reiniciar servicio específico en producción
docker compose --env-file .env.prod restart processor_lear_cable

# Reiniciar servicio específico en pre-producción
docker compose --env-file .env.pre -p ias_pre restart processor_lear_cable

# Reiniciar todo el entorno
docker compose --env-file .env.prod restart
docker compose --env-file .env.pre -p ias_pre restart
```

## Acceso a Servicios por Entorno

### Producción

| Servicio | URL |
|----------|-----|
| Web UI | http://servidor:8081/tools/lear_cable/ |
| API | http://servidor:8081/api/lear_cable/ |
| SFTPGo Admin | http://servidor:8080/ |
| SFTP | sftp://servidor:2222 |

### Pre-producción

| Servicio | URL |
|----------|-----|
| Web UI | http://servidor:8082/tools/lear_cable/ |
| API | http://servidor:8082/api/lear_cable/ |
| SFTPGo Admin | http://servidor:8083/ |
| SFTP | sftp://servidor:2223 |

## Actualizar un Entorno Específico

```bash
# Actualizar código
git pull

# Reconstruir e reiniciar solo producción
docker compose --env-file .env.prod up -d --build --force-recreate

# Reconstruir e reiniciar solo pre-producción
docker compose --env-file .env.pre -p ias_pre up -d --build --force-recreate
```

## Migrar entre Entornos

### De Pre-producción a Producción

```bash
# 1. Probar en pre-producción
# ... realizar pruebas ...

# 2. Si todo OK, actualizar producción
docker compose --env-file .env.prod down
docker compose --env-file .env.prod up -d --build

# 3. Verificar
curl http://localhost:8081/api/lear_cable/health
```

### Copiar Datos de un Entorno a Otro

```bash
# Backup de producción
sudo tar -czf /tmp/backup_prod.tar.gz /data/ias_prod/data

# Restaurar en pre (para testing con datos reales)
sudo tar -xzf /tmp/backup_prod.tar.gz -C /data/ias_pre/data --strip-components=4
sudo chown -R $USER:$USER /data/ias_pre/data
```

## Scripts de Ayuda

### Script para cambiar entre entornos

Crear `switch-env.sh`:

```bash
#!/bin/bash

ENV=${1:-prod}

if [ "$ENV" = "prod" ]; then
    export ENV_FILE=".env.prod"
    export PROJECT_NAME="ias_prod"
    echo "Usando entorno: PRODUCCIÓN"
elif [ "$ENV" = "pre" ]; then
    export ENV_FILE=".env.pre"
    export PROJECT_NAME="ias_pre"
    echo "Usando entorno: PRE-PRODUCCIÓN"
else
    echo "Uso: $0 [prod|pre]"
    exit 1
fi

# Ejecutar comando
shift
if [ $# -eq 0 ]; then
    echo "ENV_FILE=$ENV_FILE"
    echo "PROJECT_NAME=$PROJECT_NAME"
    echo ""
    echo "Comandos disponibles:"
    echo "  docker compose --env-file $ENV_FILE -p $PROJECT_NAME up -d"
    echo "  docker compose --env-file $ENV_FILE -p $PROJECT_NAME ps"
    echo "  docker compose --env-file $ENV_FILE -p $PROJECT_NAME logs -f"
else
    docker compose --env-file "$ENV_FILE" -p "$PROJECT_NAME" "$@"
fi
```

Uso:
```bash
chmod +x switch-env.sh

# Ver info del entorno
./switch-env.sh prod
./switch-env.sh pre

# Ejecutar comandos
./switch-env.sh prod ps
./switch-env.sh prod logs -f
./switch-env.sh pre up -d --build
./switch-env.sh pre down
```

## Mejores Prácticas

1. **Siempre probar en pre-producción primero**
   ```bash
   # Desplegar en pre
   ./switch-env.sh pre up -d --build
   
   # Probar
   curl http://localhost:8082/api/lear_cable/health
   
   # Si OK, desplegar en prod
   ./switch-env.sh prod up -d --build
   ```

2. **Mantener configuraciones sincronizadas**
   - Usa las mismas versiones de imagen Docker
   - Mantén configuraciones similares (excepto puertos/paths)
   - Documenta diferencias entre entornos

3. **Monitorear ambos entornos**
   ```bash
   # Ver todos los contenedores
   docker ps -a | grep ias
   
   # Monitorear recursos
   docker stats --filter "name=ias"
   ```

4. **Backups separados**
   ```bash
   # Backup de cada entorno
   sudo tar -czf backup_prod_$(date +%Y%m%d).tar.gz /data/ias_prod
   sudo tar -czf backup_pre_$(date +%Y%m%d).tar.gz /data/ias_pre
   ```

5. **Limpieza regular de pre-producción**
   ```bash
   # Limpiar datos antiguos de pre cada semana
   find /data/ias_pre/data/processed -type d -mtime +7 -exec rm -rf {} +
   ```

## Troubleshooting

### Conflicto de Puertos

Si ves errores como "port is already allocated":
```bash
# Verificar qué está usando el puerto
sudo ss -lntup | grep <puerto>

# Cambiar puerto en .env del entorno correspondiente
# Luego reiniciar
docker compose --env-file .env.pre -p ias_pre down
docker compose --env-file .env.pre -p ias_pre up -d
```

### Conflicto de Nombres de Contenedor

Si Docker se queja de nombres duplicados:
```bash
# Asegúrate de usar -p con nombre diferente
docker compose --env-file .env.pre -p ias_pre up -d
```

### Ver qué entorno está usando qué recursos

```bash
# Listar todos los recursos por proyecto
docker ps --filter "label=com.docker.compose.project=ias_prod"
docker ps --filter "label=com.docker.compose.project=ias_pre"

# Ver redes
docker network ls | grep ias

# Ver volúmenes
docker volume ls | grep ias
```

## Resumen de Comandos Rápidos

```bash
# PRODUCCIÓN
docker compose --env-file .env.prod up -d --build                    # Iniciar
docker compose --env-file .env.prod ps                               # Estado
docker compose --env-file .env.prod logs -f                          # Logs
docker compose --env-file .env.prod down                             # Detener

# PRE-PRODUCCIÓN
docker compose --env-file .env.pre -p ias_pre up -d --build          # Iniciar
docker compose --env-file .env.pre -p ias_pre ps                     # Estado
docker compose --env-file .env.pre -p ias_pre logs -f                # Logs
docker compose --env-file .env.pre -p ias_pre down                   # Detener
```
