#!/bin/bash

#############################################################################
# Invoice Automation Suite - Deployment Script for Ubuntu
#############################################################################
# Este script automatiza el despliegue del stack en una nueva máquina Ubuntu
#
# Uso:
#   sudo ./deploy.sh [--env-file <path>] [--stack-root <path>]
#
# Opciones:
#   --env-file    Ruta al archivo .env (por defecto: .env)
#   --stack-root  Directorio raíz para datos (por defecto: /data/ias_prod)
#   --skip-docker Saltar instalación de Docker/Docker Compose
#   --help        Mostrar esta ayuda
# --- END HELP ---
#############################################################################

set -e  # Salir si hay error

# Colores para output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Configuración por defecto
ENV_FILE=".env"
STACK_ROOT="/data/ias_prod"
SKIP_DOCKER=false

#############################################################################
# Funciones auxiliares
#############################################################################

print_header() {
    echo -e "\n${BLUE}========================================${NC}"
    echo -e "${BLUE}$1${NC}"
    echo -e "${BLUE}========================================${NC}\n"
}

print_success() {
    echo -e "${GREEN}✓ $1${NC}"
}

print_error() {
    echo -e "${RED}✗ Error: $1${NC}"
}

print_warning() {
    echo -e "${YELLOW}⚠ Advertencia: $1${NC}"
}

print_info() {
    echo -e "${BLUE}ℹ $1${NC}"
}

#############################################################################
# Parsear argumentos
#############################################################################

while [[ $# -gt 0 ]]; do
    case $1 in
        --env-file)
            ENV_FILE="$2"
            shift 2
            ;;
        --stack-root)
            STACK_ROOT="$2"
            shift 2
            ;;
        --skip-docker)
            SKIP_DOCKER=true
            shift
            ;;
        --help)
            sed -n '/^# Invoice Automation Suite/,/^# --- END HELP ---/p' "$0" | sed 's/^# //'
            exit 0
            ;;
        *)
            print_error "Opción desconocida: $1"
            echo "Use --help para ver las opciones disponibles"
            exit 1
            ;;
    esac
done

#############################################################################
# Verificaciones previas
#############################################################################

print_header "Verificaciones Previas"

# Verificar que se ejecuta con sudo
if [ "$EUID" -ne 0 ]; then
    print_error "Este script debe ejecutarse con sudo"
    echo "Uso: sudo ./deploy.sh"
    exit 1
fi

# Obtener el usuario real (no root)
REAL_USER=${SUDO_USER:-$USER}
if [ "$REAL_USER" == "root" ]; then
    print_warning "Ejecutando como root. Se recomienda usar un usuario normal con sudo"
    read -p "¿Continuar de todas formas? (y/n) " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        exit 1
    fi
fi

print_success "Usuario: $REAL_USER"

# Verificar sistema operativo
if [ -f /etc/os-release ]; then
    . /etc/os-release
    if [[ "$ID" != "ubuntu" ]]; then
        print_warning "Este script está optimizado para Ubuntu. SO detectado: $ID"
        read -p "¿Continuar de todas formas? (y/n) " -n 1 -r
        echo
        if [[ ! $REPLY =~ ^[Yy]$ ]]; then
            exit 1
        fi
    else
        print_success "Sistema operativo: Ubuntu $VERSION_ID"
    fi
fi

#############################################################################
# Instalar Docker y Docker Compose
#############################################################################

if [ "$SKIP_DOCKER" = false ]; then
    print_header "Instalación de Docker"

    # Verificar si Docker ya está instalado
    if command -v docker &> /dev/null; then
        DOCKER_VERSION=$(docker --version)
        print_info "Docker ya está instalado: $DOCKER_VERSION"
        read -p "¿Reinstalar Docker? (y/n) " -n 1 -r
        echo
        if [[ ! $REPLY =~ ^[Yy]$ ]]; then
            print_info "Saltando instalación de Docker"
        else
            INSTALL_DOCKER=true
        fi
    else
        INSTALL_DOCKER=true
    fi

    if [ "${INSTALL_DOCKER:-false}" = true ]; then
        print_info "Actualizando paquetes del sistema..."
        apt-get update -qq

        print_info "Instalando dependencias..."
        apt-get install -y -qq \
            apt-transport-https \
            ca-certificates \
            curl \
            software-properties-common \
            gnupg \
            lsb-release

        print_info "Añadiendo repositorio de Docker..."
        mkdir -p /etc/apt/keyrings
        curl -fsSL https://download.docker.com/linux/ubuntu/gpg | gpg --dearmor -o /etc/apt/keyrings/docker.gpg
        
        echo \
          "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu \
          $(lsb_release -cs) stable" | tee /etc/apt/sources.list.d/docker.list > /dev/null

        print_info "Instalando Docker..."
        apt-get update -qq
        apt-get install -y -qq docker-ce docker-ce-cli containerd.io docker-compose-plugin

        print_success "Docker instalado correctamente"
        docker --version
    fi

    # Verificar Docker Compose
    if command -v docker &> /dev/null && docker compose version &> /dev/null; then
        print_success "Docker Compose disponible: $(docker compose version)"
    else
        print_error "Docker Compose no está disponible"
        exit 1
    fi

    # Configurar permisos de Docker para el usuario
    print_info "Configurando permisos de Docker para el usuario $REAL_USER..."
    if ! groups "$REAL_USER" | grep -q docker; then
        usermod -aG docker "$REAL_USER"
        print_success "Usuario $REAL_USER añadido al grupo docker"
        print_warning "Necesitarás cerrar sesión y volver a entrar para que los cambios surtan efecto"
    else
        print_info "El usuario ya está en el grupo docker"
    fi

    # Iniciar y habilitar Docker
    print_info "Iniciando servicio Docker..."
    systemctl start docker
    systemctl enable docker
    print_success "Servicio Docker iniciado y habilitado"

else
    print_info "Saltando instalación de Docker (--skip-docker especificado)"
fi

#############################################################################
# Instalar herramientas adicionales
#############################################################################

print_header "Instalación de Herramientas Adicionales"

print_info "Instalando herramientas útiles..."
apt-get install -y -qq jq tree htop curl wget

print_success "Herramientas instaladas: jq, tree, htop, curl, wget"

#############################################################################
# Crear estructura de directorios
#############################################################################

print_header "Creación de Estructura de Directorios"

print_info "Directorio raíz: $STACK_ROOT"

# Crear directorios principales
mkdir -p "$STACK_ROOT"
mkdir -p "$STACK_ROOT/data"
mkdir -p "$STACK_ROOT/sftpgo"
mkdir -p "$STACK_ROOT/sftpgo_var"

# Crear subdirectorios para procesamiento
mkdir -p "$STACK_ROOT/data/inbox"
mkdir -p "$STACK_ROOT/data/processing"
mkdir -p "$STACK_ROOT/data/out"
mkdir -p "$STACK_ROOT/data/processed"
mkdir -p "$STACK_ROOT/data/error"
mkdir -p "$STACK_ROOT/data/status"

print_success "Directorios creados"

# Configurar permisos
print_info "Configurando permisos..."
chown -R "$REAL_USER:$REAL_USER" "$STACK_ROOT"
chmod -R 755 "$STACK_ROOT"

print_success "Permisos configurados para el usuario $REAL_USER"

# Mostrar estructura
print_info "Estructura de directorios:"
tree -L 3 "$STACK_ROOT" 2>/dev/null || find "$STACK_ROOT" -maxdepth 3 -type d | sed 's|[^/]*/| |g'

#############################################################################
# Configurar archivo .env
#############################################################################

print_header "Configuración de Variables de Entorno"

if [ ! -f "$ENV_FILE" ]; then
    if [ -f ".env.example" ]; then
        print_info "Creando $ENV_FILE desde .env.example..."
        cp .env.example "$ENV_FILE"
        
        # Actualizar STACK_ROOT en .env si es diferente del default
        if [ "$STACK_ROOT" != "/data/ias_prod" ]; then
            if grep -q "STACK_ROOT=" "$ENV_FILE" 2>/dev/null; then
                sed -i "s|STACK_ROOT=.*|STACK_ROOT=$STACK_ROOT|g" "$ENV_FILE"
            else
                print_warning "No se encontró STACK_ROOT en $ENV_FILE para actualizar"
            fi
            
            if grep -q "DATA_ROOT=" "$ENV_FILE" 2>/dev/null; then
                sed -i "s|DATA_ROOT=.*|DATA_ROOT=$STACK_ROOT/data|g" "$ENV_FILE"
            else
                print_warning "No se encontró DATA_ROOT en $ENV_FILE para actualizar"
            fi
        fi
        
        chown "$REAL_USER:$REAL_USER" "$ENV_FILE"
        
        print_success "Archivo $ENV_FILE creado"
        print_warning "⚠️  IMPORTANTE: Debes editar $ENV_FILE y configurar:"
        echo "   - Credenciales SMTP (SMTP_HOST, SMTP_USER, SMTP_PASS, etc.)"
        echo "   - Contraseña de administrador SFTPGo (SFTPGO_DEFAULT_ADMIN_PASSWORD)"
        echo "   - Destinatarios de email (MAIL_TO)"
        echo ""
        print_info "Ejemplo de edición:"
        echo "   nano $ENV_FILE"
        echo "   # o"
        echo "   vim $ENV_FILE"
        echo ""
    else
        print_error "No se encontró .env.example para crear $ENV_FILE"
        exit 1
    fi
else
    print_info "Archivo $ENV_FILE ya existe"
    
    # Verificar variables críticas
    if grep -q "CHANGE_ME\|INSECURE" "$ENV_FILE" 2>/dev/null; then
        print_warning "⚠️  El archivo .env contiene valores inseguros o por defecto que DEBEN ser actualizados:"
        grep -n "CHANGE_ME\|INSECURE" "$ENV_FILE" || true
    fi
    
    if grep -q "YOUR_USER" "$ENV_FILE" 2>/dev/null; then
        print_warning "⚠️  El archivo .env contiene valores por defecto que deben ser actualizados"
    fi
fi

#############################################################################
# Verificar que estamos en el directorio correcto
#############################################################################

print_header "Verificación del Proyecto"

if [ ! -f "compose.yaml" ]; then
    print_error "No se encontró compose.yaml en el directorio actual"
    print_info "Asegúrate de ejecutar este script desde el directorio raíz del proyecto"
    exit 1
fi

print_success "Archivo compose.yaml encontrado"

# Verificar estructura de servicios
if [ ! -d "services" ]; then
    print_error "No se encontró el directorio 'services'"
    exit 1
fi

if [ ! -d "apps" ]; then
    print_error "No se encontró el directorio 'apps'"
    exit 1
fi

print_success "Estructura del proyecto verificada"

#############################################################################
# Construir e iniciar servicios
#############################################################################

print_header "Construcción e Inicio de Servicios"

print_warning "⚠️  ANTES DE CONTINUAR:"
echo "   1. Asegúrate de haber editado el archivo $ENV_FILE con tus configuraciones"
echo "   2. Especialmente verifica SMTP, contraseñas, y destinatarios de email"
echo ""
read -p "¿Continuar con el despliegue? (y/n) " -n 1 -r
echo
if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    print_info "Despliegue cancelado por el usuario"
    print_info "Edita $ENV_FILE y luego ejecuta:"
    echo "   cd $(pwd)"
    echo "   docker compose up -d --build"
    exit 0
fi

print_info "Construyendo imágenes Docker..."
docker compose build

print_info "Iniciando servicios..."
docker compose up -d

print_success "Servicios iniciados"

# Esperar un momento para que los servicios se inicien
sleep 5

#############################################################################
# Verificación del despliegue
#############################################################################

print_header "Verificación del Despliegue"

print_info "Estado de los contenedores:"
docker compose ps

# Verificar que todos los servicios están running
SERVICES=("sftp_lear_cable" "processor_lear_cable" "api_lear_cable" "tools_web")
ALL_RUNNING=true

for service in "${SERVICES[@]}"; do
    if docker compose ps "$service" | grep -q "Up"; then
        print_success "Servicio $service está corriendo"
    else
        print_error "Servicio $service NO está corriendo"
        ALL_RUNNING=false
    fi
done

# Verificar salud de la API
print_info "Verificando salud de la API..."
sleep 3  # Dar un poco más de tiempo
if curl -f -s http://localhost:8081/api/lear_cable/health > /dev/null 2>&1; then
    print_success "API responde correctamente"
    curl -s http://localhost:8081/api/lear_cable/health | jq '.' 2>/dev/null || curl -s http://localhost:8081/api/lear_cable/health
else
    print_warning "API no responde aún (puede tardar unos segundos más en iniciar)"
fi

# Verificar UI
print_info "Verificando UI web..."
if curl -f -s http://localhost:8081/tools/lear_cable/ > /dev/null 2>&1; then
    print_success "UI web accesible"
else
    print_warning "UI web no responde aún"
fi

#############################################################################
# Resumen final
#############################################################################

print_header "Despliegue Completado"

if [ "$ALL_RUNNING" = true ]; then
    print_success "¡Todos los servicios están corriendo correctamente!"
else
    print_warning "Algunos servicios no están corriendo. Revisa los logs:"
    echo "   docker compose logs"
fi

echo ""
print_info "URLs de acceso:"
echo "   • Web UI:        http://$(hostname -I | awk '{print $1}'):8081/tools/lear_cable/"
echo "   • API:           http://$(hostname -I | awk '{print $1}'):8081/api/lear_cable/"
echo "   • SFTPGo Admin:  http://$(hostname -I | awk '{print $1}'):8080/"
echo "   • SFTP:          sftp://$(hostname -I | awk '{print $1}'):2222"
echo ""

print_info "Próximos pasos:"
echo "   1. Configurar usuario SFTP en SFTPGo Admin:"
echo "      - Accede a http://localhost:8080"
echo "      - Login con las credenciales del .env"
echo "      - Crea usuario 'lear_cable' con permisos apropiados"
echo ""
echo "   2. Verificar logs:"
echo "      docker compose logs -f"
echo ""
echo "   3. Probar subida de archivos:"
echo "      - Via UI: http://localhost:8081/tools/lear_cable/"
echo "      - Via API: curl -F 'file=@test.zip' http://localhost:8081/api/lear_cable/batches"
echo ""
echo "   4. Documentación:"
echo "      - Deployment: cat DEPLOYMENT.md"
echo "      - Usage: cat USAGE.md"
echo "      - Troubleshooting: cat TROUBLESHOOTING.md"
echo ""

print_warning "⚠️  RECORDATORIO DE SEGURIDAD:"
echo "   • Cambia las contraseñas por defecto en .env"
echo "   • Configura firewall (ufw) para restringir acceso"
echo "   • Considera usar HTTPS en producción"
echo "   • Haz backups regulares de $STACK_ROOT"
echo ""

print_success "Script de despliegue finalizado"
