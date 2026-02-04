# 📚 Índice de Documentación - Invoice Automation Suite

Esta página te ayuda a encontrar la documentación que necesitas según tu situación.

---

## 🆕 Nuevo Usuario / Primera Instalación

Si es la primera vez que instalas el sistema o necesitas desplegarlo en una nueva máquina:

1. **[QUICKSTART.md](./QUICKSTART.md)** ⭐ **EMPIEZA AQUÍ**
   - Guía rápida de despliegue
   - Dos opciones: automática (script) o manual
   - Resumen de comandos esenciales

2. **[DEPLOYMENT.md](./DEPLOYMENT.md)** 📖 **GUÍA COMPLETA**
   - Instrucciones paso a paso detalladas
   - Instalación de Docker y dependencias
   - Configuración completa
   - Verificación y pruebas
   - Troubleshooting
   - Seguridad y backups
   - Migración desde otro servidor

3. **[DEPLOYMENT_CHECKLIST.md](./DEPLOYMENT_CHECKLIST.md)** ✅
   - Checklist interactivo para verificar cada paso
   - Perfecto para seguir durante la instalación
   - Documenta tu despliegue

4. **[deploy.sh](./deploy.sh)** 🚀 **SCRIPT AUTOMÁTICO**
   - Script bash para instalación automatizada
   - Maneja Docker, directorios, permisos
   - Uso: `sudo ./deploy.sh`

---

## 👤 Usuario Existente

Si ya tienes el sistema instalado y necesitas usarlo:

1. **[USAGE.md](./USAGE.md)** 📘
   - Cómo usar el sistema
   - Subir facturas via UI, API, SFTP, CLI
   - Consultar estados
   - Obtener resultados

2. **[README.md](./README.md)** 📄
   - Visión general del proyecto
   - Arquitectura
   - Quickstart básico
   - Enlaces a documentación

---

## 🔧 Administrador / DevOps

Si administras el servidor o necesitas configuraciones avanzadas:

1. **[MULTI_ENVIRONMENT.md](./MULTI_ENVIRONMENT.md)** 🌍
   - Gestionar múltiples entornos (prod, pre-prod, dev)
   - Configuración de puertos y directorios
   - Scripts de ayuda
   - Mejores prácticas

2. **[TROUBLESHOOTING.md](./TROUBLESHOOTING.md)** 🔍
   - Solución de problemas comunes
   - Diagnóstico de errores
   - Comandos de debugging
   - Logs y monitoreo

3. **[.env.example](./.env.example)** ⚙️
   - Plantilla de configuración
   - Documentación de variables
   - Ejemplos de valores

---

## 📖 Guías por Caso de Uso

### Caso 1: "Necesito desplegar el stack en un nuevo servidor Ubuntu"

```
1. Leer QUICKSTART.md (10 min)
2. Ejecutar deploy.sh O seguir DEPLOYMENT.md (30-60 min)
3. Usar DEPLOYMENT_CHECKLIST.md para verificar (15 min)
4. Leer USAGE.md para aprender a usar el sistema (20 min)
```

### Caso 2: "Tengo el stack en un servidor y quiero migrarlo a otro"

```
1. Leer DEPLOYMENT.md sección "Migración desde Otro Servidor"
2. Hacer backup en servidor antiguo
3. Seguir QUICKSTART.md en servidor nuevo
4. Restaurar datos del backup
5. Verificar con DEPLOYMENT_CHECKLIST.md
```

### Caso 3: "Quiero tener producción y pre-producción separados"

```
1. Leer MULTI_ENVIRONMENT.md
2. Crear archivos .env.prod y .env.pre
3. Crear directorios separados (/data/ias_prod y /data/ias_pre)
4. Iniciar ambos entornos con diferentes puertos
```

### Caso 4: "El sistema no funciona correctamente"

```
1. Leer TROUBLESHOOTING.md
2. Verificar logs: docker compose logs -f
3. Verificar checklist: DEPLOYMENT_CHECKLIST.md
4. Verificar puertos: sudo ss -lntup
5. Verificar permisos en /data/
```

### Caso 5: "Necesito usar el sistema (subir facturas)"

```
1. Leer USAGE.md
2. Para UI: Ir a http://servidor:8081/tools/lear_cable/
3. Para API: Ver ejemplos en USAGE.md sección "API"
4. Para SFTP: Ver USAGE.md sección "SFTP"
```

### Caso 6: "Necesito actualizar el sistema"

```
1. git pull
2. docker compose down
3. docker compose up -d --build
4. Verificar: curl http://localhost:8081/api/lear_cable/health
```

---

## 📂 Estructura de Archivos del Repositorio

```
invoice-automation-suite/
├── 📄 README.md                      # Visión general del proyecto
├── 📘 QUICKSTART.md                  # ⭐ Inicio rápido
├── 📖 DEPLOYMENT.md                  # Guía completa de instalación
├── ✅ DEPLOYMENT_CHECKLIST.md        # Checklist de verificación
├── 🌍 MULTI_ENVIRONMENT.md           # Múltiples entornos
├── 🔍 TROUBLESHOOTING.md             # Solución de problemas
├── 📘 USAGE.md                       # Instrucciones de uso
├── 🚀 deploy.sh                      # Script de instalación automática
├── ⚙️ .env.example                   # Plantilla de configuración
├── 🐳 compose.yaml                   # Configuración Docker Compose
├── 📁 services/                      # Servicios Docker
│   ├── api_lear_cable/               # API REST
│   ├── processor/                    # Procesador de lotes
│   ├── web/                          # UI Web
│   └── sftp/                         # Configuración SFTP
├── 📁 apps/                          # Aplicaciones
│   └── lear_cable/
│       └── extractor/                # Extractor de datos de PDFs
└── 📁 env/                           # Configuraciones de ejemplo
    ├── prod.env                      # Ejemplo producción
    └── pre.env                       # Ejemplo pre-producción
```

---

## 🔗 Enlaces Rápidos

### Documentación Principal
- [QUICKSTART.md](./QUICKSTART.md) - Inicio rápido ⭐
- [DEPLOYMENT.md](./DEPLOYMENT.md) - Guía completa de instalación
- [USAGE.md](./USAGE.md) - Cómo usar el sistema
- [TROUBLESHOOTING.md](./TROUBLESHOOTING.md) - Solución de problemas

### Documentación Avanzada
- [MULTI_ENVIRONMENT.md](./MULTI_ENVIRONMENT.md) - Múltiples entornos
- [DEPLOYMENT_CHECKLIST.md](./DEPLOYMENT_CHECKLIST.md) - Checklist de verificación
- [README_SECTIONS.md](./README_SECTIONS.md) - Secciones adicionales

### Archivos de Configuración
- [.env.example](./.env.example) - Plantilla de variables de entorno
- [compose.yaml](./compose.yaml) - Configuración de servicios Docker

### Scripts
- [deploy.sh](./deploy.sh) - Script de instalación automática

---

## ❓ Preguntas Frecuentes (FAQ)

### ¿Por dónde empiezo?
👉 **[QUICKSTART.md](./QUICKSTART.md)**

### ¿Cómo instalo el sistema?
👉 **[DEPLOYMENT.md](./DEPLOYMENT.md)** o ejecuta `sudo ./deploy.sh`

### ¿Cómo subo facturas?
👉 **[USAGE.md](./USAGE.md)** sección "UI Web" o "API"

### ¿Cómo configuro el email?
👉 Edita `.env` y configura `SMTP_HOST`, `SMTP_USER`, `SMTP_PASS`, `MAIL_TO`

### ¿Cómo verifico que está funcionando?
👉 `curl http://localhost:8081/api/lear_cable/health`

### ¿Dónde están los archivos procesados?
👉 `/data/ias_prod/data/out/<batch_id>/`

### ¿Cómo veo los logs?
👉 `docker compose logs -f`

### ¿Cómo actualizo el sistema?
👉 `git pull && docker compose up -d --build`

### ¿Cómo hago backups?
👉 `sudo tar -czf backup.tar.gz /data/ias_prod`

### ¿Algo no funciona?
👉 **[TROUBLESHOOTING.md](./TROUBLESHOOTING.md)**

---

## 🆘 Obtener Ayuda

1. **Revisa la documentación relevante** según tu caso de uso (ver arriba)
2. **Verifica los logs**: `docker compose logs -f`
3. **Consulta TROUBLESHOOTING.md** para problemas comunes
4. **Usa el checklist**: Verifica que completaste todos los pasos en DEPLOYMENT_CHECKLIST.md
5. **Contacta al equipo de desarrollo** si persiste el problema

---

## 🎯 Resumen Ejecutivo

| Quiero... | Leo... | Tiempo |
|-----------|--------|--------|
| Instalar por primera vez | QUICKSTART.md | 10 min |
| Instalar detalladamente | DEPLOYMENT.md | 60 min |
| Migrar desde otro servidor | DEPLOYMENT.md (sección migración) | 45 min |
| Usar el sistema | USAGE.md | 20 min |
| Múltiples entornos | MULTI_ENVIRONMENT.md | 30 min |
| Resolver problemas | TROUBLESHOOTING.md | Variable |
| Verificar instalación | DEPLOYMENT_CHECKLIST.md | 15 min |

---

**Última actualización:** Febrero 2026  
**Versión de documentación:** 1.0
