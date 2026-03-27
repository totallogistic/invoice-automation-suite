# Checklist de Despliegue - Invoice Automation Suite

Use este checklist para asegurar un despliegue correcto del stack.

## Pre-instalación

- [ ] Sistema operativo: Ubuntu 20.04 LTS o superior
- [ ] Usuario con permisos sudo disponible
- [ ] Conexión a Internet activa
- [ ] Al menos 20 GB de espacio en disco disponible
- [ ] Al menos 4 GB de RAM disponible

## Instalación de Docker

- [ ] Docker instalado correctamente (`docker --version`)
- [ ] Docker Compose instalado (`docker compose version`)
- [ ] Usuario añadido al grupo docker (`groups | grep docker`)
- [ ] Servicio Docker corriendo (`systemctl status docker`)

## Estructura de Directorios

- [ ] Directorio raíz creado: `/data/ias_prod` (o tu STACK_ROOT)
- [ ] Subdirectorios creados:
  - [ ] `/data/ias_prod/data/inbox`
  - [ ] `/data/ias_prod/data/processing`
  - [ ] `/data/ias_prod/data/out`
  - [ ] `/data/ias_prod/data/processed`
  - [ ] `/data/ias_prod/data/error`
  - [ ] `/data/ias_prod/data/status`
  - [ ] `/data/ias_prod/sftpgo`
  - [ ] `/data/ias_prod/sftpgo_var`
- [ ] Permisos correctos (`ls -la /data/ias_prod`)

## Configuración

- [ ] Repositorio clonado
- [ ] Archivo `.env` creado desde `.env.example`
- [ ] Variables SMTP configuradas:
  - [ ] `SMTP_HOST`
  - [ ] `SMTP_PORT`
  - [ ] `SMTP_USER`
  - [ ] `SMTP_PASS` (App Password si es Gmail)
  - [ ] `MAIL_FROM`
  - [ ] `MAIL_TO`
- [ ] Contraseña de admin SFTPGo cambiada: `SFTPGO_DEFAULT_ADMIN_PASSWORD`
- [ ] Paths verificados (si usas diferentes a los default):
  - [ ] `STACK_ROOT`
  - [ ] `DATA_ROOT`
  - [ ] `SFTPGO_ROOT`
  - [ ] `SFTPGO_VAR`
- [ ] Puertos verificados (sin conflictos):
  - [ ] `SFTP_PORT=2222`
  - [ ] `SFTPGO_WEB_PORT=8080`
  - [ ] `TOOLS_WEB_PORT=8081`

## Construcción e Inicio

- [ ] Imágenes Docker construidas (`docker compose build`)
- [ ] Servicios iniciados (`docker compose up -d`)
- [ ] Todos los contenedores corriendo (`docker compose ps`):
  - [ ] `sftp_lear_cable` - Estado: Up
  - [ ] `processor_lear_cable` - Estado: Up
  - [ ] `api_lear_cable` - Estado: Up
  - [ ] `tools_web` - Estado: Up

## Verificación de Servicios

### API Health Check
- [ ] API responde: `curl http://localhost:8081/api/lear_cable/health`
- [ ] Respuesta JSON válida con `"status": "healthy"`

### UI Web
- [ ] UI accesible: `curl http://localhost:8081/tools/lear_cable/`
- [ ] Devuelve contenido HTML
- [ ] Página carga en navegador: `http://<IP_SERVIDOR>:8081/tools/lear_cable/`

### SFTPGo Admin
- [ ] Interfaz accesible: `http://<IP_SERVIDOR>:8080/`
- [ ] Login funciona con credenciales del `.env`
- [ ] Panel de administración carga correctamente

### Puertos
- [ ] Puerto 2222 abierto (SFTP): `sudo ss -lntup | grep 2222`
- [ ] Puerto 8080 abierto (SFTPGo Web): `sudo ss -lntup | grep 8080`
- [ ] Puerto 8081 abierto (Tools Web): `sudo ss -lntup | grep 8081`

## Configuración Post-instalación

### SFTPGo
- [ ] Acceso a SFTPGo Admin: `http://<IP_SERVIDOR>:8080/`
- [ ] Usuario SFTP creado:
  - [ ] Username: `lear_cable`
  - [ ] Password configurada
  - [ ] Home directory: `/srv/sftpgo/lear_cable/data`
  - [ ] Permisos: lectura, escritura, listar, crear directorios
- [ ] Contraseña de admin cambiada (si era la default)

### Firewall (Opcional pero Recomendado)
- [ ] UFW instalado: `sudo apt install ufw`
- [ ] Reglas configuradas:
  - [ ] SSH permitido: `sudo ufw allow 22/tcp`
  - [ ] SFTP permitido: `sudo ufw allow 2222/tcp`
  - [ ] Web UI permitido: `sudo ufw allow 8081/tcp`
  - [ ] SFTPGo Admin permitido (si necesitas): `sudo ufw allow 8080/tcp`
- [ ] Firewall habilitado: `sudo ufw enable`

## Pruebas Funcionales

### Prueba 1: Upload via API
```bash
# Crear ZIP de prueba con PDFs
# Subir via API:
curl -F "file=@test.zip" http://localhost:8081/api/lear_cable/batches
```
- [ ] API devuelve `batch_id`
- [ ] Archivos aparecen en `/data/ias_prod/data/inbox/`
- [ ] Processor detecta el lote
- [ ] Lote se procesa correctamente
- [ ] Outputs generados en `/data/ias_prod/data/out/<batch_id>/`
- [ ] Email enviado con adjuntos

### Prueba 2: Upload via UI
- [ ] Abrir: `http://<IP_SERVIDOR>:8081/tools/lear_cable/`
- [ ] Seleccionar y subir ZIP
- [ ] Ver batch_id en respuesta
- [ ] Ver progreso en UI
- [ ] Ver estado final (DONE)
- [ ] Verificar email recibido

### Prueba 3: SFTP Upload (Opcional)
```bash
sftp -P 2222 lear_cable@<IP_SERVIDOR>
# Subir archivos a inbox/
# Crear marker _DONE si es necesario
```
- [ ] Conexión SFTP exitosa
- [ ] Upload de archivos funciona
- [ ] Processor detecta y procesa archivos

## Logs y Monitoreo

- [ ] Logs accesibles: `docker compose logs -f`
- [ ] No hay errores críticos en logs
- [ ] Logs de cada servicio:
  - [ ] `docker compose logs tools_web` - Sin errores
  - [ ] `docker compose logs api_lear_cable` - Sin errores
  - [ ] `docker compose logs processor_lear_cable` - Sin errores
  - [ ] `docker compose logs sftp_lear_cable` - Sin errores

## Seguridad

- [ ] Contraseñas por defecto cambiadas
- [ ] Archivo `.env` tiene permisos restringidos: `chmod 600 .env`
- [ ] `.env` NO está en Git (verificar `.gitignore`)
- [ ] Firewall configurado (si aplica)
- [ ] Considerar HTTPS para producción (con reverse proxy)
- [ ] Backups configurados

## Documentación Revisada

- [ ] Leído [DEPLOYMENT.md](./DEPLOYMENT.md)
- [ ] Leído [QUICKSTART.md](./QUICKSTART.md)
- [ ] Leído [USAGE.md](./USAGE.md)
- [ ] Leído [TROUBLESHOOTING.md](./TROUBLESHOOTING.md)

## Backups

- [ ] Script de backup creado
- [ ] Primer backup realizado:
  ```bash
  sudo tar -czf backup_ias_prod_$(date +%Y%m%d).tar.gz /data/ias_prod
  ```
- [ ] Backup verificado (puede extraerse)
- [ ] Ubicación de backups definida
- [ ] Frecuencia de backups definida (diario/semanal)

## Mantenimiento

- [ ] Plan de actualización definido
- [ ] Plan de limpieza de datos antiguos definido
- [ ] Monitoreo de espacio en disco configurado
- [ ] Contactos de soporte identificados

## Migración (Si aplica)

Si migraste desde otro servidor:
- [ ] Backup del servidor antiguo realizado
- [ ] Datos restaurados en nuevo servidor
- [ ] Configuración `.env` migrada
- [ ] Servicios funcionando con datos migrados
- [ ] Pruebas con datos reales completadas
- [ ] Servidor antiguo apagado o en standby

---

## Resumen Final

Fecha de despliegue: _______________

Servidor: _______________

IP: _______________

Versión del stack: _______________

Responsable: _______________

Notas adicionales:
```


```

---

## Problemas Encontrados Durante el Despliegue

Si encontraste problemas, documéntalos aquí para referencia futura:

1. **Problema:** 
   **Solución:** 

2. **Problema:** 
   **Solución:** 

3. **Problema:** 
   **Solución:** 

---

✅ **Despliegue completado exitosamente**

Fecha: _______________
Firma: _______________
