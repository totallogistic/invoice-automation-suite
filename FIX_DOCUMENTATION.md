# Fix para Error HTTP 404 en Upload de Archivos

## Problema
Después de integrar la herramienta TUI (JSON Form) bajo `/tools/json_form/`, las herramientas anteriores (import_partida y lear_cable) comenzaron a fallar al hacer upload de archivos con el error:
```
Upload HTTP 404: {"detail":"Not Found"}
```

## Causa Raíz
El bloque de ubicación nginx para la herramienta TUI se configuró con el modificador `^~` (`location ^~ /tools/json_form/`), que le da mayor prioridad en el algoritmo de coincidencia de ubicaciones de nginx. Aunque esto no debería afectar directamente las rutas `/api/`, evidenció una inconsistencia en la configuración de nginx donde los bloques de ubicación de proxy de API carecían de modificadores de prioridad explícitos.

### Cómo funciona la Prioridad de Ubicación en Nginx
Nginx evalúa las ubicaciones en el siguiente orden de prioridad:
1. Coincidencia exacta (`=`)
2. Coincidencia de prefijo con `^~` (detiene la búsqueda si coincide)
3. Coincidencias de expresiones regulares en orden de aparición
4. Coincidencia de prefijo más larga

Sin el modificador `^~`, las ubicaciones de API podrían ser potencialmente interferidas por ubicaciones de expresiones regulares o cambios en el procesamiento interno de nginx.

## Solución
Se agregó el modificador de prefijo `^~` a ambos bloques de ubicación de API en `services/web/nginx.conf`:

### Antes:
```nginx
location /api/lear_cable/ {
    set $upstream_lear_cable api_lear_cable:8000;
    proxy_pass http://$upstream_lear_cable/api/lear_cable/;
    ...
}

location /api/import_partida/ {
    set $upstream_import_partida api_import_partida:8000;
    proxy_pass http://$upstream_import_partida/api/import_partida/;
    ...
}
```

### Después:
```nginx
location ^~ /api/lear_cable/ {
    set $upstream_lear_cable api_lear_cable:8000;
    proxy_pass http://$upstream_lear_cable/api/lear_cable/;
    ...
}

location ^~ /api/import_partida/ {
    set $upstream_import_partida api_import_partida:8000;
    proxy_pass http://$upstream_import_partida/api/import_partida/;
    ...
}
```

## Beneficios de este Cambio
1. **Prioridad Consistente**: Las solicitudes de API se manejan con la misma prioridad que la herramienta TUI
2. **Prevención de Interferencias**: La coincidencia de ubicación se detiene inmediatamente cuando coinciden estos prefijos, evitando cualquier interferencia potencial de ubicaciones regex u otras reglas
3. **Configuración Explícita**: Hace que la configuración sea más clara y mantenible

## Verificación
Para verificar que el fix funciona:

1. Reiniciar el contenedor nginx:
```bash
docker compose restart tools_web
```

2. Probar el upload en import_partida:
   - Navegar a http://localhost:8083/tools/import_partida/
   - Seleccionar un archivo PDF
   - Hacer clic en "Subir Archivo(s)"
   - Verificar que no aparezca el error 404 y que se muestre el progreso del batch

3. Probar el upload en lear_cable:
   - Navegar a http://localhost:8083/tools/lear_cable/
   - Seleccionar uno o varios archivos PDF (o un ZIP)
   - Hacer clic en "Subir Archivo(s)"
   - Verificar que no aparezca el error 404 y que se muestre el progreso del batch

## Archivos Modificados
- `services/web/nginx.conf`: Agregado modificador `^~` a las líneas 14 y 22

## Notas Técnicas
- El modificador `^~` significa "si este prefijo coincide, no verificar ubicaciones regex"
- Este cambio no afecta el comportamiento de las rutas estáticas servidas por `location /`
- La herramienta TUI continúa funcionando normalmente en `/tools/json_form/`
- No se requieren cambios en el código backend (FastAPI) ni en el frontend (HTML/JavaScript)
