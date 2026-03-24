#!/usr/bin/env python3
"""
Wrapper for Cuadre Asientos - Unified Processor Interface

Adapts procesar_aforos.py to work with the unified processor:
- Accepts multiple Excel files as input
- Processes each file
- Generates output Excel + text report
- Compatible with unified processor batch system
"""

SCRIPT_VERSION = "2026-03-17.v1"

SCRIPT_CHANGELOG = """
## 2026-03-17.v1

### Logica general
Procesa archivos Excel de aforos y cuadre de asientos contables,
generando un Excel de salida con el cuadre calculado mas un informe
de texto plano para auditoria.

### Entrada
- Uno o varios archivos Excel (`.xlsx`, `.xls`)
- Cada archivo representa un lote de asientos a cuadrar

### Procesamiento
- Identifica columnas de debe/haber automaticamente
- Calcula diferencias y cuadra los asientos por cuenta contable
- Detecta y reporta descuadres

### Salida
- Excel con columnas originales mas columnas de resultado del cuadre
- Informe de texto con estadisticas: total asientos, descuadres encontrados,
  cuentas afectadas
""".strip()

import argparse
import sys
from pathlib import Path
from datetime import datetime

# Import the original processor
from procesar_aforos import procesar_archivo


def generar_reporte_texto(estadisticas, archivo_entrada):
    """
    Genera reporte en texto plano (mejor compatibilidad email).
    
    Args:
        estadisticas: Dict con estadísticas del procesamiento
        archivo_entrada: Nombre del archivo procesado
        
    Returns:
        String con el reporte formateado
    """
    reporte = []
    reporte.append("=" * 80)
    reporte.append("CUADRE DE ASIENTOS CONTABLES - REPORTE DE PROCESAMIENTO")
    reporte.append("=" * 80)
    reporte.append(f"Archivo procesado: {archivo_entrada}")
    reporte.append(f"Fecha de procesamiento: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    reporte.append("")
    
    reporte.append("=" * 80)
    reporte.append("RESUMEN DEL PROCESAMIENTO")
    reporte.append("=" * 80)
    reporte.append(f"Total de registros originales:      {estadisticas['total_original']:>10,}")
    reporte.append(f"Registros con código de aforo:      {estadisticas['registros_con_aforo']:>10,}")
    reporte.append(f"Códigos de aforo únicos:             {estadisticas['aforos_unicos']:>10,}")
    reporte.append(f"Registros eliminados (compensados):  {estadisticas['registros_eliminados']:>10,}")
    reporte.append(f"Registros restantes:                 {estadisticas['registros_restantes']:>10,}")
    reporte.append(f"Porcentaje eliminado:                {estadisticas['porcentaje_eliminado']:>9.2f}%")
    reporte.append("")
    
    reporte.append("=" * 80)
    reporte.append("TOTALES FINANCIEROS - ARCHIVO ORIGINAL")
    reporte.append("=" * 80)
    reporte.append(f"Total DEBE:                          {estadisticas['total_debe_original']:>15,.2f} €")
    reporte.append(f"Total HABER:                         {estadisticas['total_haber_original']:>15,.2f} €")
    reporte.append(f"Descuadre (DEBE - HABER):            {estadisticas['descuadre_original']:>15,.2f} €")
    reporte.append("")
    
    reporte.append("=" * 80)
    reporte.append("TOTALES FINANCIEROS - ASIENTOS ELIMINADOS (Cazados)")
    reporte.append("=" * 80)
    reporte.append(f"Total DEBE eliminado:                {estadisticas['total_debe_eliminado']:>15,.2f} €")
    reporte.append(f"Total HABER eliminado:               {estadisticas['total_haber_eliminado']:>15,.2f} €")
    reporte.append(f"Descuadre (DEBE - HABER):            {estadisticas['descuadre_eliminado']:>15,.2f} €")
    
    if abs(estadisticas['descuadre_eliminado']) < 0.01:
        reporte.append("✓ Los asientos eliminados están perfectamente cuadrados")
    else:
        reporte.append(f"⚠ ADVERTENCIA: Hay un descuadre de {abs(estadisticas['descuadre_eliminado']):,.2f} €")
    reporte.append("")
    
    reporte.append("=" * 80)
    reporte.append("TOTALES FINANCIEROS - ASIENTOS RESTANTES (No Cazados)")
    reporte.append("=" * 80)
    reporte.append(f"Total DEBE restante:                 {estadisticas['total_debe_restante']:>15,.2f} €")
    reporte.append(f"Total HABER restante:                {estadisticas['total_haber_restante']:>15,.2f} €")
    reporte.append(f"Descuadre (DEBE - HABER):            {estadisticas['descuadre_restante']:>15,.2f} €")
    reporte.append("")
    
    reporte.append("=" * 80)
    reporte.append("VERIFICACIÓN DE CUADRE")
    reporte.append("=" * 80)
    
    # Verificar que Original = Eliminado + Restante
    debe_check = abs((estadisticas['total_debe_eliminado'] + estadisticas['total_debe_restante']) - 
                     estadisticas['total_debe_original'])
    haber_check = abs((estadisticas['total_haber_eliminado'] + estadisticas['total_haber_restante']) - 
                      estadisticas['total_haber_original'])
    
    if debe_check < 0.01 and haber_check < 0.01:
        reporte.append("✓ Verificación OK: ELIMINADO + RESTANTE = ORIGINAL")
    else:
        reporte.append("⚠ ERROR: Los totales no cuadran correctamente")
        reporte.append(f"  Diferencia DEBE:  {debe_check:,.2f} €")
        reporte.append(f"  Diferencia HABER: {haber_check:,.2f} €")
    
    # Mostrar el cambio en el descuadre
    mejora_descuadre = abs(estadisticas['descuadre_original']) - abs(estadisticas['descuadre_restante'])
    reporte.append("")
    reporte.append("Cambio en descuadre:")
    reporte.append(f"  Descuadre original: {abs(estadisticas['descuadre_original']):>15,.2f} €")
    reporte.append(f"  Descuadre final:    {abs(estadisticas['descuadre_restante']):>15,.2f} €")
    
    if mejora_descuadre > 0:
        reporte.append(f"  Mejora:             {mejora_descuadre:>15,.2f} € ✓")
    elif mejora_descuadre < 0:
        reporte.append(f"  Empeoramiento:      {abs(mejora_descuadre):>15,.2f} € ⚠")
    else:
        reporte.append(f"  Sin cambio:         {mejora_descuadre:>15,.2f} €")
    
    reporte.append("=" * 80)
    reporte.append("")
    reporte.append("✓ Procesamiento completado exitosamente")
    reporte.append("")
    
    return "\n".join(reporte)


def procesar_batch(archivos_entrada, output_dir):
    """
    Procesa un batch de archivos Excel.
    
    Args:
        archivos_entrada: Lista de rutas de archivos Excel
        output_dir: Directorio de salida
        
    Returns:
        Dict con archivos generados y reportes
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    resultados = []
    
    for archivo_entrada in archivos_entrada:
        archivo_path = Path(archivo_entrada)
        print(f"\n{'='*80}")
        print(f"Procesando: {archivo_path.name}")
        print(f"{'='*80}\n")
        
        # Generar nombres de salida
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_name = archivo_path.stem
        
        archivo_xlsx = output_dir / f"{base_name}_procesado_{timestamp}.xlsx"
        archivo_cazados = output_dir / f"{base_name}_cazados_{timestamp}.xlsx"
        archivo_reporte = output_dir / f"{base_name}_reporte_{timestamp}.txt"
        
        try:
            # Procesar archivo (ahora retorna 4 valores)
            df_original, df_filtrado, df_eliminados, estadisticas = procesar_archivo(
                str(archivo_path),
                str(archivo_xlsx)
            )
            
            # El procesar_archivo genera automáticamente el archivo de cazados
            # con el nombre basado en archivo_xlsx reemplazando "_procesado" por "_cazados"
            archivo_cazados_temp = str(archivo_xlsx).replace('_procesado_', '_cazados_')
            archivo_cazados_temp = archivo_cazados_temp.replace('.xlsx', '_cazados.xlsx') if '_cazados_' not in archivo_cazados_temp else archivo_cazados_temp
            
            # Copiar el archivo de cazados al nombre con timestamp deseado
            import shutil
            if Path(archivo_cazados_temp).exists():
                shutil.copy(archivo_cazados_temp, archivo_cazados)
            
            # Generar reporte de texto
            reporte_texto = generar_reporte_texto(estadisticas, archivo_path.name)
            
            # Guardar reporte
            with open(archivo_reporte, 'w', encoding='utf-8') as f:
                f.write(reporte_texto)
            
            # También imprimir a stdout para logs
            print(reporte_texto)
            
            resultados.append({
                'archivo_entrada': str(archivo_path),
                'archivo_xlsx': str(archivo_xlsx),
                'archivo_cazados': str(archivo_cazados),
                'archivo_reporte': str(archivo_reporte),
                'estadisticas': estadisticas,
                'exito': True
            })
            
            print(f"\n✓ Archivo procesado exitosamente")
            print(f"  Excel procesado: {archivo_xlsx.name}")
            print(f"  Excel cazados:   {archivo_cazados.name}")
            print(f"  Reporte generado: {archivo_reporte.name}")
            
        except Exception as e:
            print(f"\n✗ Error procesando {archivo_path.name}: {str(e)}")
            import traceback
            traceback.print_exc()
            
            resultados.append({
                'archivo_entrada': str(archivo_path),
                'exito': False,
                'error': str(e)
            })
    
    return resultados


def main(argv):
    """Función principal - interfaz unificada."""
    parser = argparse.ArgumentParser(
        description='Cuadre Asientos - Unified Processor Wrapper'
    )
    parser.add_argument('archivos', nargs='+', help='Excel files to process')
    parser.add_argument('-o', '--output', required=True, help='Output directory')
    
    args = parser.parse_args(argv)
    
    # Procesar batch
    resultados = procesar_batch(args.archivos, args.output)
    
    # Resumen final
    print(f"\n{'='*80}")
    print("RESUMEN DEL BATCH")
    print(f"{'='*80}")
    
    exitosos = sum(1 for r in resultados if r['exito'])
    fallidos = len(resultados) - exitosos
    
    print(f"Total archivos procesados: {len(resultados)}")
    print(f"Exitosos: {exitosos}")
    print(f"Fallidos: {fallidos}")
    
    if fallidos > 0:
        print("\nArchivos con error:")
        for r in resultados:
            if not r['exito']:
                print(f"  - {Path(r['archivo_entrada']).name}: {r.get('error', 'Unknown error')}")
        return 1
    
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
