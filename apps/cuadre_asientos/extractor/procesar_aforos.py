#!/usr/bin/env python3
"""
Script para procesar apuntes contables y eliminar aquellos que se compensan.

El script:
1. Lee un archivo Excel con apuntes contables
2. Extrae el código de Hoj Afo/Hoja/Aforo del campo CONCEPTO
3. Agrupa apuntes por código de aforo
4. Identifica pares que se compensan (mismo importe, uno en DEBE y otro en HABER)
5. Elimina los pares compensados
6. Genera un nuevo archivo Excel con los apuntes restantes
"""

import pandas as pd
import re
import sys
from pathlib import Path
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os


def extraer_codigo_aforo(concepto):
    """
    Extrae el código de aforo del campo concepto y lo normaliza.
    
    Soporta formatos variados:
    - Hoj Afo: 2025.00000789.1.1/... -> 202500000789
    - Hoj Afo: 2025 00079776 ... -> 202500079776
    - Hoja/Aforo 2024/00205669 ... -> 202400205669
    - Aforo 2024/204107 ... -> 202400204107
    
    IMPORTANTE: Normaliza el código a 12 dígitos (año + 8 dígitos)
    eliminando puntos, espacios y caracteres especiales.
    
    Args:
        concepto: String con el texto del concepto
        
    Returns:
        String con el código de aforo normalizado (12 dígitos) o None si no se encuentra
    """
    if pd.isna(concepto):
        return None
    
    concepto = str(concepto).strip()
    
    # Patrón más flexible que captura año (4 dígitos) seguido de más dígitos
    # con cualquier separador (puntos, espacios, barras, o nada)
    # Acepta: "Hoj Afo", "H.Afo", "H Afo", "Hoja/Aforo", "Aforo", "Parte Hoj Afo", etc.
    patron_general = r'(?:Parte\s+)?(?:H\.?\s*Afo|Hoj\s+Afo|Hoja/?Aforo|Aforo)[:\s/]*(\d{4})[\s\./]*(\d+)'
    
    match = re.search(patron_general, concepto, re.IGNORECASE)
    if match:
        anio = match.group(1)  # 4 dígitos del año (ej: 2025)
        resto = match.group(2)  # El resto de dígitos
        
        # Limpiar completamente: quitar TODOS los caracteres no numéricos
        resto_limpio = re.sub(r'\D', '', resto)
        
        # Tomar solo los primeros 8 dígitos (rellenar con ceros si es más corto)
        codigo_numerico = resto_limpio[:8].zfill(8)
        
        # Retornar código normalizado de 12 dígitos: YYYYNNNNNNNN
        codigo_normalizado = f"{anio}{codigo_numerico}"
        
        return codigo_normalizado
    
    return None


def normalizar_importe(valor):
    """
    Normaliza un importe a float, manejando diferentes formatos.
    
    Args:
        valor: Valor del importe (puede ser string, float, o vacío)
        
    Returns:
        Float con el importe o 0.0 si está vacío
    """
    if pd.isna(valor) or valor == '' or valor == ' ':
        return 0.0
    
    if isinstance(valor, (int, float)):
        return float(valor)
    
    # Si es string, limpiar y convertir
    valor_str = str(valor).strip().replace(',', '.')
    try:
        return float(valor_str)
    except ValueError:
        return 0.0


def encontrar_pares_compensados(grupo_df):
    """
    Encuentra pares de apuntes que se compensan dentro de un grupo.
    
    Un par se compensa si:
    - Uno tiene importe en DEBE y el otro en HABER
    - Los importes son EXACTAMENTE iguales (sin tolerancia)
    
    Args:
        grupo_df: DataFrame con apuntes del mismo código de aforo
        
    Returns:
        Tupla con:
        - Set con los índices de los registros a eliminar
        - Lista de tuplas (idx_debe, idx_haber) con los pares encontrados
    """
    indices_a_eliminar = set()
    pares_ordenados = []  # Lista de (idx_debe, idx_haber)
    
    # Separar en DEBE y HABER
    debe_df = grupo_df[grupo_df['DEBE_NORMALIZADO'] > 0].copy()
    haber_df = grupo_df[grupo_df['HABER_NORMALIZADO'] > 0].copy()
    
    # Para cada registro en DEBE, buscar coincidencia en HABER
    for idx_debe, row_debe in debe_df.iterrows():
        if idx_debe in indices_a_eliminar:
            continue
            
        importe_debe = row_debe['DEBE_NORMALIZADO']
        
        # Buscar en HABER un importe EXACTAMENTE igual
        for idx_haber, row_haber in haber_df.iterrows():
            if idx_haber in indices_a_eliminar:
                continue
                
            importe_haber = row_haber['HABER_NORMALIZADO']
            
            # Verificar si los importes coinciden EXACTAMENTE (sin tolerancia)
            if importe_debe == importe_haber:
                # ¡Encontramos un par! Marcar ambos para eliminar
                indices_a_eliminar.add(idx_debe)
                indices_a_eliminar.add(idx_haber)
                pares_ordenados.append((idx_debe, idx_haber))  # Guardar el par
                break  # Ya emparejamos este registro de DEBE
    
    return indices_a_eliminar, pares_ordenados


def procesar_archivo(archivo_entrada, archivo_salida=None):
    """
    Procesa el archivo Excel eliminando apuntes compensados.
    
    Args:
        archivo_entrada: Ruta del archivo Excel de entrada
        archivo_salida: Ruta del archivo Excel de salida (opcional)
        
    Returns:
        Tupla (df_original, df_filtrado, estadisticas)
    """
    print(f"Leyendo archivo: {archivo_entrada}")
    df = pd.read_excel(archivo_entrada, sheet_name=0)
    
    print(f"Total de registros leídos: {len(df)}")
    
    # Filtrar filas de totales/resumen (CUENTA contiene "TOTAL" o "HASTA FECHA")
    # Estas filas son líneas de resumen agregadas al final del Excel
    filas_antes = len(df)
    df = df[~df['CUENTA'].astype(str).str.contains('TOTAL|HASTA FECHA', case=False, na=False)]
    filas_despues = len(df)
    
    if filas_antes != filas_despues:
        print(f"Filtradas {filas_antes - filas_despues} filas de totales/resumen")
    
    print(f"Total de registros contables: {len(df)}")
    
    # Extraer códigos de aforo
    print("\nExtrayendo códigos de aforo...")
    df['CODIGO_AFORO'] = df['CONCEPTO'].apply(extraer_codigo_aforo)
    
    registros_con_aforo = df['CODIGO_AFORO'].notna().sum()
    print(f"Registros con código de aforo: {registros_con_aforo}")
    print(f"Códigos de aforo únicos: {df['CODIGO_AFORO'].nunique()}")
    
    # Normalizar importes
    print("\nNormalizando importes...")
    df['DEBE_NORMALIZADO'] = df['IMPORTE DEBE'].apply(normalizar_importe)
    df['HABER_NORMALIZADO'] = df['IMPORTE HABER'].apply(normalizar_importe)
    
    # Encontrar pares compensados
    print("\nBuscando pares compensados...")
    indices_a_eliminar = set()
    todos_los_pares = []  # Lista de todos los pares (idx_debe, idx_haber) ordenados
    
    # Agrupar por código de aforo
    for codigo_aforo, grupo in df[df['CODIGO_AFORO'].notna()].groupby('CODIGO_AFORO'):
        if len(grupo) < 2:
            continue  # No hay suficientes registros para emparejar
        
        # Buscar pares dentro del grupo
        indices_grupo, pares_grupo = encontrar_pares_compensados(grupo)
        indices_a_eliminar.update(indices_grupo)
        todos_los_pares.extend(pares_grupo)  # Agregar los pares de este grupo
    
    print(f"\nRegistros a eliminar (compensados): {len(indices_a_eliminar)}")
    
    # Calcular totales ANTES de crear el DataFrame filtrado
    # Totales del archivo original
    total_debe_original = df['DEBE_NORMALIZADO'].sum()
    total_haber_original = df['HABER_NORMALIZADO'].sum()
    
    # Totales de los registros ELIMINADOS
    # Crear DataFrame ordenado por pares (DEBE, HABER, DEBE, HABER...)
    filas_ordenadas = []
    for idx_debe, idx_haber in todos_los_pares:
        filas_ordenadas.append(idx_debe)  # Primero el DEBE
        filas_ordenadas.append(idx_haber)  # Luego el HABER
    
    df_eliminados = df.loc[filas_ordenadas]
    total_debe_eliminado = df_eliminados['DEBE_NORMALIZADO'].sum()
    total_haber_eliminado = df_eliminados['HABER_NORMALIZADO'].sum()
    
    # Guardar número de registros eliminados ANTES de agregar totales
    num_registros_eliminados = len(df_eliminados)
    
    # Crear DataFrame filtrado
    df_filtrado = df.drop(indices_a_eliminar)
    
    # Totales de los registros RESTANTES
    total_debe_restante = df_filtrado['DEBE_NORMALIZADO'].sum()
    total_haber_restante = df_filtrado['HABER_NORMALIZADO'].sum()
    
    # Guardar número de registros restantes ANTES de agregar fila de totales
    num_registros_restantes = len(df_filtrado)
    
    # Eliminar columnas auxiliares del DataFrame filtrado
    df_filtrado = df_filtrado.drop(['CODIGO_AFORO', 'DEBE_NORMALIZADO', 'HABER_NORMALIZADO'], axis=1)
    
    print(f"Registros restantes: {num_registros_restantes}")
    
    # Agregar fila de totales al final del DataFrame
    # Crear una fila vacía con la estructura del DataFrame
    fila_totales = pd.Series({col: '' for col in df_filtrado.columns})
    
    # Rellenar los valores de la fila de totales (formateados a 2 decimales)
    fila_totales['CUENTA'] = 'TOTAL......'
    fila_totales['CONCEPTO'] = 'TOTALES ASIENTOS RESTANTES'
    fila_totales['IMPORTE DEBE'] = f"{total_debe_restante:.2f}" if total_debe_restante != 0 else ''
    fila_totales['IMPORTE HABER'] = f"{total_haber_restante:.2f}" if total_haber_restante != 0 else ''
    fila_totales['SALDO'] = f"{(total_debe_restante - total_haber_restante):.2f}" if (total_debe_restante - total_haber_restante) != 0 else ''
    
    # Agregar la fila al DataFrame
    df_filtrado = pd.concat([df_filtrado, pd.DataFrame([fila_totales])], ignore_index=True)
    
    # ===== PREPARAR DATAFRAME DE ELIMINADOS (CAZADOS) =====
    # Eliminar columnas auxiliares del DataFrame eliminados
    df_eliminados_export = df_eliminados.drop(['CODIGO_AFORO', 'DEBE_NORMALIZADO', 'HABER_NORMALIZADO'], axis=1)
    
    # Agregar columna de número de par (1, 1, 2, 2, 3, 3, ...)
    # Esto ayuda a visualizar qué registros forman pares
    num_pares = []
    for i in range(len(todos_los_pares)):
        num_pares.append(i + 1)  # DEBE del par i
        num_pares.append(i + 1)  # HABER del par i
    
    # Insertar columna "PAR" al principio
    df_eliminados_export.insert(0, 'PAR', num_pares)
    
    # Agregar fila de totales al DataFrame de eliminados
    fila_totales_elim = pd.Series({col: '' for col in df_eliminados_export.columns})
    fila_totales_elim['PAR'] = ''  # Dejar vacío en la fila de totales
    fila_totales_elim['CUENTA'] = 'TOTAL......'
    fila_totales_elim['CONCEPTO'] = 'TOTALES ASIENTOS CAZADOS (CUADRADOS)'
    fila_totales_elim['IMPORTE DEBE'] = f"{total_debe_eliminado:.2f}" if total_debe_eliminado != 0 else ''
    fila_totales_elim['IMPORTE HABER'] = f"{total_haber_eliminado:.2f}" if total_haber_eliminado != 0 else ''
    fila_totales_elim['SALDO'] = f"{(total_debe_eliminado - total_haber_eliminado):.2f}" if (total_debe_eliminado - total_haber_eliminado) != 0 else '0.00'
    
    df_eliminados_export = pd.concat([df_eliminados_export, pd.DataFrame([fila_totales_elim])], ignore_index=True)
    
    # Estadísticas completas
    estadisticas = {
        'total_original': len(df),
        'registros_con_aforo': registros_con_aforo,
        'aforos_unicos': df['CODIGO_AFORO'].nunique(),
        'registros_eliminados': len(indices_a_eliminar),
        'registros_restantes': num_registros_restantes,
        'porcentaje_eliminado': (len(indices_a_eliminar) / len(df) * 100) if len(df) > 0 else 0,
        # Importes originales
        'total_debe_original': total_debe_original,
        'total_haber_original': total_haber_original,
        'descuadre_original': total_debe_original - total_haber_original,
        # Importes eliminados
        'total_debe_eliminado': total_debe_eliminado,
        'total_haber_eliminado': total_haber_eliminado,
        'descuadre_eliminado': total_debe_eliminado - total_haber_eliminado,
        # Importes restantes
        'total_debe_restante': total_debe_restante,
        'total_haber_restante': total_haber_restante,
        'descuadre_restante': total_debe_restante - total_haber_restante,
    }
    
    # Guardar archivos si se especificó ruta de salida
    if archivo_salida:
        # Guardar archivo de asientos RESTANTES
        print(f"\nGuardando archivos resultantes...")
        print(f"  - Asientos restantes: {archivo_salida}")
        df_filtrado.to_excel(archivo_salida, index=False, sheet_name='HOJA1')
        
        # Generar nombre para archivo de CAZADOS
        # Si el archivo es "archivo_procesado.xlsx", crear "archivo_cazados.xlsx"
        if archivo_salida.endswith('_procesado.xlsx'):
            archivo_cazados = archivo_salida.replace('_procesado.xlsx', '_cazados.xlsx')
        elif archivo_salida.endswith('.xlsx'):
            archivo_cazados = archivo_salida.replace('.xlsx', '_cazados.xlsx')
        else:
            archivo_cazados = archivo_salida + '_cazados.xlsx'
        
        # Guardar archivo de asientos CAZADOS
        print(f"  - Asientos cazados:   {archivo_cazados}")
        df_eliminados_export.to_excel(archivo_cazados, index=False, sheet_name='HOJA1')
        
        print("Archivos guardados exitosamente")
    
    return df, df_filtrado, df_eliminados_export, estadisticas


def mostrar_estadisticas(estadisticas):
    """Muestra las estadísticas del procesamiento."""
    print("\n" + "="*80)
    print("RESUMEN DEL PROCESAMIENTO")
    print("="*80)
    print(f"Total de registros originales:      {estadisticas['total_original']:>10,}")
    print(f"Registros con código de aforo:      {estadisticas['registros_con_aforo']:>10,}")
    print(f"Códigos de aforo únicos:             {estadisticas['aforos_unicos']:>10,}")
    print(f"Registros eliminados (compensados):  {estadisticas['registros_eliminados']:>10,}")
    print(f"Registros restantes:                 {estadisticas['registros_restantes']:>10,}")
    print(f"Porcentaje eliminado:                {estadisticas['porcentaje_eliminado']:>9.2f}%")
    print("="*80)
    
    print("\n" + "="*80)
    print("TOTALES FINANCIEROS - ARCHIVO ORIGINAL")
    print("="*80)
    print(f"Total DEBE:                          {estadisticas['total_debe_original']:>15,.2f} €")
    print(f"Total HABER:                         {estadisticas['total_haber_original']:>15,.2f} €")
    print(f"Descuadre (DEBE - HABER):            {estadisticas['descuadre_original']:>15,.2f} €")
    print("="*80)
    
    print("\n" + "="*80)
    print("TOTALES FINANCIEROS - ASIENTOS ELIMINADOS (Cazados)")
    print("="*80)
    print(f"Total DEBE eliminado:                {estadisticas['total_debe_eliminado']:>15,.2f} €")
    print(f"Total HABER eliminado:               {estadisticas['total_haber_eliminado']:>15,.2f} €")
    print(f"Descuadre (DEBE - HABER):            {estadisticas['descuadre_eliminado']:>15,.2f} €")
    if abs(estadisticas['descuadre_eliminado']) < 0.01:
        print("✓ Los asientos eliminados están perfectamente cuadrados")
    else:
        print(f"⚠ ADVERTENCIA: Hay un descuadre de {abs(estadisticas['descuadre_eliminado']):,.2f} €")
    print("="*80)
    
    print("\n" + "="*80)
    print("TOTALES FINANCIEROS - ASIENTOS RESTANTES (No Cazados)")
    print("="*80)
    print(f"Total DEBE restante:                 {estadisticas['total_debe_restante']:>15,.2f} €")
    print(f"Total HABER restante:                {estadisticas['total_haber_restante']:>15,.2f} €")
    print(f"Descuadre (DEBE - HABER):            {estadisticas['descuadre_restante']:>15,.2f} €")
    print("="*80)
    
    print("\n" + "="*80)
    print("VERIFICACIÓN DE CUADRE")
    print("="*80)
    # Verificar que Original = Eliminado + Restante
    debe_check = abs((estadisticas['total_debe_eliminado'] + estadisticas['total_debe_restante']) - estadisticas['total_debe_original'])
    haber_check = abs((estadisticas['total_haber_eliminado'] + estadisticas['total_haber_restante']) - estadisticas['total_haber_original'])
    
    if debe_check < 0.01 and haber_check < 0.01:
        print("✓ Verificación OK: ELIMINADO + RESTANTE = ORIGINAL")
    else:
        print("⚠ ERROR: Los totales no cuadran correctamente")
        print(f"  Diferencia DEBE:  {debe_check:,.2f} €")
        print(f"  Diferencia HABER: {haber_check:,.2f} €")
    
    # Mostrar el cambio en el descuadre
    mejora_descuadre = abs(estadisticas['descuadre_original']) - abs(estadisticas['descuadre_restante'])
    print(f"\nCambio en descuadre:")
    print(f"  Descuadre original: {abs(estadisticas['descuadre_original']):>15,.2f} €")
    print(f"  Descuadre final:    {abs(estadisticas['descuadre_restante']):>15,.2f} €")
    if mejora_descuadre > 0:
        print(f"  Mejora:             {mejora_descuadre:>15,.2f} € ✓")
    elif mejora_descuadre < 0:
        print(f"  Empeoramiento:      {abs(mejora_descuadre):>15,.2f} € ⚠")
    else:
        print(f"  Sin cambio:         {mejora_descuadre:>15,.2f} €")
    print("="*80)


def main():
    """Función principal."""
    import argparse
    
    # Configurar argumentos de línea de comandos
    parser = argparse.ArgumentParser(
        description='Procesa apuntes contables eliminando pares compensados.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  
  # Procesar archivo (salida automática con sufijo '_procesado')
  python procesar_aforos.py archivo.xlsx
  
  # Especificar archivo de salida
  python procesar_aforos.py archivo.xlsx -o resultado.xlsx
  
  # Ver estadísticas sin guardar archivo
  python procesar_aforos.py archivo.xlsx --no-guardar
        """
    )
    
    parser.add_argument('archivo_entrada', 
                        help='Archivo Excel de entrada (.xlsx)')
    parser.add_argument('-o', '--output', 
                        dest='archivo_salida',
                        help='Archivo Excel de salida (opcional, por defecto añade "_procesado")')
    parser.add_argument('--no-guardar', 
                        action='store_true',
                        help='Solo mostrar estadísticas sin guardar archivo')
    
    args = parser.parse_args()
    
    archivo_entrada = args.archivo_entrada
    
    # Determinar archivo de salida
    if args.no_guardar:
        archivo_salida = None
    elif args.archivo_salida:
        archivo_salida = args.archivo_salida
    else:
        # Generar nombre automático
        ruta = Path(archivo_entrada)
        archivo_salida = str(ruta.parent / f"{ruta.stem}_procesado{ruta.suffix}")
    
    # Verificar que existe el archivo de entrada
    if not Path(archivo_entrada).exists():
        print(f"ERROR: No se encuentra el archivo {archivo_entrada}")
        sys.exit(1)
    
    # Procesar archivo
    try:
        df_original, df_filtrado, df_eliminados, estadisticas = procesar_archivo(
            archivo_entrada, 
            archivo_salida
        )
        
        # Mostrar estadísticas
        mostrar_estadisticas(estadisticas)
        
        # Mostrar algunos ejemplos de lo que se eliminó
        print("\n" + "="*80)
        print("EJEMPLOS DE PARES ELIMINADOS")
        print("="*80)
        
        # Reconstruir el proceso para mostrar ejemplos
        df_original['CODIGO_AFORO'] = df_original['CONCEPTO'].apply(extraer_codigo_aforo)
        df_original['DEBE_NORMALIZADO'] = df_original['IMPORTE DEBE'].apply(normalizar_importe)
        df_original['HABER_NORMALIZADO'] = df_original['IMPORTE HABER'].apply(normalizar_importe)
        
        # Tomar el primer código de aforo con múltiples registros
        codigo_counts = df_original['CODIGO_AFORO'].value_counts()
        codigos_multiples = codigo_counts[codigo_counts > 1].head(3)
        
        for codigo in codigos_multiples.index:
            grupo = df_original[df_original['CODIGO_AFORO'] == codigo][
                ['CONCEPTO', 'IMPORTE DEBE', 'IMPORTE HABER']
            ]
            print(f"\nCódigo de aforo: {codigo}")
            print("-"*80)
            print(grupo.to_string())
            print()
        
        print("="*80)
        print("\n✓ Procesamiento completado exitosamente")
        
    except Exception as e:
        print(f"\nERROR durante el procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()