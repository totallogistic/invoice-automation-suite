#!/usr/bin/env python3
"""
INTRASTAT — Generador del CSV/TXT para AEAT
============================================

Convierte el archivo Excel que exporta el programa "Visual" (con UNA sola hoja
y las 8 columnas del DUA) en el CSV plano separado por ";" que admite la AEAT.

Replica EXACTAMENTE la lógica del modelo Excel del cliente (hojas "pegar
archivo visual" y "intrastat archivo prueba"):

    Columna A (País destino)        → tal cual                        [CASILLA 17]
    Constante                       → "11"                             (fijo)
    Columna B (Condición entrega)   → tal cual                        [CASILLA 20]
    Columna C (Naturaleza transac.) → tal cual                        [CASILLA 24]
    Constante                       → "1"                              (fijo)
    Constante                       → "1131"                           (fijo)
    Columna D (Partida estadística) → LEFT(D, 8) primeros 8           [CASILLA 33]
    Columna E (País origen)         → "MA" si MA, si no "PL"          [CASILLA 34]
    Condicional sobre E             → "2" si MA, "4" si no             (régimen)
    Columna F (Masa neta)           → F / 1000  (gramos → kg)         [CASILLA 38]
    Columna vacía                   → ""                               (separador)
    Columna G (Precio en divisas)   → IF(G<>0, G, H)                  [CASILLA 42]
    Columna H (Valor estadístico)   → tal cual                        [CASILLA 46]
    Constante                       → "PL5262654857"                   (NIF)

USO:
    python intrastat_generator.py <archivo_visual.xlsx>  [-o salida.csv]  [-n NIF]
"""

SCRIPT_VERSION = "2026-05-27.v1"
SCRIPT_CHANGELOG = """
## 2026-05-27.v1

### Lógica general
Convierte el Excel exportado de Visual (8 columnas DUA) en el CSV plano AEAT.
Replica byte a byte la salida del modelo Excel del cliente.

### Entradas
- `.xlsx`, `.ods` o `.csv` con 8 columnas (País destino, Cond. entrega,
  Naturaleza, Partida estadística, País origen, Masa neta, Precio, Valor).

### Transformaciones
- Constantes fijas: `11`, `1`, `1131`, NIF declarante.
- Partida estadística truncada a 8 caracteres (CASILLA 33).
- País origen: `MA` si extracomunitario, `PL` en otro caso (régimen 2 vs 4).
- Masa neta convertida de gramos a kg (`F/1000`).
- Valor estadístico: `G` si distinto de 0, si no `H`.

### Salida
CSV separado por `;`, CRLF, UTF-8 sin BOM, sin terminador final.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Constantes (datos fijos de la declaración)
# ---------------------------------------------------------------------------
NIF_DEFECTO = "PL5262654857"
ORIGEN_EXTRACOMUNITARIO = "MA"
DATO_FIJO_2 = "11"
DATO_FIJO_5 = "1"
DATO_FIJO_6 = "1131"


# ---------------------------------------------------------------------------
# Lectura del archivo VEA (xlsx / ods / csv)
# ---------------------------------------------------------------------------
def cargar_vea(ruta: Path) -> pd.DataFrame:
    """
    Carga el archivo exportado del programa Visual.

    Detecta automáticamente si la primera fila es cabecera (cuando hay texto
    en las columnas que deben ser numéricas) y la descarta. Devuelve un
    DataFrame con 8 columnas re-nombradas a A..H por posición.
    """
    ext = ruta.suffix.lower()

    if ext == ".csv":
        try:
            df = pd.read_csv(ruta, sep=";", header=None, dtype=object, encoding="utf-8")
            if df.shape[1] < 8:
                df = pd.read_csv(ruta, sep=",", header=None, dtype=object, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(ruta, sep=";", header=None, dtype=object, encoding="latin-1")
    elif ext == ".xlsx":
        df = pd.read_excel(ruta, header=None, dtype=object, engine="openpyxl")
    elif ext == ".ods":
        df = pd.read_excel(ruta, header=None, dtype=object, engine="odf")
    else:
        raise ValueError(f"Extensión no soportada: {ext}. Usa .xlsx, .ods o .csv")

    if df.shape[1] < 8:
        raise ValueError(
            f"El archivo debe tener al menos 8 columnas (tiene {df.shape[1]}). "
            f"Columnas esperadas: País destino, Condición entrega, Naturaleza "
            f"transacción, Partida estadística, País origen, Masa neta, Precio "
            f"en divisas, Valor estadístico."
        )

    # Quedarse con las 8 primeras columnas y re-nombrarlas A..H por posición
    df = df.iloc[:, :8]
    df.columns = ["A", "B", "C", "D", "E", "F", "G", "H"]

    # Detectar y descartar fila de cabecera: si en la primera fila
    # las columnas F (masa) o H (valor) NO son numéricas, es cabecera.
    primera = df.iloc[0]
    es_cabecera = not (_es_numero(primera["F"]) and _es_numero(primera["H"]))
    if es_cabecera:
        df = df.iloc[1:].reset_index(drop=True)

    # Eliminar filas completamente vacías
    df = df.dropna(how="all").reset_index(drop=True)

    return df


def _es_numero(valor) -> bool:
    if valor is None:
        return False
    if isinstance(valor, (int, float)):
        return not (isinstance(valor, float) and pd.isna(valor))
    try:
        float(str(valor).replace(",", "."))
        return True
    except (ValueError, TypeError):
        return False


# ---------------------------------------------------------------------------
# Formateo numérico estilo Excel español
# ---------------------------------------------------------------------------
def fmt_num(valor) -> str:
    """
    Convierte un número al formato que produce Excel al exportar a CSV
    en localización española:
      - Coma como separador decimal
      - Sin ceros finales innecesarios (3335.0 → "3335", 8042.30 → "8042,3")
      - Sin notación científica
    """
    if valor is None:
        return ""
    if isinstance(valor, str):
        s = valor.strip()
        if not s:
            return ""
        try:
            valor = float(s.replace(",", "."))
        except ValueError:
            return s
    try:
        f = float(valor)
    except (TypeError, ValueError):
        return str(valor)
    if pd.isna(f):
        return ""
    if f == int(f):
        return str(int(f))
    s = f"{f:.10f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def _safe_float(valor, default: float = 0.0) -> float:
    if valor is None:
        return default
    if isinstance(valor, str):
        s = valor.strip()
        if not s:
            return default
        try:
            return float(s.replace(",", "."))
        except ValueError:
            return default
    try:
        f = float(valor)
        return default if pd.isna(f) else f
    except (TypeError, ValueError):
        return default


# ---------------------------------------------------------------------------
# Lógica de transformación por fila (replica las fórmulas del Excel modelo)
# ---------------------------------------------------------------------------
def construir_linea(fila: pd.Series, nif: str) -> str:
    pais_destino = str(fila["A"]).strip() if pd.notna(fila["A"]) else ""
    entrega = str(fila["B"]).strip() if pd.notna(fila["B"]) else ""
    transaccion = str(fila["C"]).strip()
    # CASILLA 24 a veces viene como float (11.0). La normalizamos a entero.
    try:
        f_trans = float(transaccion)
        if f_trans == int(f_trans):
            transaccion = str(int(f_trans))
    except (TypeError, ValueError):
        pass

    # CASILLA 33 = LEFT(D, 8) — primeros 8 caracteres
    partida_raw = "" if pd.isna(fila["D"]) else str(fila["D"]).strip()
    # Si viene como número (ej. 9401992090.0), quitar el ".0" final
    try:
        f_part = float(partida_raw)
        if f_part == int(f_part):
            partida_raw = str(int(f_part))
    except (TypeError, ValueError):
        pass
    casilla_33 = partida_raw[:8]

    # CASILLA 34 = IF(E="MA","MA","PL")  +  régimen IF(MA, 2, 4)
    origen_raw = "" if pd.isna(fila["E"]) else str(fila["E"]).strip().upper()
    if origen_raw == ORIGEN_EXTRACOMUNITARIO:
        origen = "MA"
        regimen = "2"
    else:
        origen = "PL"
        regimen = "4"

    # CASILLA 38 = F / 1000  (masa neta de gramos a kilos)
    masa_gramos = _safe_float(fila["F"])
    casilla_38 = fmt_num(masa_gramos / 1000.0)

    # CASILLA 42 = IF(G<>0, G, H)  (si no hay precio en divisas, usar valor)
    precio_divisas = _safe_float(fila["G"])
    valor_factura = _safe_float(fila["H"])
    valor_estadistico = precio_divisas if precio_divisas != 0 else valor_factura
    casilla_42 = fmt_num(valor_estadistico)

    # CASILLA 46 = H tal cual
    casilla_46 = fmt_num(valor_factura)

    campos = [
        pais_destino,   # 1
        DATO_FIJO_2,    # 2  → "11"
        entrega,        # 3
        transaccion,    # 4
        DATO_FIJO_5,    # 5  → "1"
        DATO_FIJO_6,    # 6  → "1131"
        casilla_33,     # 7
        origen,         # 8
        regimen,        # 9
        casilla_38,     # 10  (F/1000)
        "",             # 11  campo vacío
        casilla_42,     # 12  (G ó H si G=0)
        casilla_46,     # 13  (H)
        nif,            # 14
    ]
    return ";".join(campos)


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------
def generar(ruta_entrada: Path, ruta_salida: Path, nif: str) -> int:
    df = cargar_vea(ruta_entrada)
    lineas = [construir_linea(fila, nif) for _, fila in df.iterrows()]

    # Formato AEAT: CRLF entre líneas, sin terminador final, UTF-8 sin BOM
    contenido = "\r\n".join(lineas)
    ruta_salida.write_bytes(contenido.encode("utf-8"))
    return len(lineas)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Genera el CSV INTRASTAT a partir del archivo exportado de Visual."
    )
    parser.add_argument("entrada", type=Path,
                        help="Archivo Visual (.xlsx, .ods o .csv) con las 8 columnas del DUA")
    parser.add_argument("-o", "--output", type=Path, default=None,
                        help="Ruta del CSV de salida (por defecto: intrastat.csv junto al de entrada)")
    parser.add_argument("-n", "--nif", default=NIF_DEFECTO,
                        help=f"NIF del declarante (por defecto: {NIF_DEFECTO})")
    args = parser.parse_args()

    if not args.entrada.exists():
        print(f"ERROR: no se encuentra el archivo {args.entrada}", file=sys.stderr)
        sys.exit(1)

    salida = args.output or args.entrada.with_name("intrastat.csv")

    try:
        n = generar(args.entrada, salida, args.nif)
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(2)

    print(f"OK -> {n} línea(s) escrita(s) en {salida}")


if __name__ == "__main__":
    main()