#!/usr/bin/env python3
from __future__ import annotations

"""
dossier_convert.py  –  Convierte el DESGLOSE de partidas al formato CSV
de carga de dossier exportación de Visualtrans (CONVERT).

Clientes soportados actualmente:
  · ALDI Melilla  (formato por defecto)

ENTRADA:  CSV separado por ";" (ISO-8859 o UTF-8), con cabecera y columnas:
            Codigo_Arancelario, Origen_Mercancia, Divisa, Precio, Masa_Bruta,
            Masa_Neta, Numero_Bultos, Tipo_Bultos, Contenedor, REGIMEN, DESCRIPTION

SALIDA:   CSV separado por ";" (UTF-8) listo para cargar en Visual.
            · Línea 1:    Cabecera con datos fijos del cliente + totales.
            · Líneas 2..N: Una partida por cada fila con PRICE > 0.

USO:
    python dossier_convert.py desglose.csv -o salida/
    python dossier_convert.py desglose1.csv desglose2.csv -o salida/ --fecha 15/04/2026
"""

import argparse
import csv
import sys
from dataclasses import dataclass
from datetime import date as Date
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path
from typing import Optional

SCRIPT_VERSION = "2026-04-15.v1"

SCRIPT_CHANGELOG = """
## 2026-04-15.v1

### Lógica general
Convierte el DESGLOSE de partidas enviado por ALDI al formato CSV de carga
de dossier de exportación de Visualtrans (CONVERT), equivalente al generado
por la herramienta SOUTO en https://cliente.souto-manager.com/total-logistic

### Clientes soportados
- ALDI Melilla (único cliente activo; se añadirán más cuando se reciban sus ficheros)

### Entrada
CSV separado por ";" (codificación ISO-8859 o UTF-8-sig) con cabecera y columnas:
  Codigo_Arancelario, Origen_Mercancia, Divisa, Precio, Masa_Bruta,
  Masa_Neta, Numero_Bultos, Tipo_Bultos, Contenedor, REGIMEN, DESCRIPTION

También acepta el formato inglés SOUTO:
  TARIC_NUMBER, COUNTRY_ORIGIN, CURRENCY, PRICE, GROSS_WEIGHT,
  NET_WEIGHT, BOXES, TYPE, TRAID

### Reglas de transformación por partida
- Si PRICE  = 0 → se omite la línea
- Si NET_WEIGHT  = 0 → se usa GROSS_WEIGHT en ambos pesos
- Si GROSS_WEIGHT = 0 → se usa NET_WEIGHT en ambos pesos
- El importe en partidas usa "," como separador decimal (formato Visual)
- Los pesos en la cabecera se redondean al entero más cercano

### Salida
  {out_dir}/Convert_-_{nombre_entrada}.csv
"""


# ---------------------------------------------------------------------------
# Configuración fija del cliente (ALDI Melilla)
# ---------------------------------------------------------------------------

@dataclass
class ClientConfig:
    """
    Datos fijos del cliente y la ruta.  Cambiar aquí para adaptar a otro cliente
    sin tocar la lógica de conversión.
    """
    # — Cabecera: campos de control —
    tipo_registro_cab: str = "1"
    modo_declaracion:  str = "E"          # E = Exportación
    regimen_cabecera:  str = "EX"
    estatuto:          str = "A Normal"
    expediente:        str = "2911"

    # — Exportador —
    exportador_nombre:  str = "ALDI DOS HERMANAS SUPERMERCADOS, S.L"
    exportador_nif:     str = "B91405142"
    exportador_dir:     str = "POLIGONO INDUSTRIAL DE LA ISLA, C/ TORRE DE LOS HERBEROS, S/N"
    exportador_cp:      str = "41703"
    exportador_ciudad:  str = "DOS HERMANAS"
    exportador_pais:    str = "ES"

    # — Destinatario —
    destinatario_nombre:  str = "ALDI DOS HERMANAS SUPERMERCADOS, S.L."
    destinatario_nif:     str = "B91405142"
    destinatario_dir:     str = "AVENIDA GENERAL POLAVIEJA, 5$"
    destinatario_cp:      str = "52006"
    destinatario_ciudad:  str = "MELILLA"
    destinatario_pais:    str = "ES"

    # — Ruta y condiciones —
    pais_expedicion:  str = "ES"
    pais_destino:     str = "MELILLA"
    hay_contenedor:   str = "0"
    incoterm:         str = "EXW"
    lugar_incoterm:   str = "SEVILLA"
    puerto_destino:   str = "RUSADIR"
    campo_cy:         str = "CY"          # terminal marítima
    divisa_nombre:    str = "EUROS"
    divisa_codigo:    str = "99"          # código Visual para EUR
    modo_transporte:  str = "1 Maritimo"
    aduana_salida:    str = "ES002911"
    aduana_despacho:  str = "ES002911"
    campo_503300:     str = "503300"      # código interno Visual
    pais_fin:         str = "ES"
    campo_fin:        str = "11"          # campo fijo final cabecera

    # — Partidas: campos fijos —
    provincia_origen: str = "29"          # provincia origen mercancía
    cod_doc_01:       str = "1003"        # tipo documento 1 (factura N380)
    cod_doc_02:       str = "1833"        # tipo documento 2 (conocimiento N705)


# Configuración por defecto = ALDI Melilla
ALDI_MELILLA = ClientConfig()


# ---------------------------------------------------------------------------
# Modelo de datos para cada fila del DESGLOSE
# ---------------------------------------------------------------------------

@dataclass
class DesgloseRow:
    taric:          str
    description:    str
    country_origin: str
    currency:       str
    price:          Decimal
    gross_weight:   Decimal
    net_weight:     Decimal
    boxes:          int
    type_box:       str
    traid:          str
    regimen:        str


# ---------------------------------------------------------------------------
# Parsing / normalización
# ---------------------------------------------------------------------------

def _parse_decimal(raw: str) -> Decimal:
    """Convierte un token numérico (con '.' o ',' como decimal) a Decimal."""
    if not raw:
        return Decimal("0")
    clean = raw.strip().replace(",", ".")
    try:
        return Decimal(clean)
    except InvalidOperation:
        return Decimal("0")


def _fmt_precio_partida(value: Decimal) -> str:
    """Precio de partida: 2 decimales, separador coma (formato Visual)."""
    rounded = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return str(rounded).replace(".", ",")


def _fmt_importe_cabecera(value: Decimal) -> str:
    """Importe total de cabecera: 2 decimales, separador punto."""
    rounded = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return str(rounded)


def _fmt_peso_cabecera(value: Decimal) -> str:
    """Peso en cabecera: redondeado al entero más cercano, sin decimales."""
    rounded = value.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return str(int(rounded))


def parse_desglose(csv_path: Path) -> list[DesgloseRow]:
    """
    Lee el DESGLOSE y devuelve la lista de filas parseadas.
    Acepta tanto cabeceras en español (Codigo_Arancelario…) como en inglés
    (TARIC_NUMBER…) para compatibilidad con la especificación SOUTO.
    """
    raw = csv_path.read_bytes()
    for enc in ("latin-1", "utf-8-sig", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(f"No se puede decodificar {csv_path.name}")

    # Normalizar separador de línea
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    reader = csv.DictReader(text.splitlines(), delimiter=";")

    # Mapeo de nombres de columna alternativos → nombre canónico interno
    COL_MAP = {
        # Formato ALDI (español)
        "Codigo_Arancelario": "taric",
        "DESCRIPTION":        "description",
        "Origen_Mercancia":   "country_origin",
        "Divisa":             "currency",
        "Precio":             "price",
        "Masa_Bruta":         "gross_weight",
        "Masa_Neta":          "net_weight",
        "Numero_Bultos":      "boxes",
        "Tipo_Bultos":        "type_box",
        "Contenedor":         "traid",
        "REGIMEN":            "regimen",
        # Formato SOUTO (inglés)
        "TARIC_NUMBER":   "taric",
        "COUNTRY_ORIGIN": "country_origin",
        "CURRENCY":       "currency",
        "PRICE":          "price",
        "GROSS_WEIGHT":   "gross_weight",
        "NET_WEIGHT":     "net_weight",
        "BOXES":          "boxes",
        "TYPE":           "type_box",
        "TRAID":          "traid",
    }

    rows: list[DesgloseRow] = []
    for i, raw_row in enumerate(reader, start=2):
        mapped: dict = {}
        for col, val in raw_row.items():
            if col is None:
                continue
            key = COL_MAP.get(col.strip())
            if key:
                mapped[key] = (val or "").strip()

        missing = {k for k in ("taric", "price", "gross_weight", "net_weight", "boxes")
                   if k not in mapped}
        if missing:
            print(f"[WARN] Fila {i}: columnas requeridas no encontradas, fila ignorada ({missing})", file=sys.stderr)
            continue

        price = _parse_decimal(mapped.get("price", "0"))
        gross = _parse_decimal(mapped.get("gross_weight", "0"))
        net   = _parse_decimal(mapped.get("net_weight", "0"))

        try:
            boxes = int(Decimal(mapped.get("boxes", "0").replace(",", "."))
                        .quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        except InvalidOperation:
            boxes = 0

        rows.append(DesgloseRow(
            taric          = mapped.get("taric", ""),
            description    = mapped.get("description", ""),
            country_origin = mapped.get("country_origin", ""),
            currency       = mapped.get("currency", ""),
            price          = price,
            gross_weight   = gross,
            net_weight     = net,
            boxes          = boxes,
            type_box       = mapped.get("type_box", ""),
            traid          = mapped.get("traid", ""),
            regimen        = mapped.get("regimen", ""),
        ))

    return rows


# ---------------------------------------------------------------------------
# Reglas de transformación
# ---------------------------------------------------------------------------

def apply_weight_rules(row: DesgloseRow) -> tuple[Decimal, Decimal]:
    """
    Devuelve (gross, net) aplicando las reglas:
      - NET  = 0  →  usar GROSS en ambos
      - GROSS = 0 →  usar NET  en ambos
    """
    gross = row.gross_weight
    net   = row.net_weight

    if net == Decimal("0"):
        net = gross
    if gross == Decimal("0"):
        gross = net

    return gross, net


def filter_valid(rows: list[DesgloseRow]) -> list[DesgloseRow]:
    """Excluye filas con PRICE = 0."""
    return [r for r in rows if r.price > Decimal("0")]


# ---------------------------------------------------------------------------
# Generación del CONVERT
# ---------------------------------------------------------------------------

def _cabecera_row(
    cfg: ClientConfig,
    valid_rows: list[DesgloseRow],
    fecha_doc: str,
) -> list[str]:
    """
    Construye la línea de cabecera del CONVERT (posición 1 del fichero).
    Los totales se calculan sumando las filas válidas con las reglas de peso.
    """
    total_price  = Decimal("0")
    total_gross  = Decimal("0")
    total_net    = Decimal("0")

    # El contenedor se toma del primer registro (en ALDI es siempre el mismo)
    contenedor = valid_rows[0].traid if valid_rows else ""

    for row in valid_rows:
        total_price += row.price
        g, n = apply_weight_rules(row)
        total_gross += g
        total_net   += n

    # 54 campos de cabecera + campos finales; posiciones vacías = ""
    # Estructura verificada contra el fichero Convert real de ALDI Melilla.
    fields = [
        cfg.tipo_registro_cab,    # [0]  tipo registro
        cfg.modo_declaracion,     # [1]  E = exportación
        cfg.regimen_cabecera,     # [2]  EX
        cfg.estatuto,             # [3]  A Normal
        "",                       # [4]  vacío
        cfg.expediente,           # [5]  número expediente Visual
        cfg.exportador_nombre,    # [6]
        cfg.exportador_nif,       # [7]
        cfg.exportador_dir,       # [8]
        cfg.exportador_cp,        # [9]
        cfg.exportador_ciudad,    # [10]
        cfg.exportador_pais,      # [11]
        cfg.destinatario_nombre,  # [12]
        cfg.destinatario_nif,     # [13]
        cfg.destinatario_dir,     # [14]
        cfg.destinatario_cp,      # [15]
        cfg.destinatario_ciudad,  # [16]
        cfg.destinatario_pais,    # [17]
        cfg.pais_expedicion,      # [18]
        cfg.pais_destino,         # [19]
        "",                       # [20]
        cfg.hay_contenedor,       # [21]
        cfg.incoterm,             # [22]
        cfg.lugar_incoterm,       # [23]
        cfg.puerto_destino,       # [24]
        cfg.campo_cy,             # [25]  terminal marítima
        cfg.divisa_nombre,        # [26]
        cfg.divisa_codigo,        # [27]
        cfg.modo_transporte,      # [28]
        cfg.aduana_salida,        # [29]
        cfg.aduana_despacho,      # [30]
        cfg.campo_503300,         # [31]
        "",                       # [32]
        _fmt_peso_cabecera(total_gross),      # [33] peso bruto total
        _fmt_peso_cabecera(total_net),        # [34] peso neto total
        contenedor,                           # [35] referencia contenedor
        cfg.exportador_nif,                   # [36] NIF (repetido)
        _fmt_importe_cabecera(total_price),   # [37] importe total factura
        # [38..63] vacíos (26 campos)
        *[""] * 26,
        cfg.pais_fin,             # [64]
        "",                       # [65]
        "",                       # [66]
        cfg.campo_fin,            # [67]
    ]

    return fields


def _partida_row(
    cfg: ClientConfig,
    row: DesgloseRow,
    fecha_doc: str,
) -> list[str]:
    """
    Construye una línea de partida del CONVERT a partir de una fila del DESGLOSE.
    Estructura verificada contra el fichero Convert real de ALDI Melilla.
    """
    fields = [
        "0",                              # [0]  tipo registro partida
        "0",                              # [1]  campo fijo
        str(row.boxes),                   # [2]  número de bultos
        f" {row.description}",            # [3]  descripción (con espacio inicial, como en el original)
        "",                               # [4]  vacío (supl. tipo)
        "",                               # [5]  vacío (supl. cantidad)
        row.taric,                        # [6]  código arancelario
        _fmt_precio_partida(row.price),   # [7]  importe (coma decimal)
        "",                               # [8]  contenedor partida
        row.country_origin,               # [9]  país origen
        cfg.provincia_origen,             # [10] provincia origen
        "",                               # [11] doc. cargo cabecera
        "",                               # [12] doc. cargo partida
        row.regimen,                      # [13] régimen aduanero
        "",                               # [14] no usar
        "",                               # [15] no usar
        "",                               # [16] no usar
        cfg.cod_doc_01,                   # [17] tipo documento 1
        "",                               # [18] número documento 1
        fecha_doc,                        # [19] fecha documento 1
        cfg.cod_doc_02,                   # [20] tipo documento 2
        "",                               # [21] número documento 2
        fecha_doc,                        # [22] fecha documento 2
        # [23..end] vacíos
        *[""] * 41,
    ]

    return fields


def generate_convert(
    desglose_path: Path,
    out_dir: Path,
    cfg: ClientConfig = ALDI_MELILLA,
    fecha_doc: Optional[str] = None,
) -> Path:
    """
    Función principal de conversión.

    Lee el DESGLOSE, aplica las reglas de transformación y escribe el CONVERT
    en out_dir con el nombre  Convert_-_{stem_entrada}.csv

    Returns:
        Path del fichero generado.
    """
    if fecha_doc is None:
        today = Date.today()
        fecha_doc = f"{today.day}/{today.month}/{today.year}"

    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"Convert_-_{desglose_path.stem}.csv"

    # 1. Parsear entrada
    all_rows   = parse_desglose(desglose_path)
    valid_rows = filter_valid(all_rows)

    skipped = len(all_rows) - len(valid_rows)
    if skipped:
        print(f"[INFO] {skipped} filas omitidas (PRICE = 0)", file=sys.stderr)

    if not valid_rows:
        print("[WARN] No hay filas válidas. Se generará el fichero de salida solo con cabecera.", file=sys.stderr)

    # 2. Construir filas de salida
    cabecera = _cabecera_row(cfg, valid_rows, fecha_doc)
    partidas = [_partida_row(cfg, r, fecha_doc) for r in valid_rows]

    # 3. Escribir CSV
    with out_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, delimiter=";", quoting=csv.QUOTE_MINIMAL)
        writer.writerow(cabecera)
        for p in partidas:
            writer.writerow(p)

    # 4. Resumen
    total_price = sum(r.price for r in valid_rows)
    total_gross = sum(apply_weight_rules(r)[0] for r in valid_rows)
    total_net   = sum(apply_weight_rules(r)[1] for r in valid_rows)

    total_price_r = total_price.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    total_gross_r = total_gross.quantize(Decimal("1"),    rounding=ROUND_HALF_UP)
    total_net_r   = total_net.quantize(Decimal("1"),      rounding=ROUND_HALF_UP)

    print(f"[INFO] Fichero:             {desglose_path.name}")
    print(f"[INFO] Partidas procesadas: {len(valid_rows)}")
    print(f"[INFO] Total importe:       {total_price_r:.2f}")
    print(f"[INFO] Total peso bruto:    {int(total_gross_r)}")
    print(f"[INFO] Total peso neto:     {int(total_net_r)}")
    print(f"[INFO] Fecha documentos:    {fecha_doc}")

    return out_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description=(
            "Convierte el DESGLOSE de partidas al formato CSV "
            "de carga de dossier exportación de Visualtrans."
        )
    )
    ap.add_argument(
        "inputs",
        nargs="+",
        type=Path,
        help="Fichero(s) DESGLOSE de entrada (CSV separado por ';')",
    )
    ap.add_argument(
        "-o", "--out",
        type=Path,
        default=Path("out"),
        help="Carpeta de salida (por defecto: ./out)",
    )
    ap.add_argument(
        "--fecha",
        metavar="DD/MM/YYYY",
        default=None,
        help=(
            "Fecha a usar en los campos de documento de cada partida "
            "(por defecto: fecha de hoy)"
        ),
    )
    ap.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {SCRIPT_VERSION}",
    )
    args = ap.parse_args()

    errors = 0
    for input_path in args.inputs:
        if not input_path.exists():
            print(f"ERROR: no se encuentra el fichero: {input_path}", file=sys.stderr)
            errors += 1
            continue

        try:
            out_path = generate_convert(
                desglose_path=input_path,
                out_dir=args.out,
                cfg=ALDI_MELILLA,
                fecha_doc=args.fecha,
            )
            print(f"OK -> {out_path}")
        except Exception as exc:
            print(f"ERROR [{input_path.name}]: {exc}", file=sys.stderr)
            errors += 1
            raise

    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
