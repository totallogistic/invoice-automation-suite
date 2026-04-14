#!/usr/bin/env python3
from __future__ import annotations

"""
export_visual.py  –  Convierte el fichero de partidas de un cliente al formato
CSV de carga de dossier exportación de Visualtrans.

ENTRADA:  CSV separado por ";" con las partidas (columnas según cliente).
SALIDA:   CSV separado por ";" listo para importar en Visual.
            · Línea 1:    Cabecera con datos fijos del cliente + totales.
            · Líneas 2..N: Una partida por cada fila con PRICE > 0.

USO:
    python export_visual.py desglose.csv -o salida/ --cliente aldi --fecha 15/04/2026
    python export_visual.py desglose.csv -o salida/ --cliente aldi
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
Convierte el fichero de partidas enviado por un cliente al formato CSV de carga
de dossier de exportación de Visualtrans. Soporta múltiples clientes mediante
una configuración por cliente (ClientConfig). La estructura de salida es siempre
la misma; lo que varía por cliente son los datos fijos de cabecera y el mapeo
de columnas de entrada.

### Clientes soportados
- aldi: ALDI Dos Hermanas → Melilla (export marítimo vía RUSADIR)

### Entrada (ALDI)
CSV separado por ";" (ISO-8859) con columnas:
  Codigo_Arancelario, Origen_Mercancia, Divisa, Precio, Masa_Bruta,
  Masa_Neta, Numero_Bultos, Tipo_Bultos, Contenedor, REGIMEN, DESCRIPTION

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
# Configuraciones de cliente
# ---------------------------------------------------------------------------

@dataclass
class ClientConfig:
    """Datos fijos del cliente y la ruta. Un objeto por cliente."""
    # Cabecera Visual
    tipo_registro_cab: str = "1"
    modo_declaracion:  str = "E"
    regimen_cabecera:  str = "EX"
    estatuto:          str = "A Normal"
    expediente:        str = ""

    # Exportador
    exportador_nombre:  str = ""
    exportador_nif:     str = ""
    exportador_dir:     str = ""
    exportador_cp:      str = ""
    exportador_ciudad:  str = ""
    exportador_pais:    str = "ES"

    # Destinatario
    destinatario_nombre:  str = ""
    destinatario_nif:     str = ""
    destinatario_dir:     str = ""
    destinatario_cp:      str = ""
    destinatario_ciudad:  str = ""
    destinatario_pais:    str = "ES"

    # Ruta
    pais_expedicion:  str = "ES"
    pais_destino:     str = ""
    hay_contenedor:   str = "0"
    incoterm:         str = ""
    lugar_incoterm:   str = ""
    puerto_destino:   str = ""
    campo_cy:         str = "CY"
    divisa_nombre:    str = "EUROS"
    divisa_codigo:    str = "99"
    modo_transporte:  str = "1 Maritimo"
    aduana_salida:    str = ""
    aduana_despacho:  str = ""
    campo_503300:     str = "503300"
    pais_fin:         str = "ES"
    campo_fin:        str = "11"

    # Partidas
    provincia_origen: str = "29"
    cod_doc_01:       str = "1003"
    cod_doc_02:       str = "1833"

    # Mapeo de columnas de entrada (nombre columna → campo interno)
    col_taric:         str = "Codigo_Arancelario"
    col_description:   str = "DESCRIPTION"
    col_country_origin:str = "Origen_Mercancia"
    col_currency:      str = "Divisa"
    col_price:         str = "Precio"
    col_gross_weight:  str = "Masa_Bruta"
    col_net_weight:    str = "Masa_Neta"
    col_boxes:         str = "Numero_Bultos"
    col_type_box:      str = "Tipo_Bultos"
    col_traid:         str = "Contenedor"
    col_regimen:       str = "REGIMEN"


# ---------------------------------------------------------------------------
# Registro de clientes
# ---------------------------------------------------------------------------

CLIENTES: dict[str, ClientConfig] = {

    "aldi": ClientConfig(
        expediente         = "2911",
        exportador_nombre  = "ALDI DOS HERMANAS SUPERMERCADOS, S.L",
        exportador_nif     = "B91405142",
        exportador_dir     = "POLIGONO INDUSTRIAL DE LA ISLA, C/ TORRE DE LOS HERBEROS, S/N",
        exportador_cp      = "41703",
        exportador_ciudad  = "DOS HERMANAS",
        exportador_pais    = "ES",
        destinatario_nombre= "ALDI DOS HERMANAS SUPERMERCADOS, S.L.",
        destinatario_nif   = "B91405142",
        destinatario_dir   = "AVENIDA GENERAL POLAVIEJA, 5$",
        destinatario_cp    = "52006",
        destinatario_ciudad= "MELILLA",
        destinatario_pais  = "ES",
        pais_expedicion    = "ES",
        pais_destino       = "MELILLA",
        incoterm           = "EXW",
        lugar_incoterm     = "SEVILLA",
        puerto_destino     = "RUSADIR",
        aduana_salida      = "ES002911",
        aduana_despacho    = "ES002911",
        provincia_origen   = "29",
        cod_doc_01         = "1003",
        cod_doc_02         = "1833",
        campo_fin          = "11",
    ),

    # Añadir nuevos clientes aquí con su propio ClientConfig
    # "mercadona": ClientConfig(
    #     expediente = "XXXX",
    #     ...
    # ),
}


# ---------------------------------------------------------------------------
# Modelo de datos para cada fila de entrada
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
    if not raw:
        return Decimal("0")
    clean = raw.strip().replace(",", ".")
    try:
        return Decimal(clean)
    except InvalidOperation:
        return Decimal("0")


def _fmt_precio_partida(value: Decimal) -> str:
    """Precio de partida con coma decimal (formato Visual)."""
    return str(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)).replace(".", ",")


def _fmt_importe_cabecera(value: Decimal) -> str:
    """Importe total cabecera con punto decimal."""
    return str(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def _fmt_peso_cabecera(value: Decimal) -> str:
    """Peso cabecera: entero sin decimales."""
    return str(int(value.quantize(Decimal("1"), rounding=ROUND_HALF_UP)))


def parse_desglose(csv_path: Path, cfg: ClientConfig) -> list[DesgloseRow]:
    """Lee el CSV de entrada y devuelve las filas parseadas según el mapeo del cliente."""
    raw = csv_path.read_bytes()
    for enc in ("latin-1", "utf-8-sig", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(f"No se puede decodificar {csv_path.name}")

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    reader = csv.DictReader(text.splitlines(), delimiter=";")

    # Soporte adicional para columnas en inglés (spec SOUTO)
    ALT_MAP = {
        "TARIC_NUMBER":   cfg.col_taric,
        "COUNTRY_ORIGIN": cfg.col_country_origin,
        "CURRENCY":       cfg.col_currency,
        "PRICE":          cfg.col_price,
        "GROSS_WEIGHT":   cfg.col_gross_weight,
        "NET_WEIGHT":     cfg.col_net_weight,
        "BOXES":          cfg.col_boxes,
        "TYPE":           cfg.col_type_box,
        "TRAID":          cfg.col_traid,
    }

    rows: list[DesgloseRow] = []
    for i, raw_row in enumerate(reader, start=2):
        # Normalizar nombre de columnas
        normalized: dict[str, str] = {}
        for col, val in raw_row.items():
            if col is None:
                continue
            col_clean = col.strip()
            # Si la columna está en el mapeo alternativo, traducir
            mapped_col = ALT_MAP.get(col_clean, col_clean)
            normalized[mapped_col] = (val or "").strip()

        def get(col_name: str) -> str:
            return normalized.get(col_name, "")

        price = _parse_decimal(get(cfg.col_price))
        gross = _parse_decimal(get(cfg.col_gross_weight))
        net   = _parse_decimal(get(cfg.col_net_weight))

        try:
            boxes_str = get(cfg.col_boxes).replace(",", ".")
            boxes = int(Decimal(boxes_str).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        except (InvalidOperation, ValueError):
            boxes = 0

        rows.append(DesgloseRow(
            taric          = get(cfg.col_taric),
            description    = get(cfg.col_description),
            country_origin = get(cfg.col_country_origin),
            currency       = get(cfg.col_currency),
            price          = price,
            gross_weight   = gross,
            net_weight     = net,
            boxes          = boxes,
            type_box       = get(cfg.col_type_box),
            traid          = get(cfg.col_traid),
            regimen        = get(cfg.col_regimen),
        ))

    return rows


# ---------------------------------------------------------------------------
# Reglas de transformación
# ---------------------------------------------------------------------------

def apply_weight_rules(row: DesgloseRow) -> tuple[Decimal, Decimal]:
    gross, net = row.gross_weight, row.net_weight
    if net   == Decimal("0"): net   = gross
    if gross == Decimal("0"): gross = net
    return gross, net


def filter_valid(rows: list[DesgloseRow]) -> list[DesgloseRow]:
    return [r for r in rows if r.price > Decimal("0")]


# ---------------------------------------------------------------------------
# Generación del CSV de salida
# ---------------------------------------------------------------------------

def _cabecera_row(cfg: ClientConfig, valid_rows: list[DesgloseRow]) -> list[str]:
    total_price = Decimal("0")
    total_gross = Decimal("0")
    total_net   = Decimal("0")
    contenedor  = valid_rows[0].traid if valid_rows else ""

    for row in valid_rows:
        total_price += row.price
        g, n = apply_weight_rules(row)
        total_gross += g
        total_net   += n

    return [
        cfg.tipo_registro_cab,
        cfg.modo_declaracion,
        cfg.regimen_cabecera,
        cfg.estatuto,
        "",
        cfg.expediente,
        cfg.exportador_nombre,
        cfg.exportador_nif,
        cfg.exportador_dir,
        cfg.exportador_cp,
        cfg.exportador_ciudad,
        cfg.exportador_pais,
        cfg.destinatario_nombre,
        cfg.destinatario_nif,
        cfg.destinatario_dir,
        cfg.destinatario_cp,
        cfg.destinatario_ciudad,
        cfg.destinatario_pais,
        cfg.pais_expedicion,
        cfg.pais_destino,
        "",
        cfg.hay_contenedor,
        cfg.incoterm,
        cfg.lugar_incoterm,
        cfg.puerto_destino,
        cfg.campo_cy,
        cfg.divisa_nombre,
        cfg.divisa_codigo,
        cfg.modo_transporte,
        cfg.aduana_salida,
        cfg.aduana_despacho,
        cfg.campo_503300,
        "",
        _fmt_peso_cabecera(total_gross),
        _fmt_peso_cabecera(total_net),
        contenedor,
        cfg.exportador_nif,
        _fmt_importe_cabecera(total_price),
        *[""] * 26,
        cfg.pais_fin,
        "",
        "",
        cfg.campo_fin,
    ]


def _partida_row(cfg: ClientConfig, row: DesgloseRow, fecha_doc: str) -> list[str]:
    return [
        "0",
        "0",
        str(row.boxes),
        f" {row.description}",
        "",
        "",
        row.taric,
        _fmt_precio_partida(row.price),
        "",
        row.country_origin,
        cfg.provincia_origen,
        "",
        "",
        row.regimen,
        "",
        "",
        "",
        cfg.cod_doc_01,
        "",
        fecha_doc,
        cfg.cod_doc_02,
        "",
        fecha_doc,
        *[""] * 41,
    ]


def generate_convert(
    desglose_path: Path,
    out_dir: Path,
    cfg: ClientConfig,
    fecha_doc: Optional[str] = None,
) -> Path:
    if fecha_doc is None:
        today = Date.today()
        fecha_doc = f"{today.day}/{today.month}/{today.year}"

    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"Convert_-_{desglose_path.stem}.csv"

    all_rows   = parse_desglose(desglose_path, cfg)
    valid_rows = filter_valid(all_rows)

    skipped = len(all_rows) - len(valid_rows)
    if skipped:
        print(f"[INFO] {skipped} filas omitidas (PRICE = 0)", file=sys.stderr)
    if not valid_rows:
        print("[WARN] No hay filas válidas.", file=sys.stderr)

    cabecera = _cabecera_row(cfg, valid_rows)
    partidas = [_partida_row(cfg, r, fecha_doc) for r in valid_rows]

    with out_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh, delimiter=";", quoting=csv.QUOTE_MINIMAL)
        writer.writerow(cabecera)
        for p in partidas:
            writer.writerow(p)

    total_price = sum((r.price for r in valid_rows), Decimal("0"))
    total_gross = sum((apply_weight_rules(r)[0] for r in valid_rows), Decimal("0"))
    total_net   = sum((apply_weight_rules(r)[1] for r in valid_rows), Decimal("0"))

    print(f"[INFO] Cliente:             {cfg.expediente or '—'}")
    print(f"[INFO] Partidas procesadas: {len(valid_rows)}")
    print(f"[INFO] Total importe:       {total_price.quantize(Decimal('0.01'))}")
    print(f"[INFO] Total peso bruto:    {int(total_gross.quantize(Decimal('1'), rounding=ROUND_HALF_UP))}")
    print(f"[INFO] Total peso neto:     {int(total_net.quantize(Decimal('1'), rounding=ROUND_HALF_UP))}")
    print(f"[INFO] Fecha documentos:    {fecha_doc}")

    return out_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Convierte el fichero de partidas de un cliente al formato CSV de Visualtrans."
    )
    ap.add_argument("input",    type=Path, help="Fichero de entrada (CSV separado por ';')")
    ap.add_argument("-o", "--out", type=Path, default=Path("out"), help="Carpeta de salida")
    ap.add_argument(
        "--cliente",
        choices=list(CLIENTES.keys()),
        default="aldi",
        help=f"Cliente a procesar (por defecto: aldi). Disponibles: {', '.join(CLIENTES.keys())}",
    )
    ap.add_argument(
        "--fecha",
        metavar="DD/MM/YYYY",
        default=None,
        help="Fecha de los documentos de partida (por defecto: hoy)",
    )
    ap.add_argument("--version", action="version", version=f"%(prog)s {SCRIPT_VERSION}")
    args = ap.parse_args()

    if not args.input.exists():
        print(f"ERROR: fichero no encontrado: {args.input}", file=sys.stderr)
        return 1

    cfg = CLIENTES[args.cliente]

    try:
        out_path = generate_convert(
            desglose_path=args.input,
            out_dir=args.out,
            cfg=cfg,
            fecha_doc=args.fecha,
        )
        print(f"OK -> {out_path}")
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise


if __name__ == "__main__":
    raise SystemExit(main())