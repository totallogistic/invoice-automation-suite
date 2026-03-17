# Croton ODS Processor

Processes Croton packing-list files against a factura/mapeo de productos and
returns the original file enriched with a new **Resumen_Partidas** tab.

## Overview

This tool receives:

1. A **packing-list** (`.ods` or `.xlsx`) (`Hoja5` format — one row per roll/bulto)
2. A **factura / mapeo de productos** file (`.ods`, `.xlsx` or `.xls`) used to
   allocate customs values and classify product references

It produces the **same file** with an additional tab `Resumen_Partidas` that
aggregates rolls by customs tariff (`PARTIDA`), which is then emailed to the
configured recipients.

## Input

| File | Format | Role |
|------|--------|------|
| Packing list | `.ods`, `.xlsx` | Source data — Hoja5 sheet |
| Mapeo de productos | `.ods`, `.xlsx`, `.xls` | Lookup table / factura (REFERENCIA → VALOR / PARTIDA) |

Both files must be uploaded together in one batch.

## Output

- **Format**: ODS spreadsheet (the input ODS + a new `Resumen_Partidas` tab)
- **Filename**: `CROTON_{input_stem}.ods` (or `.xlsx` when packing is XLSX)

## Tab columns (`Resumen_Partidas`)

| Column | Description |
|--------|-------------|
| MERCANCIA | Goods description |
| PARTIDA | Customs tariff code |
| Suma - BX | Total rolls/boxes |
| Suma - VALOR | Declared customs value |
| Suma - BRUTO | Total gross weight (kg) |
| Suma - NETO | Total net weight (kg) |
| Suma - M2 | Total area (m²) |

## Usage

### Standalone

```bash
python3 extract_croton.py packing_list.ods --factura mapeo_productos.ods -o output.ods
python3 extract_croton.py packing_list.xlsx --factura mapeo_productos.xlsx -o output.xlsx
```

### Via Unified Processor

Upload both files at once through the web UI at `/tools/croton/`.
The processor automatically identifies packing vs factura/mapeo by extension
and filename keywords (`packing`/`parking` for the packing list,
`factura`/`invoice`/`mapeo`/`mapping` for the second file).

## Dependencies

```bash
pip install odfpy openpyxl
```

## Fallback CSV mapping

If no second file is uploaded, the bundled `product_mapping.csv` is used as a
fallback for product classification (useful for testing). For production,
always provide an up-to-date mapping file.

## Version

2026-03-17.v2
