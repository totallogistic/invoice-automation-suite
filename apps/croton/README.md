# Croton ODS Processor

Processes Croton packing-list ODS files against an XLSX product mapping and
returns the original ODS enriched with a new **Resumen_Partidas** tab.

## Overview

This tool receives:

1. An **ODS packing-list** (`Hoja5` format — one row per roll/bulto)
2. An **XLSX product-mapping** file (columns: `CODIGO`, `DESCRIPCION_CONTAINS`,
   `MERCANCIA`, `PARTIDA`)

It produces the **same ODS** with an additional tab `Resumen_Partidas` that
aggregates rolls by customs tariff (`PARTIDA`), which is then emailed to the
configured recipients.

## Input

| File | Format | Role |
|------|--------|------|
| Packing list | `.ods` | Source data — Hoja5 sheet |
| Product mapping | `.xlsx` | Lookup table (CODIGO → MERCANCIA / PARTIDA) |

Both files must be uploaded together in one batch.

## Output

- **Format**: ODS spreadsheet (the input ODS + a new `Resumen_Partidas` tab)
- **Filename**: `CROTON_{input_stem}.ods`

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
python3 extract_croton.py packing_list.ods product_mapping.xlsx -o /output/dir
```

### Via Unified Processor

Upload both files at once through the web UI at `/tools/croton/`.
The processor automatically identifies the ODS and XLSX by extension.

## Dependencies

```bash
pip install odfpy openpyxl
```

## Fallback CSV mapping

If no XLSX is uploaded, the bundled `product_mapping.csv` is used as a
fallback (useful for testing). For production, always provide an up-to-date
XLSX mapping file.

## Version

2026-03-09.v1
