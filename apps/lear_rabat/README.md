# Lear Rabat Invoice Extractor

Extracts customs clearance data from Lear Automotive Morocco (Rabat) invoices.

## Overview

This tool processes Lear Rabat PDF invoices and generates ODS files populated with DUA (customs) data ready for import processing.

## Input

- **Format**: PDF invoices from Lear Automotive Morocco
- **Source**: Lear invoices with project-based line items (e.g., O_BCP21, O_BR213, O_C1A, etc.)

## Output

- **Format**: ODS spreadsheet (LibreOffice Calc / Excel compatible)
- **Filename**: `COMPLETADO_{invoice_number}.ods`
- **Contents**: 
  - Populated template with extracted data mapped to customs codes
  - Line items grouped by project code
  - Totals, weights, pallet counts
  - OPR Material costs per project

## Data Extracted

- Invoice number
- Line items with:
  - Part numbers
  - Customs codes (CODIGO)
  - HS codes (PARTIDA)
  - Quantities (UN)
  - Subtotal prices (VALOR DUA)
  - OPR Material costs
  - Weights (PESO NET, PESO BRUT)
  - Pallet counts (PALETS)

## Template

The tool requires a template file `COMPLETADO_TEMPLATE.ods` which contains:
- Pre-configured product codes and customs classifications
- Formulas for calculations
- Standard format for DUA processing

Place the template in the extractor directory or specify with `-t` flag.

## Usage

### Standalone
```bash
python3 extract_lear_rabat.py invoice1.pdf invoice2.pdf -o /output/dir
python3 extract_lear_rabat.py invoice.pdf -o /output -t custom_template.ods
```

### Via Unified Processor
The tool is integrated with the unified processor and will be called automatically when PDFs are placed in the inbox directory.

## Dependencies

- pdfplumber
- odfpy

## Version

2026-02-16.v1
