# Croton ODS Processor

Processes Croton ODS input files and generates populated ODS output files for delivery via email.

## Overview

This tool receives one or more ODS spreadsheets, applies the Croton extraction/transformation logic, and produces a processed ODS file per input that is automatically emailed to the configured recipients.

## Input

- **Format**: ODS spreadsheet (`.ods`)
- **Source**: Croton data files

## Output

- **Format**: ODS spreadsheet (LibreOffice Calc compatible)
- **Filename**: `CROTON_{input_stem}.ods`

## Usage

### Standalone
```bash
python3 extract_croton.py input.ods -o /output/dir
python3 extract_croton.py file1.ods file2.ods -o /output/dir
```

### Via Unified Processor
The tool is integrated with the unified processor and will be called automatically when ODS files are placed in the inbox directory (`/data/croton/inbox/`).

## Dependencies

- odfpy

```bash
pip install odfpy
```

## Implementation Note

The extraction/transformation logic lives in the `process_ods()` function inside `extract_croton.py`. Replace the `TODO` section with the actual processing rules.

## Version

2026-03-02.v1
