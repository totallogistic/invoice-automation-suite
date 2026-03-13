# Lear Rabat Integration Guide

## Overview

This document describes the integration of the Lear Rabat DUA processor into the unified invoice automation suite.

## What was added

### 1. Extractor Script
**Location**: `apps/lear_rabat/extractor/extract_lear_rabat.py`

- Extracts data from Lear Automotive Morocco PDF invoices
- Generates populated ODS files for DUA (customs clearance)
- Uses pdfplumber for PDF text extraction (no OCR)
- Uses odfpy for ODS file manipulation
- Follows unified processor interface: `script.py <pdfs...> -o <output_dir>`

**Key Features**:
- Extracts invoice metadata (number, totals, weights, pallets)
- Parses line items grouped by project (e.g., O_BCP21, O_BR213, etc.)
- Maps extracted data to customs codes in template
- Calculates V.E values (VALOR DUA + OPR Material Cost)
- Handles multiple invoices in batch mode

### 2. Template File
**Location**: `apps/lear_rabat/extractor/COMPLETADO_TEMPLATE.ods`

Pre-configured ODS template with:
- Product codes and customs classifications
- Formula columns for automatic calculations
- Standard DUA format

### 3. Configuration
**Modified**: `config/tools.yaml`

Added Lear Rabat tool configuration:
```yaml
- name: lear_rabat
  enabled: true
  display_name: "Lear Rabat DUA Processor"
  description: "Extracts DUA data from Lear Automotive Morocco invoices"
  
  extractor:
    type: python
    path: /apps/lear_rabat/extractor/extract_lear_rabat.py
  
  input:
    formats: [pdf]
    inbox: /data/lear_rabat/inbox
  
  output:
    directory: /data/lear_rabat/out
    artifacts:
      - COMPLETADO_*.ods
  
  email:
    subject_template: "[Lear Rabat] Batch {batch_id} - DUA files ready"
```

### 4. Web Interface
**Locations**:
- `services/web/sites/dev/tools/lear_rabat/index.html`
- `services/web/sites/prod/tools/lear_rabat/index.html`

Added web upload interface for both dev and prod environments.

**Updated main pages**:
- `services/web/sites/dev/index.html` - Added Lear Rabat link
- `services/web/sites/prod/index.html` - Added Lear Rabat link

## Directory Structure

```
apps/lear_rabat/
├── extractor/
│   ├── extract_lear_rabat.py      # Main extractor script
│   └── COMPLETADO_TEMPLATE.ods    # ODS template
└── README.md                       # Tool documentation
```

## Deployment Steps

### 1. Deploy Updated Code

```bash
# On your Raspberry Pi
cd ~/invoice-automation-suite
git pull origin feature/unified-architecture

# Or upload the modified files via scp/rsync
```

### 2. Install Dependencies

The extractor requires:
- pdfplumber
- odfpy

These should already be in your environment if you've been using the suite, but verify:

```bash
pip list | grep -E 'pdfplumber|odfpy'
```

If missing:
```bash
pip install pdfplumber odfpy
```

### 3. Create Data Directories

```bash
# Create inbox and output directories for Lear Rabat
mkdir -p /data/lear_rabat/inbox
mkdir -p /data/lear_rabat/out
mkdir -p /data/lear_rabat/archive
```

### 4. Restart Services

```bash
# If using Docker Compose
docker-compose down
docker-compose up -d

# Or restart individual services
docker-compose restart unified_processor
docker-compose restart web
```

### 5. Verify Integration

#### Test the extractor standalone:
```bash
python3 apps/lear_rabat/extractor/extract_lear_rabat.py \
  test_invoice.pdf \
  -o /tmp/test_output
```

#### Test via unified processor:
```bash
# Place a PDF in inbox
cp test_invoice.pdf /data/lear_rabat/inbox/

# Check processor logs
docker-compose logs -f unified_processor

# Check output directory
ls -la /data/lear_rabat/out/
```

#### Test web interface:
1. Open browser: `http://your-raspberry-pi:8081/tools/lear_rabat/`
2. Upload a test PDF
3. Verify ODS file is generated and downloadable

## Usage

### Via Web Interface
1. Navigate to: http://your-ip:8081/tools/lear_rabat/
2. Upload one or more Lear Rabat PDF invoices
3. Click "Procesar"
4. Download generated COMPLETADO_*.ods files

### Via File Watcher (Automated)
1. Drop PDF files into `/data/lear_rabat/inbox/`
2. Processor automatically detects and processes them
3. Output ODS files appear in `/data/lear_rabat/out/`
4. Email is sent with download links
5. Processed PDFs move to `/data/lear_rabat/archive/`

### Via Command Line
```bash
python3 apps/lear_rabat/extractor/extract_lear_rabat.py \
  invoice1.pdf invoice2.pdf \
  -o /output/directory \
  -t /path/to/custom_template.ods  # optional
```

## Output Format

For each invoice PDF, generates:
- **Filename**: `COMPLETADO_{invoice_number}.ods`
- **Contents**: Populated customs clearance data
  - Product codes mapped to template
  - Quantities, weights, prices
  - OPR Material costs per project
  - Calculated V.E totals
  - Invoice-level summary data

## Email Notifications

When batch processing completes, email is sent with:
- Subject: `[Lear Rabat] Batch {batch_id} - DUA files ready`
- Attachments: All generated COMPLETADO_*.ods files
- Body: Summary of processed files

## Monitoring

Check processor logs:
```bash
docker-compose logs -f unified_processor | grep lear_rabat
```

Check for errors:
```bash
grep ERROR /data/lear_rabat/out/*.log
```

## Troubleshooting

### Issue: Template not found
**Solution**: Verify `COMPLETADO_TEMPLATE.ods` exists in `apps/lear_rabat/extractor/`

### Issue: Extraction errors
**Solution**: Check PDF is text-based (not scanned image). Verify PDF structure matches expected Lear Rabat format.

### Issue: Missing data in output
**Solution**: Check that product codes in PDF match codes in template. Review extractor logs for mapping warnings.

### Issue: ODS file won't open
**Solution**: Verify odfpy version is compatible. Try opening with LibreOffice Calc instead of Excel.

## Maintenance

### Updating the Template
1. Modify `apps/lear_rabat/extractor/COMPLETADO_TEMPLATE.ods`
2. Test with sample invoices
3. Deploy updated template
4. No code changes needed if column positions remain same

### Adding New Product Codes
1. Edit template ODS file
2. Add new rows with product codes and classifications
3. Extractor will automatically map to new codes

## Integration with Other Tools

The Lear Rabat tool follows the same architecture as:
- **Lear Cable**: Batch invoice extraction to XLSX
- **Import Partida**: B/L processing to CSV

All three tools share:
- Unified processor interface
- Common email notification system
- Standardized web upload interface
- Consistent file watching and archival

## Version

Integration Version: 2026-02-16.v1
Extractor Version: 2026-02-16.v1
