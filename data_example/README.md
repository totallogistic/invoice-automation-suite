# Import Partida Extractor - Test Data

This directory contains test data for the import_partida extractor.

## Files

- **NPOS56802.pdf**: Sample Bill of Lading PDF with NPOS format B/L number
- **26XH009BL-1.pdf**: Sample Bill of Lading PDF with alphanumeric B/L number
- **expected_output.csv**: Expected CSV output format (based on NPOS56802.pdf)

## Usage

Run the extractor on a sample PDF:

```bash
python3 apps/import_partida/extractor/extract_import_partida_fields.py \
    "data_example/NPOS56802.pdf" \
    "/tmp/import_partida_out"
```

This will generate:
- `/tmp/import_partida_out/import_partida.csv` - CSV with extracted data
- `/tmp/import_partida_out/import_partida.xlsx` - Excel with extracted data
- `/tmp/import_partida_out/extracted.json` - Raw JSON with extracted fields

## Output Format

The CSV output has the following format:

**Headers** (lowercase with underscores, unquoted):
```
bl_no,shipper,consignee,vessel,port_of_loading,port_of_discharge,contain,weight,measurement,mrsu
```

**Values** (selective quoting):
- First field (`bl_no`): unquoted
- Middle fields (shipper through measurement): quoted
- Last field (`mrsu`): unquoted

**Units included in values**:
- `contain`: "4245 PACKAGES" (not just "4245")
- `weight`: "3310,000 KGS" (not just "3310,000")
- `measurement`: "58,0000 CBM" (not just "58,0000")

## Example Output

```csv
bl_no,shipper,consignee,vessel,port_of_loading,port_of_discharge,contain,weight,measurement,mrsu
NPOS56802,"NINGBO YINZHOU SUNEVER FASHION","ALVARO MORENO RETAIL S.L.U.","BERLIN MAERSK 603W","NINGBO, CHINA","Valencia,Spain","4245 PACKAGES","3310,000 KGS","58,0000 CBM",MRSU5144291
```

## Verify Output

Compare the generated CSV with the expected output:

```bash
diff data_example/expected_output.csv /tmp/import_partida_out/import_partida.csv
```

If the command produces no output, the files match exactly.

## Sample Data in NPOS56802.pdf

The NPOS56802 PDF contains the following Bill of Lading information:

- **B/L No**: NPOS56802
- **Shipper**: NINGBO YINZHOU SUNEVER FASHION
- **Consignee**: ALVARO MORENO RETAIL S.L.U.
- **Vessel**: BERLIN MAERSK 603W
- **Port of Loading**: NINGBO, CHINA
- **Port of Discharge**: Valencia,Spain
- **Packages**: 4245 PACKAGES
- **Weight**: 3310,000 KGS
- **Measurement**: 58,0000 CBM
- **Container**: MRSU5144291
