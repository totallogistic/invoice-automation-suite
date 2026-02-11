# Import Partida Extractor - Test Data

This directory contains test data for the import_partida extractor.

## Files

- **26XH009BL-1.pdf**: Sample Bill of Lading PDF for testing
- **expected_output.csv**: Expected CSV output when extracting from the sample PDF

## Usage

Run the extractor on the sample PDF:

```bash
python3 apps/import_partida/extractor/extract_import_partida_fields.py \
    "data_example/26XH009BL-1.pdf" \
    "/tmp/import_partida_out"
```

This will generate:
- `/tmp/import_partida_out/import_partida.csv` - CSV with extracted data
- `/tmp/import_partida_out/import_partida.xlsx` - Excel with extracted data
- `/tmp/import_partida_out/extracted.json` - Raw JSON with extracted fields

## Verify Output

Compare the generated CSV with the expected output:

```bash
diff data_example/expected_output.csv /tmp/import_partida_out/import_partida.csv
```

If the command produces no output, the files match exactly.

## Sample Data

The sample PDF contains the following Bill of Lading information:

- **B/L No**: 26XH009BL
- **Shipper**: ABC SHIPPING CO LTD, 123 MAIN STREET, CITY COUNTRY
- **Consignee**: XYZ IMPORTS INC
- **Vessel**: MSC MEDITERRANEAN
- **Port of Loading**: SHANGHAI CHINA
- **Port of Discharge**: LONG BEACH USA
- **Packages**: 100
- **Weight**: 15000 KGS
- **Measurement**: 25.5 CBM
- **Container**: MRSU1234567
