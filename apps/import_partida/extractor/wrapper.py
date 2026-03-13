#!/usr/bin/env python3
"""
Wrapper to adapt import_partida extractor to unified processor interface.
Converts: file1.pdf file2.pdf -o /output
To: extract_import_partida_fields.py file1.pdf /output (called for each file)
"""
import sys
import subprocess
from pathlib import Path

def main():
    # Parse args: wrapper.py file1.pdf file2.pdf ... -o /output/dir
    args = sys.argv[1:]
    
    if '-o' not in args:
        print("ERROR: Missing -o flag", file=sys.stderr)
        return 1
    
    o_index = args.index('-o')
    pdfs = args[:o_index]
    output_dir = args[o_index + 1] if o_index + 1 < len(args) else None
    
    if not pdfs or not output_dir:
        print("Usage: wrapper.py <pdf files...> -o <output_dir>", file=sys.stderr)
        return 1
    
    # Get path to original extractor
    extractor = Path(__file__).parent / "extract_import_partida_fields.py"
    
    # Process each PDF
    for pdf in pdfs:
        print(f"Processing: {pdf}")
        result = subprocess.run(
            ["python3", str(extractor), pdf, output_dir],
            capture_output=True,
            text=True
        )
        
        if result.stdout:
            print(result.stdout, end='')
        if result.stderr:
            print(result.stderr, end='', file=sys.stderr)
        
        if result.returncode != 0:
            return result.returncode
    
    print(f"✓ Processed {len(pdfs)} PDF(s)")
    return 0

if __name__ == "__main__":
    sys.exit(main())
