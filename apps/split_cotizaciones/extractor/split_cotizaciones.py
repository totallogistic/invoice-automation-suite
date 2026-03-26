"""
split_cotizaciones.py
---------------------
Extractor para el stack IASuite (unified_processor).

Divide PDFs de cotizaciones de la Seguridad Social en archivos individuales
por empresa/tipo, agrupando páginas consecutivas del mismo documento.

Tipos de documento detectados:
  - RLC: Recibo de Liquidación de Cotizaciones
  - RNT: Relación Nominal de Trabajadores

Naming: MMYY-RLC/RNT-EMPRESA[-CIUDAD].pdf
  Ejemplos: 0126-RLC-TLS-AGP.pdf  /  1225-RNT-JAML.pdf

Calling convention (igual que el resto de extractores del stack):
    python3 split_cotizaciones.py file1.pdf [file2.pdf ...] -o /output/dir

Dependencias (ya presentes en services/unified_processor/requirements.txt):
    pypdf
    pdfplumber
"""

import argparse
import re
import sys
from pathlib import Path

SCRIPT_VERSION = "2026-03-26.v1"

SCRIPT_CHANGELOG = """
## 2026-03-26.v1

### Logica general
Divide PDFs multi-página de cotizaciones de la Seguridad Social en archivos
individuales por empresa y tipo de documento.

### Tipos de documento
- RLC: Recibo de Liquidación de Cotizaciones
- RNT: Relación Nominal de Trabajadores

### Extracción por página
- Detecta tipo de documento por texto en la página
- Extrae Código de Cuenta de Cotización (patrón 0111 XXXXXXXXXXX)
- Extrae período de liquidación (MM/YYYY → MMYY)

### Agrupación
Las páginas consecutivas con el mismo (tipo, cuenta, período) se fusionan
en un único PDF de salida. Esto cubre los casos de RNT multi-página
(ej. TLS CAD con Página 1 de 3, 2 de 3, 3 de 3).

### Nomenclatura de salida
MMYY-RLC/RNT-EMPRESA[-CIUDAD].pdf
Ejemplos:
  0126-RLC-TLS-AGP.pdf
  0126-RNT-TE-ACAD.pdf
  1225-RNT-JAML.pdf
  1225-RLC-ASOC.pdf

### Mapeo de códigos de cuenta
- 0111 52101622831 → ASOC
- 0111 52100999910 → JAML
- 0111 52101627982 → TE ACAD
- 0111 52100846326 → TE MLN
- 0111 11103655876 → TLS CAD
- 0111 41118113171 → TLS SVQ
- 0111 29109338672 → TLS AGP
"""

try:
    import pdfplumber
    from pypdf import PdfReader, PdfWriter
except ImportError as e:
    print(f"[split_cotizaciones] ERROR: dependencia no instalada -> {e}", flush=True)
    print("[split_cotizaciones]   pip install pypdf pdfplumber", flush=True)
    sys.exit(1)


# ── Mapeo código de cuenta → (empresa, ciudad) ────────────────────────────────
# ciudad vacía = empresa única, no se añade sufijo
ACCOUNT_MAP = {
    "011152101622831": ("ASOC", ""),
    "011152100999910": ("JAML", ""),
    "011152101627982": ("TE",   "ACAD"),
    "011152100846326": ("TE",   "MLN"),
    "011111103655876": ("TLS",  "CAD"),
    "011141118113171": ("TLS",  "SVQ"),
    "011129109338672": ("TLS",  "AGP"),
}


def colapsar_espaciado(texto: str) -> str:
    """Comprime 'R L C' → 'RLC' (artefacto ocasional de pdfplumber)."""
    return re.sub(
        r"(?<=[A-Za-z0-9áéíóúüñÁÉÍÓÚÜÑ]) (?=[A-Za-z0-9áéíóúüñÁÉÍÓÚÜÑ])", "", texto
    )


def normalizar_cuenta(raw: str) -> str:
    """Elimina espacios del código de cuenta para comparación."""
    return re.sub(r"\s+", "", raw)


def extraer_datos_pagina(texto: str) -> tuple:
    """
    Devuelve (doc_type, cuenta_normalizada, periodo_mmyy).

    doc_type: 'RLC' | 'RNT' | 'UNKNOWN'
    cuenta:   e.g. '011152100999910'
    periodo:  e.g. '0126'
    """
    texto_col = colapsar_espaciado(texto)

    # ── Tipo de documento ─────────────────────────────────────────────────────
    if re.search(r"Recibo de Liquidaci[oó]n de Cotizaciones", texto, re.IGNORECASE):
        doc_type = "RLC"
    elif re.search(r"RELACI[OÓ]N NOMINAL DE TRABAJADORES", texto, re.IGNORECASE):
        doc_type = "RNT"
    else:
        doc_type = "UNKNOWN"

    # ── Código de cuenta de cotización ────────────────────────────────────────
    # Formato: "0111 XXXXXXXXXXX"  (4 dígitos + espacio + 11 dígitos)
    m_cuenta = re.search(r"(0111\s+\d{11})", texto)
    cuenta = normalizar_cuenta(m_cuenta.group(1)) if m_cuenta else "DESCONOCIDA"

    # ── Período de liquidación ────────────────────────────────────────────────
    # Formatos: "01/2026 - 01/2026"  o  "01/2026-01/2026"  o colapsado
    m_periodo = re.search(
        r"(\d{2})[/\-](\d{4})\s*[-–]\s*\d{2}[/\-]\d{4}",
        texto_col,
    )
    if m_periodo:
        periodo = m_periodo.group(1) + m_periodo.group(2)[2:]   # "01/2026" → "0126"
    else:
        periodo = "0000"

    return doc_type, cuenta, periodo


def build_nombre(doc_type: str, cuenta: str, periodo: str) -> str:
    """Construye el nombre base del fichero de salida."""
    info = ACCOUNT_MAP.get(cuenta)
    if info:
        empresa, ciudad = info
        partes = [periodo, doc_type, empresa]
        if ciudad:
            partes.append(ciudad)
        return "-".join(partes)
    else:
        # Cuenta desconocida: usar el código directamente
        return f"{periodo}-{doc_type}-{cuenta}"


def sanitizar(nombre: str) -> str:
    return re.sub(r"[^\w\-]", "_", nombre, flags=re.UNICODE)


def split_pdf(input_pdf: Path, output_dir: Path) -> list:
    output_dir.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(input_pdf))
    n = len(reader.pages)
    print(f"[split_cotizaciones] {input_pdf.name}: {n} páginas", flush=True)

    # ── Paso 1: extraer metadatos de cada página ──────────────────────────────
    page_data = []
    with pdfplumber.open(str(input_pdf)) as plumb:
        for i, page in enumerate(plumb.pages):
            texto = page.extract_text() or ""
            doc_type, cuenta, periodo = extraer_datos_pagina(texto)
            page_data.append((i, doc_type, cuenta, periodo))
            print(
                f"[split_cotizaciones]   pág {i+1:>3}: {doc_type} | {cuenta} | {periodo}",
                flush=True,
            )

    if not page_data:
        return []

    # ── Paso 2: agrupar páginas consecutivas con la misma clave ───────────────
    groups = []
    current_key = (page_data[0][1], page_data[0][2], page_data[0][3])
    current_pages = [page_data[0][0]]

    for i, doc_type, cuenta, periodo in page_data[1:]:
        key = (doc_type, cuenta, periodo)
        if key == current_key:
            current_pages.append(i)
        else:
            groups.append((current_key, current_pages))
            current_key = key
            current_pages = [i]
    groups.append((current_key, current_pages))

    # ── Paso 3: escribir un PDF por grupo ─────────────────────────────────────
    creados = []
    for (doc_type, cuenta, periodo), pages in groups:
        base  = sanitizar(build_nombre(doc_type, cuenta, periodo))
        fname = base + ".pdf"
        dest  = output_dir / fname

        # Evitar colisiones (no debería ocurrir con datos bien formados)
        c = 1
        while dest.exists():
            fname = f"{base}_{c}.pdf"
            dest  = output_dir / fname
            c += 1

        w = PdfWriter()
        for page_idx in pages:
            w.add_page(reader.pages[page_idx])

        with open(dest, "wb") as f:
            w.write(f)

        creados.append(dest)
        npags = len(pages)
        print(
            f"[split_cotizaciones]   → {fname}  ({npags} pág{'s' if npags > 1 else ''})",
            flush=True,
        )

    return creados


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("files", nargs="+", help="PDF(s) de cotizaciones")
    parser.add_argument("-o", "--output-dir", required=True, help="Directorio de salida")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    total = []

    for fpath in args.files:
        p = Path(fpath)
        if not p.exists():
            print(f"[split_cotizaciones] AVISO: no existe '{p}', se omite", flush=True)
            continue
        total.extend(split_pdf(p, output_dir))

    print(
        f"[split_cotizaciones] OK: {len(total)} PDFs generados en {output_dir}",
        flush=True,
    )


if __name__ == "__main__":
    main()