"""
split_nominas_extractor.py
--------------------------
Extractor para el stack IASuite (unified_processor).

Calling convention (igual que el resto de extractores del stack):
    python3 split_nominas_extractor.py file1.pdf [file2.pdf ...] -o /output/dir

Dependencias (añadir a services/unified_processor/requirements.txt):
    pypdf
    pdfplumber
"""

import argparse
import re
import sys
from pathlib import Path

try:
    import pdfplumber
    from pypdf import PdfReader, PdfWriter
except ImportError as e:
    print(f"[split_nominas] ERROR: dependencia no instalada -> {e}", flush=True)
    print("[split_nominas]   pip install pypdf pdfplumber", flush=True)
    sys.exit(1)


MESES = {
    "enero": "01", "febrero": "02", "marzo": "03",
    "abril": "04", "mayo": "05", "junio": "06",
    "julio": "07", "agosto": "08", "septiembre": "09",
    "octubre": "10", "noviembre": "11", "diciembre": "12",
}

# Mapeo empresa → iniciales  (orden: más específico primero)
# Cada entrada: (patron_regex, iniciales, sufijo_forma_juridica_regex_opcional)
EMPRESAS = [
    (r"TOTAL\s+LOGISTIC\s+SERVICES", "TLS"),
    (r"TOTAL\s+LOGISTIC",            "TLS"),
    (r"TOTAL\s+ENGINEERING",         "TE"),
    (r"ASOC(?:IACI[ÓO]N)?\.\s+MELILLA\s+INTEGRA", "ASOC"),
]

# Sufijos jurídicos a eliminar al extraer el nombre del trabajador
_SUFIJOS_JURIDICOS = r"(?:\s+S\.?\s*L\.?|\s+S\.?\s*A\.?|\s+S\.?\s*L\.?\s*U\.?)?"


def colapsar_espaciado(texto: str) -> str:
    """Comprime 'E n e r o' -> 'Enero' (artefacto habitual de pdfplumber)."""
    return re.sub(
        r"(?<=[A-Za-záéíóúüñÁÉÍÓÚÜÑ\d]) (?=[A-Za-záéíóúüñÁÉÍÓÚÜÑ\d])", "", texto
    )


def extraer_datos_pagina(texto: str) -> tuple:
    """
    Devuelve (mes, anyo, nombre, empresa_iniciales).
    Ejemplo: ("02", "26", "BAEZA_TORRES_GONZALO", "TLS")

    Soporta dos formatos de nómina:
      - Formato TE:  primera línea = "EMPRESA S.L. CC ACAD APELLIDO NOMBRE"
      - Formato TLS: primera línea = "EMPRESA S.L. APELLIDO NOMBRE"
    """
    mes     = "00"
    anyo    = "00"
    nombre  = "DESCONOCIDO"
    empresa = "XX"

    primera_linea = texto.split("\n")[0] if texto else ""
    texto_col     = colapsar_espaciado(texto)

    # ── Empresa (primera línea) ───────────────────────────────────────────
    for patron, iniciales in EMPRESAS:
        if re.search(patron, primera_linea, re.IGNORECASE):
            empresa = iniciales
            break

    # ── Nombre: dos formatos ──────────────────────────────────────────────
    # Formato TE:  "... CC ACAD APELLIDO NOMBRE"
    m_nombre = re.search(
        r"CC\s+ACAD\s+([A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ ]+?)(?:\s*$)",
        primera_linea, re.IGNORECASE,
    )
    if m_nombre:
        nombre = "_".join(m_nombre.group(1).strip().split())
    else:
        # Formato TLS/AMI: "EMPRESA [S.L.] APELLIDO NOMBRE"
        # Eliminamos el nombre de la empresa + sufijo jurídico opcional
        nombre_raw = primera_linea
        for patron, _ in EMPRESAS:
            nombre_raw = re.sub(
                patron + _SUFIJOS_JURIDICOS, "", nombre_raw, flags=re.IGNORECASE
            ).strip()
        # Eliminar puntos y comas residuales al inicio
        nombre_raw = re.sub(r"^[\s.,]+", "", nombre_raw)
        m_resto = re.match(r"([A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ ]+)", nombre_raw)
        if m_resto:
            nombre = "_".join(m_resto.group(1).strip().split())

    # ── Mes (texto colapsado) ─────────────────────────────────────────────
    m_mes = re.search("(" + "|".join(MESES.keys()) + ")", texto_col, re.IGNORECASE)
    if m_mes:
        mes = MESES[m_mes.group(1).lower()]

    # ── Año: buscamos el año JUNTO al nombre del mes para evitar falsos
    #    positivos de números de cotización como "29109338672 01 28/01/02"
    #    que al colapsar generan secuencias como "2012".
    #    Patrón: "de{MES}de{YYYY}" o "{MES}de{YYYY}" o "{MES}{YYYY}"
    m_anyo = re.search(
        r"(?:" + "|".join(MESES.keys()) + r")(?:de)?(20\d{2})",
        texto_col, re.IGNORECASE,
    )
    if m_anyo:
        anyo = m_anyo.group(1)[2:]   # "2026" → "26"

    return mes, anyo, nombre, empresa


def sanitizar(nombre: str) -> str:
    return re.sub(r"[^\w\-]", "_", nombre, flags=re.UNICODE)


def split_pdf(input_pdf: Path, output_dir: Path) -> list:
    output_dir.mkdir(parents=True, exist_ok=True)

    reader = PdfReader(str(input_pdf))
    n = len(reader.pages)
    print(f"[split_nominas] {input_pdf.name}: {n} páginas", flush=True)

    creados = []
    with pdfplumber.open(str(input_pdf)) as plumb:
        for i, page in enumerate(plumb.pages):
            texto = page.extract_text() or ""
            mes, anyo, nombre, empresa = extraer_datos_pagina(texto)

            # Formato: MMYY-APELLIDO_APELLIDO_NOMBRE-EMPRESA
            # Ejemplo: 0226-BAEZA_TORRES_GONZALO-TLS
            base  = sanitizar(f"{mes}{anyo}-{nombre}-{empresa}")
            fname = base + ".pdf"
            dest  = output_dir / fname

            c = 1
            while dest.exists():
                fname = f"{base}_{c}.pdf"
                dest  = output_dir / fname
                c += 1

            w = PdfWriter()
            w.add_page(reader.pages[i])
            with open(dest, "wb") as f:
                w.write(f)

            creados.append(dest)
            print(f"[split_nominas]   pág {i+1:>3} -> {fname}", flush=True)

    return creados


def main():
    # Convention: positional file args + "-o /output/dir"  (same as all other extractors)
    parser = argparse.ArgumentParser()
    parser.add_argument("files", nargs="+", help="PDF(s) de nóminas")
    parser.add_argument("-o", "--output-dir", required=True, help="Directorio de salida")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    total = []

    for fpath in args.files:
        p = Path(fpath)
        if not p.exists():
            print(f"[split_nominas] AVISO: no existe '{p}', se omite", flush=True)
            continue
        total.extend(split_pdf(p, output_dir))

    print(f"[split_nominas] OK: {len(total)} PDFs generados en {output_dir}", flush=True)


if __name__ == "__main__":
    main()