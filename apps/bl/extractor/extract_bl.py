"""
Extractor de Conocimientos de Embarque (BL)
============================================
Navieras soportadas : AML · Balearia · DFDS · RFS · Trasmediterranea
Campos extraídos    : num_bl, nombre, puerto_destino, matricula, buque, fecha, hora
Filtro              : solo se procesan BLs con puerto_destino = ALGECIRAS

Uso standalone:
    python extract_bl.py archivo1.pdf archivo2.pdf ...
    python extract_bl.py /carpeta/con/pdfs/

Uso integrado (llamado por unified_processor):
    python extract_bl.py file1.pdf file2.pdf -o /data/bl/out/{batch_id}/
"""

SCRIPT_VERSION = "1.1.0"

import re
import os
import csv
import sys
import glob
import logging
import argparse
from datetime import datetime, timezone
from dataclasses import dataclass, fields, asdict
from pathlib import Path
from typing import Optional

import pdfplumber

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

DESTINO_FILTRO = "ALGECIRAS"


# ── Modelo de datos ──────────────────────────────────────────────────────────
@dataclass
class RegistroBL:
    archivo:         str = ""
    naviera:         str = ""
    num_bl:          str = ""
    nombre:          str = ""
    puerto_destino:  str = ""
    matricula:       str = ""
    buque:           str = ""
    fecha:           str = ""
    hora:            str = ""
    ts_extraccion:   str = ""


# ── Utilidades ───────────────────────────────────────────────────────────────
def _txt(path: str) -> str:
    with pdfplumber.open(path) as pdf:
        return pdf.pages[0].extract_text() or ""


def _re1(pattern: str, text: str, group: int = 1, flags=re.IGNORECASE) -> str:
    m = re.search(pattern, text, flags)
    return " ".join(m.group(group).split()).strip() if m else ""


def _fecha(raw: str) -> str:
    raw = raw.strip()
    meses_fr = {
        "janvier":"01","fevrier":"02","mars":"03","avril":"04",
        "mai":"05","juin":"06","juillet":"07","aout":"08",
        "septembre":"09","octobre":"10","novembre":"11","decembre":"12",
    }
    texto = raw.lower()
    for mes, num in meses_fr.items():
        if mes in texto:
            texto = re.sub(mes, num, texto)
            texto = re.sub(r"[^\d/]", "/", texto)
            texto = re.sub(r"/+", "/", texto).strip("/")
            raw = texto
            break
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return raw


# ── Parsers ──────────────────────────────────────────────────────────────────

class ParserAML:
    NAVIERA = "AML"

    def parse(self, path: str) -> RegistroBL:
        text = _txt(path)
        lineas = text.splitlines()
        rec = RegistroBL(
            archivo=os.path.basename(path),
            naviera=self.NAVIERA,
            ts_extraccion=datetime.now(timezone.utc).isoformat(),
        )
        if lineas:
            m = re.match(r"^(.+?)\s+(\d{3})\s*$", lineas[0].strip())
            if m:
                rec.nombre = m.group(1).strip()
                rec.num_bl = m.group(2)

        for i, linea in enumerate(lineas):
            if re.match(r"^\d+[A-Z]+\s+\d{2}:\d{2}\s+\S", linea):
                rec.buque = lineas[i - 1].strip() if i > 0 else ""
                m = re.search(r"(\d{2}:\d{2})", linea)
                rec.hora = m.group(1) if m else ""
                if i + 2 < len(lineas):
                    rec.puerto_destino = lineas[i + 2].strip().upper()
                break

        m = re.search(r"^(\d+)\s+\d+\s*x\s+\w+\s+disant", text, re.MULTILINE | re.IGNORECASE)
        if m:
            rec.matricula = m.group(1)

        m = re.search(r"(\d{1,2})\s*/\s*([\w]+)\s*/\s*(\d{4})", text)
        if m:
            rec.fecha = _fecha(f"{m.group(1)}/{m.group(2)}/{m.group(3)}")

        return rec


class ParserBalearia:
    NAVIERA = "BALEARIA"

    def parse(self, path: str) -> RegistroBL:
        text = _txt(path)
        rec = RegistroBL(
            archivo=os.path.basename(path),
            naviera=self.NAVIERA,
            ts_extraccion=datetime.now(timezone.utc).isoformat(),
        )
        rec.num_bl = _re1(r"N[o\u00ba\u00b0]\s*CTO:\s*(\S+)", text)
        if rec.num_bl:
            rec.num_bl = str(rec.num_bl)
        rec.nombre = _re1(r"Remitente o embarcador:\s*(.+)", text)

        m = re.search(r"a bordo del buque\s+(.+?)\s*,\s*las", text, re.DOTALL)
        if m:
            tokens = m.group(1).split()
            tokens = [t for t in tokens if not re.match(r"^[A-Z]?\d{7,}$", t)]
            rec.buque = " ".join(tokens)

        m = re.search(r"con destino al puerto\s+de\s+(\w+)", text, re.IGNORECASE)
        if m:
            rec.puerto_destino = m.group(1).upper()

        m = re.search(r"Fecha y hora de la salida:\s*(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})", text)
        if m:
            rec.fecha = _fecha(m.group(1))
            rec.hora  = m.group(2)

        seccion = re.search(
            r"Matr.cula.+?\n(.+?)(?:TOTAL|MRN:|Cliente pagador)",
            text, re.DOTALL | re.IGNORECASE
        )
        if seccion:
            matriculas = re.findall(r"^([A-Z0-9]{7,9})\b", seccion.group(1), re.MULTILINE)
            matriculas = [m for m in matriculas if re.search(r"[A-Z]", m) and re.search(r"\d", m)]
            rec.matricula = " / ".join(matriculas) if matriculas else ""

        return rec


class ParserDFDS:
    NAVIERA = "DFDS"

    def parse(self, path: str) -> RegistroBL:
        text = _txt(path)
        rec = RegistroBL(
            archivo=os.path.basename(path),
            naviera=self.NAVIERA,
            ts_extraccion=datetime.now(timezone.utc).isoformat(),
        )
        m = re.search(r"(\d+)\s*\nCONOCIMIENTO DE EMBARQUE N", text, re.IGNORECASE)
        if m:
            rec.num_bl = m.group(1).strip()

        m = re.search(r"merc.nc.as que m.s\s*\n([^\n]+)\nabajo", text, re.IGNORECASE)
        if m:
            rec.nombre = " ".join(m.group(1).split())

        m = re.search(r"FECHA Y HORA\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})", text)
        if m:
            rec.fecha = _fecha(m.group(1))
            rec.hora  = m.group(2)

        m = re.search(r"BUQUE / VIAJE\s+(.+?)\s*[-\u2013]\s*\S+\s*\n", text)
        if m:
            rec.buque = m.group(1).strip()

        rec.puerto_destino = _re1(r"PUERTO DESTINO\s+(\S+)", text).upper()

        m = re.search(r"^([A-Z0-9]{6,10})\s+\d+\s+\S", text, re.MULTILINE)
        if m:
            rec.matricula = m.group(1)

        return rec


class ParserRFS:
    NAVIERA = "RFS"

    def parse(self, path: str) -> RegistroBL:
        rec = ParserDFDS().parse(path)
        rec.naviera = self.NAVIERA
        text = _txt(path)
        matriculas_solas = re.findall(r"^([A-Z0-9]{6,10})\s*$", text, re.MULTILINE)
        if matriculas_solas:
            rec.matricula = " / ".join(matriculas_solas)
        return rec


class ParserTrasme:
    NAVIERA = "TRASME"

    def parse(self, path: str) -> RegistroBL:
        text = _txt(path)
        rec = RegistroBL(
            archivo=os.path.basename(path),
            naviera=self.NAVIERA,
            ts_extraccion=datetime.now(timezone.utc).isoformat(),
        )
        rec.num_bl = _re1(r"NUMERO:\s*(\S+)", text)
        rec.nombre = _re1(r"Nombre:\s+(.+?)(?:\s{2,}|Puerto origen:|\n)", text)
        rec.buque = _re1(r"Buque:\s+(\S+)", text)
        rec.puerto_destino = _re1(r"Puerto destino:\s+(\S+)", text).upper()

        m = re.search(r"Fecha:\s*(\d{2}/\d{2}/\d{4})\s+Hora:\s*(\d{2}:\d{2})", text)
        if m:
            rec.fecha = _fecha(m.group(1))
            rec.hora  = m.group(2)

        m = re.search(
            r"MATRICULA\s+TIPO.+?\n([A-Z0-9]{4,10})\s+\w+",
            text, re.DOTALL | re.IGNORECASE
        )
        if m:
            rec.matricula = m.group(1)

        return rec


# ── Detección de naviera ─────────────────────────────────────────────────────

FIRMAS = {
    "disant contenir":     "AML",
    "BALEARIA EUROLINEAS": "BALEARIA",
    "BALEARIA":            "BALEARIA",   # variante sin "EUROLINEAS" (ej. EUROMAROC)
    "DFDS IBERIA":         "DFDS",
    "Red Fish Speedlines": "RFS",
    "TRASMEDITERRANEA":    "TRASME",
}

PARSERS = {
    "AML":      ParserAML(),
    "BALEARIA": ParserBalearia(),
    "DFDS":     ParserDFDS(),
    "RFS":      ParserRFS(),
    "TRASME":   ParserTrasme(),
}


def detectar_naviera(path: str) -> Optional[str]:
    texto = _txt(path)
    for firma, clave in FIRMAS.items():
        if firma.lower() in texto.lower():
            return clave
    return None


def procesar(path: str) -> Optional[RegistroBL]:
    log.info("Procesando: %s", path)
    naviera = detectar_naviera(path)
    if not naviera:
        log.warning("  Naviera no reconocida — omitiendo %s", path)
        return None
    log.info("  Naviera: %s", naviera)

    try:
        rec = PARSERS[naviera].parse(path)
    except Exception as exc:
        log.error("  Error al extraer: %s", exc, exc_info=True)
        return None

    if DESTINO_FILTRO not in rec.puerto_destino.upper():
        log.info("  Puerto destino '%s' != %s — omitido", rec.puerto_destino, DESTINO_FILTRO)
        return None

    log.info("  Destino: %s — INCLUIDO", rec.puerto_destino)
    return rec


# ── CSV ──────────────────────────────────────────────────────────────────────

CAMPOS_CSV = [f.name for f in fields(RegistroBL)]


def _clave_bl(naviera: str, num_bl: str, fecha: str,
              nombre: str, buque: str, matricula: str) -> tuple:
    """Clave de deduplicación — independiente del nombre de fichero."""
    return (
        naviera.strip().upper(),
        num_bl.strip().upper(),
        fecha.strip(),
        nombre.strip().upper(),
        buque.strip().upper(),
        matricula.strip().upper(),
    )


def guardar_csv(registros: list, ruta: str) -> None:
    # Leer claves ya existentes para evitar duplicados
    claves_existentes: set[tuple] = set()
    if os.path.isfile(ruta):
        with open(ruta, newline="", encoding="utf-8") as fh:
            for row in csv.DictReader(fh):
                claves_existentes.add(_clave_bl(
                    row.get("naviera",   ""),
                    row.get("num_bl",    ""),
                    row.get("fecha",     ""),
                    row.get("nombre",    ""),
                    row.get("buque",     ""),
                    row.get("matricula", ""),
                ))

    nuevos = []
    for rec in registros:
        clave = _clave_bl(rec.naviera, rec.num_bl, rec.fecha,
                          rec.nombre, rec.buque, rec.matricula)
        if clave in claves_existentes:
            log.warning("Duplicado ignorado: %s / %s / %s", rec.naviera, rec.num_bl, rec.fecha)
        else:
            nuevos.append(rec)
            claves_existentes.add(clave)

    if not nuevos:
        log.info("Nada nuevo que guardar — todos los registros ya existían en %s", ruta)
        return

    existe = os.path.isfile(ruta)
    with open(ruta, "a", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=CAMPOS_CSV, quoting=csv.QUOTE_ALL)
        if not existe:
            writer.writeheader()
        for rec in nuevos:
            fila = {k: str(v) if v is not None else "" for k, v in asdict(rec).items()}
            writer.writerow(fila)
    log.info("Guardados %d registro(s) nuevos en %s (ignorados %d duplicados)",
             len(nuevos), ruta, len(registros) - len(nuevos))


# ── CLI ──────────────────────────────────────────────────────────────────────

def recopilar_pdfs(paths: list) -> list:
    pdfs = []
    for arg in paths:
        if os.path.isdir(arg):
            pdfs.extend(sorted(glob.glob(os.path.join(arg, "*.pdf"))))
            pdfs.extend(sorted(glob.glob(os.path.join(arg, "*.PDF"))))
        elif any(c in arg for c in "*?"):
            pdfs.extend(sorted(glob.glob(arg)))
        elif os.path.isfile(arg):
            pdfs.append(arg)
        else:
            log.warning("Ruta no encontrada: %s", arg)
    return pdfs


def main():
    parser = argparse.ArgumentParser(description="Extractor de Conocimientos de Embarque")
    parser.add_argument(
        "pdfs", nargs="*",
        help="Ficheros PDF o directorios a procesar"
    )
    parser.add_argument(
        "-o", "--output-dir", default=None,
        help="Directorio de salida para el CSV (por defecto: junto al script)"
    )
    args = parser.parse_args()

    # Sin argumentos: usar carpeta de muestra
    if not args.pdfs:
        muestra = "/mnt/user-data/uploads"
        if os.path.isdir(muestra):
            args.pdfs = [muestra]
        else:
            parser.print_help()
            sys.exit(0)

    pdfs = recopilar_pdfs(args.pdfs)
    if not pdfs:
        log.error("No se encontraron PDFs.")
        sys.exit(1)

    log.info("PDFs encontrados: %d", len(pdfs))

    registros, omitidos = [], []
    for pdf in pdfs:
        rec = procesar(pdf)
        if rec:
            registros.append(rec)
        else:
            omitidos.append(os.path.basename(pdf))

    # Determinar directorio de salida
    if args.output_dir:
        output_dir = Path(args.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    else:
        output_dir = Path(__file__).parent

    etiqueta = datetime.now(timezone.utc).strftime("%Y%m%d")
    ruta_csv = output_dir / f"bl_{etiqueta}.csv"

    if registros:
        guardar_csv(registros, str(ruta_csv))

    # Resumen (stdout — el processor lo guarda como reporte_{batch_id}.txt)
    print(f"\n{'─'*55}")
    print(f"  Version          : {SCRIPT_VERSION}")
    print(f"  PDFs encontrados : {len(pdfs)}")
    print(f"  Incluidos        : {len(registros)}  (destino = {DESTINO_FILTRO})")
    print(f"  Omitidos         : {len(omitidos)}")
    for f in omitidos:
        print(f"    · {f}")
    if registros:
        print(f"  CSV              : {ruta_csv.name}")
    print(f"{'─'*55}\n")


if __name__ == "__main__":
    main()