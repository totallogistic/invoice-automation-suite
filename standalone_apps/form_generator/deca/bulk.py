"""Importación masiva de DeCA desde CSV.

Genera un DeCA por cada línea de un CSV de repartos (formato de exportación de
Visual Trans: codificación cp1252, delimitador ';', números en formato español
—el punto es separador de miles). Aplica las MISMAS reglas que el formulario web:

  - Se OMITE la fila si el embalaje es «Caja»/«Cajas» (no procede DeCA).
  - Se OMITE la fila si no hay matrícula (vacía o "0"), igual que el formulario
    rechaza el alta sin matrícula.

Los nombres de empresa del CSV (expedidor / cargador / transportista) se
enriquecen con su NIF y domicilio a partir del catálogo del grupo (el mismo que
usa el desplegable del formulario); el destinatario se toma tal cual (nombre del
cliente final). La fecha de transporte es el día de emisión (hoy).

Devuelve un ZIP con un PDF por DeCA generado + un índice de control
(00_indice.csv) que lista TODAS las filas: las generadas (con UUID y URL) y las
omitidas (con el motivo), de modo que el resumen sirva también como control.
"""

from __future__ import annotations

import csv
import datetime as dt
import io
import re
import unicodedata
import zipfile

from .models import DecaInput

# El CSV de Visual Trans viene en Windows-1252 (cp1252).
CSV_ENCODING = "cp1252"

# ── Catálogo de empresas del grupo (alineado con EMPRESAS del formulario web) ──
_TL_NIF = "B29961059"
_TL_NOMBRE = "Total Logistic Services, S.L."
# El domicilio del expedidor/cargador es el del almacén de SALIDA → se elige por
# el ORIGEN del transporte.
_TL_DOMICILIOS = {
    "ALGECIRAS": "Avda. del Estrecho, parcela 5.15, Pol. Ind. La Menacha, 11204 Algeciras (Cádiz)",
    "MALAGA": "CTM Centro de Transporte de Mercancías, C/ Gluck s/n módulo 4, 29590 Campanillas (Málaga)",
    "PUERTO REAL": "C/ Chile s/n, Parcela I-4 nave B6, Bajo de la Cabezuela, 11519 Puerto Real (Cádiz)",
}
_TL_DOM_DEFAULT = _TL_DOMICILIOS["MALAGA"]

_CARMELO_NIF = "B29905940"
_CARMELO_NOMBRE = "CARMELO MARTINEZ RODRIGUEZ, S.L."
_CARMELO_DOM = "P.I. Sepes - C/ Dalia nave 25, 52006 Melilla"

# Reglas de omisión (mismas que el formulario).
_EMBALAJES_OMITIR = {"caja", "cajas"}
_MATRICULAS_VACIAS = {"", "0", "-", "N/A", "NA", "SIN", "SIN MATRICULA"}


# ── Utilidades ────────────────────────────────────────────────────────────────
def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def _norm(s) -> str:
    """minúsculas, sin acentos, sin espacios sobrantes (para comparar)."""
    return _strip_accents((s or "").strip().lower())


def _slug(s, maxlen: int = 40) -> str:
    s = _strip_accents((s or "").strip()).lower()
    s = re.sub(r"[^a-z0-9]+", "_", s).strip("_")
    return s[:maxlen] or "sin_nombre"


def _parse_peso(s) -> float:
    """Formato español: el punto es separador de miles, la coma decimal.
    '1.119' -> 1119.0 ; '383' -> 383.0 ; '1.119,50' -> 1119.5"""
    t = (s or "").strip().replace(".", "").replace(",", ".")
    if not t:
        return 0.0
    try:
        return float(t)
    except ValueError:
        t2 = re.sub(r"[^0-9.]", "", t)
        return float(t2) if t2 else 0.0


def _col(row: dict, *candidatos) -> str:
    """Lee una columna por nombre de cabecera, tolerante a acentos/espacios."""
    for c in candidatos:
        cn = _norm(c)
        for k, v in row.items():
            if k is not None and _norm(k) == cn:
                return (v or "").strip()
    return ""


def _empresa(nombre, origen: str = ""):
    """(nombre, nif, domicilio, reconocida) para una empresa del grupo.
    Total Logistic toma el domicilio del almacén de origen del transporte."""
    n = _norm(nombre)
    if "total logistic" in n:
        dom = _TL_DOM_DEFAULT
        o = _norm(origen)
        for clave, key in (("puerto real", "PUERTO REAL"), ("algeciras", "ALGECIRAS"), ("malaga", "MALAGA")):
            if clave in o:
                dom = _TL_DOMICILIOS[key]
                break
        return (_TL_NOMBRE, _TL_NIF, dom, True)
    if "carmelo" in n:
        return (_CARMELO_NOMBRE, _CARMELO_NIF, _CARMELO_DOM, True)
    return ((nombre or "").strip(), "", "", False)


# ── Motor ───────────────────────────────────────────────────────────────────
def generar_desde_csv(file_bytes: bytes, svc, fecha: dt.date | None = None,
                      telefono_default: str | None = None) -> dict:
    """Procesa el CSV y genera un DeCA por fila (salvo las omitidas).

    Devuelve un dict con: zip_bytes, nombre_zip, total, generados, omitidos, filas.
    `svc` es un DecaService (se usa svc.create(data, publico=True) por fila).
    """
    fecha = fecha or dt.date.today()
    texto = file_bytes.decode(CSV_ENCODING, errors="replace")
    reader = csv.DictReader(io.StringIO(texto), delimiter=";")

    filas: list[dict] = []
    generados: list[tuple[str, bytes]] = []

    for i, row in enumerate(reader, start=1):
        expedidor_raw = _col(row, "EXPEDIDOR")
        dest_raw = _col(row, "DESTINATARIO/CONSIGNATARIO", "DESTINATARIO", "CONSIGNATARIO")
        cargador_raw = _col(row, "CARGADOR CONTRACTUAL", "CARGADOR")
        transp_raw = _col(row, "TRANSPORTISTA EFECTIVO", "TRANSPORTISTA")
        origen = _col(row, "ORIGEN")
        destino = _col(row, "DESTINO")
        matricula = _col(row, "MATRICULA", "MATRÍCULA")
        naturaleza = _col(row, "NATURALEZA")
        peso_raw = _col(row, "PESO")
        bultos = _col(row, "BULTOS")
        telefono = _col(row, "TELEFONO", "TELÉFONO") or (telefono_default or "")
        embalaje = _col(row, "EMBALAJE")

        # Fila totalmente vacía → se ignora en silencio (no cuenta).
        if not any([expedidor_raw, dest_raw, cargador_raw, transp_raw, origen,
                    destino, naturaleza, matricula]):
            continue

        base = {
            "fila": i, "destinatario": dest_raw, "naturaleza": naturaleza,
            "embalaje": embalaje, "matricula": matricula, "peso": peso_raw,
            "bultos": bultos, "origen": origen, "destino": destino,
        }

        # 1) Reglas de omisión (mismas que el formulario).
        motivos = []
        if _norm(embalaje) in _EMBALAJES_OMITIR:
            motivos.append(f"embalaje «{embalaje}» (no procede DeCA)")
        if matricula.strip().upper() in _MATRICULAS_VACIAS:
            motivos.append("sin matrícula")
        if motivos:
            filas.append({**base, "estado": "OMITIDO", "motivo": "; ".join(motivos),
                          "uuid": "", "url_publica": "", "archivo": ""})
            continue

        # 2) Enriquecer empresas (nombre → NIF/domicilio del catálogo del grupo).
        exp_n, exp_nif, exp_dom, exp_ok = _empresa(expedidor_raw or cargador_raw, origen)
        car_n, car_nif, car_dom, car_ok = _empresa(cargador_raw or expedidor_raw, origen)
        tr_n, tr_nif, tr_dom, tr_ok = _empresa(transp_raw, origen)
        if not car_ok or not tr_ok:
            desconocidas = []
            if not car_ok:
                desconocidas.append(f"cargador «{cargador_raw}»")
            if not tr_ok:
                desconocidas.append(f"transportista «{transp_raw}»")
            filas.append({**base, "estado": "OMITIDO",
                          "motivo": "empresa no reconocida: " + ", ".join(desconocidas),
                          "uuid": "", "url_publica": "", "archivo": ""})
            continue

        # 3) Construir el DeCA (fecha = día de emisión; destinatario = cliente final).
        data = DecaInput(
            tipo_documento="deca",
            expedidor={"nombre": exp_n, "nif": exp_nif, "domicilio": exp_dom},
            cargador_contractual={"nombre": car_n, "nif": car_nif, "domicilio": car_dom},
            transportista_efectivo={"nombre": tr_n, "nif": tr_nif},
            origen=origen, destino=destino,
            mercancia={"naturaleza": naturaleza, "peso_kg": _parse_peso(peso_raw),
                       "bultos": bultos or None, "embalaje": embalaje or None},
            fecha_transporte=fecha,
            matricula_tractora=matricula,
            telefono_conductor=telefono or None,
            destinatario={"nombre": dest_raw} if dest_raw else None,
        )

        # 4) Generar (público: bucket + QR), tolerando errores por fila.
        try:
            rec = svc.create(data, publico=True)
        except Exception as e:  # noqa: BLE001 (una fila mala no debe tumbar el lote)
            filas.append({**base, "estado": "OMITIDO", "motivo": f"error al generar: {e}",
                          "uuid": "", "url_publica": "", "archivo": ""})
            continue

        archivo = f"{i:02d}_{_slug(dest_raw)}_{rec.uuid[:8]}.pdf"
        generados.append((archivo, rec.pdf_bytes))
        filas.append({**base, "estado": "GENERADO", "motivo": "",
                      "uuid": rec.uuid, "url_publica": rec.url_publica or "", "archivo": archivo})

    # ── Índice de control (CSV ';' con BOM para que Excel lo abra bien) ──
    idx = io.StringIO()
    idx.write("﻿")
    w = csv.writer(idx, delimiter=";", lineterminator="\r\n")
    w.writerow(["Fila", "Estado", "Motivo", "Destinatario", "Naturaleza", "Embalaje",
                "Matrícula", "Peso (kg)", "Bultos", "Origen", "Destino",
                "UUID", "URL pública", "Archivo PDF"])
    for f in filas:
        w.writerow([f["fila"], f["estado"], f["motivo"], f["destinatario"], f["naturaleza"],
                    f["embalaje"], f["matricula"], f["peso"], f["bultos"], f["origen"],
                    f["destino"], f["uuid"], f["url_publica"], f["archivo"]])
    indice_bytes = idx.getvalue().encode("utf-8")

    # ── Empaquetar ZIP ──
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("00_indice.csv", indice_bytes)
        for nombre, pdf in generados:
            z.writestr(nombre, pdf)
    zip_bytes = buf.getvalue()

    n_gen = sum(1 for f in filas if f["estado"] == "GENERADO")
    stamp = dt.datetime.now().strftime("%H%M%S")
    return {
        "zip_bytes": zip_bytes,
        "nombre_zip": f"deca_lote_{fecha.strftime('%Y%m%d')}_{stamp}.zip",
        "total": len(filas),
        "generados": n_gen,
        "omitidos": len(filas) - n_gen,
        "filas": filas,
    }
