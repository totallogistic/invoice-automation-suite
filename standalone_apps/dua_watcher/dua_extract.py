#!/usr/bin/env python3
"""
dua_extract.py — Extraccion y validacion del numero DUA de un PDF escaneado.

Los DUA de aduanas llegan como PDF ESCANEADO (sin capa de texto), asi que hay
que usar OCR (tesseract). El numero DUA tiene formato fijo:

    AA ES NNNNNN X NNN X N R N        (18 caracteres)
    ^^ ^^ -------------------
    año ES  resto alfanumerico
    ej.: 26ES001131I040P6R0

El OCR sobre escaneos de baja calidad tiene un SESGO sistematico conocido:
confunde I<->1 y O<->0. Por eso NO se confia ciegamente:

  - Se hace OCR de TODAS las paginas y se recogen todos los candidatos a DUA.
  - El mismo DUA suele aparecer en varias paginas -> validacion cruzada.
  - Se reconstruye por voto posicion-a-posicion entre paginas.
  - Se calcula un nivel de CONFIANZA:
      * ALTA   : formato exacto + sin posiciones en discrepancia + >=2 paginas
                 coinciden + sin caracteres ambiguos sin resolver.
      * DUDOSO : cualquier otra cosa (chars ambiguos, discrepancia, 1 sola pagina,
                 formato no exacto) -> requiere revision humana.

Uso como libreria:
    from dua_extract import extract_dua
    res = extract_dua("/ruta/al.pdf")
    # res = {dua, confianza, n_paginas, lecturas_por_pagina, motivo}

Uso CLI (para probar):
    python3 dua_extract.py fichero.pdf
"""

import re
import sys
import json
from collections import Counter

import pytesseract
from pdf2image import convert_from_path

# ── Formato del DUA ───────────────────────────────────────────────────────────
# 2 digitos (año) + ES + 14 alfanumericos = 18 chars.
DUA_LEN = 18
DUA_CORE_RE = re.compile(r"\d{2}ES[0-9A-Z]{14}")
# Para capturar en texto OCR (puede venir pegado a texto vecino; recortamos a 18).
DUA_FIND_RE = re.compile(r"\d{2}ES[0-9A-Z]{14,}")

OCR_LANG = "eng"   # el DUA es alfanumerico; no necesita diccionario español
OCR_DPI = 300

# Caracteres que el OCR confunde en escaneos (para marcar ambiguedad).
AMBIGUOUS = {"I": "1", "1": "I", "O": "0", "0": "O", "B": "8", "8": "B"}


def _ocr_pages(pdf_path, dpi=OCR_DPI):
    """Devuelve lista de (n_pagina, texto_ocr)."""
    out = []
    for i, img in enumerate(convert_from_path(pdf_path, dpi=dpi)):
        txt = pytesseract.image_to_string(img, lang=OCR_LANG)
        out.append((i + 1, txt))
    return out


def _candidates(txt):
    """Candidatos a DUA en un texto OCR, recortados a 18 chars."""
    flat = txt.replace(" ", "").replace("\n", "")
    return [m.group(0)[:DUA_LEN] for m in DUA_FIND_RE.finditer(flat)]


def _skeleton(s):
    """Colapsa caracteres ambiguos a una forma canonica para agrupar lecturas
    del MISMO DUA que difieren solo en I/1, O/0."""
    return s.replace("I", "1").replace("O", "0").replace("B", "8")


def extract_dua(pdf_path, dpi=OCR_DPI):
    """Extrae el DUA de un PDF escaneado con nivel de confianza.

    Returns dict:
      dua               : mejor reconstruccion (str) o None si no se hallo nada
      confianza         : "alta" | "dudoso"
      motivo            : explicacion del nivel (str)
      n_paginas         : nº de paginas del PDF
      paginas_con_dua   : nº de paginas donde aparecio el DUA ganador
      lecturas          : {pagina: [candidatos]} para auditoria
      posiciones_ambiguas: lista de indices donde las paginas discreparon
    """
    pages = _ocr_pages(pdf_path, dpi=dpi)
    n_paginas = len(pages)

    lecturas = {}
    todos = []
    for pg, txt in pages:
        cands = _candidates(txt)
        if cands:
            lecturas[pg] = cands
            todos.extend((pg, c) for c in cands)

    if not todos:
        return {
            "dua": None, "confianza": "dudoso",
            "motivo": "No se detecto ningun patron de DUA por OCR",
            "n_paginas": n_paginas, "paginas_con_dua": 0,
            "lecturas": lecturas, "posiciones_ambiguas": [],
        }

    # Agrupar por esqueleto (mismo DUA pese a I/1, O/0)
    groups = {}
    for pg, c in todos:
        groups.setdefault(_skeleton(c), []).append((pg, c))

    # Grupo ganador = el que aparece en mas paginas distintas
    def paginas_distintas(variants):
        return len(set(pg for pg, _ in variants))
    best_k = max(groups, key=lambda k: paginas_distintas(groups[k]))
    variants = groups[best_k]
    paginas_con_dua = paginas_distintas(variants)

    # Reconstruccion por voto posicion-a-posicion
    L = max(len(c) for _, c in variants)
    recon = []
    ambiguas = []
    for i in range(L):
        chars = [c[i] for _, c in variants if len(c) > i]
        cnt = Counter(chars)
        ganador = cnt.most_common(1)[0][0]
        recon.append(ganador)
        # discrepancia entre paginas en esta posicion
        if len(set(chars)) > 1:
            ambiguas.append(i)
    dua = "".join(recon)

    # ── Nivel de confianza ──
    motivos = []
    es_formato = bool(DUA_CORE_RE.fullmatch(dua))
    if not es_formato:
        motivos.append(f"formato no exacto (len={len(dua)}, esperado {DUA_LEN})")
    if ambiguas:
        motivos.append(f"paginas discrepan en posiciones {ambiguas}")
    if paginas_con_dua < 2:
        motivos.append("DUA visto en una sola pagina (sin validacion cruzada)")
    # caracteres intrinsecamente ambiguos en el resultado (riesgo de sesgo)
    chars_ambiguos = [i for i, ch in enumerate(dua) if ch in AMBIGUOUS]
    if chars_ambiguos:
        motivos.append(f"contiene caracteres ambiguos I/1,O/0,B/8 en {chars_ambiguos}")

    # ALTA solo si: formato exacto + sin discrepancia + >=2 paginas + sin ambiguos.
    # NOTA: por el sesgo sistematico del OCR en estos escaneos, en la practica casi
    # siempre habra algun caracter ambiguo -> casi todo cae en DUDOSO. Es DELIBERADO:
    # para compliance preferimos revision humana a un nombre incorrecto.
    if es_formato and not ambiguas and paginas_con_dua >= 2 and not chars_ambiguos:
        confianza = "alta"
        motivo = "formato exacto, confirmado en varias paginas, sin caracteres ambiguos"
    else:
        confianza = "dudoso"
        motivo = "; ".join(motivos) if motivos else "revision recomendada"

    return {
        "dua": dua,
        "confianza": confianza,
        "motivo": motivo,
        "n_paginas": n_paginas,
        "paginas_con_dua": paginas_con_dua,
        "lecturas": lecturas,
        "posiciones_ambiguas": ambiguas,
    }


def main():
    if len(sys.argv) != 2:
        print("Uso: dua_extract.py <fichero.pdf>", file=sys.stderr)
        sys.exit(2)
    res = extract_dua(sys.argv[1])
    print(json.dumps(res, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
