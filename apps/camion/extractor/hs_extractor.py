#!/usr/bin/env python3
"""
hs_extractor.py  –  Extracción de HS codes desde DOC.pdf (OCR)
==============================================================================
El DOC.pdf que envía el cliente es un PDF combinado escaneado (sin texto
embebido) que contiene, por cada DAE:

    [Parte de Entrada]   ← slip interno de Sostmeier
    [Declaración EX]     ← contiene el MRN
    [Factura]            ← contiene "Tar.doua.: NNNNNNNN" (HS code 8-10 dígitos)
    ...
    [Siguiente Parte de Entrada]  ← inicio del bloque siguiente

Algoritmo:
    1. Render cada página a 200 dpi.
    2. OCR con tesseract (lang='spa+fra+eng').
    3. Para cada página detecta:
         - ¿Es "Parte de Entrada"? → marca inicio de bloque
         - MRNs presentes (regex 26[A-Z]{2}[A-Z0-9]{14,16})
         - HS codes (regex 'Tar.doua.: NNNNNNNN' → primeros 4 dígitos)
    4. Agrupa páginas en bloques [Parte → siguiente Parte) y mapea
       MRN → HS codes encontrados en el mismo bloque.
    5. Cachea resultados por MD5 del PDF en /tmp/hs_cache_<hash>.json
       para evitar re-OCR en re-runs.

Uso standalone:
    python hs_extractor.py DOC1.pdf DOC2.pdf

Uso integrado:
    from hs_extractor import extract_hs_map
    hs_map = extract_hs_map(['DOC1.pdf', 'DOC2.pdf'], dpi=200)
    # hs_map = {'26FR10002980612MB3': '3926', ...}
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Iterable

import fitz                # PyMuPDF
import pytesseract
from PIL import Image

# ─────────────────────────────────────────────────────────────────────────────
# Regex
# ─────────────────────────────────────────────────────────────────────────────

# MRN exportación: "26ES000841511882J0" → 2 dígitos año + 2 letras país + 14–16 alfanum
# Patrón estricto — lo usamos para mapear MRN→HS (requiere que el MRN esté bien leído)
MRN_RE = re.compile(r'\b(2\d[A-Z]{2}[A-Z0-9]{14,16})\b')

# Patrón laxo — solo se usa para DETECTAR presencia de "algún MRN" en una página
# y así abrir un bloque nuevo. Tolera errores OCR típicos como:
#   26ITQTG01CL44788A4  →  25ITOTGO1CLA788A4   (5→6, Q→O, 44→A, falta 1 char)
# Requisitos:
#   - 18 chars alfanuméricos mayúsculos, empezando por 2 dígitos
#   - al menos 2 letras en los 4 primeros chars (para distinguir de números puros)
#   - permite hasta 4 "agujeros" donde un dígito esperado es letra o viceversa
MRN_LAX_RE = re.compile(
    r'\b([12]\d[A-Z0-9]{2}[A-Z0-9]{14,16})\b'
)

# HS codes en facturas: múltiples idiomas/layouts observados en DOCs reales.
# OCR a veces confunde letras (ll→li, ff→f, etc.) por lo que los patrones
# son tolerantes. Todos capturan 8–10 dígitos seguidos; el HS de 4 dígitos
# son los primeros 4.
#
# Hay DOS familias de patrones, diferenciadas por nivel de confianza:
#
#   STRONG_HS_LABELS — etiqueta directa del HS: el label SOLO precede a HS
#     ("Tar.doua.: 39263000", "Zollnr/Cust.tariffno.: 60063200", "Commodity
#     code: 85472000"). El número que sigue ES el HS, sin ambigüedad.
#     Score base: 10.
#
#   WEAK_HS_WINDOWS — header de TABLA cuyo contenido es el HS pero también
#     puede contener otros números (Nomenclature, Warennummer, Nomenclatura
#     Combinata, Cod de nomenclatură combinată). El primer número 8-10 dígitos
#     tras el header puede ser un código de referencia/albarán, no el HS real.
#     Capturamos TODOS los candidatos en una ventana de 1500 chars y filtramos
#     por contexto léxico (ANTI_LABEL_RE backward + POSITIVE_FWD_RE forward).
#     Score base: 5, +5 si POSITIVE_FWD match.

STRONG_HS_LABELS = [
    # Francés: "Tar.doua.: 39263000"                          (ADEX/ITW, Aplix)
    re.compile(r'tar\s*\.?\s*doua\s*\.?\s*:?\s*(\d{8,10})', re.I),

    # Italiano: "Tariffa doganale/Customs rate: 42050090"    (Mastrotto)
    re.compile(r'tariff?a\s*dogana[le]*(?:\s*/\s*customs?\s*rate)?\s*:?\s*(\d{8,10})', re.I),

    # Alemán: "Zollnr./Cust.tariffno.: 60063200"             (Strähle, Guetermann)
    # OCR frecuentemente entrega "Zolinr." en vez de "Zollnr."
    re.compile(r'zoll?\w{0,3}\.?\s*/\s*cust\.?\s*tarif[fno\.\s]*:?\s*(\d{8,10})', re.I),

    # Alemán standalone: "Zollnr:" o "Zolltarif:"
    re.compile(r'zoll?\s*(?:nr|tarif)\w*\s*\.?\s*:?\s*(\d{8,10})', re.I),

    # Inglés: "Customs tariff no." / "Customs rate:"         (fallback)
    re.compile(r'customs?\s*(?:rate|tariff\s*(?:no\.?|number)?)\s*:?\s*(\d{8,10})', re.I),

    # Warennummer (EX alemán, campo [18 09])
    re.compile(r'waren\s*nummer[\s\[\]\d]*(\d{8,10})', re.I),

    # Commodity code / Code douanier (genérico UE)
    re.compile(r'(?:commodity\s*code|code\s*douanier)\s*:?\s*(\d{8,10})', re.I),

    # CTC prefix: "CTC 54011018 CN"                          (Gütermann)
    # Negative lookahead (?!0\d) descarta números que empiezan con 0X (capítulo
    # HS 00 no existe). Esto evita capturar "CT 1 0014045..." de los EAD
    # alemanes (Hoffmann), donde "CT 1" se confunde con "CTC " en OCR.
    re.compile(r'\bCTC\s+(?!0\d)(\d{8,10})\b', re.I),

    # "CUSTOM TARIFF: 84669400"                              (Mecal IT)
    # Variante inglesa sin barra ni "rate"
    re.compile(r'custom\s+tariff\s*:?\s*(\d{8,10})', re.I),

    # "HS CODE: 4821109000" / "HS-Code : 39269097"           (Micros ES, Raymond DE)
    # Label genérico universal, usado por proveedores de varias jurisdicciones.
    re.compile(r'HS[\s\-]*code\s*:?\s*(\d{8,10})', re.I),

    # "Tariff Code: 84199085"                                (Mecalbi PT)
    # Variante portuguesa/inglesa sin "Customs" delante.
    re.compile(r'tariff\s*code\s*:?\s*(\d{8,10})', re.I),
]

# Headers de TABLA — la ventana de 1500 chars tras el header puede contener
# números que NO son HS (referencias, albaranes, números de orden). Por eso
# se aplica filtrado contextual (anti-label backward + positive forward).
WEAK_HS_WINDOWS = [
    # Francés EAD/EX: "Nomenclature [18 09]" ... 58061000     (Aplix EAD)
    re.compile(r'Nomenclature(.{1,1500})', re.I | re.DOTALL),

    # Alemán Ausfuhrbegleitdokument
    re.compile(r'Waren\s*nummer(.{1,1500})', re.I | re.DOTALL),

    # Rumano: "Cod de nomenclatură combinată ..."             (Delfingen RO)
    # Clase [áaăã] cubre realizaciones del OCR para la "ă" sin lang=ron.
    re.compile(
        r'cod\s*de?\s*nomenclatur[áaăã]\s*combinat[áaăã](.{1,1500})',
        re.I | re.DOTALL,
    ),

    # Italiano: "Nomenclatura Combinata (8 cifre) [18 09]"    (MTA, otros EAD IT)
    re.compile(r'Nomenclatura\s*Combinata(.{1,1500})', re.I | re.DOTALL),
]

# Número 8-10 dígitos NO precedido por palabra (\w = [A-Za-z0-9_]) ni seguido
# por dígito. Esto rechaza referencias con prefijo letra pegado tipo
# "E20589300", pero permite "RE6288EU E20589300" (porque entre espacio y E
# hay separación) — esos casos los descarta luego ANTI_LABEL_RE.
HS_NUM_RE = re.compile(r'(?<!\w)(?!0\d)(\d{8,10})(?!\d)')

# Variante con 1 espacio interno: el OCR de MTA pg 108 entrega "853690 10"
# por errores de kerning. Normalizamos quitando el espacio antes de validar.
HS_NUM_SPACED_RE = re.compile(r'(?<!\w)(?!0\d)(\d{4,8})\s(\d{1,4})(?!\d)')

# ANTI-LABEL: tokens léxicos que indican "el número que sigue NO es HS".
# Es propiedad estructural del documento (etiqueta de tipo "número de cosa"),
# 100% dinámico: NO depende del valor del HS ni de listas de partidas WCO.
# Se evalúa sobre los ~25 chars previos al candidato dentro de la ventana.
ANTI_LABEL_RE = re.compile(
    r'(?:'
    r'CT\s*\d+\s*|'                                # Molex: "CT 1 587097309..."
    r'order\s*(?:no|number)?\s*\.?\s*:?\s*|'       # "Order no.: 377294501"
    r'consignment\s*(?:no|number)?\s*\.?\s*|'      # "Consignment no. 377294501"
    r'referen[zţt]\w*\s*\w*\s*|'                   # Referenznummer, Referinţă
    r'UCR\s*[\[\]\d\s]*|'                          # UCR [12 08]
    r'spediteur\s*-?\s*nr\s*\.?\s*|'
    r'lieferant\w*\s*-?\s*nr\s*\.?\s*|'
    r'versender\s*-?\s*\w*\s*|'
    r'sendungs?\s*[\/]?\s*ladungs?\s*-?\s*\w*\s*|'
    r'customer\s*(?:material|order)?\s*no\.?\s*|'
    r'supplier\s*(?:material|order)?\s*no\.?\s*|'
    r'EORI\s*:?\s*|'
    r'IBAN\s*\w*\s*|'
    r'tax\s*number\s*:?\s*|'
    r'registration\s*number\s*:?\s*|'
    r'SWIFT\s*:?\s*\w*\s*|'
    r'CUI\s*:?\s*\w*\s*|'                          # CUI rumano (registro fiscal)
    r'konto\s*:?\s*|'                              # IBAN/Konto bancario
    r'fournisseur\s*:?\s*|'
    r'frachtauftr\w*\s*-?\s*\w*\s*|'
    r'autorisation\s+\w+\s+|'
    r'ROREX\w*\s*|'                                # Autorizaciones RO/EX
    r'N\d{3}\s*[—\-\/]\s*|'                        # Códigos documento aduanal N380, N864
    r'20\d{2}\s*-\s*[A-Z]{2}\s*-\s*|'              # 2026-IT-... refs MRN-like
    r'\d{1,2}[\./]\d{1,2}[\./]20\d{2}\s*[\/—\-]\s*'  # Fechas DD.MM.YYYY /
    r')',
    re.I,
)

# POSITIVE-FORWARD: tokens que CONFIRMAN que el número precedente es HS.
# Aparecen inmediatamente después: % de TVA, código país ISO-2 EU, "VAT".
_COUNTRY_ISO2 = (
    r'(?:DE|FR|IT|ES|AT|CZ|PT|NL|BE|RO|HU|SK|PL|GB|IE|DK|SE|FI|EE|LV|LT|'
    r'SI|HR|BG|GR|MT|CY|LU|CH|NO)'
)
POSITIVE_FWD_RE = re.compile(
    r'^\s*(?:'
    r'\d{1,2}\s*[,\.]\s*\d{1,2}\s*%|'              # 0,00% TVA
    + _COUNTRY_ISO2 + r'\b|'                        # IT, DE, etc. como país
    r'-\s*' + _COUNTRY_ISO2 + r'\b|'                # " - RO -> ..."
    r'EU\s+VAT|'
    r'VAT\b'
    r')',
    re.I,
)

# JUNK-FORWARD: si el número está inmediatamente seguido (en ≤15 chars sin
# newline) por un identificador alfanumérico con 3+ letras mayúsculas, es muy
# probable que sea parte de un part-number / referencia que el OCR partió por
# mitad. Ej: "37729450 1WCCC001" — el "37729450" parece HS válido pero realmente
# es "377294501WCCC001" partido por OCR en dos líneas.
#
# Discriminación vs POSITIVE_FWD_RE: los códigos país ISO-2 (IT, DE, FR, ...)
# son SOLO 2 letras, no matchean [A-Z]{3,}. Los part-numbers tienen 3+ letras
# (WCCC, REEU, MIDI, etc.) o letras + dígitos mezclados.
JUNK_FWD_RE = re.compile(r'^\s{0,3}\d{0,4}[A-Z]{3,}', re.I)

PARTE_RE = re.compile(r'Parte\s+de\s+Entrada', re.I)


# ─────────────────────────────────────────────────────────────────────────────
# Cache helpers
# ─────────────────────────────────────────────────────────────────────────────

def _file_md5(path: Path) -> str:
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def _cache_path(pdf_path: Path) -> Path:
    """
    La clave de caché incluye:
      - MD5 del PDF (cambia → re-OCR, obvio)
      - Hash de las reglas de extracción (HS_PATTERNS y SHIPPER_KEYWORDS)

    Si cambian los patrones o se añaden shippers, la caché se invalida sola
    (cualquier run con código modificado escribe en un fichero distinto).
    Esto evita el bug clásico: tocas una regex, lanzas el script, lee caché
    vieja del run anterior, no aplica la regex nueva.
    """
    pdf_digest = _file_md5(pdf_path)
    rules_digest = _rules_signature()
    return Path('/tmp') / f'hs_cache_{pdf_digest}_{rules_digest}.json'


def _rules_signature() -> str:
    """
    Hash MD5 corto (8 hex) de las reglas de extracción que afectan al
    output cacheado. Si esto cambia, la caché es inválida.
    """
    # Lookup lazy del módulo (las constantes pueden estar definidas más abajo
    # en el archivo). Si por alguna razón no están disponibles, falla con
    # un hash conocido para no bloquear el módulo.
    g = globals()
    strong   = g.get('STRONG_HS_LABELS', [])
    weak     = g.get('WEAK_HS_WINDOWS', [])
    anti     = g.get('ANTI_LABEL_RE', None)
    pos_fwd  = g.get('POSITIVE_FWD_RE', None)
    junk_fwd = g.get('JUNK_FWD_RE', None)
    shippers = g.get('SHIPPER_KEYWORDS', [])

    h = hashlib.md5()
    for p in (*strong, *weak):
        h.update(p.pattern.encode('utf-8'))
        h.update(str(p.flags).encode('utf-8'))
    for p in (anti, pos_fwd, junk_fwd):
        if p is not None:
            h.update(p.pattern.encode('utf-8'))
    for key, kws in shippers:
        h.update(key.encode('utf-8'))
        for kw in kws:
            h.update(kw.encode('utf-8'))
    return h.hexdigest()[:8]


# Shippers conocidos y sus palabras clave distintivas. Se usa para recuperar
# bloques huérfanos (HS detectado pero MRN no legible en OCR) emparejándolos
# al MRN correspondiente del XLSX por nombre de shipper.
# Cada entrada: (shipper_key, [keywords_to_match])
# La key se compara case-insensitive y substring contra el shipper del XLSX.
SHIPPER_KEYWORDS: list[tuple[str, list[str]]] = [
    ('mastrotto',   ['MASTROTTO', 'GRUPPO MASTROTTO']),
    ('strahle',     ['STRÄHLE', 'STRAHLE', 'STRÄHLE+HESS', 'STRAHLE+HESS']),
    ('guetermann',  ['GÜTERMANN', 'GUETERMANN', 'GUTERMANN']),
    ('aplix',       ['APLIX']),
    ('adex',        ['ADEX', 'ITW EF&C', 'ITW ENGINEERED']),
    ('mssl',        ['MSSL', 'SUMITOMO', 'MOTHERSON']),
    ('ykk',         ['YKK']),
    ('autajon',     ['AUTAJON']),
    ('aptiv',       ['APTIV']),
    ('micros',      ['MICROS']),
    ('velcro',      ['VELCRO']),
    ('rodavigo',    ['RODAVIGO']),
    ('lamitex',     ['LAMITEX', 'MIKO']),
    ('nuova',       ['NUOVA F.NT', 'FABBRICA NONTESSUTI']),
    ('mecal',       ['MECAL']),
    ('hoffmann',    ['HOFFMANN', 'HOFFMANN SUPPLY CHAIN']),
    # ── Shippers añadidos tras validación lote Kenitra 5424-011 ────────────
    # Cubren los casos donde OCR detecta HS bien pero el shipper no estaba
    # en el diccionario, dejando el bloque huérfano y sin recuperar.
    ('delfingen',     ['DELFINGEN']),
    ('molex',         ['MOLEX']),
    ('tti',           ['TTI INC', 'TTI ELECTRONICS']),
    ('scapa',         ['SCAPA', 'GROUPE SCAPA']),
    ('lisi',          ['LISI AUTOMOTIVE']),
    ('iriso',         ['IRISO']),
    # Raymond: varias grafías porque OCR mete/quita espacios y puntos en el
    # punto inicial "A." (apellido es "Raymond", marca "A. Raymond").
    ('raymond',       ['A. RAYMOND', 'A.RAYMOND', 'RAYMOND BAGL', 'RAYMOND A.']),
    # MTA: nombre genérico (3 letras), riesgo de match fortuito en texto OCR.
    # Solo aceptar formas con sufijo corporativo explícito.
    ('mta',           ['MTA S.P.A', 'MTA SPA']),
    # Lear Vyškov (planta CZ): el OCR pierde la háček con frecuencia → ambas
    # grafías. "MAURICE WARD" es el agente logístico que aparece en sus DOCs.
    ('lear_vyskov',   ['LEAR VYSZKOW', 'VYŠKOV', 'VYSKOV', 'MAURICE WARD']),
    ('schleuniger',   ['SCHLEUNIGER']),
    ('mecalbi',       ['MECALBI']),
    ('elastomer',     ['ELASTOMER SOLUTIONS']),
    ('df_szerszam',   ['SZERSZAMGYARTO']),
]


def _detect_shipper(text: str) -> str | None:
    """
    Detecta el nombre del shipper (key) en el texto OCR, o None.

    Estrategia: longest-match-wins. Se prefiere el keyword más largo que
    matchee, no el primero que aparezca en la lista. Esto evita colisiones
    cuando un keyword es substring de otro (p.ej. 'MECAL' dentro de
    'MECALBI'). Sin esto, el shipper declarado primero en SHIPPER_KEYWORDS
    "absorbe" al que tenga un prefijo común y declarado después.
    """
    upper = text.upper()
    best_key: str | None = None
    best_len = 0
    for key, kws in SHIPPER_KEYWORDS:
        for kw in kws:
            if kw in upper and len(kw) > best_len:
                best_key = key
                best_len = len(kw)
    return best_key

def _ocr_pdf(pdf_path: Path, dpi: int = 200, lang: str = 'spa+fra+eng+ita+deu+ron',
             verbose: bool = False) -> list[dict]:
    """
    OCR todas las páginas y extrae:
        page_no, is_parte (bool), mrns (list), hs_codes (list de 4-dígitos)
    """
    doc = fitz.open(str(pdf_path))
    pages_info: list[dict] = []

    for pnum in range(len(doc)):
        page = doc[pnum]
        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)

        try:
            text = pytesseract.image_to_string(img, lang=lang)
        except pytesseract.TesseractError:
            # Idiomas no instalados → cadena de fallback progresiva
            for fallback in ('spa+fra+eng+ita', 'spa+fra+eng', 'eng'):
                try:
                    text = pytesseract.image_to_string(img, lang=fallback)
                    break
                except pytesseract.TesseractError:
                    continue
            else:
                text = ''   # ningún fallback funcionó

        is_parte = bool(PARTE_RE.search(text))
        mrns = sorted(set(MRN_RE.findall(text)))

        # Detector laxo: ¿hay algo que PARECE un MRN aunque esté mangled?
        # Se usa solo para abrir un bloque nuevo cuando "Parte de Entrada" y el
        # MRN estricto fallan juntos (caso: EADs italianos como Mastrotto).
        # Requiere además que aparezca alguna pista de documento EX/EAD para
        # reducir falsos positivos en páginas de factura normales.
        has_mrn_hint = (
            not mrns                                        # no hay MRN estricto
            and bool(MRN_LAX_RE.search(text))               # sí hay algo mrn-like
            and bool(re.search(
                r'(?:ACCOMPAGNAMENTO|ESPORTAZION|'
                r'DGDDI|CONEX\s+DELTA|MODELE\s+DGDDI|'     # EAD FR
                r'EXPORT\s*DECLARATION|TYPE\s*DE\s*D[EÉ]CLARATION|'
                r'TIPO\s*DI\s*DICHIARAZIONE|'              # EAD IT
                r'AUSFUHRBEGLEITDOKUMENT)',                # EAD DE
                text, re.I))
        )

        hs_raw: list[tuple[str, int, bool]] = []   # (hs8_or_10, score, is_strong)

        # Pase 1 — labels fuertes: el número que sigue ES HS sin ambigüedad.
        for pat in STRONG_HS_LABELS:
            for m in pat.finditer(text):
                hs_raw.append((m.group(1), 10, True))

        # Pase 2 — headers de tabla con ventana + filtro contextual.
        # Para cada candidato número 8-10 dígitos dentro de la ventana:
        #   - Si la "celda lógica" inmediatamente previa (≤25 chars, tras
        #     último \n o |) contiene un anti-label → DESCARTAR
        #   - Si los ≤25 chars siguientes contienen un positive-fwd → +5 score
        # Limitar el back a la "celda" evita falsos positivos cuando hay un
        # anti-label irrelevante en alguna línea anterior dentro de la ventana
        # (ej. "CT 1 0005019980\n\n4,550\n\n39269097" — el "CT 1" no se refiere
        # al 39269097 sino al número de bulto en la línea previa).
        def _adjacent_back(window: str, pos: int) -> str:
            raw = window[max(0, pos - 25):pos]
            # Quedarnos solo con la última celda: tras el último \n o |
            return re.split(r'[\n\|]', raw)[-1]

        for wpat in WEAK_HS_WINDOWS:
            for wm in wpat.finditer(text):
                window = wm.group(1)
                # Candidatos sin espacio interno
                for nm in HS_NUM_RE.finditer(window):
                    num   = nm.group(1)
                    back  = _adjacent_back(window, nm.start())
                    fwd   = window[nm.end():nm.end() + 25]
                    if ANTI_LABEL_RE.search(back):
                        continue
                    if JUNK_FWD_RE.match(fwd):
                        continue
                    score = 5 + (5 if POSITIVE_FWD_RE.match(fwd) else 0)
                    hs_raw.append((num, score, False))
                # Candidatos con 1 espacio interno (OCR mangled, ej. "853690 10")
                for sm in HS_NUM_SPACED_RE.finditer(window):
                    clean = sm.group(1) + sm.group(2)
                    if not (8 <= len(clean) <= 10):
                        continue
                    back = _adjacent_back(window, sm.start())
                    fwd  = window[sm.end():sm.end() + 25]
                    if ANTI_LABEL_RE.search(back):
                        continue
                    if JUNK_FWD_RE.match(fwd):
                        continue
                    score = 5 + (5 if POSITIVE_FWD_RE.match(fwd) else 0)
                    hs_raw.append((clean, score, False))

        # Sanity check: en el sistema HS no existe el capítulo 00. Cualquier
        # número de 8 dígitos que empiece por "00" es ruido OCR (típicamente
        # números de albarán o lote captados por error).
        # Acumulamos score por hs4 dentro de la página (un mismo hs4 capturado
        # por varios patrones suma sus scores). Además trackeamos qué hs4 tienen
        # al menos una contribución STRONG (label directo, no ambiguo).
        hs_cands:        dict[str, int]  = {}
        hs_strong_seen:  set[str]        = set()
        for code, score, is_strong in hs_raw:
            if code.startswith('00'):
                continue
            hs4 = code[:4]
            hs_cands[hs4] = hs_cands.get(hs4, 0) + score
            if is_strong:
                hs_strong_seen.add(hs4)
        # Vista plana retrocompatible para el log verbose y otros consumidores
        hs4 = sorted(hs_cands.keys())

        shipper = _detect_shipper(text)

        pages_info.append({
            'page': pnum + 1,
            'is_parte': is_parte,
            'has_mrn_hint': has_mrn_hint,
            'mrns': mrns,
            'hs4': hs4,
            'hs_cands': hs_cands,         # {hs4: score_total_pagina}
            'hs_strong': hs_strong_seen,  # set de hs4 con al menos 1 STRONG match
            'shipper': shipper,
        })

        if verbose:
            tag = 'PARTE' if is_parte else ('MRN?' if has_mrn_hint else '')
            ship = f' ship={shipper}' if shipper else ''
            print(f'  pg {pnum+1:>3} | {tag:<5} | MRNs={mrns or "-"} | HS4={hs4 or "-"}{ship}',
                  file=sys.stderr, flush=True)

    return pages_info


# ─────────────────────────────────────────────────────────────────────────────
# Agrupación de páginas → bloques → mapping MRN→HS
# ─────────────────────────────────────────────────────────────────────────────

def _group_into_blocks(pages_info: list[dict]) -> list[dict]:
    """
    Agrupa páginas en bloques. Abre bloque nuevo cuando:
      1. La página es un "Parte de Entrada" (slip Sostmeier)
      2. Introduce un MRN nuevo (distinto al último visto)
      3. Tiene pinta de ser una declaración EAD/EX pero OCR no leyó bien el MRN
         (has_mrn_hint). Cubre el caso de EADs italianos donde el MRN del código
         de barras sale mangled en el OCR.

    Cada bloque acumula MRNs y HS codes de todas sus páginas.
    """
    blocks: list[dict] = []
    current = None
    last_mrn_seen: str | None = None

    for info in pages_info:
        # ¿Página introduce un MRN nuevo (estricto)?
        new_mrn = next(
            (mrn for mrn in info['mrns'] if mrn != last_mrn_seen),
            None,
        )

        opens_block = (
            info['is_parte']
            or new_mrn is not None
            or info.get('has_mrn_hint', False)
        )

        if opens_block:
            if current is not None:
                blocks.append(current)
            current = {'pages': [], 'mrns': [], 'hs4': [], 'hs_cands': {},
                       'hs_strong': set()}

        if current is None:
            current = {'pages': [], 'mrns': [], 'hs4': [], 'hs_cands': {},
                       'hs_strong': set()}

        current['pages'].append(info['page'])
        current['mrns'].extend(info['mrns'])
        current['hs4'].extend(info['hs4'])
        # Acumular scores: sumamos por hs4 a lo largo de las páginas del bloque
        for hs4, sc in info.get('hs_cands', {}).items():
            current['hs_cands'][hs4] = current['hs_cands'].get(hs4, 0) + sc
        # Set de hs4 con al menos una contribución STRONG en el bloque
        current['hs_strong'] |= info.get('hs_strong', set())
        if info.get('shipper'):
            current.setdefault('shippers', []).append(info['shipper'])

        if info['mrns']:
            last_mrn_seen = info['mrns'][-1]

    if current is not None:
        blocks.append(current)

    return blocks


def _resolve_block_hs(hs_cands: dict[str, int],
                      hs_strong: set[str] | None = None) -> str | None:
    """
    Dado el diccionario {hs4: score_total} de un bloque, decide el HS final:

      - Si ≥2 HS4 distintos tienen al menos UNA contribución STRONG (label
        directo, no header de tabla) → multi-HS legítimo → '9999'. Esta regla
        sobrescribe el ratio porque dos labels directos en el mismo bloque
        son evidencia explícita de mercancía mixta (ej. Scapa: caucho 4005
        + papel adhesivo 4811, ambos con "Commodity Code").
      - Si solo hay 1 hs4 con score > 0 → ese.
      - Si el top tiene score > 2× el del segundo → top gana ("dominancia clara").
      - Si no hay dominancia → multi-HS → '9999'.
      - Si no hay ningún candidato → None.

    El ratio 2× distingue:
      - HS real + ruido OCR (3926:35 vs 2901:15 → top wins)
      - dos HS por WEAK con scores comparables (8536:15 vs 8547:15 → 9999)
    """
    if not hs_cands:
        return None
    hs_strong = hs_strong or set()
    if len(hs_strong) >= 2:
        return '9999'
    ranked = sorted(hs_cands.items(), key=lambda kv: kv[1], reverse=True)
    top_hs, top_sc = ranked[0]
    if len(ranked) == 1:
        return top_hs
    _, second_sc = ranked[1]
    if top_sc > 2 * second_sc:
        return top_hs
    return '9999'


def _blocks_to_hs_map(blocks: list[dict]) -> tuple[dict[str, str], dict[str, str]]:
    """
    Convierte la lista de bloques a dos mappings:
      - mrn_map:     {MRN: HS4}           — emparejamiento directo
      - shipper_map: {shipper_key: HS4}   — para bloques huérfanos (HS sin MRN
                                             legible, típico de EADs italianos)

    Heurísticas:
      - Si el bloque tiene 1 MRN + HS → asociar al MRN.
      - Si tiene N MRNs + HS → asociar a cada uno (marca ambigüedad en stderr).
      - Si tiene 0 MRNs pero SÍ shipper detectado + HS → volcar al shipper_map.
      - HS por bloque se resuelve con _resolve_block_hs (scoring + ratio 2×).

    Anti-falsos-positivos en shipper_map:
      - Solo se promueve un shipper→HS si el bloque resuelve a un HS != 9999
        con score >= MIN_SHIPPER_SCORE. Un solo HS aislado puede ser metadata
        de otra factura (p.ej. aptiv mencionando el HS de ADEX en cabecera).
    """
    from collections import Counter, defaultdict

    MIN_SHIPPER_SCORE = 10   # score mínimo total para confiar en shipper-match

    # Para cada MRN, acumulamos scores por hs4 a través de TODOS los bloques que
    # lo contienen (suma de evidencia). También acumulamos el set de hs4 con al
    # menos una contribución STRONG.
    mrn_score:     dict[str, dict[str, int]]     = defaultdict(lambda: defaultdict(int))
    mrn_strong:    dict[str, set[str]]           = defaultdict(set)
    shipper_score: dict[str, dict[str, int]]     = defaultdict(lambda: defaultdict(int))
    shipper_strong:dict[str, set[str]]           = defaultdict(set)

    for block in blocks:
        mrns      = sorted(set(block['mrns']))
        hs_cands  = block.get('hs_cands', {})
        hs_strong = block.get('hs_strong', set())
        if not hs_cands:
            continue

        if mrns:
            for mrn in mrns:
                for hs4, sc in hs_cands.items():
                    mrn_score[mrn][hs4] += sc
                mrn_strong[mrn] |= hs_strong
            if len(mrns) > 1:
                print(f'  ⚠ Bloque páginas {block["pages"]} con {len(mrns)} MRNs '
                      f'({mrns}) — HS cands {dict(hs_cands)} asociados a todos',
                      file=sys.stderr)
        else:
            # Bloque huérfano: sin MRN legible. Usar shipper como fallback.
            shippers = block.get('shippers', [])
            if shippers:
                top_shipper = Counter(shippers).most_common(1)[0][0]
                for hs4, sc in hs_cands.items():
                    shipper_score[top_shipper][hs4] += sc
                shipper_strong[top_shipper] |= hs_strong

    # Construir mrn_map aplicando ratio 2× a los scores acumulados
    mrn_map: dict[str, str] = {}
    for mrn, score_dict in mrn_score.items():
        resolved = _resolve_block_hs(dict(score_dict), mrn_strong.get(mrn))
        if resolved is None:
            continue
        mrn_map[mrn] = resolved
        if resolved == '9999':
            print(f'  ℹ MRN={mrn!r} con HS ambiguos {dict(score_dict)} → 9999',
                  file=sys.stderr)

    # Construir shipper_map con salvaguardas anti-ruido
    shipper_map: dict[str, str] = {}
    for ship, score_dict in shipper_score.items():
        resolved = _resolve_block_hs(dict(score_dict), shipper_strong.get(ship))
        if resolved is None or resolved == '9999':
            print(f'  ⊘ shipper={ship!r} descartado: HS ambiguo o vacío '
                  f'(scores {dict(score_dict)})',
                  file=sys.stderr)
            continue

        # Salvaguarda: el HS dominante debe tener score mínimo. Filtra
        # menciones aisladas (p.ej. un HS suelto en metadata de otra factura).
        top_score = score_dict[resolved]
        if top_score < MIN_SHIPPER_SCORE:
            print(f'  ⊘ shipper={ship!r} descartado: HS {resolved} con score {top_score} '
                  f'< MIN_SHIPPER_SCORE={MIN_SHIPPER_SCORE}',
                  file=sys.stderr)
            continue

        shipper_map[ship] = resolved
        print(f'  ℹ shipper={ship!r} → HS {resolved} (score {top_score} en bloque huérfano)',
              file=sys.stderr)

    return mrn_map, shipper_map


# ─────────────────────────────────────────────────────────────────────────────
# API pública
# ─────────────────────────────────────────────────────────────────────────────

# Cambio de forma en v3 → versiona la caché para invalidar automáticamente
# resultados de versiones anteriores del extractor (v2 → v3: scoring + ratio2x).
_CACHE_VERSION = 3


def extract_hs_from_pdf(pdf_path: str | Path, dpi: int = 200,
                        verbose: bool = False, use_cache: bool = True
                        ) -> dict[str, dict[str, str]]:
    """
    Extrae mapeo {MRN: HS4} y {shipper_key: HS4} de un único DOC.pdf,
    con caché por MD5 + hash de reglas.

    La caché se invalida automáticamente cuando se modifican HS_PATTERNS
    o SHIPPER_KEYWORDS (el nombre del fichero incluye un hash de las reglas).

    Returns:
        {
          'mrn':     {MRN: HS4, ...},
          'shipper': {shipper_key: HS4, ...},   # bloques huérfanos
        }
    """
    pdf_path = Path(pdf_path)
    cache = _cache_path(pdf_path)

    if use_cache and cache.exists():
        try:
            data = json.loads(cache.read_text(encoding='utf-8'))
            if data.get('_v') == _CACHE_VERSION:
                return {'mrn': data['mrn'], 'shipper': data['shipper']}
            # Caché de versión antigua → ignorar y re-OCR
        except Exception:
            pass

    # Limpia cachés viejas del MISMO PDF con OTRA firma de reglas, para que
    # /tmp no acumule ficheros obsoletos a medida que se evolucionan los
    # patrones. La caché actual se escribirá fresco abajo.
    if use_cache:
        try:
            pdf_digest = _file_md5(pdf_path)
            for old in Path('/tmp').glob(f'hs_cache_{pdf_digest}_*.json'):
                if old != cache:
                    old.unlink()
        except Exception:
            pass

    print(f'  🔎 OCR {pdf_path.name} ({dpi} dpi)...', file=sys.stderr, flush=True)
    pages_info = _ocr_pdf(pdf_path, dpi=dpi, verbose=verbose)
    blocks = _group_into_blocks(pages_info)
    mrn_map, shipper_map = _blocks_to_hs_map(blocks)

    result = {'mrn': mrn_map, 'shipper': shipper_map}

    if use_cache:
        try:
            payload = {'_v': _CACHE_VERSION, **result}
            cache.write_text(json.dumps(payload, indent=2), encoding='utf-8')
        except Exception as e:
            print(f'  ⚠ No se pudo cachear: {e}', file=sys.stderr)

    return result


def extract_hs_map(pdf_paths: Iterable[str | Path], dpi: int = 200,
                   verbose: bool = False, use_cache: bool = True
                   ) -> dict[str, dict[str, str]]:
    """
    Combina resultados de varios DOC.pdfs en los dos mappings unificados.
    Si hay conflicto, el primer valor gana y se avisa por stderr.
    """
    combined_mrn:     dict[str, str] = {}
    combined_shipper: dict[str, str] = {}

    for path in pdf_paths:
        partial = extract_hs_from_pdf(path, dpi=dpi, verbose=verbose, use_cache=use_cache)

        for mrn, hs in partial['mrn'].items():
            if mrn in combined_mrn and combined_mrn[mrn] != hs:
                print(f'  ⚠ Conflicto HS para MRN {mrn}: {combined_mrn[mrn]} vs {hs} '
                      f'(en {Path(path).name}) — me quedo con {combined_mrn[mrn]}',
                      file=sys.stderr)
                continue
            combined_mrn[mrn] = hs

        for ship, hs in partial['shipper'].items():
            if ship in combined_shipper and combined_shipper[ship] != hs:
                print(f'  ⚠ Conflicto HS para shipper {ship}: {combined_shipper[ship]} '
                      f'vs {hs} (en {Path(path).name}) — me quedo con {combined_shipper[ship]}',
                      file=sys.stderr)
                continue
            combined_shipper[ship] = hs

    return {'mrn': combined_mrn, 'shipper': combined_shipper}


# ─────────────────────────────────────────────────────────────────────────────
# CLI standalone
# ─────────────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(
        description='OCR de DOC.pdf para extraer HS codes por MRN')
    parser.add_argument('pdfs', nargs='+', help='DOC.pdf paths')
    parser.add_argument('--dpi', type=int, default=200)
    parser.add_argument('--no-cache', action='store_true')
    parser.add_argument('--verbose', action='store_true')
    args = parser.parse_args()

    result = extract_hs_map(args.pdfs, dpi=args.dpi,
                            verbose=args.verbose, use_cache=not args.no_cache)

    mrn_map     = result.get('mrn', {})
    shipper_map = result.get('shipper', {})

    print('\n=== Resultado ===')
    if mrn_map:
        print(f'\n{len(mrn_map)} MRNs con HS directo:')
        for mrn, hs in sorted(mrn_map.items()):
            print(f'  {mrn}  →  {hs}')
    if shipper_map:
        print(f'\n{len(shipper_map)} HS por shipper (bloques con MRN no legible):')
        for ship, hs in sorted(shipper_map.items()):
            print(f'  shipper={ship!r}  →  {hs}')
    if not mrn_map and not shipper_map:
        print('(ningún HS code encontrado)')

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
