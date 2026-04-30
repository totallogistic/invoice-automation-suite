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
HS_PATTERNS = [
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

    # Nomenclature column in EAD/EX francés: "Nomenclature [18 09]" ...
    # muchas líneas después ... "58061000 LAPLIX..."         (Aplix EAD)
    # Ventana amplia (1500 chars) porque entre el label y el primer valor
    # aparecen todas las cabeceras de columna de la tabla EAD (medida real
    # en OCR 200 dpi: ~680 chars). El negative lookahead (?!0\d) salta
    # números de albarán/lote que empiecen por 0X y siguen buscando.
    re.compile(r'Nomenclature.{1,1500}?(?<!\d)(?!0\d)(\d{8,10})(?!\d)', re.I | re.DOTALL),

    # Warennummer en Ausfuhrbegleitdokument alemán, mismo esquema que
    # Nomenclature pero en alemán                            (EX alemán)
    re.compile(r'Waren\s*nummer.{1,1500}?(?<!\d)(?!0\d)(\d{8,10})(?!\d)', re.I | re.DOTALL),
]

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
    patterns = g.get('HS_PATTERNS', [])
    shippers = g.get('SHIPPER_KEYWORDS', [])

    h = hashlib.md5()
    for p in patterns:
        h.update(p.pattern.encode('utf-8'))
        h.update(str(p.flags).encode('utf-8'))
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
]


def _detect_shipper(text: str) -> str | None:
    """Detecta el nombre del shipper (key) en el texto OCR, o None."""
    upper = text.upper()
    for key, kws in SHIPPER_KEYWORDS:
        for kw in kws:
            if kw in upper:
                return key
    return None

def _ocr_pdf(pdf_path: Path, dpi: int = 200, lang: str = 'spa+fra+eng+ita+deu',
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

        hs_raw: list[str] = []
        for pat in HS_PATTERNS:
            for m in pat.finditer(text):
                hs_raw.append(m.group(1))
        # Sanity check: en el sistema HS no existe el capítulo 00. Cualquier
        # número de 8 dígitos que empiece por "00" es ruido OCR (típicamente
        # números de albarán o lote captados por error).
        hs4 = sorted({code[:4] for code in hs_raw if not code.startswith('00')})

        shipper = _detect_shipper(text)

        pages_info.append({
            'page': pnum + 1,
            'is_parte': is_parte,
            'has_mrn_hint': has_mrn_hint,
            'mrns': mrns,
            'hs4': hs4,
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
            current = {'pages': [], 'mrns': [], 'hs4': []}

        if current is None:
            current = {'pages': [], 'mrns': [], 'hs4': []}

        current['pages'].append(info['page'])
        current['mrns'].extend(info['mrns'])
        current['hs4'].extend(info['hs4'])
        if info.get('shipper'):
            current.setdefault('shippers', []).append(info['shipper'])

        if info['mrns']:
            last_mrn_seen = info['mrns'][-1]

    if current is not None:
        blocks.append(current)

    return blocks


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
      - HS más frecuente dentro del bloque gana.

    Anti-falsos-positivos en shipper_map:
      - Solo se promueve un shipper→HS si el HS aparece al menos MIN_SHIPPER_HITS
        veces en el bloque (2 por defecto). Un solo HS aislado puede ser metadata
        de otra factura (p.ej. aptiv mencionando el HS de ADEX en cabecera).
      - Si hay empate entre varios HS, se descarta el match (ambiguo).
    """
    from collections import Counter, defaultdict

    MIN_SHIPPER_HITS = 2   # cuántas veces debe aparecer un HS para confiar en él

    mrn_hs_counts:     dict[str, Counter] = defaultdict(Counter)
    shipper_hs_counts: dict[str, Counter] = defaultdict(Counter)

    for block in blocks:
        mrns    = sorted(set(block['mrns']))
        hs_list = block['hs4']
        if not hs_list:
            continue

        if mrns:
            for mrn in mrns:
                for hs in hs_list:
                    mrn_hs_counts[mrn][hs] += 1
            if len(mrns) > 1:
                print(f'  ⚠ Bloque páginas {block["pages"]} con {len(mrns)} MRNs '
                      f'({mrns}) — HS {hs_list} asociados a todos',
                      file=sys.stderr)
        else:
            # Bloque huérfano: sin MRN legible. Usar shipper como fallback.
            shippers = block.get('shippers', [])
            if shippers:
                top_shipper = Counter(shippers).most_common(1)[0][0]
                for hs in hs_list:
                    shipper_hs_counts[top_shipper][hs] += 1

    # Construir mrn_map: HS más frecuente gana
    mrn_map = {mrn: c.most_common(1)[0][0] for mrn, c in mrn_hs_counts.items()}

    # Construir shipper_map con salvaguardas anti-ruido
    shipper_map: dict[str, str] = {}
    for ship, counter in shipper_hs_counts.items():
        top_list = counter.most_common(2)
        top_hs, top_hits = top_list[0]

        # Salvaguarda 1: frecuencia mínima (evita HS aislados que son metadata ajena)
        if top_hits < MIN_SHIPPER_HITS:
            print(f'  ⊘ shipper={ship!r} descartado: HS {top_hs} solo aparece {top_hits} vez '
                  f'(umbral MIN_SHIPPER_HITS={MIN_SHIPPER_HITS})',
                  file=sys.stderr)
            continue

        # Salvaguarda 2: si hay empate en primer lugar, es ambiguo → descartar
        if len(top_list) > 1 and top_list[0][1] == top_list[1][1]:
            print(f'  ⊘ shipper={ship!r} descartado: empate entre {top_list[0]} y {top_list[1]}',
                  file=sys.stderr)
            continue

        shipper_map[ship] = top_hs
        print(f'  ℹ shipper={ship!r} → HS {top_hs} ({top_hits} ocurrencias en bloque huérfano)',
              file=sys.stderr)

    return mrn_map, shipper_map


# ─────────────────────────────────────────────────────────────────────────────
# API pública
# ─────────────────────────────────────────────────────────────────────────────

# Cambio de forma en v2 → versiona la caché para invalidar automáticamente
# resultados de versiones anteriores del extractor.
_CACHE_VERSION = 2


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