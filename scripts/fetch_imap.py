#!/usr/bin/env python3
"""
Conecta vía IMAP a un buzón, descarga los mensajes de una carpeta y
extrae las referencias #EX{ref}# para construir/mergear _processed.json
del endpoint de Hoja Control Expedientes.

CREDENCIALES por variables de entorno (no las pases como args):
    IMAP_HOST   p.ej. imap.serviciodecorreo.es
    IMAP_USER   p.ej. miguel.pino@codeengtools.eu
    IMAP_PASS   contraseña / app password
    IMAP_PORT   opcional, default 993

USO:
    # Listar carpetas disponibles (útil para confirmar nombre exacto):
    python3 fetch_imap.py --list-folders

    # Descargar y procesar carpeta "Reportes", mergear en _processed.json:
    python3 fetch_imap.py _processed.json Reportes

    # Otra carpeta con prefijo (algunos servidores las namespace bajo INBOX):
    python3 fetch_imap.py _processed.json INBOX/Reportes

Idempotente: puedes ejecutarlo varias veces. Si ya existe el JSON, hace
merge conservando la entrada con timestamp más reciente.
"""

import os
import sys
import json
import re
import imaplib
import email
import argparse
from email.utils import parsedate_to_datetime
from email.header import decode_header
from collections import Counter
from datetime import datetime

RE_SUBJ = re.compile(r'#EX(\w+)#')
RE_TIPO = re.compile(r'Tipo:\s*(HOJA CONTROL [^\n\r]+?)(?:\s*Usuario|$)', re.MULTILINE)
RE_USR  = re.compile(r'Usuario:\s*([^\n\r]+?)(?:\s*Fecha|$)', re.MULTILINE)

BATCH_SIZE = 100   # mensajes por lote (evita timeouts con carpetas grandes)


def decode_subject(raw: str) -> str:
    """Decodifica subjects en RFC 2047 (UTF-8 base64, latin1, etc.)."""
    if not raw:
        return ''
    parts = decode_header(raw)
    out = []
    for chunk, enc in parts:
        if isinstance(chunk, bytes):
            try:
                out.append(chunk.decode(enc or 'utf-8', errors='ignore'))
            except Exception:
                out.append(chunk.decode('utf-8', errors='ignore'))
        else:
            out.append(chunk)
    return ''.join(out)


def get_body(msg) -> str:
    """Extrae text/plain del mensaje, manejando multipart."""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                try:
                    return part.get_payload(decode=True).decode('utf-8', errors='ignore')
                except Exception:
                    pass
        return ''
    try:
        return msg.get_payload(decode=True).decode('utf-8', errors='ignore')
    except Exception:
        return ''


def connect():
    host = os.getenv('IMAP_HOST')
    user = os.getenv('IMAP_USER')
    pwd  = os.getenv('IMAP_PASS')
    port = int(os.getenv('IMAP_PORT', '993'))
    if not (host and user and pwd):
        print("❌ Faltan IMAP_HOST / IMAP_USER / IMAP_PASS en el entorno", file=sys.stderr)
        sys.exit(2)
    print(f"🔌 Conectando a {host}:{port} como {user}...")
    M = imaplib.IMAP4_SSL(host, port)
    M.login(user, pwd)
    print("✅ Login OK")
    return M


def list_folders(M):
    print("\n📁 Carpetas disponibles:")
    typ, data = M.list()
    if typ != 'OK':
        print(f"❌ Error listando carpetas: {data}")
        return
    for line in data:
        # Formato típico: b'(\\HasNoChildren) "/" "INBOX"'
        try:
            decoded = line.decode('utf-8', errors='ignore')
        except Exception:
            decoded = str(line)
        # Extrae el nombre entre las últimas comillas
        m = re.search(r'"([^"]*)"\s*$', decoded)
        name = m.group(1) if m else decoded
        print(f"  {name}")


def fetch_and_extract(M, folder: str) -> dict:
    """Selecciona una carpeta, descarga todos los mensajes y devuelve
    {ref: entry} solo de los que tienen #EX{ref}# en subject."""
    print(f"\n📂 Seleccionando carpeta {folder!r}...")
    typ, data = M.select(folder, readonly=True)
    if typ != 'OK':
        print(f"❌ No se pudo seleccionar {folder}: {data}", file=sys.stderr)
        sys.exit(3)
    total = int(data[0])
    print(f"   {total} mensajes en la carpeta")

    typ, data = M.search(None, 'ALL')
    if typ != 'OK':
        print(f"❌ Error en SEARCH: {data}", file=sys.stderr)
        sys.exit(3)
    ids = data[0].split()
    print(f"   {len(ids)} IDs devueltos por SEARCH\n")

    entries = {}
    matched = 0
    skipped = 0

    for i in range(0, len(ids), BATCH_SIZE):
        batch = ids[i:i + BATCH_SIZE]
        id_set = b','.join(batch)
        # BODY.PEEK[] = descarga sin marcar como leído
        typ, msg_data = M.fetch(id_set, '(BODY.PEEK[])')
        if typ != 'OK':
            print(f"⚠️  Error fetch batch {i}-{i+len(batch)}: {msg_data}", file=sys.stderr)
            continue

        for resp in msg_data:
            if not isinstance(resp, tuple) or len(resp) < 2:
                continue
            raw_msg = resp[1]
            try:
                msg = email.message_from_bytes(raw_msg)
            except Exception:
                continue
            subject = decode_subject(msg.get('Subject', ''))
            m = RE_SUBJ.search(subject)
            if not m:
                skipped += 1
                continue
            ref = m.group(1)
            try:
                ts = parsedate_to_datetime(msg.get('Date', '')).isoformat()
            except Exception:
                ts = None
            body = get_body(msg)
            tm = RE_TIPO.search(body)
            um = RE_USR.search(body)
            entry = {
                'timestamp': ts,
                'sheet':     tm.group(1).strip() if tm else None,
                'usuario':   um.group(1).strip() if um else None,
                'xlsx_file': None,
                'pdf_file':  None,
                'source':    f'migrated_from_imap ({folder})',
            }
            # Conserva el más reciente si hay duplicado en el inbox
            if ref in entries:
                cur_ts = entries[ref].get('timestamp') or ''
                if (ts or '') > cur_ts:
                    entries[ref] = entry
            else:
                entries[ref] = entry
                matched += 1

        progress = min(i + BATCH_SIZE, len(ids))
        print(f"   ⏳ Procesados {progress}/{len(ids)}  ({matched} hojas extraídas)")

    print(f"\n📦 Extracción terminada:")
    print(f"   Mensajes totales:       {len(ids)}")
    print(f"   Con #EX{{ref}}# (hojas): {matched}")
    print(f"   Sin patrón hoja:        {skipped}  (otras tools: Lear, Consolidar, etc.)")
    return entries


def merge_into_json(out_path: str, new_entries: dict):
    merged = {}
    if os.path.exists(out_path):
        with open(out_path, 'r', encoding='utf-8') as f:
            merged = json.load(f)
        print(f"\n📂 JSON existente: {len(merged)} entradas previas")

    added = updated = 0
    for ref, entry in new_entries.items():
        if ref in merged:
            cur_ts = merged[ref].get('timestamp') or ''
            new_ts = entry.get('timestamp') or ''
            if new_ts > cur_ts:
                merged[ref] = entry
                updated += 1
        else:
            merged[ref] = entry
            added += 1

    tipos = Counter(v['sheet'] for v in merged.values() if v['sheet'])
    fechas = sorted(v['timestamp'] for v in merged.values() if v['timestamp'])

    print(f"\n📊 Total final: {len(merged)} referencias únicas  (+{added} nuevas, {updated} actualizadas)")
    print(f"Distribución por tipo:")
    for t, n in tipos.most_common():
        print(f"  {n:>4}  {t}")
    sin = sum(1 for v in merged.values() if not v['sheet'])
    if sin:
        print(f"  {sin:>4}  (sin tipo)")
    if fechas:
        print(f"Rango fechas: {fechas[0][:10]}  →  {fechas[-1][:10]}")

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(merged, f, indent=2, ensure_ascii=False)
    print(f"\n✅ Escrito {out_path}")


def main():
    p = argparse.ArgumentParser(
        description="Fetch IMAP y extracción de refs Hoja Control",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument('--list-folders', action='store_true',
                   help='Lista las carpetas IMAP y sale')
    p.add_argument('out_json',  nargs='?', help='Ruta del JSON destino')
    p.add_argument('folder',    nargs='?', help='Nombre exacto de la carpeta IMAP')
    args = p.parse_args()

    M = connect()
    try:
        if args.list_folders:
            list_folders(M)
            return
        if not (args.out_json and args.folder):
            print("❌ Faltan argumentos: out_json folder  (o usa --list-folders)", file=sys.stderr)
            sys.exit(1)
        entries = fetch_and_extract(M, args.folder)
        merge_into_json(args.out_json, entries)
    finally:
        try:
            M.close()
        except Exception:
            pass
        M.logout()
        print("\n👋 Conexión IMAP cerrada")


if __name__ == '__main__':
    main()