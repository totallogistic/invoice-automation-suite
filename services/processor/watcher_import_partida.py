from __future__ import annotations

import json
import os
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict


SERVICE_ROOT = Path(os.getenv("SERVICE_ROOT", "/data"))
INBOX_DIR = Path(os.getenv("INBOX_DIR", str(SERVICE_ROOT / "import_partida" / "inbox")))
STATUS_DIR = Path(os.getenv("STATUS_DIR", str(SERVICE_ROOT / "import_partida" / "status")))

POLL_SECONDS = int(os.getenv("POLL_SECONDS", "3"))
BATCH_QUIET_SECONDS = int(os.getenv("BATCH_QUIET_SECONDS", "5"))  # corto para dev

ALLOWED_EXT = {".pdf"}


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def read_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def atomic_write_json(path: Path, data: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def update_status(batch_id: str, **fields: Any) -> None:
    status_file = STATUS_DIR / batch_id / "status.json"
    st: Dict[str, Any]
    if status_file.exists():
        st = read_json(status_file)
    else:
        st = {"batch_id": batch_id}

    st.update(fields)
    st["updated_at"] = now_iso()
    atomic_write_json(status_file, st)


def is_quiet(dirpath: Path, quiet_seconds: int) -> bool:
    # evita pillar uploads a medio escribir
    latest = 0.0
    for p in dirpath.rglob("*"):
        if p.is_file():
            latest = max(latest, p.stat().st_mtime)
    if latest == 0.0:
        return False
    return (time.time() - latest) >= quiet_seconds


def main() -> None:
    INBOX_DIR.mkdir(parents=True, exist_ok=True)
    STATUS_DIR.mkdir(parents=True, exist_ok=True)

    print(f"[INFO] [import_partida] SERVICE_ROOT={SERVICE_ROOT}")
    print(f"[INFO] [import_partida] INBOX_DIR={INBOX_DIR}")
    print(f"[INFO] [import_partida] STATUS_DIR={STATUS_DIR}")
    print(f"[INFO] [import_partida] POLL_SECONDS={POLL_SECONDS} | BATCH_QUIET_SECONDS={BATCH_QUIET_SECONDS}")

    while True:
        try:
            batch_dirs = sorted([p for p in INBOX_DIR.iterdir() if p.is_dir()])
            for bdir in batch_dirs:
                batch_id = bdir.name
                status_file = STATUS_DIR / batch_id / "status.json"

                # Solo procesar si hay status y está en WAITING/UPLOADED
                if not status_file.exists():
                    continue

                st = read_json(status_file)
                state = (st.get("state") or "").upper()
                stage = (st.get("stage") or "").upper()

                if state in {"DONE", "ERROR"}:
                    continue
                if stage not in {"WAITING"}:
                    continue

                if not is_quiet(bdir, BATCH_QUIET_SECONDS):
                    continue

                pdfs = [p for p in bdir.iterdir() if p.is_file() and p.suffix.lower() in ALLOWED_EXT]
                if len(pdfs) != 1:
                    update_status(
                        batch_id,
                        state="ERROR",
                        stage="ERROR",
                        message=f"Se esperaba exactamente 1 PDF en el batch, pero hay {len(pdfs)}.",
                    )
                    continue

                # Marcamos EXTRACTING (stub)
                update_status(batch_id, state="RUNNING", stage="EXTRACTING", message="Procesando (stub watcher)...")

                # STUB: aquí luego llamaremos al extractor y generaremos el XLSX
                time.sleep(0.5)

                # Marcamos DONE (stub)
                update_status(
                    batch_id,
                    state="DONE",
                    stage="DONE",
                    processed_files=1,
                    message="Procesado OK (stub). Próximo paso: generar XLSX y email.",
                )

        except Exception as e:
            print(f"[ERROR] [import_partida] loop error: {e}")

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()