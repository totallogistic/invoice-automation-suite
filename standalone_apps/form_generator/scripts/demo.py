"""Genera un DeCA de ejemplo sin levantar el servidor.

Uso:
    cd deca-prototype
    python -m scripts.demo

Produce el PDF (en modo local, bajo _deca_store/d/<uuid>.pdf) y lo copia a
./ejemplo-deca.pdf para inspección. Si defines DECA_S3_* en el entorno, lo sube
al bucket real en vez de a disco.
"""

from __future__ import annotations

import datetime as dt
import pathlib
import shutil

from deca.models import Cargador, DecaInput, Mercancia, Transportista
from deca.service import DecaService


def main() -> None:
    data = DecaInput(
        cargador_contractual=Cargador(
            nombre="Envases del Sur, S.L.",
            nif="B11223344",
            domicilio="Pol. Ind. Las Salinas, 12 — 11500 El Puerto de Santa María (Cádiz)",
        ),
        transportista_efectivo=Transportista(
            nombre="Total Logistic Services, S.L.",
            nif="B11000111",
        ),
        origen="El Puerto de Santa María (Cádiz)",
        destino="Getafe (Madrid)",
        mercancia=Mercancia(naturaleza="Envases de vidrio paletizados", peso_kg=18450),
        fecha_transporte=dt.date(2026, 10, 6),
        matricula_tractora="1234-JKL",
        matricula_remolque="R-5678-BCD",
        autorizaciones_especiales=None,
        observaciones="Mercancía frágil. 22 palets. Entrega en muelle 4, cita 09:30.",
        servicio_inicio=dt.datetime(2026, 10, 6, 7, 0, tzinfo=dt.timezone.utc),
        servicio_fin=dt.datetime(2026, 10, 6, 15, 0, tzinfo=dt.timezone.utc),
    )

    svc = DecaService()
    rec = svc.create(data)

    print(f"DeCA creado")
    print(f"  UUID            : {rec.uuid}")
    print(f"  URL pública     : {rec.url_publica}")
    print(f"  Modo storage    : {svc.storage.mode}")
    print(f"  URL activa hasta: {rec.url_activa_hasta}")

    # Copia para inspección
    src = pathlib.Path("_deca_store") / rec.object_key()
    if src.exists():
        shutil.copy(src, "ejemplo-deca.pdf")
        print(f"  PDF de ejemplo  : ejemplo-deca.pdf ({src.stat().st_size/1024:.0f} KB)")


if __name__ == "__main__":
    main()
