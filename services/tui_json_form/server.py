from textual_serve.server import Server

# OJO: este comando se ejecuta *dentro del contenedor*
# Ajusta paths para que coincidan con los mounts que definimos en compose.
CMD = (
    "python3 /apps/tui_json_form/tui_json_form.py "
    "--schema /apps/tui_json_form/eu_customs_clearance_dossier_flat.schema.json "
    "--out /data/forms/out/dossier.draft.json"
)

server = Server(
    CMD,
    host="0.0.0.0",
    port=8000,
    public_url="/tools/json_form",
)
server.serve()