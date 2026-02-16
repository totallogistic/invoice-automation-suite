#!/usr/bin/env python3
from textual_serve.server import Server

server = Server(
    "python3 tui_json_form.py --schema /schemas/schema.json --out /tmp/packing_list.json",
    host="0.0.0.0",
    port=8000
)
server.serve()
