#!/usr/bin/env python3
"""tui_json_form.py (arrays + object navigation)

Keyboard-first TUI to generate JSON from JSON Schema (draft 2020-12).

Supports:
- object navigation (Open group)
- arrays (manager screen): add/remove/open items
  - items of type object: full recursive editing
  - items of type array: recursive editing
  - items of scalar types: basic editing per-row (string/number/int/bool/enum)
- leaf fields: string/integer/number/boolean/enum
- required + basic constraints (min/max, lengths, pattern)
- draft load/save + final save
- final validation using jsonschema Draft202012Validator

Keys:
- Tab / Shift+Tab: focus navigation
- Enter/Space: activate focused button / open
- Esc / Backspace: back
- Ctrl+S: save
- Ctrl+Q: quit
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from jsonschema import Draft202012Validator
from textual import on
from textual.app import App, ComposeResult
from textual.containers import Container, VerticalScroll, Horizontal
from textual.widgets import Button, Footer, Header, Input, Label, Select, Static, Switch

PathSeg = Union[str, int]
PathKey = Tuple[PathSeg, ...]


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def dump_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def is_object_schema(s: Dict[str, Any]) -> bool:
    return s.get("type") == "object" and isinstance(s.get("properties"), dict)


def is_array_schema(s: Dict[str, Any]) -> bool:
    return s.get("type") == "array" and isinstance(s.get("items"), dict)


def schema_title(schema: Dict[str, Any]) -> str:
    return schema.get("title") or schema.get("$id") or "Form"


@dataclass(frozen=True)
class FieldSpec:
    key: str
    schema: Dict[str, Any]
    required: bool


def extract_fields(schema: Dict[str, Any]) -> List[FieldSpec]:
    required = set(schema.get("required", []))
    props = schema.get("properties", {})
    fields: List[FieldSpec] = []
    for key, subschema in props.items():
        fields.append(FieldSpec(key=key, schema=subschema, required=(key in required)))
    fields.sort(key=lambda f: (not f.required, f.key))
    return fields


def field_hint(field_schema: Dict[str, Any]) -> str:
    t = field_schema.get("type", "any")
    parts = [t]
    if "enum" in field_schema:
        parts.append(f"enum={len(field_schema['enum'])}")
    if "minLength" in field_schema:
        parts.append(f"minLen={field_schema['minLength']}")
    if "maxLength" in field_schema:
        parts.append(f"maxLen={field_schema['maxLength']}")
    if "minimum" in field_schema:
        parts.append(f"min={field_schema['minimum']}")
    if "maximum" in field_schema:
        parts.append(f"max={field_schema['maximum']}")
    if "pattern" in field_schema:
        parts.append("pattern")
    return " • ".join(parts)


def validate_locally(field_schema: Dict[str, Any], value: Any) -> Optional[str]:
    t = field_schema.get("type")

    if value is None:
        return None

    if t == "string":
        if not isinstance(value, str):
            return "Expected string."
        if "minLength" in field_schema and len(value) < field_schema["minLength"]:
            return f"Too short (min {field_schema['minLength']})."
        if "maxLength" in field_schema and len(value) > field_schema["maxLength"]:
            return f"Too long (max {field_schema['maxLength']})."
        if "enum" in field_schema and value not in field_schema["enum"]:
            return "Value not in enum."
        if "pattern" in field_schema:
            try:
                if not re.match(field_schema["pattern"], value):
                    return "Does not match pattern."
            except re.error:
                return None

    elif t == "integer":
        if not isinstance(value, int):
            return "Expected integer."
        if "minimum" in field_schema and value < field_schema["minimum"]:
            return f"Too small (min {field_schema['minimum']})."
        if "maximum" in field_schema and value > field_schema["maximum"]:
            return f"Too large (max {field_schema['maximum']})."

    elif t == "number":
        if not isinstance(value, (int, float)):
            return "Expected number."
        if "minimum" in field_schema and value < field_schema["minimum"]:
            return f"Too small (min {field_schema['minimum']})."
        if "maximum" in field_schema and value > field_schema["maximum"]:
            return f"Too large (max {field_schema['maximum']})."

    elif t == "boolean":
        if not isinstance(value, bool):
            return "Expected boolean."

    return None


def coerce_value(field_schema: Dict[str, Any], raw: Any) -> Any:
    t = field_schema.get("type")
    if raw is None:
        return None

    if t == "boolean":
        return bool(raw)
    if t == "integer":
        if raw == "":
            return None
        return int(raw)
    if t == "number":
        if raw == "":
            return None
        return float(raw)

    s = str(raw)
    if s == "":
        return None
    return s


def get_at_path(root: Any, path: PathKey) -> Any:
    cur = root
    for seg in path:
        if isinstance(seg, int):
            if not isinstance(cur, list) or seg < 0 or seg >= len(cur):
                return None
            cur = cur[seg]
        else:
            if not isinstance(cur, dict) or seg not in cur:
                return None
            cur = cur[seg]
    return cur


def ensure_container(parent: Any, seg: PathSeg, next_is_index: bool) -> Any:
    if isinstance(seg, int):
        if not isinstance(parent, list):
            raise TypeError("Parent must be list for int seg")
        while len(parent) <= seg:
            parent.append([] if next_is_index else {})
        return parent[seg]
    else:
        if not isinstance(parent, dict):
            raise TypeError("Parent must be dict for str seg")
        if seg not in parent or parent[seg] is None:
            parent[seg] = [] if next_is_index else {}
        return parent[seg]


def set_at_path(root: Any, path: PathKey, value: Any) -> None:
    if not path:
        raise ValueError("Cannot set root directly")
    cur = root
    for i, seg in enumerate(path[:-1]):
        nxt = path[i + 1]
        cur = ensure_container(cur, seg, next_is_index=isinstance(nxt, int))
    last = path[-1]
    if isinstance(last, int):
        if not isinstance(cur, list):
            raise TypeError("Expected list at parent")
        while len(cur) <= last:
            cur.append(None)
        cur[last] = value
    else:
        if not isinstance(cur, dict):
            raise TypeError("Expected dict at parent")
        cur[last] = value


def delete_at_path(root: Any, path: PathKey) -> None:
    if not path:
        return
    parent = get_at_path(root, path[:-1])
    last = path[-1]
    if isinstance(last, int) and isinstance(parent, list) and 0 <= last < len(parent):
        parent.pop(last)
    elif isinstance(last, str) and isinstance(parent, dict) and last in parent:
        parent.pop(last, None)


def build_default_for_schema(s: Dict[str, Any]) -> Any:
    if "default" in s:
        return s["default"]
    t = s.get("type")
    if t == "object":
        out = {}
        props = s.get("properties", {}) if isinstance(s.get("properties"), dict) else {}
        for k, sub in props.items():
            if isinstance(sub, dict) and "default" in sub:
                out[k] = sub["default"]
        return out
    if t == "array":
        return []
    if t == "boolean":
        return False
    return None


def schema_at_path(root_schema: Dict[str, Any], path: PathKey) -> Dict[str, Any]:
    s = root_schema
    for seg in path:
        if isinstance(seg, int):
            if s.get("type") == "array" and isinstance(s.get("items"), dict):
                s = s["items"]
            else:
                return {}
        else:
            if s.get("type") == "object" and isinstance(s.get("properties"), dict):
                s = s["properties"].get(seg, {})
            else:
                return {}
    return s if isinstance(s, dict) else {}


def compute_missing_required(root_schema: Dict[str, Any], data: Any, obj_path: PathKey) -> List[str]:
    obj_schema = schema_at_path(root_schema, obj_path)
    if not is_object_schema(obj_schema):
        return []
    required = obj_schema.get("required", [])
    props = obj_schema.get("properties", {})
    missing: List[str] = []
    for k in required:
        subschema = props.get(k, {})
        v = get_at_path(data, obj_path + (k,))
        if is_object_schema(subschema):
            if not isinstance(v, dict) or compute_missing_required(root_schema, data, obj_path + (k,)):
                missing.append(k)
        elif is_array_schema(subschema):
            if not isinstance(v, list) or len(v) == 0:
                missing.append(k)
        else:
            if v is None or (isinstance(v, str) and v.strip() == ""):
                missing.append(k)
    return missing


class JsonSchemaFormApp(App):
    CSS = """
    Screen { layout: vertical; }
    .root { height: 1fr; padding: 1 2; }
    .meta { color: $text-muted; }
    .crumb { margin: 0 0 1 0; }

    .row { layout: horizontal; height: auto; margin: 0 0 1 0; }
    .k { width: 28; content-align: right middle; padding: 0 1 0 0; }
    .hint { width: 1fr; color: $text-muted; }

    .err { color: $error; padding-left: 2; }
    .ok { color: $success; padding-left: 2; }

    .actions { height: auto; margin-top: 1; layout: horizontal; }
    .actions Button { margin-right: 1; }

    .groupbtn { width: 1fr; }
    .arrayitem { layout: horizontal; height: auto; margin: 0 0 1 0; }
    .arrayitem Button { margin-right: 1; }
    .idx { width: 8; content-align: right middle; padding-right: 1; color: $text-muted; }
    """

    BINDINGS = [
        ("ctrl+s", "save", "Save JSON"),
        ("ctrl+q", "quit", "Quit"),
        ("escape", "back", "Back"),
        ("backspace", "back", "Back"),
        ("enter", "open", "Open"),
        ("space", "open", "Open"),
    ]

    def __init__(self, schema_path: Path, out_path: Path, draft_path: Optional[Path] = None):
        super().__init__()
        self.schema_path = schema_path
        self.out_path = out_path
        self.draft_path = draft_path

        self.schema: Dict[str, Any] = {}
        self.validator: Optional[Draft202012Validator] = None

        self.data: Any = {}
        self.nav_path: List[PathSeg] = []

        self.widgets_by_path: Dict[PathKey, Any] = {}
        self.error_by_path: Dict[PathKey, Label] = {}

        self.scroll: Optional[VerticalScroll] = None
        self.status: Optional[Label] = None
        self.crumb: Optional[Static] = None

    def compose(self) -> ComposeResult:
        yield Header()
        with Container(classes="root"):
            yield Static(f"Schema: {self.schema_path.name}", classes="meta")
            yield Static("Keyboard: Tab/Shift+Tab • Enter/Space open • Esc/Backspace back • Ctrl+S save • Ctrl+Q quit", classes="meta")
            self.crumb = Static("", classes="crumb")
            yield self.crumb

            self.scroll = VerticalScroll(id="form_scroll")
            yield self.scroll

            with Container(classes="actions"):
                yield Button("Back (Esc)", id="btn_back")
                yield Button("Save (Ctrl+S)", id="btn_save", variant="success")
                yield Button("Quit (Ctrl+Q)", id="btn_quit", variant="error")

            self.status = Label("")
            yield self.status

        yield Footer()

    def on_mount(self) -> None:
        self.schema = load_json(self.schema_path)
        self.validator = Draft202012Validator(self.schema)

        if self.draft_path and self.draft_path.exists():
            try:
                self.data = load_json(self.draft_path)
            except Exception:
                self.data = {}
        else:
            self.data = {}

        if not isinstance(self.data, dict):
            self.data = {}

        self.nav_path = []
        self._render_current()

    def _location_str(self) -> str:
        if not self.nav_path:
            return "<root>"
        parts = []
        for seg in self.nav_path:
            if isinstance(seg, int):
                parts.append(f"[{seg}]")
            else:
                parts.append(seg)
        return "/".join(parts)

    def _crumb_text(self) -> str:
        title = schema_title(self.schema)
        cur_schema = schema_at_path(self.schema, tuple(self.nav_path))
        t = cur_schema.get("type", "unknown")
        return f"{title}  •  Location: {self._location_str()}  •  {t}"

    def _clear_scroll(self) -> None:
        assert self.scroll is not None
        for child in list(self.scroll.children):
            child.remove()

    def action_back(self) -> None:
        if not self.nav_path:
            self.action_quit()
            return
        self._commit_visible_fields_to_data()
        self.nav_path.pop()
        self._render_current()

    def action_quit(self) -> None:
        self._commit_visible_fields_to_data()
        self._write_draft()
        self.exit()

    def action_save(self) -> None:
        self._commit_visible_fields_to_data()
        self._save()

    def action_open(self) -> None:
        focused = self.focused
        if isinstance(focused, Button) and focused.id:
            path = self._parse_nav_btn_id(focused.id)
            if path is not None:
                self._open_path(path)
                return
        assert self.scroll is not None
        for btn in self.scroll.query(Button):
            if btn.id:
                path = self._parse_nav_btn_id(btn.id)
                if path is not None:
                    self._open_path(path)
                    return

    def _open_path(self, path: PathKey) -> None:
        self._commit_visible_fields_to_data()
        self.nav_path = list(path)
        self._render_current()

    def _nav_btn_id(self, path: PathKey) -> str:
        enc = []
        for seg in path:
            if isinstance(seg, int):
                enc.append(f"#{seg}")
            else:
                enc.append(seg)
        return "nav-" + "__".join(enc)

    def _parse_nav_btn_id(self, widget_id: str) -> Optional[PathKey]:
        if not widget_id.startswith("nav-"):
            return None
        rest = widget_id[4:]
        if not rest:
            return None
        parts = rest.split("__")
        out: List[PathSeg] = []
        for p in parts:
            if p.startswith("#"):
                try:
                    out.append(int(p[1:]))
                except ValueError:
                    return None
            else:
                out.append(p)
        return tuple(out)

    def _render_current(self) -> None:
        if self.crumb:
            self.crumb.update(self._crumb_text())
        self.widgets_by_path.clear()
        self.error_by_path.clear()
        self._clear_scroll()

        cur_schema = schema_at_path(self.schema, tuple(self.nav_path))

        if is_array_schema(cur_schema):
            self._render_array(cur_schema)
        elif is_object_schema(cur_schema):
            self._render_object(cur_schema)
        else:
            assert self.scroll is not None
            self.scroll.mount(Label("Current location is not an object/array.", classes="err"))

        self._update_status()

    def _render_object(self, obj_schema: Dict[str, Any]) -> None:
        assert self.scroll is not None
        fields = extract_fields(obj_schema)

        for f in fields:
            subschema = f.schema
            k_label = f"{f.key}{' *' if f.required else ''}"
            row = Container(classes="row")
            self.scroll.mount(row)
            row.mount(Label(k_label, classes="k"))
            row.mount(Static(field_hint(subschema), classes="hint"))

            pk = tuple(self.nav_path + [f.key])

            if is_object_schema(subschema) or is_array_schema(subschema):
                if is_object_schema(subschema):
                    missing = compute_missing_required(self.schema, self.data, pk)
                    badge = "OK" if not missing else f"missing: {', '.join(missing)}"
                else:
                    arr = get_at_path(self.data, pk)
                    n = len(arr) if isinstance(arr, list) else 0
                    badge = f"{n} item(s)"
                btn = Button(f"Open {f.key}  [{badge}]", id=self._nav_btn_id(pk), classes="groupbtn")
                self.scroll.mount(btn)
                err = Label("", classes="err")
                self.scroll.mount(err)
                self.error_by_path[pk] = err
                continue

            cur_val = get_at_path(self.data, pk)
            default = subschema.get("default")
            if cur_val is None and default is not None:
                cur_val = default
                set_at_path(self.data, pk, cur_val)

            w = self._make_leaf_widget(subschema, cur_val)
            self.widgets_by_path[pk] = w
            self.scroll.mount(w)
            err = Label("", classes="err")
            self.error_by_path[pk] = err
            self.scroll.mount(err)
            self._validate_leaf(pk, subschema, f.required)

        for btn in self.scroll.query(Button):
            if btn.id and self._parse_nav_btn_id(btn.id):
                btn.focus()
                return
        for w in self.scroll.query(Input, Select, Switch):
            w.focus()
            return

    def _render_array(self, arr_schema: Dict[str, Any]) -> None:
        assert self.scroll is not None
        items_schema = arr_schema.get("items", {})
        pk = tuple(self.nav_path)

        cur_arr = get_at_path(self.data, pk)
        if cur_arr is None or not isinstance(cur_arr, list):
            cur_arr = []
            if pk:
                set_at_path(self.data, pk, cur_arr)
            else:
                self.data = cur_arr  # unlikely in your schemas

        controls = Horizontal()
        self.scroll.mount(controls)
        controls.mount(Button("Add item (+)", id="btn_add_item", variant="success"))
        controls.mount(Button("Delete last (-)", id="btn_del_last", variant="error"))
        self.scroll.mount(Static(f"Items: {len(cur_arr)}", classes="meta"))

        for idx, _ in enumerate(cur_arr):
            row = Container(classes="arrayitem")
            self.scroll.mount(row)
            row.mount(Label(f"[{idx}]", classes="idx"))
            item_path = tuple(self.nav_path + [idx])

            if is_object_schema(items_schema) or is_array_schema(items_schema):
                if is_object_schema(items_schema):
                    missing = compute_missing_required(self.schema, self.data, item_path)
                    badge = "OK" if not missing else f"missing: {', '.join(missing)}"
                else:
                    v = get_at_path(self.data, item_path)
                    badge = f"{len(v) if isinstance(v, list) else 0} item(s)"
                row.mount(Button(f"Open item {idx}  [{badge}]", id=self._nav_btn_id(item_path)))
            else:
                val = get_at_path(self.data, item_path)
                w = self._make_leaf_widget(items_schema, val)
                self.widgets_by_path[item_path] = w
                row.mount(w)
                err = Label("", classes="err")
                self.error_by_path[item_path] = err
                row.mount(err)
                self._validate_leaf(item_path, items_schema, required=False)

        try:
            self.query_one("#btn_add_item", Button).focus()
        except Exception:
            pass

    def _make_leaf_widget(self, schema: Dict[str, Any], cur_val: Any):
        t = schema.get("type")
        if t == "boolean":
            return Switch(value=bool(cur_val) if cur_val is not None else bool(schema.get("default", False)))
        if "enum" in schema and isinstance(schema["enum"], list):
            options = [(str(v), str(v)) for v in schema["enum"]]
            initial = str(cur_val) if cur_val is not None else (options[0][0] if options else "")
            return Select(options=options, value=initial)
        placeholder = "" if schema.get("default") is None else str(schema.get("default"))
        return Input(value="" if cur_val is None else str(cur_val), placeholder=placeholder)

    def _validate_leaf(self, pk: PathKey, schema: Dict[str, Any], required: bool) -> None:
        w = self.widgets_by_path.get(pk)
        if w is None:
            return
        raw = w.value if not isinstance(w, Switch) else w.value
        try:
            coerced = coerce_value(schema, raw)
        except Exception:
            coerced = "__COERCE_ERROR__"
        if required and coerced is None:
            msg = "Required."
        elif coerced == "__COERCE_ERROR__":
            msg = "Invalid value."
        else:
            msg = validate_locally(schema, coerced) or ""
        if pk in self.error_by_path:
            self.error_by_path[pk].update(msg)

    def _commit_visible_fields_to_data(self) -> None:
        for pk, w in list(self.widgets_by_path.items()):
            s = schema_at_path(self.schema, pk)
            raw = w.value if not isinstance(w, Switch) else w.value
            try:
                val = coerce_value(s, raw)
            except Exception:
                continue
            if val is None:
                if pk and isinstance(pk[-1], str):
                    delete_at_path(self.data, pk)
                else:
                    set_at_path(self.data, pk, None)
            else:
                set_at_path(self.data, pk, val)

    def _update_status(self) -> None:
        if not self.status:
            return
        self.status.update("OK")
        self.status.set_class(True, "ok")

    def _write_draft(self) -> None:
        if not self.draft_path:
            return
        try:
            dump_json(self.draft_path, self.data)
        except Exception as e:
            if self.status:
                self.status.update(f"Draft write failed: {e}")
                self.status.set_class(False, "err")

    def _save(self) -> None:
        assert self.validator is not None
        errors = sorted(self.validator.iter_errors(self.data), key=lambda e: list(e.path))
        if errors:
            e0 = errors[0]
            path = ".".join([str(p) for p in e0.path]) if e0.path else "<root>"
            if self.status:
                self.status.update(f"Schema validation failed at {path}: {e0.message}")
                self.status.set_class(False, "err")
            return
        try:
            dump_json(self.out_path, self.data)
            if self.status:
                self.status.update(f"Saved: {self.out_path}")
                self.status.set_class(True, "ok")
            self._write_draft()
        except Exception as e:
            if self.status:
                self.status.update(f"Save failed: {e}")
                self.status.set_class(False, "err")

    @on(Button.Pressed)
    def _on_button_pressed(self, event: Button.Pressed) -> None:
        bid = event.button.id or ""
        if bid == "btn_save":
            self.action_save(); return
        if bid == "btn_quit":
            self.action_quit(); return
        if bid == "btn_back":
            self.action_back(); return
        if bid == "btn_add_item":
            self._add_array_item(); return
        if bid == "btn_del_last":
            self._del_array_last(); return
        path = self._parse_nav_btn_id(bid)
        if path is not None:
            self._open_path(path)

    @on(Input.Changed)
    def _on_input_changed(self, event: Input.Changed) -> None:
        pk = self._find_path_for_widget(event.input)
        if pk is None: return
        s = schema_at_path(self.schema, pk)
        self._validate_leaf(pk, s, self._is_required_leaf(pk))

    @on(Select.Changed)
    def _on_select_changed(self, event: Select.Changed) -> None:
        pk = self._find_path_for_widget(event.select)
        if pk is None: return
        s = schema_at_path(self.schema, pk)
        self._validate_leaf(pk, s, self._is_required_leaf(pk))

    @on(Switch.Changed)
    def _on_switch_changed(self, event: Switch.Changed) -> None:
        pk = self._find_path_for_widget(event.switch)
        if pk is None: return
        s = schema_at_path(self.schema, pk)
        self._validate_leaf(pk, s, self._is_required_leaf(pk))

    def _find_path_for_widget(self, widget: Any) -> Optional[PathKey]:
        for pk, w in self.widgets_by_path.items():
            if w is widget:
                return pk
        return None

    def _is_required_leaf(self, pk: PathKey) -> bool:
        if not pk: return False
        if isinstance(pk[-1], int): return False
        parent_schema = schema_at_path(self.schema, pk[:-1])
        if not is_object_schema(parent_schema): return False
        return pk[-1] in set(parent_schema.get("required", []))

    def _add_array_item(self) -> None:
        cur_schema = schema_at_path(self.schema, tuple(self.nav_path))
        if not is_array_schema(cur_schema): return
        items_schema = cur_schema.get("items", {})
        arr = get_at_path(self.data, tuple(self.nav_path))
        if not isinstance(arr, list):
            arr = []
            set_at_path(self.data, tuple(self.nav_path), arr)
        arr.append(build_default_for_schema(items_schema))
        self._render_current()

    def _del_array_last(self) -> None:
        arr = get_at_path(self.data, tuple(self.nav_path))
        if isinstance(arr, list) and arr:
            arr.pop()
            self._render_current()


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--schema", required=True, type=Path)
    p.add_argument("--out", required=True, type=Path)
    p.add_argument("--draft", type=Path, default=None)
    args = p.parse_args()
    app = JsonSchemaFormApp(schema_path=args.schema, out_path=args.out, draft_path=args.draft)
    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

