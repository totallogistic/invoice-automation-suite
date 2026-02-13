#!/usr/bin/env python3
"""
TUI JSON Form Builder for EU Customs Clearance Dossier

This application provides a text-based user interface (TUI) for creating
JSON files based on a schema. It uses the Textual library to create an
interactive form in the terminal/browser.
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict

from textual.app import App, ComposeResult
from textual.containers import Container, Horizontal, Vertical, VerticalScroll
from textual.widgets import Button, Footer, Header, Input, Label, Select, Static
from textual.binding import Binding


class JSONFormApp(App):
    """A Textual app to create JSON files from a schema."""
    
    CSS = """
    Screen {
        background: $surface;
    }
    
    #title {
        background: $primary;
        color: $text;
        text-align: center;
        padding: 1;
        text-style: bold;
    }
    
    #form-container {
        background: $panel;
        border: tall $primary;
        padding: 1 2;
        margin: 1 2;
        height: auto;
    }
    
    .field-group {
        layout: vertical;
        margin: 1 0;
        height: auto;
    }
    
    .field-label {
        color: $accent;
        text-style: bold;
        margin-bottom: 1;
    }
    
    .field-input {
        margin-bottom: 1;
    }
    
    .field-description {
        color: $text-muted;
        text-style: italic;
        margin-bottom: 1;
    }
    
    #button-container {
        layout: horizontal;
        align: center middle;
        padding: 1;
        height: auto;
    }
    
    Button {
        margin: 0 1;
    }
    
    #status {
        background: $panel;
        color: $warning;
        padding: 1;
        text-align: center;
    }
    """
    
    BINDINGS = [
        Binding("ctrl+s", "save", "Save", priority=True),
        Binding("ctrl+q", "quit", "Quit", priority=True),
    ]
    
    def __init__(self, schema_path: str, output_path: str):
        super().__init__()
        self.schema_path = Path(schema_path)
        self.output_path = Path(output_path)
        self.schema: Dict[str, Any] = {}
        self.field_widgets: Dict[str, Input] = {}
        self.load_schema()
        
    def load_schema(self):
        """Load the JSON schema from file."""
        try:
            with open(self.schema_path, 'r', encoding='utf-8') as f:
                self.schema = json.load(f)
        except Exception as e:
            print(f"Error loading schema: {e}", file=sys.stderr)
            sys.exit(1)
    
    def compose(self) -> ComposeResult:
        """Create child widgets for the app."""
        yield Header()
        
        # Title
        title = self.schema.get('title', 'JSON Form')
        yield Static(title, id="title")
        
        # Form container with scroll
        with VerticalScroll(id="form-container"):
            properties = self.schema.get('properties', {})
            required_fields = self.schema.get('required', [])
            
            for field_name, field_spec in properties.items():
                field_title = field_spec.get('title', field_name)
                field_desc = field_spec.get('description', '')
                field_type = field_spec.get('type', 'string')
                is_required = field_name in required_fields
                
                # Mark required fields
                label_text = f"{field_title}" + (" *" if is_required else "")
                
                with Container(classes="field-group"):
                    yield Label(label_text, classes="field-label")
                    
                    if field_desc:
                        yield Static(field_desc, classes="field-description")
                    
                    # Handle enum fields with Select widget
                    if 'enum' in field_spec:
                        options = [(str(opt), str(opt)) for opt in field_spec['enum']]
                        default = field_spec.get('default', '')
                        select_widget = Select(
                            options,
                            value=default if default else None,
                            id=f"field_{field_name}",
                            classes="field-input"
                        )
                        self.field_widgets[field_name] = select_widget
                        yield select_widget
                    else:
                        # Regular input field
                        default_value = str(field_spec.get('default', ''))
                        input_widget = Input(
                            value=default_value,
                            placeholder=f"Enter {field_title.lower()}",
                            id=f"field_{field_name}",
                            classes="field-input"
                        )
                        self.field_widgets[field_name] = input_widget
                        yield input_widget
        
        # Buttons
        with Horizontal(id="button-container"):
            yield Button("Save (Ctrl+S)", variant="primary", id="save-btn")
            yield Button("Clear Form", variant="warning", id="clear-btn")
            yield Button("Quit (Ctrl+Q)", variant="error", id="quit-btn")
        
        # Status message
        yield Static("", id="status")
        
        yield Footer()
    
    def on_button_pressed(self, event: Button.Pressed) -> None:
        """Handle button press events."""
        if event.button.id == "save-btn":
            self.action_save()
        elif event.button.id == "clear-btn":
            self.clear_form()
        elif event.button.id == "quit-btn":
            self.action_quit()
    
    def action_save(self) -> None:
        """Save the form data to JSON file."""
        data = self.collect_form_data()
        
        # Validate required fields
        required_fields = self.schema.get('required', [])
        missing_fields = [f for f in required_fields if not data.get(f)]
        
        if missing_fields:
            status = self.query_one("#status", Static)
            status.update(f"❌ Missing required fields: {', '.join(missing_fields)}")
            return
        
        # Save to file
        try:
            # Ensure output directory exists
            self.output_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.output_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            
            status = self.query_one("#status", Static)
            status.update(f"✅ Form saved successfully to {self.output_path}")
        except Exception as e:
            status = self.query_one("#status", Static)
            status.update(f"❌ Error saving file: {e}")
    
    def clear_form(self) -> None:
        """Clear all form fields."""
        for widget in self.field_widgets.values():
            if isinstance(widget, Input):
                widget.value = ""
            elif isinstance(widget, Select):
                widget.value = Select.BLANK
        
        status = self.query_one("#status", Static)
        status.update("🔄 Form cleared")
    
    def collect_form_data(self) -> Dict[str, Any]:
        """Collect data from all form fields."""
        data = {}
        properties = self.schema.get('properties', {})
        
        for field_name, widget in self.field_widgets.items():
            field_spec = properties.get(field_name, {})
            field_type = field_spec.get('type', 'string')
            
            # Get the value from the widget
            if isinstance(widget, Select):
                value = widget.value
                if value == Select.BLANK:
                    value = None
            else:
                value = widget.value
            
            # Skip empty values
            if not value:
                continue
            
            # Type conversion
            try:
                if field_type == 'number':
                    data[field_name] = float(value)
                elif field_type == 'integer':
                    data[field_name] = int(value)
                elif field_type == 'boolean':
                    data[field_name] = value.lower() in ('true', '1', 'yes')
                else:
                    data[field_name] = value
            except (ValueError, AttributeError):
                # If conversion fails, store as string
                data[field_name] = value
        
        return data


def main():
    """Main entry point for the application."""
    parser = argparse.ArgumentParser(
        description='TUI JSON Form Builder for EU Customs Clearance Dossier'
    )
    parser.add_argument(
        '--schema',
        required=True,
        help='Path to the JSON schema file'
    )
    parser.add_argument(
        '--out',
        required=True,
        help='Path to save the output JSON file'
    )
    
    args = parser.parse_args()
    
    app = JSONFormApp(args.schema, args.out)
    app.run()


if __name__ == '__main__':
    main()
