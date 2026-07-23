"""Tool registry - loads and manages tool configurations."""
from __future__ import annotations

import yaml
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional


@dataclass
class ToolConfig:
    """Configuration for a single tool."""
    name: str
    display_name: str
    description: str
    input_formats: List[str]
    inbox_dir: Path
    status_dir: Path
    output_dir: Path
    output_artifacts: List[str]
    extractor_path: Path
    email_subject_template: str
    max_file_size_mb: int = 200
    enabled: bool = True
    # ── Inject mode (e.g. Croton) ────────────────────────────────────────────
    # When True, the extractor receives:
    #   - the primary ODS as the first positional argument
    #   - the XLSX invoice via --factura <path>
    #   - --inject  (appends Resumen_Partidas sheet to the ODS instead of
    #                writing a separate output file)
    # input_formats must include both accepted extensions, e.g. ["ods", "xlsx"]
    inject_mode: bool = False
    filename_pattern: str = ""   # e.g. "CROTON_{stem}.ods"
    # ── Camion mode ──────────────────────────────────────────────────────────
    # When True, the extractor receives named file arguments:
    #   --xlsx <packing_list.xlsx>
    #   --t1   <T1_1.pdf> [<T1_2.pdf> ...]
    #   --doc  <doc.pdf>
    #   -o     <output_dir>
    # T1 files are identified by "t1" in the filename; the DOC file by "doc".
    camion_mode: bool = False
    croton_import_mode: bool = False
    bl_mode: bool = False
    export_visual_mode: bool = False
    email_mail_from: Optional[str] = None   # override MAIL_FROM global para esta tool
    email_mode_send_email: Optional[str] = None  # 'postfix' | 'custom' (override del MAIL_SEND_MODE global)

class ToolRegistry:
    """Registry of all available tools."""
    
    def __init__(self, tools: List[ToolConfig]):
        self.tools = tools
        self._tools_by_name = {t.name: t for t in tools}
    
    @classmethod
    def from_yaml(cls, config_path: Path | str) -> ToolRegistry:
        """Load tool registry from YAML."""
        config_path = Path(config_path)
        
        if not config_path.exists():
            raise FileNotFoundError(f"Config not found: {config_path}")
        
        with open(config_path) as f:
            config_data = yaml.safe_load(f)
        
        tools = []
        data_root = Path(config_data.get("data_root", "/data"))
        
        for tool_data in config_data.get("tools", []):
            if not tool_data.get("enabled", True):
                continue
            
            tool_name = tool_data["name"]
            tool_root = data_root / tool_name
            
            tool = ToolConfig(
                name=tool_name,
                display_name=tool_data.get("display_name", tool_name),
                description=tool_data.get("description", ""),
                input_formats=tool_data.get("input", {}).get("formats", ["pdf"]),
                inbox_dir=Path(tool_data.get("input", {}).get("inbox", tool_root / "inbox")),
                status_dir=Path(tool_data.get("status_dir", tool_root / "status")),
                output_dir=Path(tool_data.get("output", {}).get("directory", tool_root / "out")),
                output_artifacts=tool_data.get("output", {}).get("artifacts", []),
                extractor_path=Path(tool_data.get("extractor", {}).get("path", "")),
                email_subject_template=tool_data.get("email", {}).get(
                    "subject_template",
                    f"[{tool_name}] Batch {{batch_id}} processed"
                ),
                max_file_size_mb=tool_data.get("max_file_size_mb", 200),
                enabled=True,
                inject_mode=tool_data.get("extractor", {}).get("inject_mode", False),
                filename_pattern=tool_data.get("output", {}).get("filename_pattern", ""),
                camion_mode=tool_data.get("extractor", {}).get("camion_mode", False),
                croton_import_mode=tool_data.get("extractor", {}).get("croton_import_mode", False),
                bl_mode=tool_data.get("extractor", {}).get("bl_mode", False),
                export_visual_mode=tool_data.get("extractor", {}).get("export_visual_mode", False),
                email_mail_from=tool_data.get("email", {}).get("mail_from") or None,
                email_mode_send_email=tool_data.get("email", {}).get("mode_send_email") or None,
            )
            
            # Ensure directories exist
            tool.inbox_dir.mkdir(parents=True, exist_ok=True)
            tool.status_dir.mkdir(parents=True, exist_ok=True)
            tool.output_dir.mkdir(parents=True, exist_ok=True)
            
            tools.append(tool)
        
        return cls(tools)
    
    def get_tool(self, name: str) -> Optional[ToolConfig]:
        """Get tool by name."""
        return self._tools_by_name.get(name)
    
    def list_tools(self) -> List[str]:
        """List all tool names."""
        return list(self._tools_by_name.keys())