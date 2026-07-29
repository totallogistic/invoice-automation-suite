"""Tool registry - loads and manages tool configurations."""
from __future__ import annotations

import yaml
from dataclasses import dataclass
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
    email_mail_from: Optional[str] = None   # override MAIL_FROM global para esta tool
    email_mode_send_email: Optional[str] = None  # 'postfix' | 'custom' (override del MAIL_SEND_MODE global)
    version_file: Optional[Path] = None      # fichero de versión si difiere de extractor.path (p.ej. intrastat)
    max_file_size_mb: int = 200
    enabled: bool = True


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
                version_file=(
                    Path(tool_data["extractor"]["version_file"])
                    if tool_data.get("extractor", {}).get("version_file") else None
                ),
                email_subject_template=tool_data.get("email", {}).get(
                    "subject_template",
                    f"[{tool_name}] Batch {{batch_id}} processed"
                ),
                email_mail_from=tool_data.get("email", {}).get("mail_from") or None,
                email_mode_send_email=tool_data.get("email", {}).get("mode_send_email") or None,
                max_file_size_mb=tool_data.get("max_file_size_mb", 200),
                enabled=True
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
