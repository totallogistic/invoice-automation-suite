"""Unified processor - handles all tools."""
from __future__ import annotations

import logging
import os
import subprocess
import time
from pathlib import Path
from typing import List

from iasuite_common.batch import BatchOperations
from iasuite_common.email import EmailService, EmailConfig
from iasuite_common.status import StatusManager
from tool_registry import ToolRegistry, ToolConfig

logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s"
)
logger = logging.getLogger("processor")


class UnifiedProcessor:
    """Multi-tool processor."""
    
    def __init__(
        self,
        registry: ToolRegistry,
        email_service: EmailService,
        poll_seconds: int = 3,
        batch_quiet_seconds: int = 180
    ):
        self.registry = registry
        self.email_service = email_service
        self.poll_seconds = poll_seconds
        self.batch_quiet_seconds = batch_quiet_seconds
        
        # Setup directories
        self.processing_dirs = {}
        self.processed_dirs = {}
        self.error_dirs = {}
        
        for tool in registry.tools:
            tool_root = tool.inbox_dir.parent
            self.processing_dirs[tool.name] = tool_root / "processing"
            self.processed_dirs[tool.name] = tool_root / "processed"
            self.error_dirs[tool.name] = tool_root / "error"
            
            self.processing_dirs[tool.name].mkdir(parents=True, exist_ok=True)
            self.processed_dirs[tool.name].mkdir(parents=True, exist_ok=True)
            self.error_dirs[tool.name].mkdir(parents=True, exist_ok=True)
        
        logger.info(f"Initialized for {len(registry.tools)} tools")
    
    def run(self):
        """Main loop."""
        logger.info("Starting unified processor...")
        logger.info(f"Tools: {self.registry.list_tools()}")
        
        while True:
            try:
                for tool in self.registry.tools:
                    self._process_tool(tool)
            except Exception as e:
                logger.exception(f"Loop error: {e}")
            
            time.sleep(self.poll_seconds)
    
    def _process_tool(self, tool: ToolConfig):
        """Process batches for a tool."""
        batches = BatchOperations.find_batches(tool.inbox_dir)
        
        for batch_dir in batches:
            batch_id = batch_dir.name
            
            # Initialize status if needed
            status_mgr = StatusManager(tool.status_dir)
            if not status_mgr.get_status(batch_id):
                # Count files with the tool's accepted formats
                file_count = len(BatchOperations.get_files(batch_dir, tool.input_formats))
                status_mgr.create_status(
                    batch_id=batch_id,
                    total_files=file_count,
                    message="Waiting"
                )
            
            # Check if ready
            if not BatchOperations.is_batch_ready(
                batch_dir,
                quiet_seconds=self.batch_quiet_seconds
            ):
                continue
            
            logger.info(f"[{tool.name}] Processing: {batch_id}")
            
            processing_path = BatchOperations.move_batch(
                batch_dir,
                self.processing_dirs[tool.name]
            )
            
            try:
                self._process_batch(tool, batch_id, processing_path)
                BatchOperations.move_batch(
                    processing_path,
                    self.processed_dirs[tool.name]
                )
                logger.info(f"[{tool.name}] ✓ Done: {batch_id}")
            
            except Exception as e:
                logger.error(f"[{tool.name}] ✗ Failed: {batch_id} - {e}")
                status_mgr.update_status(
                    batch_id=batch_id,
                    state="ERROR",
                    stage="ERROR",
                    message=str(e)
                )
                BatchOperations.move_batch(
                    processing_path,
                    self.error_dirs[tool.name]
                )
    
    def _process_batch(self, tool: ToolConfig, batch_id: str, processing_path: Path):
        """Process a single batch."""
        status_mgr = StatusManager(tool.status_dir)
        
        # Update: processing
        file_count = len(BatchOperations.get_files(processing_path, tool.input_formats))
        status_mgr.update_status(
            batch_id=batch_id,
            state="PROCESSING",
            stage="RUNNING_EXTRACTOR",
            message="Running extractor"
        )
        
        # Run extractor
        output_path = self._run_extractor(tool, batch_id, processing_path)
        
        # Collect artifacts - support glob patterns (e.g., COMPLETADO_*.ods)
        artifacts = []
        for artifact_pattern in tool.output_artifacts:
            if '*' in artifact_pattern or '?' in artifact_pattern:
                # Use glob for wildcard patterns
                matched = list(output_path.glob(artifact_pattern))
                artifacts.extend(matched)
                logger.info(f"[{tool.name}] Pattern '{artifact_pattern}' matched {len(matched)} file(s)")
            else:
                # Exact filename match
                artifact_path = output_path / artifact_pattern
                if artifact_path.exists():
                    artifacts.append(artifact_path)
                    logger.info(f"[{tool.name}] Found artifact: {artifact_pattern}")
                else:
                    logger.warning(f"[{tool.name}] Artifact not found: {artifact_pattern}")
        
        if not artifacts:
            logger.warning(f"[{tool.name}] No artifacts found for patterns: {tool.output_artifacts}")
        
        # Send email
        status_mgr.update_status(
            batch_id=batch_id,
            stage="SENDING_EMAIL",
            processed_files=file_count
        )
        
        recipients = self._get_recipients(tool.name)
        if recipients:
            subject = tool.email_subject_template.format(batch_id=batch_id)
            
            # Check for text report to use as email body
            body = self._build_email_body(batch_id, file_count, output_path, artifacts)
            
            self.email_service.send(recipients, subject, body, artifacts)
        
        # Done
        status_mgr.update_status(
            batch_id=batch_id,
            state="DONE",
            stage="DONE",
            message="Complete"
        )
    
    def _run_extractor(self, tool: ToolConfig, batch_id: str, processing_path: Path) -> Path:
        """Run extractor."""
        output_path = tool.output_dir / batch_id
        output_path.mkdir(parents=True, exist_ok=True)
        
        # Get files with the tool's accepted formats
        files = BatchOperations.get_files(processing_path, tool.input_formats)
        
        cmd = [
            "python3",
            str(tool.extractor_path)
        ] + [str(p) for p in files] + [
            "-o", str(output_path)
        ]
        
        logger.info(f"[{tool.name}] Running extractor...")
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode != 0:
            raise RuntimeError(f"Extractor failed: {result.stderr}")
        
        return output_path
    
    def _get_recipients(self, tool_name: str) -> List[str]:
        """Get recipients: per-tool first, then global fallback."""
        tool_key = f"MAIL_TO_{tool_name.upper()}"
        mail_to = os.getenv(tool_key, "").strip()
        if not mail_to:
            mail_to = os.getenv("MAIL_TO", "").strip()
        return [e.strip() for e in mail_to.split(",") if e.strip()]
    
    def _build_email_body(self, batch_id: str, file_count: int, output_path: Path, artifacts: List[Path]) -> str:
        """
        Build email body. If a .txt report file exists in artifacts, use it as body.
        Otherwise use default message.
        """
        # Look for text report file (e.g., *_reporte_*.txt)
        txt_reports = [a for a in artifacts if a.suffix == '.txt' and 'report' in a.name.lower() or 'reporte' in a.name.lower()]
        
        if txt_reports:
            # Use the first text report as email body
            try:
                report_content = txt_reports[0].read_text(encoding='utf-8')
                return report_content
            except Exception as e:
                logger.warning(f"Failed to read report file {txt_reports[0]}: {e}")
        
        # Default body if no report found
        return f"Batch {batch_id} processed.\n{file_count} file(s) processed."


def main():
    """Entry point."""
    config_path = os.getenv("CONFIG_PATH", "/config/tools.yaml")
    registry = ToolRegistry.from_yaml(config_path)
    
    email_config = EmailConfig(
        host=os.getenv("SMTP_HOST", ""),
        port=int(os.getenv("SMTP_PORT", "587")),
        user=os.getenv("SMTP_USER", ""),
        password=os.getenv("SMTP_PASS", ""),
        mail_from=os.getenv("MAIL_FROM", ""),
        use_ssl=(os.getenv("SMTP_PORT") == "465")
    )
    
    processor = UnifiedProcessor(
        registry=registry,
        email_service=EmailService(email_config),
        poll_seconds=int(os.getenv("POLL_SECONDS", "3")),
        batch_quiet_seconds=int(os.getenv("BATCH_QUIET_SECONDS", "180"))
    )
    
    processor.run()


if __name__ == "__main__":
    main()

