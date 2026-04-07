"""Shared status management utilities."""
from __future__ import annotations

import json
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional


@dataclass
class BatchStatus:
    """Standardized batch status structure."""
    batch_id: str
    state: str  # UPLOADED, PROCESSING, DONE, ERROR
    stage: str  # WAITING, RUNNING_EXTRACTOR, SENDING_EMAIL, DONE, ERROR
    total_files: int
    processed_files: int
    started_at: str
    updated_at: str
    recipients: List[str]
    message: str
    error_details: Optional[str] = None


class StatusManager:
    """Centralized status management for all tools."""
    
    def __init__(self, status_root: Path):
        self.status_root = Path(status_root)
        self.status_root.mkdir(parents=True, exist_ok=True)
    
    def _status_path(self, batch_id: str) -> Path:
        return self.status_root / batch_id / "status.json"
    
    def _now_iso(self) -> str:
        return datetime.now(timezone.utc).isoformat()
    
    def create_status(
        self,
        batch_id: str,
        state: str = "UPLOADED",
        stage: str = "WAITING",
        total_files: int = 0,
        recipients: List[str] = None,
        message: str = ""
    ) -> BatchStatus:
        """Create initial status for a new batch."""
        status = BatchStatus(
            batch_id=batch_id,
            state=state,
            stage=stage,
            total_files=total_files,
            processed_files=0,
            started_at=self._now_iso(),
            updated_at=self._now_iso(),
            recipients=recipients or [],
            message=message
        )
        self._write_status(status)
        return status
    
    def update_status(
        self,
        batch_id: str,
        state: Optional[str] = None,
        stage: Optional[str] = None,
        processed_files: Optional[int] = None,
        message: Optional[str] = None,
        error_details: Optional[str] = None
    ) -> BatchStatus:
        """Update existing status."""
        current = self.get_status(batch_id)
        if not current:
            raise ValueError(f"Status not found for batch: {batch_id}")
        
        if state is not None:
            current.state = state
        if stage is not None:
            current.stage = stage
        if processed_files is not None:
            current.processed_files = processed_files
        if message is not None:
            current.message = message
        if error_details is not None:
            current.error_details = error_details
        
        current.updated_at = self._now_iso()
        self._write_status(current)
        return current
    
    def get_status(self, batch_id: str) -> Optional[BatchStatus]:
        """Get current status for a batch."""
        path = self._status_path(batch_id)
        if not path.exists():
            return None
        
        data = json.loads(path.read_text(encoding="utf-8"))
        return BatchStatus(**data)
    
    def _write_status(self, status: BatchStatus):
        """Atomic write of status JSON."""
        path = self._status_path(status.batch_id)
        path.parent.mkdir(parents=True, exist_ok=True)
        
        tmp = path.with_suffix(".tmp")
        tmp.write_text(
            json.dumps(asdict(status), ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
        tmp.replace(path)
