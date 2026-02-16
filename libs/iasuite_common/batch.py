"""Shared batch handling utilities."""
from __future__ import annotations

import shutil
import time
from pathlib import Path
from typing import List
import logging

logger = logging.getLogger(__name__)


class BatchOperations:
    """Utility class for batch file operations."""
    
    @staticmethod
    def find_batches(inbox_dir: Path) -> List[Path]:
        """Find all batch directories in inbox."""
        if not inbox_dir.exists():
            return []
        return sorted([p for p in inbox_dir.iterdir() if p.is_dir()])
    
    @staticmethod
    def is_batch_ready(
        batch_dir: Path,
        done_marker: str = "_DONE",
        quiet_seconds: int = 180
    ) -> bool:
        """Check if batch is ready for processing."""
        # Check for explicit done marker
        if (batch_dir / done_marker).exists():
            return True
        
        # Check quiet time
        latest_mtime = BatchOperations._get_latest_mtime(batch_dir)
        if latest_mtime == 0.0:
            return False
        
        idle_time = time.time() - latest_mtime
        return idle_time >= quiet_seconds
    
    @staticmethod
    def _get_latest_mtime(directory: Path) -> float:
        """Get latest modification time."""
        latest = 0.0
        for file_path in directory.rglob("*"):
            if file_path.is_file():
                try:
                    latest = max(latest, file_path.stat().st_mtime)
                except FileNotFoundError:
                    continue
        return latest
    
    @staticmethod
    def move_batch(src: Path, dest_dir: Path) -> Path:
        """Move batch directory."""
        dest = dest_dir / src.name
        dest_dir.mkdir(parents=True, exist_ok=True)
        
        if dest.exists():
            shutil.rmtree(dest)
        
        shutil.move(str(src), str(dest))
        logger.info(f"Moved: {src.name} -> {dest_dir.name}/")
        return dest
    
    @staticmethod
    def get_pdfs(directory: Path) -> List[Path]:
        """Get all PDF files (kept for backward compatibility)."""
        return sorted(directory.rglob("*.pdf"))
    
    @staticmethod
    def get_files(directory: Path, extensions: List[str] = None) -> List[Path]:
        """
        Get all files with specified extensions.
        
        Args:
            directory: Directory to search
            extensions: List of extensions (without dot), e.g., ['pdf', 'xlsx', 'xls']
                       If None, returns all files except hidden and _DONE marker
        
        Returns:
            Sorted list of file paths
        """
        files = []
        
        if extensions:
            # Get files with specific extensions
            for ext in extensions:
                files.extend(directory.rglob(f"*.{ext}"))
        else:
            # Get all files, excluding hidden and markers
            for file_path in directory.rglob("*"):
                if file_path.is_file() and not file_path.name.startswith(('.', '_')):
                    files.append(file_path)
        
        return sorted(files)

