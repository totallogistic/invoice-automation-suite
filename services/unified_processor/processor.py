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

BL_CSV_DIR = Path(os.getenv("BL_CSV_DIR", "/data/bl/csv"))

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
        
        # Count input files
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
                matched = list(output_path.glob(artifact_pattern))
                artifacts.extend(matched)
                logger.info(f"[{tool.name}] Pattern '{artifact_pattern}' matched {len(matched)} file(s)")
            else:
                artifact_path = output_path / artifact_pattern
                if artifact_path.exists():
                    artifacts.append(artifact_path)
                    logger.info(f"[{tool.name}] Found artifact: {artifact_pattern}")
                else:
                    logger.warning(f"[{tool.name}] Artifact not found: {artifact_pattern}")
        
        if not artifacts:
            logger.warning(f"[{tool.name}] No artifacts found for patterns: {tool.output_artifacts}")
        
        # For tools that deliver via direct download (no email), processed_files
        # reflects the number of output artifacts so the UI can display it.
        # For email-based tools it reflects the number of input files processed.
        reported_count = len(artifacts) if not tool.email_subject_template else file_count

        # Send email
        status_mgr.update_status(
            batch_id=batch_id,
            stage="SENDING_EMAIL",
            processed_files=reported_count
        )
        
        recipients = self._get_recipients(tool.name)
        if recipients and tool.email_subject_template:
            subject = tool.email_subject_template.format(batch_id=batch_id)
            body = self._build_email_body(batch_id, file_count, output_path, artifacts)
            self.email_service.send(recipients, subject, body, artifacts)
        
        if tool.bl_mode:
            try:
                report_script = Path("/apps/bl/extractor/bl_report.py")
                if report_script.exists():
                    result_report = subprocess.run(
                        ["python3", str(report_script)],
                        capture_output=True, text=True,
                        env={
                            **os.environ,
                            "BL_CSV_DIR": os.getenv("BL_CSV_DIR",
                                str(Path(os.getenv("BL_DATA_ROOT", "/data/bl")) / "csv")),
                        }
                    )
                    if result_report.returncode == 0:
                        logger.info(f"[{tool.name}] Reporte BL enviado")
                    else:
                        logger.warning(f"[{tool.name}] Reporte BL falló: {result_report.stderr.strip()}")
                else:
                    logger.warning(f"[{tool.name}] bl_report.py no encontrado en {report_script}")
            except Exception as e:
                logger.warning(f"[{tool.name}] Error en post-proceso BL: {e}")
       
        # Done — processed_files carries the final count into the DONE state
        status_mgr.update_status(
            batch_id=batch_id,
            state="DONE",
            stage="DONE",
            processed_files=reported_count,
            message="Complete"
        )
    
    def _run_extractor(self, tool: ToolConfig, batch_id: str, processing_path: Path) -> Path:
        """Run extractor."""
        output_path = tool.output_dir / batch_id
        output_path.mkdir(parents=True, exist_ok=True)
        
        # Get files with the tool's accepted formats
        files = BatchOperations.get_files(processing_path, tool.input_formats)

        if not files:
            raise RuntimeError(
                f"No files found with extensions {tool.input_formats} in {processing_path}. "
                f"Present: {[f.name for f in processing_path.iterdir() if not f.name.startswith('_')]}"
            )
        
        if tool.inject_mode:
            cmd = self._build_inject_cmd(tool, files, output_path)
        elif tool.croton_import_mode:
            cmd = self._build_croton_import_cmd(tool, files, output_path)
        elif tool.camion_mode:
            skip_validation = (processing_path / "_SKIP_VALIDATION").exists()
            skip_t1  = (processing_path / "_SKIP_T1").exists()
            skip_dae = (processing_path / "_SKIP_DAE").exists()
            cmd = self._build_camion_cmd(
                tool, files, output_path,
                skip_validation=skip_validation,
                skip_t1=skip_t1,
                skip_dae=skip_dae,
            )
        elif tool.bl_mode:
            cmd = self._build_bl_cmd(tool, files)

        elif tool.export_visual_mode:
            cmd = self._build_export_visual_cmd(tool, files, output_path)

        else:
            cmd = [
                "python3",
                str(tool.extractor_path)
            ] + [str(p) for p in files] + [
                "-o", str(output_path)
            ]
        
        logger.info(f"[{tool.name}] Running extractor...")
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            raise RuntimeError(
                f"Extractor failed:\n"
                f"  stdout: {result.stdout.strip()}\n"
                f"  stderr: {result.stderr.strip()}"
            )
        
        report_path = output_path / f"reporte_{batch_id}.txt"
        report_path.write_text(result.stdout, encoding="utf-8")

        return output_path
    
    def _build_inject_cmd(self, tool: ToolConfig, files: List[Path], output_path: Path) -> List[str]:
        ods_files = [f for f in files if f.suffix.lower() == ".ods"]
        excel_files = [f for f in files if f.suffix.lower() in (".xlsx", ".xls")]

        packing_file = None
        factura_file = None

        _packing_kw = ("packing", "parking", "packing_list", "parking_list")
        _factura_kw = ("factura", "invoice", "mapeo", "mapping")

        if len(ods_files) == 1 and len(excel_files) == 1:
            ods_name = ods_files[0].name.lower()
            xls_name = excel_files[0].name.lower()
            if any(k in xls_name for k in _packing_kw) and any(k in ods_name for k in _factura_kw):
                packing_file = excel_files[0]
                factura_file = ods_files[0]
            else:
                packing_file = ods_files[0]
                factura_file = excel_files[0]

        elif len(ods_files) == 2:
            for f in ods_files:
                name = f.name.lower()
                if any(k in name for k in _packing_kw):
                    packing_file = f
                elif any(k in name for k in _factura_kw):
                    factura_file = f

            if not packing_file or not factura_file:
                raise RuntimeError(
                    f"[{tool.name}] Could not infer packing/factura from ODS filenames. "
                    f"Use names containing 'packing'/'parking' and 'factura'/'invoice'/'mapeo'/'mapping'."
                )

        elif len(ods_files) == 0 and len(excel_files) == 2:
            for f in excel_files:
                name = f.name.lower()
                if any(k in name for k in _packing_kw):
                    packing_file = f
                elif any(k in name for k in _factura_kw):
                    factura_file = f

            if not packing_file or not factura_file:
                raise RuntimeError(
                    f"[{tool.name}] Could not infer packing/factura from filenames. "
                    f"Use names containing 'packing'/'parking' and 'factura'/'invoice'/'mapeo'/'mapping'."
                )

        else:
            raise RuntimeError(
                f"[{tool.name}] inject_mode expects exactly 2 files "
                f"(ODS and/or XLSX/XLS: packing + factura/mapeo)."
            )

        stem = packing_file.stem
        out_ext = packing_file.suffix.lower()
        out_filename = (
            tool.filename_pattern.format(stem=stem, ext=out_ext)
            if tool.filename_pattern
            else f"CROTON_{stem}{out_ext}"
        )
        out_file = output_path / out_filename

        return [
            "python3", str(tool.extractor_path),
            str(packing_file),
            "--factura", str(factura_file),
            "--inject",
            "-o", str(out_file),
        ]

    def _build_bl_cmd(self, tool: ToolConfig, files: List[Path]) -> List[str]:
        bl_csv_dir = Path(os.getenv("BL_CSV_DIR", "/data/bl/csv"))
        bl_csv_dir.mkdir(parents=True, exist_ok=True)
        return [
            "python3",
            str(tool.extractor_path),
            *[str(f) for f in files],
            "-o", str(bl_csv_dir),
        ]

    def _build_camion_cmd(
        self,
        tool: ToolConfig,
        files: List[Path],
        output_path: Path,
        skip_validation: bool = False,
        skip_t1:  bool = False,
        skip_dae: bool = False,
    ) -> List[str]:
        xlsx_files = [f for f in files if f.suffix.lower() == ".xlsx"]
        pdf_files = [f for f in files if f.suffix.lower() == ".pdf"]

        if len(xlsx_files) != 1:
            raise RuntimeError(
                f"[{tool.name}] camion_mode expects exactly 1 XLSX file "
                f"(packing list), got {len(xlsx_files)}."
            )
        xlsx_file = xlsx_files[0]

        t1_files = [f for f in pdf_files if "t1" in f.name.lower()]
        doc_files = [f for f in pdf_files if "doc" in f.name.lower()]
        other_pdfs = [f for f in pdf_files if f not in t1_files and f not in doc_files]

        if not skip_validation:
            if other_pdfs:
                if not doc_files and len(other_pdfs) == 1:
                    doc_files = other_pdfs
                    other_pdfs = []
                else:
                    raise RuntimeError(
                        f"[{tool.name}] Cannot classify PDF(s): "
                        f"{[f.name for f in other_pdfs]}. "
                        "Use filenames containing 't1' or 'doc'."
                    )

            if len(doc_files) != 1:
                raise RuntimeError(
                    f"[{tool.name}] Expected exactly 1 DOC PDF file, "
                    f"got {len(doc_files)}. Filename must contain 'doc'."
                )

        cmd = [
            "python3", str(tool.extractor_path),
            "--xlsx", str(xlsx_file),
            "-o", str(output_path),
        ]

        if not skip_validation:
            cmd.extend(["--doc", str(doc_files[0])])

        if t1_files:
            cmd.extend(["--t1", *[str(f) for f in t1_files]])

        if skip_t1:
            cmd.append("--skip-t1")
        if skip_dae:
            cmd.append("--skip-dae")

        return cmd

    def _build_croton_import_cmd(self, tool, files, output_path):
        pdf_files  = [f for f in files if f.suffix.lower() == '.pdf']
        xlsx_files = [f for f in files if f.suffix.lower() == '.xlsx']
        if len(pdf_files) != 1:
            raise RuntimeError(f"[{tool.name}] esperaba 1 PDF, recibió {len(pdf_files)}")
        if len(xlsx_files) != 1:
            raise RuntimeError(f"[{tool.name}] esperaba 1 XLSX, recibió {len(xlsx_files)}")
        return [
            "python3", str(tool.extractor_path),
            "--pdf",  str(pdf_files[0]),
            "--xlsx", str(xlsx_files[0]),
            "--output", str(output_path),
        ]

    def _build_export_visual_cmd(self, tool: ToolConfig, files: List[Path], output_path: Path) -> List[str]:
            """
            Modo export_visual: lee el CSV de partidas del cliente y genera el CONVERT.
            El cliente y la fecha opcional se leen de ficheros marcadores en el batch.
            """
            csv_files = [f for f in files if f.suffix.lower() == ".csv"]
            if len(csv_files) != 1:
                raise RuntimeError(
                    f"[{tool.name}] export_visual_mode espera exactamente 1 CSV, "
                    f"recibió {len(csv_files)}: {[f.name for f in csv_files]}"
                )
    
            # Leer cliente del marcador _CLIENTE.txt (por defecto: aldi)
            cliente_marker = output_path.parent.parent / "processing" / output_path.name / "_CLIENTE.txt"
            # El processing_path es el directorio del batch; buscamos el marcador allí
            # Nota: en _run_extractor se llama con processing_path como base,
            # pero aquí recibimos output_path. Buscamos el marcador via el inbox original.
            # Solución simple: el marcador se guarda como fichero en el mismo batch junto al CSV.
            batch_dir = csv_files[0].parent
            cliente_file = batch_dir / "_CLIENTE.txt"
            cliente = cliente_file.read_text(encoding="utf-8").strip() if cliente_file.exists() else "aldi"
    
            fecha_file = batch_dir / "_FECHA.txt"
            fecha = fecha_file.read_text(encoding="utf-8").strip() if fecha_file.exists() else None
    
            cmd = [
                "python3", str(tool.extractor_path),
                str(csv_files[0]),
                "-o", str(output_path),
                "--cliente", cliente,
            ]
            if fecha:
                cmd += ["--fecha", fecha]
    
            return cmd

    def _get_recipients(self, tool_name: str) -> List[str]:
        tool_key = f"MAIL_TO_{tool_name.upper()}"
        mail_to = os.getenv(tool_key, "").strip()
        if not mail_to:
            mail_to = os.getenv("MAIL_TO", "").strip()
        return [e.strip() for e in mail_to.split(",") if e.strip()]
    
    def _build_email_body(self, batch_id: str, file_count: int, output_path: Path, artifacts: List[Path]) -> str:
        report_path = output_path / f"reporte_{batch_id}.txt"
        if report_path.exists():
            try:
                return report_path.read_text(encoding="utf-8")
            except Exception as e:
                logger.warning(f"Failed to read report file {report_path}: {e}")
        return f"Batch {batch_id} processed.\n{file_count} file(s) processed."


def main():
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