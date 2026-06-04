"""Shared email service."""
import smtplib
import ssl
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from typing import List, Optional
import logging

logger = logging.getLogger(__name__)


@dataclass
class EmailConfig:
    """Email configuration."""
    host: str
    port: int
    user: str
    password: str
    mail_from: str
    use_ssl: bool = False
    timeout: int = 30


class EmailService:
    """Centralized email sending service."""
    
    def __init__(self, config: EmailConfig):
        self.config = config
    
    def send(
        self,
        to: List[str],
        subject: str,
        body: str,
        attachments: List[Path] = None,
        inline_images: List[Path] = None,
    ) -> bool:
        """Send email with attachments. Returns True if successful.

        Si body comienza con <html o <!DOCTYPE se envía como text/html.
        inline_images: lista de ficheros imagen que se embeben como CID inline.
        El CID de cada imagen es el nombre del fichero sin extensión
        (e.g. logo.gif → cid:logo en el HTML).
        """
        if not to:
            logger.warning("No recipients, skipping email")
            return False

        try:
            msg = self._build_message(
                to, subject, body,
                attachments or [],
                inline_images or [],
            )
            self._send_message(msg, to)
            logger.info(f"Email sent to {len(to)} recipients")
            return True
        except Exception as e:
            logger.error(f"Failed to send email: {e}")
            return False

    def _build_message(
        self,
        to: List[str],
        subject: str,
        body: str,
        attachments: List[Path],
        inline_images: List[Path] = None,
    ) -> EmailMessage:
        """Build email message (HTML o texto plano según el cuerpo)."""
        msg = EmailMessage()
        msg["From"] = self.config.mail_from
        msg["To"] = ", ".join(to)
        msg["Subject"] = subject

        is_html = body.strip().lower().startswith(("<html", "<!doctype"))

        if is_html:
            msg.set_content(body, subtype="html")
            # Añadir imágenes inline (CID = nombre sin extensión)
            for img in (inline_images or []):
                if not img.exists():
                    logger.warning(f"Inline image not found: {img}")
                    continue
                data = img.read_bytes()
                subtype = img.suffix.lstrip(".").lower() or "octet-stream"
                cid = img.stem          # logo.gif → cid "logo"
                msg.add_related(
                    data,
                    maintype="image",
                    subtype=subtype,
                    cid=f"<{cid}>",
                    disposition="inline",
                )
                logger.debug(f"Inline image added: cid:{cid} ({img.name})")
        else:
            msg.set_content(body)

        for path in attachments:
            if not path.exists():
                logger.warning(f"Attachment not found: {path}")
                continue

            data = path.read_bytes()
            maintype, subtype = self._get_mime_type(path)
            msg.add_attachment(
                data,
                maintype=maintype,
                subtype=subtype,
                filename=path.name,
            )

        return msg

    def _send_message(self, msg: EmailMessage, recipients: List[str]):
        """Send via SMTP."""
        ctx = ssl.create_default_context()
        
        if self.config.use_ssl:
            with smtplib.SMTP_SSL(
                self.config.host,
                self.config.port,
                timeout=self.config.timeout,
                context=ctx
            ) as smtp:
                if self.config.user:                              # ← NUEVO
                    smtp.login(self.config.user, self.config.password)
                smtp.send_message(msg, from_addr=self.config.mail_from, to_addrs=recipients)
        else:
            with smtplib.SMTP(
                self.config.host,
                self.config.port,
                timeout=self.config.timeout
            ) as smtp:
                smtp.ehlo()
                if self.config.user:                              # ← NUEVO (envuelve starttls + login)
                    smtp.starttls(context=ctx)
                    smtp.ehlo()
                    smtp.login(self.config.user, self.config.password)
                smtp.send_message(msg, from_addr=self.config.mail_from, to_addrs=recipients)
        
    @staticmethod
    def _get_mime_type(path: Path) -> tuple:
        """Get MIME type for file."""
        suffix = path.suffix.lower()
        types = {
            ".json": ("application", "json"),
            ".csv": ("text", "csv"),
            ".txt": ("text", "plain"),
            ".xlsx": ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            ".xls": ("application", "vnd.ms-excel"),
            ".ods": ("application", "vnd.oasis.opendocument.spreadsheet"),
            ".pdf": ("application", "pdf"),
        }
        return types.get(suffix, ("application", "octet-stream"))
