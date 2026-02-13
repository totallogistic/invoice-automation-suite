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
        attachments: List[Path] = None
    ) -> bool:
        """Send email with attachments. Returns True if successful."""
        if not to:
            logger.warning("No recipients, skipping email")
            return False
        
        try:
            msg = self._build_message(to, subject, body, attachments or [])
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
        attachments: List[Path]
    ) -> EmailMessage:
        """Build email message."""
        msg = EmailMessage()
        msg["From"] = self.config.mail_from
        msg["To"] = ", ".join(to)
        msg["Subject"] = subject
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
                filename=path.name
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
                smtp.login(self.config.user, self.config.password)
                smtp.send_message(msg, from_addr=self.config.mail_from, to_addrs=recipients)
        else:
            with smtplib.SMTP(
                self.config.host,
                self.config.port,
                timeout=self.config.timeout
            ) as smtp:
                smtp.ehlo()
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
            ".xlsx": ("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            ".pdf": ("application", "pdf"),
        }
        return types.get(suffix, ("application", "octet-stream"))
