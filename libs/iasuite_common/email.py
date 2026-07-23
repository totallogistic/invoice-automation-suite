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

    @classmethod
    def from_env(cls, mode: Optional[str] = None) -> "EmailConfig":
        """Construye la config según el MODO DE ENVÍO.

        mode: 'postfix' (relay local del host, sin autenticación) | 'custom'
        (cuenta SMTP externa autenticada). Si es None, se usa la variable de
        entorno MAIL_SEND_MODE (por defecto 'postfix').

        Cada modo lee sus propias variables namespaced; si no están definidas,
        cae a las planas SMTP_HOST/PORT/USER/PASS (retrocompatibilidad: sin
        tocar el .env, 'postfix' reproduce el comportamiento actual).
        """
        import os
        mode = (mode or os.getenv("MAIL_SEND_MODE", "postfix") or "postfix").lower()
        mail_from = os.getenv("MAIL_FROM", "")
        if mode == "custom":
            host = os.getenv("SMTP_CUSTOM_HOST") or os.getenv("SMTP_HOST", "")
            port = int(os.getenv("SMTP_CUSTOM_PORT") or os.getenv("SMTP_PORT") or "587")
            user = os.getenv("SMTP_CUSTOM_USER") or os.getenv("SMTP_USER", "")
            pwd = os.getenv("SMTP_CUSTOM_PASS") or os.getenv("SMTP_PASS", "")
        else:  # 'postfix' (relay local, sin auth)
            host = os.getenv("SMTP_POSTFIX_HOST") or os.getenv("SMTP_HOST", "") or "172.18.0.1"
            port = int(os.getenv("SMTP_POSTFIX_PORT") or os.getenv("SMTP_PORT") or "25")
            user = os.getenv("SMTP_POSTFIX_USER", "")   # normalmente vacío (sin auth)
            pwd = os.getenv("SMTP_POSTFIX_PASS", "")
        return cls(host=host, port=port, user=user, password=pwd,
                   mail_from=mail_from, use_ssl=(port == 465))


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
        """Send email with optional attachments and inline images.

        inline_images: lista de imágenes que se embeben como CID (cid:<stem>).
        Requiere body HTML. Crea estructura multipart/related igual que Thunderbird.
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
        """Build email message.

        Con inline_images construye estructura MIME igual a Thunderbird:
            multipart/mixed
            ├── multipart/related
            │   ├── text/html
            │   └── image/gif  (Content-ID: <tls_logo>)
            └── application/pdf  (adjunto)
        Sin inline_images usa EmailMessage estándar.
        """
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.mime.image import MIMEImage
        from email.mime.base import MIMEBase
        from email import encoders as _enc

        is_html = body.strip().lower().startswith(("<html", "<!doctype"))

        if is_html:
            outer = MIMEMultipart("mixed")
            outer["From"]    = self.config.mail_from
            outer["To"]      = ", ".join(to)
            outer["Subject"] = subject

            related = MIMEMultipart("related")
            related.attach(MIMEText(body, "html", "utf-8"))

            for img_path in (inline_images or []):
                if not img_path.exists():
                    logger.warning(f"Inline image not found: {img_path}")
                    continue
                img_subtype = img_path.suffix.lstrip(".").lower() or "octet-stream"
                mime_img = MIMEImage(img_path.read_bytes(), img_subtype)
                mime_img["Content-ID"] = f"<{img_path.stem}>"
                mime_img["Content-Disposition"] = f'inline; filename="{img_path.name}"'
                related.attach(mime_img)
                logger.debug(f"Inline image: cid:{img_path.stem} ({img_path.name})")

            outer.attach(related)

            for path in attachments:
                if not path.exists():
                    logger.warning(f"Attachment not found: {path}")
                    continue
                maintype, subtype = self._get_mime_type(path)
                part = MIMEBase(maintype, subtype)
                part.set_payload(path.read_bytes())
                _enc.encode_base64(part)
                part["Content-Disposition"] = f'attachment; filename="{path.name}"'
                outer.attach(part)

            return outer

        else:
            msg = EmailMessage()
            msg["From"]    = self.config.mail_from
            msg["To"]      = ", ".join(to)
            msg["Subject"] = subject
            msg.set_content(body)

            for path in attachments:
                if not path.exists():
                    logger.warning(f"Attachment not found: {path}")
                    continue
                data = path.read_bytes()
                maintype, subtype = self._get_mime_type(path)
                msg.add_attachment(
                    data, maintype=maintype, subtype=subtype, filename=path.name
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