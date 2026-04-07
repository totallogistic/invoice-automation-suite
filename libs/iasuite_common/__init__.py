"""Invoice Automation Suite - Common Utilities."""

from .status import StatusManager, BatchStatus
from .email import EmailService, EmailConfig
from .batch import BatchOperations

__version__ = "1.0.0"
__all__ = [
    "StatusManager",
    "BatchStatus",
    "EmailService",
    "EmailConfig",
    "BatchOperations",
]
