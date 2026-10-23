"""Multi-source consistency auditor engine."""

from .profile import ProfileError, load_profile
from .rules import AuditEngine

__all__ = ["AuditEngine", "ProfileError", "load_profile"]
__version__ = "0.1.0"
