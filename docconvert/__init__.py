"""docconvert - a small document converter (docx, pdf, md, txt, html)."""

from .core import SUPPORTED, convert

__all__ = ["SUPPORTED", "convert"]
__version__ = "0.1.0"
