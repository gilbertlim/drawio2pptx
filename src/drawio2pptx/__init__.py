"""Convert draw.io diagrams into individually editable PowerPoint objects."""

__version__ = "0.1.0"

from .build import convert  # noqa: E402,F401

__all__ = ["convert", "__version__"]
