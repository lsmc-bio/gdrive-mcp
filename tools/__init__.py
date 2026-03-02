"""Tool modules for gdrive-mcp."""

from . import drive, docs, sheets, slides, comments, scripts, management, gmail


def register_all(mcp):
    """Register all tool modules."""
    drive.register(mcp)
    docs.register(mcp)
    sheets.register(mcp)
    slides.register(mcp)
    comments.register(mcp)
    scripts.register(mcp)
    management.register(mcp)
    gmail.register(mcp)
