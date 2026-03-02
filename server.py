"""
Google Drive MCP Server for Claude

A local MCP server that gives Claude full read-write access to Google Drive,
Docs, Sheets, Slides, and Apps Script.

Run with:  python server.py
Or via Claude Code config pointing to this file.
"""

from mcp.server.fastmcp import FastMCP
from tools import register_all

mcp = FastMCP("gdrive_mcp")
register_all(mcp)

if __name__ == "__main__":
    mcp.run()
