"""Excel MCP Server - Excel operations via Model Context Protocol"""

from .server import mcp

__version__ = "0.1.0"


def main():
    """Main entry point for CLI"""
    mcp.run()


__all__ = ["mcp", "main"]
