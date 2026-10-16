from .server import mcp

def main() -> None:
    mcp.run()

__all__ = ["main", "mcp"]
