"""
Version module - reads version from pyproject.toml
"""

import tomllib
from pathlib import Path

def get_version() -> str:
    """
    Read version from pyproject.toml

    Returns:
        str: Version string (e.g., "0.1.0")
    """
    try:
        # Find pyproject.toml
        project_root = Path(__file__).parent.parent.parent
        pyproject_path = project_root / "pyproject.toml"

        if not pyproject_path.exists():
            return "unknown"

        if tomllib is not None:
            with open(pyproject_path, "rb") as f:
                data = tomllib.load(f)
                return data.get("project", {}).get("version", "unknown")
        else:
            # Fallback: simple regex parsing if tomli not available
            with open(pyproject_path, "r") as f:
                for line in f:
                    if line.startswith("version"):
                        # Extract version between quotes
                        return line.split('"')[1]
            return "unknown"
    except Exception:
        return "unknown"