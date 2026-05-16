"""
    File System Utilities - VCM
    Handles src structure creation and validation for VBA exports
"""

from pathlib import Path
from constants import EXPORT_FOLDERS

class FileSystem:
    """Handles project folder structure for VCM"""

    def __init__(self, base_dir: Path | None = None):
        self.base_dir = base_dir or Path.cwd()
        self.src_dir = self.base_dir / "src"

    def ensure_src(self) -> Path:
        """Ensure src directory exists"""
        self.src_dir.mkdir(parents=True, exist_ok=True)
        return self.src_dir

    def ensure_structure(self) -> Path:
        """Ensure full VBA export folder structure exists"""
        self.ensure_src()

        for folder in set(EXPORT_FOLDERS.values()):
            (self.src_dir / folder).mkdir(parents=True, exist_ok=True)

        return self.src_dir

    def is_structure_valid(self) -> bool:
        """Check if src structure is properly set up"""
        if not self.src_dir.exists():
            return False

        required_folders = set(EXPORT_FOLDERS.values())

        for folder in required_folders:
            if not (self.src_dir / folder).exists():
                return False

        return True

    def validate_or_create(self) -> bool:
        """
        Ensure structure is valid, create if missing
        Returns True if already valid, False if it had to create
        """
        if self.is_structure_valid():
            return True

        self.ensure_structure()
        return False

    def reset_src(self) -> None:
        """Delete and recreate src folder"""
        if self.src_dir.exists():
            for item in self.src_dir.iterdir():
                if item.is_dir():
                    import shutil
                    shutil.rmtree(item)
                else:
                    item.unlink()

        self.ensure_structure()

def ensure_project_structure(base_dir: Path | None = None) -> Path:
    """Quick helper to ensure full structure"""
    return FileSystem(base_dir).ensure_structure()


def validate_project_structure(base_dir: Path | None = None) -> bool:
    """Quick helper to validate structure"""
    return FileSystem(base_dir).is_structure_valid()
