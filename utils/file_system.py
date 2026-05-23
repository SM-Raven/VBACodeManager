from pathlib import Path
import shutil
from constants import EXPORT_FOLDERS


class FileSystem:
    """Handles project folder structure for VCM"""

    def __init__(self, base_dir: Path | None = None):
        self.base_dir: Path = base_dir if base_dir is not None else Path.cwd()
        self.src_dir: Path = self.base_dir / "src"
        self.required_folders = set(EXPORT_FOLDERS.values())

    def ensure_src(self) -> Path:
        """Ensure src directory exists"""
        self.src_dir.mkdir(parents=True, exist_ok=True)
        return self.src_dir

    def ensure_structure(self) -> Path:
        """Ensure full VBA export folder structure exists"""
        self.ensure_src()

        for folder in self.required_folders:
            (self.src_dir / folder).mkdir(parents=True, exist_ok=True)

        return self.src_dir

    def is_structure_valid(self) -> bool:
        """Check if src structure is properly set up"""

        if not self.src_dir.exists():
            return False

        return all(
            (self.src_dir / folder).exists()
            for folder in self.required_folders
        )

    def validate_or_create(self) -> bool:
        """
        Ensure structure is valid, create if missing.

        Returns:
            True  -> already valid
            False -> structure had to be created
        """

        if self.is_structure_valid():
            return True

        self.ensure_structure()
        return False

    def reset_src(self) -> None:
        """Delete and recreate src folder"""

        if self.src_dir.exists():
            shutil.rmtree(self.src_dir, ignore_errors=True)

        self.ensure_structure()


def ensure_project_structure(base_dir: Path | None = None) -> Path:
    """Ensure project structure exists"""
    return FileSystem(base_dir).ensure_structure()


def is_valid_project_structure(base_dir: Path | None = None) -> bool:
    """Validate project structure"""
    return FileSystem(base_dir).is_structure_valid()