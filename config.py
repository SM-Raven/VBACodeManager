"""
Configuration module - Central place for VCM settings
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict


@dataclass
class VCMConfig:
    """VCM configuration settings"""
    
    # Folder configuration
    src_folder: Path = field(default_factory=lambda: Path("src"))
    
    # Component type mappings
    subfolder_mapping: Dict[str, str] = field(
        default_factory=lambda: {
            "cls": "Class Modules",
            "mod": "Standard Modules",
            "frm": "Forms",
            "exl": "Excel Objects"
        }
    )
    
    # Supported file extensions
    supported_extensions: tuple = (".xlsm", ".xla", ".xls")
    
    # File extensions for each component type
    extension_mapping: Dict[str, str] = field(
        default_factory=lambda: {
            "cls": ".cls",
            "mod": ".bas",
            "frm": ".frm",
            "exl": ".cls"
        }
    )
    
    # Formatting configuration
    indent_size: int = 4
    max_line_length: int = 100
    
    def get_component_folder(self, component_type: str) -> Path:
        """
        Get the folder for a component type.
        
        Args:
            component_type: Type code (cls, mod, frm, exl)
            
        Returns:
            Path to the folder
        """
        if component_type not in self.subfolder_mapping:
            raise ValueError(f"Unknown component type: {component_type}")
        
        return self.src_folder / component_type
    
    def get_extension(self, component_type: str) -> str:
        """
        Get the file extension for a component type.
        
        Args:
            component_type: Type code (cls, mod, frm, exl)
            
        Returns:
            File extension (e.g., ".cls")
        """
        if component_type not in self.extension_mapping:
            raise ValueError(f"Unknown component type: {component_type}")
        
        return self.extension_mapping[component_type]
    
    def is_supported_extension(self, file_path: str) -> bool:
        """
        Check if file has supported Excel extension.
        
        Args:
            file_path: Path or filename
            
        Returns:
            True if supported
        """
        return file_path.lower().endswith(self.supported_extensions)
    
    def get_all_component_types(self) -> list:
        """
        Get all component type codes.
        
        Returns:
            List of type codes
        """
        return list(self.subfolder_mapping.keys())


# Global config instance
config = VCMConfig()


if __name__ == "__main__":
    # Test configuration
    cfg = VCMConfig()
    
    print("Component types:", cfg.get_all_component_types())
    print("Folder for cls:", cfg.get_component_folder("cls"))
    print("Extension for mod:", cfg.get_extension("mod"))
    print("Is .xlsm supported:", cfg.is_supported_extension("book.xlsm"))
    print("Is .xlsx supported:", cfg.is_supported_extension("book.xlsx"))
