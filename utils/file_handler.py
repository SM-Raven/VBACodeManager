"""
File Handler - Reusable file I/O for VBA components
Used by both export and import commands
"""

from pathlib import Path
from typing import List, Tuple, Optional
import shutil


class FileHandler:
    """Handles file operations for VBA components"""
    
    @staticmethod
    def ensure_folder(folder_path: Path) -> None:
        """
        Create folder if it doesn't exist.
        
        Args:
            folder_path: Path to folder to create
        """
        folder_path.mkdir(parents=True, exist_ok=True)
    
    @staticmethod
    def delete_folder(folder_path: Path) -> None:
        """
        Recursively delete a folder and all its contents.
        
        Args:
            folder_path: Path to folder to delete
        """
        if folder_path.exists():
            shutil.rmtree(folder_path)
    
    @staticmethod
    def write_component(
        folder_path: Path,
        component_name: str,
        code: str,
        extension: str
    ) -> Path:
        """
        Write a VBA component to a file.
        
        Args:
            folder_path: Path to folder where file will be written
            component_name: Name of the component (without extension)
            code: VBA code content
            extension: File extension (e.g., '.cls', '.bas', '.frm')
            
        Returns:
            Path to the written file
            
        Raises:
            IOError: If write fails
        """
        try:
            # Ensure folder exists
            FileHandler.ensure_folder(folder_path)
            
            # Create file path
            file_path = folder_path / f"{component_name}{extension}"
            
            # Write code to file
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(code)
            
            return file_path
        
        except Exception as e:
            raise IOError(
                f"Failed to write component '{component_name}' to '{folder_path}': {str(e)}"
            ) from e
    
    @staticmethod
    def read_component(file_path: Path) -> str:
        """
        Read VBA code from a file.
        
        Args:
            file_path: Path to the component file
            
        Returns:
            VBA code content
            
        Raises:
            FileNotFoundError: If file doesn't exist
            IOError: If read fails
        """
        try:
            if not file_path.exists():
                raise FileNotFoundError(f"Component file not found: {file_path}")
            
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        
        except FileNotFoundError:
            raise
        except Exception as e:
            raise IOError(
                f"Failed to read component from '{file_path}': {str(e)}"
            ) from e
    
    @staticmethod
    def component_exists(folder_path: Path, component_name: str, extension: str) -> bool:
        """
        Check if a component file exists.
        
        Args:
            folder_path: Path to folder
            component_name: Component name (without extension)
            extension: File extension (e.g., '.cls')
            
        Returns:
            True if file exists
        """
        file_path = folder_path / f"{component_name}{extension}"
        return file_path.exists()
    
    @staticmethod
    def get_components_in_folder(folder_path: Path) -> List[Tuple[str, Path]]:
        """
        Get all component files in a folder.
        
        Args:
            folder_path: Path to folder (e.g., src/cls)
            
        Returns:
            List of tuples (component_name, file_path)
        """
        components = []
        
        if not folder_path.exists():
            return components
        
        valid_extensions = {'.bas', '.cls', '.frm'}
        
        for file_path in folder_path.iterdir():
            if file_path.is_file() and file_path.suffix in valid_extensions:
                component_name = file_path.stem
                components.append((component_name, file_path))
        
        return components
    
    @staticmethod
    def get_all_components_in_src(src_path: Path) -> List[Tuple[str, str, Path]]:
        """
        Get all components across all subfolders (cls, mod, frm, exl).
        
        Args:
            src_path: Path to src folder
            
        Returns:
            List of tuples (component_type, component_name, file_path)
            
        Example:
            [('cls', 'MyClass', Path('src/cls/MyClass.cls')),
             ('mod', 'MyModule', Path('src/mod/MyModule.bas'))]
        """
        components = []
        
        for type_folder in ['cls', 'mod', 'frm', 'exl']:
            folder = src_path / type_folder
            for component_name, file_path in FileHandler.get_components_in_folder(folder):
                components.append((type_folder, component_name, file_path))
        
        return components
    
    @staticmethod
    def parse_onefile_path(onefile_arg: str) -> Tuple[str, str, str]:
        """
        Parse --onefile argument into components.
        
        Args:
            onefile_arg: Input like 'cls/MyClass' or 'mod/Module.bas'
            
        Returns:
            Tuple of (component_type, component_name, extension)
            
        Raises:
            ValueError: If format is invalid
        """
        valid_types = {'cls', 'mod', 'frm', 'exl'}
        
        # Parse the path
        if '/' not in onefile_arg:
            raise ValueError(
                f"Invalid format: '{onefile_arg}'\n"
                "Expected format: 'type/ComponentName' or 'type/ComponentName.ext'\n"
                "Example: 'cls/MyClass' or 'mod/MyModule.bas'"
            )
        
        parts = onefile_arg.split('/', 1)
        component_type = parts[0].lower()
        filename = parts[1]
        
        # Validate type
        if component_type not in valid_types:
            raise ValueError(
                f"Invalid component type: '{component_type}'\n"
                f"Valid types: {', '.join(sorted(valid_types))}"
            )
        
        # Get extension
        type_extensions = {
            'cls': '.cls',
            'mod': '.bas',
            'frm': '.frm',
            'exl': '.cls'
        }
        
        if '.' in filename:
            component_name = filename.rsplit('.', 1)[0]
            extension = '.' + filename.rsplit('.', 1)[1]
        else:
            component_name = filename
            extension = type_extensions[component_type]
        
        return component_type, component_name, extension
    
    @staticmethod
    def validate_component_exists(
        src_path: Path,
        component_type: str,
        component_name: str,
        extension: str
    ) -> bool:
        """
        Check if a component file exists in src folder.
        
        Args:
            src_path: Path to src folder
            component_type: Type folder (cls, mod, frm, exl)
            component_name: Component name
            extension: File extension
            
        Returns:
            True if exists
        """
        file_path = src_path / component_type / f"{component_name}{extension}"
        return file_path.exists()
    
    @staticmethod
    def create_src_structure(src_path: Path, force: bool = False) -> None:
        """
        Create or reset the src folder structure.
        
        Args:
            src_path: Path to src folder
            force: If True, delete and recreate; if False, just ensure exists
        """
        if force and src_path.exists():
            FileHandler.delete_folder(src_path)
        
        for type_folder in ['cls', 'mod', 'frm', 'exl']:
            FileHandler.ensure_folder(src_path / type_folder)
    
    @staticmethod
    def delete_file(file_path: Path) -> None:
        """
        Delete a single file.
        
        Args:
            file_path: Path to file to delete
        """
        if file_path.exists():
            file_path.unlink()


if __name__ == "__main__":
    # Test usage
    from pathlib import Path
    
    # Test parse_onefile_path
    test_cases = [
        'cls/MyClass',
        'mod/MyModule.bas',
        'frm/MyForm',
        'exl/ThisWorkbook.cls'
    ]
    
    for test in test_cases:
        try:
            type_, name, ext = FileHandler.parse_onefile_path(test)
            print(f"'{test}' -> type='{type_}', name='{name}', ext='{ext}'")
        except ValueError as e:
            print(f"'{test}' -> ERROR: {e}")
