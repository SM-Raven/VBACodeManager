"""
VBA Extractor - Reusable VBA component extraction logic
Used by both export command (all and onefile variants)
"""

from typing import List, Optional, Tuple
from enum import Enum


class ComponentType(Enum):
    """VBA component types with their folder and file extension mappings"""
    CLASS = ("cls", ".cls")
    MODULE = ("mod", ".bas")
    FORM = ("frm", ".frm")
    EXCEL_OBJECT = ("exl", ".cls")

    def get_folder(self) -> str:
        return self.value[0]

    def get_extension(self) -> str:
        return self.value[1]


class VBAComponent:
    """Represents a single VBA component"""

    def __init__(self, name: str, code: str, component_type: ComponentType):
        self.name = name
        self.code = code
        self.type = component_type

    def get_folder(self) -> str:
        """Get target folder for this component"""
        return self.type.get_folder()

    def get_extension(self) -> str:
        """Get file extension for this component"""
        return self.type.get_extension()

    def get_filename(self) -> str:
        """Get full filename with extension"""
        return f"{self.name}{self.get_extension()}"

    def __repr__(self) -> str:
        return f"VBAComponent({self.name}, {self.type.name})"


class VBAExtractor:
    """Extracts VBA components from Excel workbooks"""

    # VBA component type codes from pywin32
    VBA_COMPONENT_TYPES = {
        1: ComponentType.MODULE,      # Standard module
        2: ComponentType.CLASS,       # Class module
        3: ComponentType.FORM,        # Form
        100: ComponentType.EXCEL_OBJECT,  # Worksheet Objects
    }

    def __init__(self, workbook):
        """
        Initialize extractor with workbook.

        Args:
            workbook: Excel workbook COM object
        """
        self.workbook = workbook
        self.workbook_name = workbook.Name

    def extract_all_components(self) -> List[VBAComponent]:
        """
        Extract all VBA components from workbook.

        Returns:
            List of VBAComponent objects

        Raises:
            ValueError: If VBA project is not accessible
        """
        try:
            vba_project = self.workbook.VBProject
            components = []

            for vba_component in vba_project.VBComponents:
                try:
                    component = self._extract_single_component(vba_component)
                    if component:
                        components.append(component)
                except Exception as e:
                    # Log but continue with other components
                    print(f"⚠️  Skipped component (error reading): {str(e)}")
                    continue

            return components

        except Exception as e:
            raise ValueError(
                f"Cannot access VBA components in '{self.workbook_name}'.\n"
                f"Ensure macros are enabled and file is not read-only.\n"
                f"Error: {str(e)}"
            ) from e

    def extract_component_by_name(self, component_name: str) -> Optional[VBAComponent]:
        """
        Extract a single component by name.

        Args:
            component_name: Name of the component to extract

        Returns:
            VBAComponent if found, None if not found

        Raises:
            ValueError: If VBA project is not accessible
        """
        try:
            vba_project = self.workbook.VBProject

            for vba_component in vba_project.VBComponents:
                if vba_component.Name == component_name:
                    return self._extract_single_component(vba_component)

            # Component not found
            return None

        except Exception as e:
            raise ValueError(
                f"Cannot access VBA components in '{self.workbook_name}'.\n"
                f"Error: {str(e)}"
            ) from e

    def _extract_single_component(self, vba_component) -> Optional[VBAComponent]:
        """
        Extract code from a single VBA component.

        Args:
            vba_component: pywin32 VBComponent object

        Returns:
            VBAComponent object or None if cannot extract
        """
        try:
            component_name = vba_component.Name

            # Get component type
            component_type = self.VBA_COMPONENT_TYPES.get(
                vba_component.Type,
                ComponentType.MODULE
            )

            # Extract code
            code_module = vba_component.CodeModule
            line_count = code_module.CountOfLines

            if line_count > 0:
                code = code_module.Lines(1, line_count)
            else:
                code = ""

            return VBAComponent(component_name, code, component_type)

        except Exception as e:
            raise RuntimeError(
                f"Failed to extract component '{vba_component.Name}': {str(e)}"
            ) from e

    def find_component(self, component_name: str) -> bool:
        """
        Check if a component exists in the workbook.

        Args:
            component_name: Name to search for

        Returns:
            True if found, False otherwise
        """
        try:
            vba_project = self.workbook.VBProject
            for component in vba_project.VBComponents:
                if component.Name == component_name:
                    return True
            return False
        except Exception:
            return False

    def list_components(self) -> List[str]:
        """
        Get list of all component names in workbook.

        Returns:
            List of component names
        """
        try:
            vba_project = self.workbook.VBProject
            return [comp.Name for comp in vba_project.VBComponents]
        except Exception:
            return []