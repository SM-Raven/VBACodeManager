"""
VBA Injector - Reusable VBA component injection logic
Used by both import command (all and onefile variants)
"""

from typing import Optional, List
from .extractor_template import ComponentType, VBAComponent


class VBAInjector:
    """Injects VBA components into Excel workbooks"""
    
    def __init__(self, workbook):
        """
        Initialize injector with workbook.
        
        Args:
            workbook: Excel workbook COM object
        """
        self.workbook = workbook
        self.workbook_name = workbook.Name
    
    def inject_component(
        self,
        vba_component: VBAComponent,
        overwrite: bool = True
    ) -> bool:
        """
        Add or update a VBA component in the workbook.
        
        For excel objects (exl type), only updates code.
        For other types, adds new or replaces existing.
        
        Args:
            vba_component: VBAComponent object to inject
            overwrite: If True, replace existing; if False, skip existing
            
        Returns:
            True if successfully injected, False if skipped
            
        Raises:
            ValueError: If injection fails
        """
        try:
            vba_project = self.workbook.VBProject
            component_name = vba_component.name
            is_excel_object = vba_component.type == ComponentType.EXCEL_OBJECT
            
            # Check if component exists
            existing_component = None
            for comp in vba_project.VBComponents:
                if comp.Name == component_name:
                    existing_component = comp
                    break
            
            if existing_component:
                if is_excel_object:
                    # For Excel objects, only update code
                    self._update_component_code(existing_component, vba_component.code)
                    return True
                elif overwrite:
                    # Replace existing non-Excel component
                    self._remove_component(existing_component)
                    self._add_component(vba_project, vba_component)
                    return True
                else:
                    # Skip existing component
                    return False
            else:
                # Component doesn't exist - add it
                if not is_excel_object:
                    # Can only add non-Excel objects
                    self._add_component(vba_project, vba_component)
                    return True
                else:
                    # Cannot add Excel objects - they must exist
                    raise ValueError(
                        f"Cannot create Excel object '{component_name}'.\n"
                        "Excel objects (ThisWorkbook, Sheets, etc.) cannot be created via VBA.\n"
                        "Component was skipped."
                    )
        
        except ValueError:
            raise
        except Exception as e:
            raise ValueError(
                f"Failed to inject component '{vba_component.name}': {str(e)}"
            ) from e
    
    def inject_code_only(self, component_name: str, code: str) -> bool:
        """
        Update only the code of an existing component (for Excel objects).
        
        Does not create the component if it doesn't exist.
        
        Args:
            component_name: Name of component to update
            code: New VBA code
            
        Returns:
            True if updated, False if component not found
            
        Raises:
            ValueError: If update fails
        """
        try:
            vba_project = self.workbook.VBProject
            
            for component in vba_project.VBComponents:
                if component.Name == component_name:
                    self._update_component_code(component, code)
                    return True
            
            # Component not found
            return False
        
        except Exception as e:
            raise ValueError(
                f"Failed to update code for '{component_name}': {str(e)}"
            ) from e
    
    def remove_component(self, component_name: str) -> bool:
        """
        Remove a component from the workbook.
        
        Never removes Excel objects (ThisWorkbook, Sheets, etc.).
        
        Args:
            component_name: Name of component to remove
            
        Returns:
            True if removed, False if not found or is Excel object
            
        Raises:
            ValueError: If removal fails
        """
        try:
            vba_project = self.workbook.VBProject
            
            for component in vba_project.VBComponents:
                if component.Name == component_name:
                    # Never delete Excel objects
                    if self._is_excel_object_type(component.Type):
                        return False
                    
                    self._remove_component(component)
                    return True
            
            return False
        
        except Exception as e:
            raise ValueError(
                f"Failed to remove component '{component_name}': {str(e)}"
            ) from e
    
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
    
    # ==================== Private Helper Methods ====================
    
    def _add_component(self, vba_project, vba_component: VBAComponent):
        """
        Add a new VBA component to the project.
        
        Args:
            vba_project: VBProject object
            vba_component: VBAComponent to add
        """
        # Determine the type code for pywin32
        type_code = self._get_component_type_code(vba_component.type)
        
        # Add the component
        new_component = vba_project.VBComponents.Add(type_code)
        new_component.Name = vba_component.name
        
        # Set the code
        code_module = new_component.CodeModule
        code_module.AddFromString(vba_component.code)
    
    def _remove_component(self, component):
        """Remove a component from the project"""
        vba_project = self.workbook.VBProject
        vba_project.VBComponents.Remove(component)
    
    def _update_component_code(self, component, new_code: str):
        """
        Replace all code in a component.
        
        Args:
            component: pywin32 VBComponent object
            new_code: New code to set
        """
        code_module = component.CodeModule
        
        # Delete all existing lines
        if code_module.CountOfLines > 0:
            code_module.DeleteLines(1, code_module.CountOfLines)
        
        # Add new code
        if new_code.strip():
            code_module.AddFromString(new_code)
    
    def _get_component_type_code(self, component_type: ComponentType) -> int:
        """
        Get pywin32 VBComponent type code for a ComponentType.
        
        Args:
            component_type: ComponentType enum value
            
        Returns:
            Type code for pywin32 VBComponents.Add()
        """
        type_mapping = {
            ComponentType.MODULE: 1,      # vbext_ct_StdModule
            ComponentType.CLASS: 2,       # vbext_ct_ClassModule
            ComponentType.FORM: 3,        # vbext_ct_Document (Form)
        }
        
        if component_type not in type_mapping:
            raise ValueError(
                f"Cannot create component of type {component_type.name}"
            )
        
        return type_mapping[component_type]
    
    def _is_excel_object_type(self, type_code: int) -> bool:
        """
        Check if a component type code represents an Excel object.
        
        Args:
            type_code: pywin32 VBComponent Type value
            
        Returns:
            True if it's an Excel object (cannot be created/deleted)
        """
        excel_object_types = [100, 120, 13]  # ThisWorkbook, Sheets, Documents
        return type_code in excel_object_types


if __name__ == "__main__":
    # Test usage
    from workbook_template import WorkbookManager
    from extractor_template import VBAExtractor
    
    try:
        wb = WorkbookManager.validate_single_workbook()
        
        # Extract a component
        extractor = VBAExtractor(wb)
        components = extractor.extract_all_components()
        
        if components:
            # Try to inject it
            injector = VBAInjector(wb)
            test_component = components[0]
            
            # Example: Update code only
            new_code = test_component.code + "\n' Added by injector"
            injector.inject_code_only(test_component.name, new_code)
            
            print(f"Injected code into: {test_component.name}")
    
    except Exception as e:
        print(f"Error: {e}")
