"""
Custom exceptions for VCM
"""


class VCMException(Exception):
    """Base exception for all VCM errors"""
    pass


class NoWorkbookOpenError(VCMException):
    """Raised when no workbook is currently open in Excel"""
    def __init__(self):
        super().__init__(
            "❌ Error: No Excel workbook is currently open.\n"
            "Please open an Excel file (.xlsm, .xla, or .xls) before running this command."
        )


class MultipleWorkbooksError(VCMException):
    """Raised when multiple workbooks are open"""
    def __init__(self, count: int):
        super().__init__(
            f"❌ Error: {count} workbooks are currently open.\n"
            "VCM requires exactly one workbook to be open. "
            "Please close all but one workbook and try again."
        )


class UnsupportedExtensionError(VCMException):
    """Raised for unsupported Excel file types"""
    def __init__(self, filename: str):
        super().__init__(
            f"❌ Error: '{filename}' has an unsupported format.\n"
            "Supported formats: .xlsm, .xla, .xls"
        )


class ComponentNotFoundError(VCMException):
    """Raised when a VBA component is not found"""
    def __init__(self, component_name: str, workbook_name: str):
        super().__init__(
            f"❌ Error: Component '{component_name}' not found in '{workbook_name}'."
        )


class FileNotFoundError(VCMException):
    """Raised when a source file is not found"""
    def __init__(self, file_path: str):
        super().__init__(
            f"❌ Error: File not found: '{file_path}'"
        )


class InvalidComponentTypeError(VCMException):
    """Raised when component type is invalid"""
    def __init__(self, component_type: str):
        super().__init__(
            f"❌ Error: Invalid component type '{component_type}'.\n"
            "Valid types: cls, mod, frm, exl"
        )


class SourceFolderNotFoundError(VCMException):
    """Raised when src folder structure doesn't exist"""
    def __init__(self):
        super().__init__(
            "❌ Error: 'src' folder not found.\n"
            "Please run 'vcm export' first to create the source structure."
        )


class NoComponentsFoundError(VCMException):
    """Raised when no components are found to process"""
    def __init__(self, location: str = "src"):
        super().__init__(
            f"❌ Error: No VBA components found in '{location}' folder."
        )


class ExcelAPIError(VCMException):
    """Raised when Excel COM API fails"""
    def __init__(self, message: str):
        super().__init__(
            f"❌ Error: Failed to communicate with Excel.\n{message}\n"
            "Make sure Excel is running and pywin32 is properly installed."
        )
