"""
Workbook Manager - Handles Excel workbook COM automation and validation
Enforces single supported workbook in the folder where VCM is called
"""

from pathlib import Path
from typing import Tuple
from win32com.client import Dispatch, GetObject

SUPPORTED_EXTENSIONS = (".xlsm", ".xla", ".xls")


class WorkbookManager:
    """Manages Excel workbook COM interactions with single workbook validation"""

    @staticmethod
    def get_call_folder() -> str:
        """
        Get the folder where VCM is called from.

        Returns:
            Path to the current working directory
        """
        return str(Path.cwd())

    @staticmethod
    def get_excel_application():
        """
        Get the Excel application object (COM).

        Returns:
            Excel.Application COM object

        Raises:
            RuntimeError: If Excel is not running or not installed
        """
        try:
            # Try to get running Excel instance
            excel = GetObject(Class="Excel.Application")
            return excel
        except Exception:
            try:
                # If not running, create new instance
                excel = Dispatch("Excel.Application")
                return excel
            except Exception as e:
                raise RuntimeError(
                    "Excel is not running or not installed.\n"
                    "Please start Excel and try again."
                ) from e

    @staticmethod
    def _is_in_call_folder(workbook) -> bool:
        """
        Check if workbook is in the folder where VCM was called.

        Args:
            workbook: Excel workbook COM object

        Returns:
            True if workbook is in call folder, False otherwise
        """
        try:
            workbook_path = Path(workbook.FullName)
            call_folder = Path(WorkbookManager.get_call_folder())
            return workbook_path.parent == call_folder
        except Exception:
            return False

    @staticmethod
    def _is_supported(workbook) -> bool:
        """
        Check if workbook has a supported extension.

        Args:
            workbook: Excel workbook COM object

        Returns:
            True if workbook extension is supported
        """
        try:
            full_path = workbook.FullName
            return full_path.lower().endswith(SUPPORTED_EXTENSIONS)
        except Exception:
            return False

    @staticmethod
    def validate_single_workbook():
        """
        Validate exactly one supported workbook is open in the call folder.

        Returns:
            The single open workbook COM object

        Raises:
            ValueError: If validation fails
            RuntimeError: If unable to get workbooks
        """
        try:
            all_workbooks = WorkbookManager.get_excel_application().Workbooks
            call_folder = WorkbookManager.get_call_folder()

            # Filter: workbooks in call folder with supported extensions
            valid_workbooks = [
                wb for wb in all_workbooks
                if WorkbookManager._is_in_call_folder(wb)
                and WorkbookManager._is_supported(wb)
            ]

            if len(valid_workbooks) == 0:
                raise ValueError(
                    f"❌ Error: No supported workbook open in this folder.\n"
                    f"Folder: {call_folder}\n"
                    f"Supported formats: {', '.join(SUPPORTED_EXTENSIONS)}"
                )

            if len(valid_workbooks) > 1:
                names = ", ".join([wb.Name for wb in valid_workbooks])
                raise ValueError(
                    f"❌ Error: {len(valid_workbooks)} workbooks are open in this folder.\n"
                    f"Open: {names}\n"
                    f"Only one workbook allowed. Please close the extras and try again."
                )

            return valid_workbooks[0]

        except ValueError:
            raise
        except Exception as e:
            raise RuntimeError(f"Failed to validate workbook: {str(e)}") from e

    @staticmethod
    def get_workbook_info(workbook) -> Tuple[str, str, str]:
        """
        Get information about a workbook.

        Args:
            workbook: Excel workbook COM object

        Returns:
            Tuple of (name, full_path, extension)
        """
        try:
            name = workbook.Name
            full_path = workbook.FullName
            extension = full_path.split(".")[-1].lower()
            return name, full_path, extension
        except Exception as e:
            raise RuntimeError(f"Failed to get workbook info: {str(e)}") from e

    @staticmethod
    def get_vba_project(workbook):
        """
        Get the VBA project from a workbook.

        Args:
            workbook: Excel workbook COM object

        Returns:
            VBProject object

        Raises:
            ValueError: If VBA project is not accessible
        """
        try:
            vba_project = workbook.VBProject
            if vba_project is None:
                raise ValueError(
                    f"Cannot access VBA project in '{workbook.Name}'.\n"
                    "Ensure the workbook is not read-only and has macros enabled."
                )
            return vba_project
        except Exception as e:
            raise ValueError(
                f"Cannot access VBA project in '{workbook.Name}'.\n"
                f"Error: {str(e)}"
            ) from e
