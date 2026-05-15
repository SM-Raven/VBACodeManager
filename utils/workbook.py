"""
Workbook Manager - Handles Excel workbook COM automation and validation
Reusable across export, import, and format commands
"""

from typing import Optional, Tuple
from win32com.client import Dispatch, GetObject
import win32com.client as win32


class WorkbookManager:
    """Manages Excel workbook COM interactions"""

    @staticmethod
    def get_excel_application():
        """
        Get the Excel application object (COM).

        Returns:
            Excel.Application COM object

        Raises:
            RuntimeError: If Excel is not running
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
    def get_open_workbooks():
        """
        Get list of all open workbooks.

        Returns:
            List of Excel workbook COM objects
        """
        try:
            excel = WorkbookManager.get_excel_application()
            return list(excel.Workbooks)
        except Exception as e:
            raise RuntimeError(f"Failed to get open workbooks: {str(e)}") from e

    @staticmethod
    def validate_single_workbook():
        """
        Ensure exactly one workbook is open.

        Returns:
            The single open workbook COM object

        Raises:
            ValueError: If zero or multiple workbooks are open
        """
        try:
            workbooks = WorkbookManager.get_open_workbooks()

            if len(workbooks) == 0:
                raise ValueError(
                    "❌ Error: No Excel workbook is currently open.\n"
                    "Please open an Excel file (.xlsm, .xla, or .xls) before running this command."
                )

            if len(workbooks) > 1:
                names = ", ".join([wb.Name for wb in workbooks])
                raise ValueError(
                    f"❌ Error: {len(workbooks)} workbooks are currently open.\n"
                    f"Open workbooks: {names}\n"
                    "VCM requires exactly one workbook to be open.\n"
                    "Please close all but one workbook and try again."
                )

            return workbooks[0]
        except ValueError:
            raise
        except Exception as e:
            raise RuntimeError(f"Failed to validate workbooks: {str(e)}") from e

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

    @staticmethod
    def is_workbook_supported(full_path: str) -> bool:
        """
        Check if workbook has supported extension.

        Args:
            full_path: Full path to workbook file

        Returns:
            True if supported (.xlsm, .xla, .xls)
        """
        supported = (".xlsm", ".xla", ".xls")
        return full_path.lower().endswith(supported)


if __name__ == "__main__":
    # Test
    try:
        wb = WorkbookManager.validate_single_workbook()
        name, path, ext = WorkbookManager.get_workbook_info(wb)
        print(f"Workbook: {name}")
        print(f"Path: {path}")
        print(f"Extension: {ext}")
    except Exception as e:
        print(f"Error: {e}")
