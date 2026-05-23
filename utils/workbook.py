import win32com.client
import typer
from pathlib import Path
from typing import Any
from constants import SUPPORTED_VBA_PROJECTS


def get_active_workbook() -> Any:
    """
        Returns instance of valid VBA Project Workbook.
        Error when multiple workbook is open in current directory.
        Error when no open workbook.
        This ensure to return one valid Worbook Open in current directory.
    """

    try:
        excel: Any = win32com.client.GetObject(Class="Excel.Application")
    except Exception:
        typer.echo("Excel is not running")
        raise typer.Exit(code=1)

    root_directory: Path = Path.cwd().resolve()
    match: Any | None = None

    for workbook in excel.Workbooks:
        try:
            wb_path: Path = Path(workbook.FullName).resolve()
        except Exception:
            continue

        if wb_path.suffix.lower() not in SUPPORTED_VBA_PROJECTS:
            continue

        if wb_path.parent == root_directory:
            if match is not None:
                typer.echo("Multiple workbooks open in this directory.")
                raise typer.Exit(code=1)
            match = workbook

    if match is None:
        typer.echo("No supported workbook open in current directory")
        raise typer.Exit(code=1)

    return match