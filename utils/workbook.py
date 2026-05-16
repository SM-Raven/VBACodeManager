from pathlib import Path
from constants import SUPPORTED_VBA_PROJECTS
import win32com.client
import typer

def get_active_workbook():
    try:
        excel = win32com.client.GetObject(Class="Excel.Application")
    except Exception:
        typer.echo("❌ Excel is not running")
        raise typer.Exit(code=1)

    cwd = Path.cwd().resolve()
    match = None

    for wb in excel.Workbooks:
        try:
            wb_path = Path(wb.FullName).resolve()
        except Exception:
            continue  # skip unsaved or invalid workbooks

        # 🔹 Validate file type
        if wb_path.suffix.lower() not in SUPPORTED_VBA_PROJECTS:
            continue

        # 🔹 Match current working directory
        if wb_path.parent == cwd:
            if match is not None:
                typer.echo("❌ Multiple workbooks open in this directory")
                raise typer.Exit(code=1)
            match = wb

    if match is None:
        typer.echo("❌ No supported workbook open in current directory")
        raise typer.Exit(code=1)

    return match