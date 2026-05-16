"""
Export Command - Export VBA components from Excel workbook to src folder

Supports:
    vcm export
    vcm export --onefile cls/MyClass
    vcm export --force
    vcm export --onefile cls/MyClass --force
"""

import shutil
from pathlib import Path
from typing import Optional

import typer

from constants import (
    ComponentType,
    EXPORT_EXTENSIONS,
    EXPORT_FOLDERS,
    SUPPORTED_COMPONENT_TYPES,
)
from utils.workbook import get_active_workbook

app = typer.Typer(help="Export VBA components from Excel workbook")


@app.command(name="export")
def export_command(
    onefile: Optional[str] = typer.Option(
        None,
        "--onefile",
        "-o",
        help="Export only one component (example: cls/MyClass)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Overwrite existing files",
    ),
):
    """
    Export Command - Export VBA components from Excel workbook to src folder

    Supports:
        vcm export
        vcm export --onefile cls/MyClass
        vcm export --force
        vcm export --onefile cls/MyClass --force
    """

    try:
        wb = get_active_workbook()
        vb_components = wb.VBProject.VBComponents

        src_dir = Path.cwd() / "src"

        # Clear src folder if exporting all with --force
        if force and onefile is None and src_dir.exists():
            shutil.rmtree(src_dir)

        src_dir.mkdir(parents=True, exist_ok=True)

        exported_count = 0
        skipped_count = 0

        target_component = None

        if onefile:
            target_component = onefile.replace("\\", "/").strip()

        for component in vb_components:

            component_type = component.Type
            component_name = component.Name

            # Skip unsupported component types
            if component_type not in SUPPORTED_COMPONENT_TYPES:
                continue

            component_type = ComponentType(component_type)

            folder_name = EXPORT_FOLDERS[component_type]
            extension = EXPORT_EXTENSIONS[component_type]

            component_identifier = f"{folder_name}/{component_name}"

            # Skip if not matching --onefile target
            if target_component and component_identifier != target_component:
                continue

            export_folder = src_dir / folder_name
            export_folder.mkdir(parents=True, exist_ok=True)

            export_path = export_folder / f"{component_name}{extension}"

            # Skip existing unless --force
            if export_path.exists() and not force:
                typer.echo(f"⏭️ Skipped existing: {component_identifier}")
                skipped_count += 1
                continue

            try:

                # Standard modules, classes, forms
                if component_type in {
                    ComponentType.STANDARD_MODULE,
                    ComponentType.CLASS_MODULE,
                    ComponentType.USER_FORM,
                }:
                    component.Export(str(export_path))

                # Document modules (ThisWorkbook, Sheet1, etc.)
                elif component_type == ComponentType.DOCUMENT_MODULE:

                    code_module = component.CodeModule
                    total_lines = code_module.CountOfLines

                    code = code_module.Lines(1, total_lines)

                    with open(export_path, "w", encoding="utf-8") as file:
                        file.write(code)

                typer.echo(f"✅ Exported: {component_identifier}")
                exported_count += 1

            except Exception as e:
                typer.echo(
                    f"❌ Failed to export {component_name}: {e}",
                    err=True,
                )

        # No component matched --onefile
        if target_component and exported_count == 0 and skipped_count == 0:
            typer.echo(
                f"❌ Component not found: {target_component}",
                err=True,
            )
            raise typer.Exit(code=1)

        typer.echo(
            f"\n📦 Export complete | Exported: {exported_count} | Skipped: {skipped_count}"
        )

    except Exception as e:
        typer.echo(f"❌ Export failed: {e}", err=True)
        raise typer.Exit(code=1)