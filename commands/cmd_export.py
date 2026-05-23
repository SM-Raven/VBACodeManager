import typer
from typing import Optional
from utils.workbook import get_active_workbook
from utils.file_system import FileSystem
from constants import (
    ComponentType,
    EXPORT_EXTENSIONS,
    EXPORT_FOLDERS,
    SUPPORTED_COMPONENT_TYPES,
)

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
        vcm export                          (skip existing, export new)
        vcm export --onefile cls/MyClass    (ask user if file exists)
        vcm export --force                  (overwrite all - deletes src folder first)
        vcm export --onefile cls/MyClass --force (overwrite immediately)
    """

    try:

        fs = FileSystem()

        # Delete src folder if force is True and onefile is not specified
        if force and not onefile:
            fs.reset_src()
        else:
            fs.ensure_structure()

        workbook_project = get_active_workbook()
        vb_components = workbook_project.VBProject.VBComponents

        exported_count = 0
        skipped_count = 0
        failed_count = 0
        target_component = None

        if onefile:
            target_component = onefile.replace("\\", "/").strip()

        for component in vb_components:
            component_type = component.Type
            component_name = component.Name

            if component_type not in SUPPORTED_COMPONENT_TYPES:
                continue

            component_type = ComponentType(component_type)
            folder_name = EXPORT_FOLDERS[component_type]
            extension = EXPORT_EXTENSIONS[component_type]
            component_identifier = f"{folder_name}/{component_name}"

            if target_component and component_identifier != target_component:
                continue

            export_path = fs.src_dir / folder_name / f"{component_name}{extension}"

            if export_path.exists():
                if onefile and not force:
                    if not typer.confirm(
                        f"{component_identifier} already exists. Overwrite and potentially lose progress?"):
                        typer.echo(f"Skipped: {component_identifier}")
                        skipped_count += 1
                        continue
                elif not onefile and not force:
                    typer.echo(f"Skipped existing: {component_identifier}")
                    skipped_count += 1
                    continue

            try:
                code_module = component.CodeModule
                total_lines = code_module.CountOfLines
                code = clean_vba_code(code_module.Lines(1, total_lines))

                with open(export_path, "w", encoding="utf-8") as file:
                    file.write(code)

                typer.echo(f"Exported: {component_identifier}")
                exported_count += 1

            except Exception as e:
                typer.echo(
                    f"Failed to export {component_name}: {e}",
                    err=True,
                )
                failed_count += 1

        if target_component and exported_count == 0 and skipped_count == 0:
            typer.echo(
                f"Component not found: {target_component}",
                err=True,
            )
            raise typer.Exit(code=1)

        typer.echo(
            f"\nExport complete | Exported: {exported_count} | Skipped: {skipped_count} | Failed: {failed_count}\n")

        if failed_count > 0:
            raise typer.Exit(code=1)

    except Exception as e:
        typer.echo(f"Export failed: {e}", err=True)
        raise typer.Exit(code=1)

def clean_vba_code(code: str) -> str:
    lines = code.splitlines()

    cleaned = []
    prev_blank = True  # important: prevents leading blank lines

    for line in lines:
        is_blank = line == ""

        if is_blank:
            if not prev_blank:
                cleaned.append("")
            prev_blank = True
        else:
            cleaned.append(line)
            prev_blank = False

    # remove trailing blank lines
    while cleaned and cleaned[-1] == "":
        cleaned.pop()

    return "\n".join(cleaned)

if __name__ == "__main__":
    app()