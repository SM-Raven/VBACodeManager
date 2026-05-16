import typer
from typing import Optional, List, Tuple
from utils.workbook import get_active_workbook
from utils.file_system import FileSystem
from constants import (
    ComponentType,
    EXPORT_EXTENSIONS,
    EXPORT_FOLDERS
)

app = typer.Typer(help="Import VBA components to Excel workbook")

@app.command(name="import")
def import_command(
    onefile: Optional[str] = typer.Option(
        None,
        "--onefile",
        "-o",
        help="Import only one component (example: cls/MyClass)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Delete components from workbook that don't exist in src",
    ),
):
    """
    Import VBA components from src folder to Excel workbook

    Supports:
        vcm import                          (import all components)
        vcm import --onefile cls/MyClass    (import one component)
        vcm import --force                  (import all, delete missing)
    """

    try:
        # Setup file system
        fs = FileSystem()

        # Validate src structure exists
        if not fs.is_structure_valid():
            typer.echo("❌ No src folder structure found. Run 'vcm export' first.", err=True)
            raise typer.Exit(code=1)

        # Get active workbook
        wb = get_active_workbook()

        imported_count = 0
        deleted_count = 0
        failed_count = 0
        target_component = None

        if onefile:
            target_component = onefile.replace("\\", "/").strip()

        # Determine which files to import
        if target_component:
            components_to_import = _get_single_component(fs, target_component)
        else:
            components_to_import = _get_all_components(fs)

        if not components_to_import:
            typer.echo("⚠️ No components found to import")
            return

        # Import components
        for folder_name, component_name, file_path in components_to_import:
            try:
                component_identifier = f"{folder_name}/{component_name}"

                # Determine component type from folder name
                component_type = _get_component_type_from_folder(folder_name)

                # Type 1, 2, 3: Normal import (create/add component)
                if component_type in {ComponentType.STANDARD_MODULE, ComponentType.CLASS_MODULE, ComponentType.USER_FORM}:
                    # Remove if component already exists
                    try:
                        existing = wb.VBProject.VBComponents(component_name)
                        wb.VBProject.VBComponents.Remove(existing)
                    except Exception:
                        pass

                    # Import component using VBA's built-in Import method
                    vb_component = wb.VBProject.VBComponents.Import(str(file_path))
                    typer.echo(f"✅ Imported: {component_identifier}")

                # Type 100: Document module (import code only)
                elif component_type == ComponentType.DOCUMENT_MODULE:
                    code = file_path.read_text(encoding="utf-8")
                    try:
                        vb_component = wb.VBProject.VBComponents(component_name)
                        vb_component.CodeModule.DeleteLines(1, vb_component.CodeModule.CountOfLines)
                        vb_component.CodeModule.AddFromString(code)
                        typer.echo(f"✅ Updated: {component_identifier}")
                    except Exception as e:
                        typer.echo(f"❌ Document module not found: {component_name}", err=True)
                        failed_count += 1
                        continue

                imported_count += 1

            except Exception as e:
                typer.echo(f"❌ Failed to import {component_name}: {e}", err=True)
                failed_count += 1

        # If force, delete missing components
        if force:
            deleted_count = _delete_missing_components(wb, fs)
            typer.echo(f"🗑️ Removed: {deleted_count} components")

        typer.echo(
            f"\n📦 Import complete | Imported: {imported_count} | Deleted: {deleted_count} | Failed: {failed_count}"
        )

    except Exception as e:
        typer.echo(f"❌ Import failed: {e}", err=True)
        raise typer.Exit(code=1)


def _get_single_component(fs: FileSystem, target: str) -> List[Tuple[str, str, object]]:
    """
    Get a single component file for import.

    Args:
        fs: FileSystem instance
        target: Component specification (e.g., 'cls/MyClass')

    Returns:
        List with single tuple (folder_name, component_name, file_path)
    """
    parts = target.split("/")
    if len(parts) != 2:
        typer.echo(f"❌ Invalid format: {target}. Use 'cls/MyClass'", err=True)
        raise typer.Exit(code=1)

    folder_name, component_name = parts

    # Validate folder name exists in EXPORT_FOLDERS
    if folder_name not in EXPORT_FOLDERS.values():
        typer.echo(f"❌ Invalid component type: {folder_name}", err=True)
        raise typer.Exit(code=1)

    # Find extension for this folder
    extension = None
    for ctype, fname in EXPORT_FOLDERS.items():
        if fname == folder_name:
            extension = EXPORT_EXTENSIONS[ctype]
            break

    file_path = fs.src_dir / folder_name / f"{component_name}{extension}"

    if not file_path.exists():
        typer.echo(f"❌ Component file not found: {file_path}", err=True)
        raise typer.Exit(code=1)

    return [(folder_name, component_name, file_path)]


def _get_all_components(fs: FileSystem) -> List[Tuple[str, str, object]]:
    """
    Get all components from src folder for import.

    Args:
        fs: FileSystem instance

    Returns:
        List of tuples (folder_name, component_name, file_path)
    """
    components = []

    for folder_name in EXPORT_FOLDERS.values():
        folder_path = fs.src_dir / folder_name

        if not folder_path.exists():
            continue

        for file_path in sorted(folder_path.iterdir()):
            if file_path.is_file():
                # Extract component name (remove extension)
                component_name = file_path.stem
                components.append((folder_name, component_name, file_path))

    return components


def _get_component_type_from_folder(folder_name: str) -> ComponentType:
    """
    Get ComponentType enum from folder name.

    Args:
        folder_name: Folder name like 'cls', 'mod', 'frm', 'doc'

    Returns:
        ComponentType enum value
    """
    for ctype, fname in EXPORT_FOLDERS.items():
        if fname == folder_name:
            return ctype

    raise ValueError(f"Invalid component type: {folder_name}")


def _delete_missing_components(wb, fs: FileSystem) -> int:
    """
    Delete components from workbook that don't exist in src folder.

    Never deletes Excel objects (ThisWorkbook, Sheets, etc.).

    Args:
        wb: Excel workbook object
        fs: FileSystem instance

    Returns:
        Number of components deleted
    """
    # Get component names that exist in src
    existing_in_src = set()

    for folder_name in EXPORT_FOLDERS.values():
        folder_path = fs.src_dir / folder_name

        if folder_path.exists():
            for file_path in folder_path.iterdir():
                if file_path.is_file():
                    existing_in_src.add(file_path.stem)

    # Get all components in workbook
    vb_components = wb.VBProject.VBComponents
    deleted = 0

    # Build list of components to remove (can't modify during iteration)
    to_remove = []
    for component in vb_components:
        if component.Name not in existing_in_src:
            # Skip Excel objects (documents) - type 100
            if component.Type != ComponentType.DOCUMENT_MODULE:
                to_remove.append(component)

    # Remove the components
    for component in to_remove:
        try:
            vb_components.Remove(component)
            typer.echo(f"✓ Removed: {component.Name}")
            deleted += 1
        except Exception as e:
            typer.echo(f"✗ Failed to remove {component.Name}: {e}", err=True)

    return deleted

if __name__ == "__main__":
    app()