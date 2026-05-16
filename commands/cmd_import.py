"""
Import Command - Import VBA components from src folder to Excel workbook
Supports:
  - vcm import (all components)
  - vcm import --onefile component_path (single component)
  - vcm import --force (remove missing components)
  - vcm import --onefile component_path --force (one file, remove if doesn't match)
"""

import typer
from pathlib import Path
from typing import Optional, List, Tuple
from utils.workbook import WorkbookManager
from utils.exceptions import (
    SourceFolderNotFoundError,
    FileNotFoundError,
)
from vba.injector import VBAInjector, VBAComponent, ComponentType
from utils.file_handler import FileHandler
from config import VCMConfig

app = typer.Typer(help="Import VBA components to Excel workbook")


@app.command()
def import_command(
    onefile: Optional[str] = typer.Option(
        None,
        "--onefile",
        "-o",
        help="Import only one component (format: cls/MyClass or mod/Module.bas)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Remove components from workbook that don't exist in src",
    ),
):
    """
    Import VBA components from src folder to the currently open Excel workbook.

    By default, adds or updates components in the workbook.
    Use --force to also remove components not present in the src folder.

    For excel objects (exl folder), only code is updated - objects are never deleted.

    Examples:
        vcm import                              # Import all components
        vcm import --onefile cls/MyClass        # Import one component
        vcm import --force                      # Import all, remove missing
        vcm import -o mod/Module1 -f            # One file + remove missing
    """
    try:
        typer.echo("📁 Checking for open Excel workbook...")

        # Step 1: Get and validate workbook
        workbook = WorkbookManager.validate_single_workbook()
        wb_name, wb_path, wb_ext = WorkbookManager.get_workbook_info(workbook)

        typer.echo(f"✓ Found workbook: {wb_name}")

        # Step 2: Check src folder exists
        config = VCMConfig()
        src_path = Path(config.src_folder)

        if not src_path.exists():
            raise SourceFolderNotFoundError()

        # Step 3: Determine what to import
        if onefile:
            components_to_import = _get_single_component_for_import(src_path, onefile)
        else:
            components_to_import = _get_all_components_for_import(src_path)

        if not components_to_import:
            typer.echo("⚠️  No components found to import")
            return

        # Step 4: Import components
        typer.echo(f"📥 Importing {len(components_to_import)} component(s)...")

        injector = VBAInjector(workbook)
        imported_count = _import_components_from_files(
            injector,
            components_to_import,
            force_overwrite=True
        )

        # Step 5: If force, delete missing components
        if force:
            typer.echo("🗑️  Cleaning up components not in src folder...")
            deleted_count = _delete_missing_components(injector, src_path, config)
            typer.echo(f"   Removed: {deleted_count} components")

        # Step 6: Summary
        typer.echo("")
        typer.echo("="*50)
        typer.echo(f"✅ Import completed!")
        typer.echo(f"   Imported: {imported_count} components")
        typer.echo("="*50)

    except (ValueError, SourceFolderNotFoundError, FileNotFoundError) as e:
        typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(code=1)
    except Exception as e:
        typer.echo(f"❌ Unexpected error: {str(e)}", err=True)
        raise typer.Exit(code=1)


# ==================== Reusable Import Functions ====================

def _get_single_component_for_import(
    src_path: Path,
    onefile: str
) -> List[Tuple[str, VBAComponent]]:
    """
    Get a single component file for import.

    Args:
        src_path: Path to src folder
        onefile: Component specification (e.g., 'cls/MyClass')

    Returns:
        List with single tuple (component_type, VBAComponent)

    Raises:
        FileNotFoundError: If component file doesn't exist
    """
    # Parse the argument
    try:
        component_type, component_name, extension = FileHandler.parse_onefile_path(onefile)
    except ValueError as e:
        raise ValueError(f"Invalid component specification: {str(e)}")

    # Build file path
    file_path = src_path / component_type / f"{component_name}{extension}"

    if not file_path.exists():
        typer.echo(
            f"❌ Component file not found: {file_path}\n\n"
            f"Available files in '{component_type}' folder:\n"
        )

        available_folder = src_path / component_type
        if available_folder.exists():
            for file in sorted(available_folder.iterdir()):
                typer.echo(f"  - {file.name}")
        else:
            typer.echo(f"  (folder '{component_type}' doesn't exist)")

        raise FileNotFoundError(str(file_path))

    # Read component code
    code = FileHandler.read_component(file_path)

    # Create VBAComponent
    type_enum = _parse_component_type(component_type)
    vba_component = VBAComponent(component_name, code, type_enum)

    return [(component_type, vba_component)]


def _get_all_components_for_import(src_path: Path) -> List[Tuple[str, VBAComponent]]:
    """
    Get all components from src folder for import.

    Args:
        src_path: Path to src folder

    Returns:
        List of tuples (component_type, VBAComponent)
    """
    components = []

    # Get all components from all subfolders
    all_files = FileHandler.get_all_components_in_src(src_path)

    for component_type, component_name, file_path in all_files:
        try:
            code = FileHandler.read_component(file_path)
            type_enum = _parse_component_type(component_type)
            vba_component = VBAComponent(component_name, code, type_enum)
            components.append((component_type, vba_component))
        except Exception as e:
            typer.echo(f"  ⚠️  Skipped {file_path.name}: {str(e)}", err=True)
            continue

    return components


def _import_components_from_files(
    injector: VBAInjector,
    components: List[Tuple[str, VBAComponent]],
    force_overwrite: bool = True
) -> int:
    """
    Import a list of VBA components into the workbook.

    This is the core reusable import logic.

    Args:
        injector: VBAInjector instance
        components: List of (component_type, VBAComponent) tuples
        force_overwrite: If True, overwrite existing components

    Returns:
        Number of successfully imported components
    """
    imported = 0

    for component_type, vba_component in components:
        try:
            # Inject the component
            success = injector.inject_component(vba_component, overwrite=force_overwrite)

            if component_type == "exl":
                action = "Updated code"
            else:
                action = "Imported" if success else "Skipped"

            if success or component_type == "exl":
                typer.echo(f"  ✓ {action}: {vba_component.name}")
                imported += 1
            else:
                typer.echo(f"  ⊘ {action}: {vba_component.name} (already exists)")

        except ValueError as e:
            typer.echo(f"  ✗ Failed {vba_component.name}: {str(e)}", err=True)
        except Exception as e:
            typer.echo(f"  ✗ Error with {vba_component.name}: {str(e)}", err=True)

    return imported


def _delete_missing_components(
    injector: VBAInjector,
    src_path: Path,
    config: VCMConfig
) -> int:
    """
    Delete components from workbook that don't exist in src folder.

    Never deletes Excel objects (ThisWorkbook, Sheets, etc.).

    Args:
        injector: VBAInjector instance
        src_path: Path to src folder
        config: VCMConfig instance

    Returns:
        Number of components deleted
    """
    # Get component names that exist in src
    existing_in_src = set()

    for component_type, component_name, file_path in FileHandler.get_all_components_in_src(src_path):
        existing_in_src.add(component_name)

    # Get all components in workbook
    workbook_components = injector.list_components()

    deleted = 0

    for component_name in workbook_components:
        if component_name not in existing_in_src:
            # Try to remove it
            try:
                removed = injector.remove_component(component_name)
                if removed:
                    typer.echo(f"  ✓ Removed: {component_name}")
                    deleted += 1
                else:
                    # Not removed because it's an Excel object
                    typer.echo(f"  ⊘ Kept: {component_name} (Excel object)")
            except Exception as e:
                typer.echo(f"  ✗ Failed to remove {component_name}: {str(e)}", err=True)

    return deleted


# ==================== Helper Functions ====================

def _parse_component_type(type_str: str) -> ComponentType:
    """
    Convert string component type to ComponentType enum.

    Args:
        type_str: String like 'cls', 'mod', 'frm', 'exl'

    Returns:
        ComponentType enum value

    Raises:
        ValueError: If invalid type
    """
    mapping = {
        'cls': ComponentType.CLASS,
        'mod': ComponentType.MODULE,
        'frm': ComponentType.FORM,
        'exl': ComponentType.EXCEL_OBJECT,
    }

    if type_str not in mapping:
        raise ValueError(f"Invalid component type: '{type_str}'")

    return mapping[type_str]


if __name__ == "__main__":
    app()
