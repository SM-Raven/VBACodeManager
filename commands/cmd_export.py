"""
Export Command - Export VBA components from Excel workbook to src folder
Supports both:
  - vcm export (all components)
  - vcm export --onefile component_name (single component)
  - vcm export --force (overwrite all)
  - vcm export --onefile component_name --force (overwrite single)
"""

import typer
from pathlib import Path
from typing import Optional
from ..utils.workbook_manager import WorkbookManager
from ..utils.exceptions import (
    NoWorkbookOpenError,
    MultipleWorkbooksError,
    ComponentNotFoundError,
)
from ..vba.extractor import VBAExtractor
from ..utils.file_handler import FileHandler
from ..config import VCMConfig

app = typer.Typer(help="Export VBA components from Excel workbook")


@app.command()
def export_command(
    onefile: Optional[str] = typer.Option(
        None,
        "--onefile",
        "-o",
        help="Export only one component (format: cls/MyClass or mod/Module.bas)",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Overwrite existing components (or delete src folder if exporting all)",
    ),
):
    """
    Export VBA components from the currently open Excel workbook.
    
    By default, skips components that already exist in the src folder.
    Use --force to overwrite existing components.
    
    Examples:
        vcm export                              # Export all, skip existing
        vcm export --force                      # Export all, overwrite all
        vcm export --onefile cls/MyClass        # Export single component
        vcm export --onefile cls/MyClass -f     # Export single, overwrite
    """
    try:
        typer.echo("📁 Checking for open Excel workbook...")
        
        # Step 1: Get and validate workbook
        workbook = WorkbookManager.validate_single_workbook()
        wb_name, wb_path, wb_ext = WorkbookManager.get_workbook_info(workbook)
        
        typer.echo(f"✓ Found workbook: {wb_name}")
        
        # Step 2: Setup src folder structure
        config = VCMConfig()
        src_path = Path(config.src_folder)
        
        # If exporting all (not onefile), prepare folder structure
        if not onefile:
            if force:
                typer.echo("🗑️  Deleting existing src folder...")
                FileHandler.delete_folder(src_path)
            
            typer.echo("📂 Creating src folder structure...")
            FileHandler.create_src_structure(src_path, force=False)
            typer.echo("✓ Folder structure ready")
        
        # Step 3: Extract components
        typer.echo(f"📤 Exporting from {wb_name}...")
        extractor = VBAExtractor(workbook)
        
        if onefile:
            # Export single component
            _export_single_component(extractor, src_path, onefile, force)
        else:
            # Export all components
            _export_all_components(extractor, src_path, force)
    
    except (ValueError, TypeError, ComponentNotFoundError) as e:
        typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(code=1)
    except Exception as e:
        typer.echo(f"❌ Unexpected error: {str(e)}", err=True)
        raise typer.Exit(code=1)


def _export_single_component(
    extractor: VBAExtractor,
    src_path: Path,
    onefile: str,
    force: bool
) -> None:
    """
    Export a single component by name.
    
    Args:
        extractor: VBAExtractor instance
        src_path: Path to src folder
        onefile: Component specification (e.g., 'cls/MyClass')
        force: If True, overwrite existing
        
    Raises:
        ComponentNotFoundError: If component doesn't exist
    """
    # Parse the onefile argument
    try:
        component_type, component_name, extension = FileHandler.parse_onefile_path(onefile)
    except ValueError as e:
        raise TypeError(f"Invalid component specification: {str(e)}")
    
    # Extract the component
    component = extractor.extract_component_by_name(component_name)
    
    if component is None:
        # Component doesn't exist in workbook
        typer.echo(
            f"❌ Component '{component_name}' not found in workbook.\n"
            f"Available components:\n"
        )
        
        available = extractor.list_components()
        if available:
            for name in sorted(available):
                typer.echo(f"  - {name}")
        else:
            typer.echo("  (no components found)")
        
        raise ComponentNotFoundError(component_name, extractor.workbook_name)
    
    # Write component to file
    target_folder = src_path / component.get_folder()
    target_file = target_folder / f"{component_name}{extension}"
    
    # Check if already exists
    if target_file.exists() and not force:
        typer.echo(
            f"❌ Component already exists: {target_file.relative_to(Path.cwd())}\n"
            f"Use --force to overwrite."
        )
        raise typer.Exit(code=1)
    
    # Ensure folder structure
    FileHandler.ensure_folder(target_folder)
    
    # Write the file
    FileHandler.write_component(target_folder, component_name, component.code, extension)
    
    action = "Overwrote" if target_file.exists() else "Exported"
    typer.echo(f"✓ {action}: {component_name}")
    typer.echo(f"  Location: {target_file.relative_to(Path.cwd())}")


def _export_all_components(
    extractor: VBAExtractor,
    src_path: Path,
    force: bool
) -> None:
    """
    Export all components from workbook.
    
    Args:
        extractor: VBAExtractor instance
        src_path: Path to src folder
        force: If True, overwrite all existing
    """
    # Extract all components
    components = extractor.extract_all_components()
    
    if not components:
        typer.echo("⚠️  No VBA components found in workbook")
        return
    
    # Export each component
    skipped = 0
    exported = 0
    
    for component in components:
        target_folder = src_path / component.get_folder()
        target_file = target_folder / component.get_filename()
        
        # Ensure folder exists
        FileHandler.ensure_folder(target_folder)
        
        if target_file.exists() and not force:
            # Skip existing
            typer.echo(f"  ⊘ Skipped: {component.name} (already exists)")
            skipped += 1
        else:
            # Write component
            FileHandler.write_component(
                target_folder,
                component.name,
                component.code,
                component.get_extension()
            )
            action = "Overwrote" if target_file.exists() else "Exported"
            typer.echo(f"  ✓ {action}: {component.name}")
            exported += 1
    
    # Summary
    typer.echo("")
    typer.echo("="*50)
    typer.echo(f"✅ Export completed!")
    typer.echo(f"   Exported: {exported} components")
    if skipped > 0:
        typer.echo(f"   Skipped:  {skipped} components")
    typer.echo(f"   Location: {src_path.absolute()}")
    typer.echo("="*50)


if __name__ == "__main__":
    app()
