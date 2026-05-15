"""
Format Command - Format VBA code in src folder
Supports:
  - vcm format --all (format all components)
  - vcm format --onefile component_path (format single component)
"""

import typer
from pathlib import Path
from typing import Optional, List, Tuple
from ..utils.file_handler import FileHandler
from ..vba.formatter import VBAFormatter
from ..config import VCMConfig

app = typer.Typer(help="Format VBA code in src folder")


@app.command()
def format_command(
    all: bool = typer.Option(
        False,
        "--all",
        "-a",
        help="Format all components in src folder",
    ),
    onefile: Optional[str] = typer.Option(
        None,
        "--onefile",
        "-o",
        help="Format only one component (format: cls/MyClass or mod/Module.bas)",
    ),
):
    """
    Format VBA code in the src folder.
    
    Formats either all components or a single component.
    Uses consistent indentation, spacing, and style rules.
    
    Examples:
        vcm format --all                    # Format all components
        vcm format --onefile cls/MyClass    # Format one component
    """
    try:
        config = VCMConfig()
        src_path = Path(config.src_folder)
        
        # Validate src folder exists
        if not src_path.exists():
            raise ValueError(
                "❌ Error: 'src' folder not found.\n"
                "Please run 'vcm export' first to create the source structure."
            )
        
        typer.echo("🎨 Starting VBA code formatting...")
        
        if all and onefile:
            raise ValueError(
                "❌ Error: Cannot use both --all and --onefile.\n"
                "Choose one: 'vcm format --all' or 'vcm format --onefile cls/MyClass'"
            )
        
        if not all and not onefile:
            raise ValueError(
                "❌ Error: Please specify either --all or --onefile.\n"
                "Examples:\n"
                "  vcm format --all\n"
                "  vcm format --onefile cls/MyClass"
            )
        
        if all:
            formatted_count = _format_all_components(src_path)
        else:
            formatted_count = _format_single_component(src_path, onefile)
        
        # Summary
        typer.echo("")
        typer.echo("="*50)
        typer.echo(f"✅ Formatting completed!")
        typer.echo(f"   Formatted: {formatted_count} component(s)")
        typer.echo("="*50)
    
    except ValueError as e:
        typer.echo(f"❌ {str(e)}", err=True)
        raise typer.Exit(code=1)
    except Exception as e:
        typer.echo(f"❌ Unexpected error: {str(e)}", err=True)
        raise typer.Exit(code=1)


def _format_all_components(src_path: Path) -> int:
    """
    Format all VBA components in src folder.
    
    Args:
        src_path: Path to src folder
        
    Returns:
        Number of components formatted
    """
    # Get all components using reusable FileHandler
    all_components = FileHandler.get_all_components_in_src(src_path)
    
    if not all_components:
        typer.echo("⚠️  No VBA components found in src folder")
        return 0
    
    formatted = 0
    
    for component_type, component_name, file_path in all_components:
        try:
            # Read component
            code = FileHandler.read_component(file_path)
            
            # Format code
            formatted_code = VBAFormatter.format_code(code)
            
            # Check if code changed
            if formatted_code != code:
                # Write formatted code back
                FileHandler.write_component(
                    file_path.parent,
                    component_name,
                    formatted_code,
                    file_path.suffix
                )
                typer.echo(f"  ✓ Formatted: {component_name}")
            else:
                typer.echo(f"  ⊘ Already formatted: {component_name}")
            
            formatted += 1
        
        except Exception as e:
            typer.echo(f"  ✗ Failed to format {component_name}: {str(e)}", err=True)
    
    return formatted


def _format_single_component(src_path: Path, onefile: str) -> int:
    """
    Format a single VBA component.
    
    Args:
        src_path: Path to src folder
        onefile: Component specification (e.g., 'cls/MyClass')
        
    Returns:
        Number of components formatted (1 or 0)
    """
    # Parse the argument using reusable FileHandler
    try:
        component_type, component_name, extension = FileHandler.parse_onefile_path(onefile)
    except ValueError as e:
        raise ValueError(f"Invalid component specification: {str(e)}")
    
    # Build file path
    file_path = src_path / component_type / f"{component_name}{extension}"
    
    # Check if file exists
    if not file_path.exists():
        typer.echo(
            f"❌ Component file not found: {file_path}\n\n"
            f"Available files in '{component_type}' folder:\n"
        )
        
        available_folder = src_path / component_type
        if available_folder.exists():
            for file in sorted(available_folder.iterdir()):
                if file.is_file() and file.suffix in ['.bas', '.cls', '.frm']:
                    typer.echo(f"  - {file.name}")
        else:
            typer.echo(f"  (folder '{component_type}' doesn't exist)")
        
        raise ValueError(f"Component not found: {file_path}")
    
    try:
        # Read component
        code = FileHandler.read_component(file_path)
        
        # Format code
        formatted_code = VBAFormatter.format_code(code)
        
        # Check if code changed
        if formatted_code != code:
            # Write formatted code back
            FileHandler.write_component(
                file_path.parent,
                component_name,
                formatted_code,
                extension
            )
            typer.echo(f"  ✓ Formatted: {component_name}")
            typer.echo(f"  Location: {file_path.relative_to(Path.cwd())}")
            return 1
        else:
            typer.echo(f"  ⊘ Already formatted: {component_name}")
            return 1
    
    except Exception as e:
        raise ValueError(f"Failed to format component: {str(e)}")


if __name__ == "__main__":
    app()
