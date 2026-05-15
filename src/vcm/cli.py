"""
VCM CLI - Main entry point
Registers all commands: export, import, format
"""

import typer
from typing import Optional
from . import version
from .commands import export_refactored, import_refactored, format_refactored

# Create main Typer app
app = typer.Typer(
    name="vcm",
    help="VBA Component Manager - Export/Import VBA components from Excel workbooks",
    no_args_is_help=True,
    rich_markup_mode="rich"
)

# Add commands
app.command(name="export")(export_refactored.export_command)
app.command(name="import")(import_refactored.import_command)
app.command(name="format")(format_refactored.format_command)


@app.callback(invoke_without_command=True)
def main(
    version_flag: Optional[bool] = typer.Option(
        None,
        "--version",
        "-v",
        help="Show version and exit",
        is_flag=True,
        callback=None,
    ),
):
    """
    VBA Component Manager
    
    Manage VBA components in Excel workbooks with ease.
    
    [bold cyan]Commands:[/bold cyan]
    
    [yellow]vcm export[/yellow]               Export VBA components from workbook
    [yellow]vcm import[/yellow]               Import VBA components to workbook
    [yellow]vcm format[/yellow]               Format VBA code
    
    [bold cyan]Examples:[/bold cyan]
    
    [green]# Export all components[/green]
    vcm export
    
    [green]# Export single component[/green]
    vcm export --onefile cls/MyClass
    
    [green]# Import all components[/green]
    vcm import
    
    [green]# Import and remove missing[/green]
    vcm import --force
    
    [green]# Format all code[/green]
    vcm format --all
    
    For more help on a command:
    vcm export --help
    vcm import --help
    vcm format --help
    """
    if version_flag:
        version_str = version.get_version()
        typer.echo(f"vcm version {version_str}")
        raise typer.Exit()


if __name__ == "__main__":
    app()
