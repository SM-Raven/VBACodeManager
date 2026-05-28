"""
    Main Application refer to Typer App
"""

import typer
import click
from typing import Optional
from commands.cmd_export import export_command
from commands.cmd_import import import_command
from commands.cmd_format import format_command

VCM_VERSION: str = "0.3.0"

app = typer.Typer(
    name="vcm",
    help="VBA Code Manager - Export/Import/Format VBA Code.",
    no_args_is_help=False,
    rich_markup_mode="rich"
)

app.command(name="export")(export_command)
app.command(name="import")(import_command)
app.command(name="format")(format_command)

@app.callback(invoke_without_command=True)
def main(
    ctx: typer.Context,
    version_flag: Optional[bool] = typer.Option(
        None,
        "--version",
        "-v",
        help="Show version and exit",
        is_flag=True,
    ),
    help_flag: Optional[bool] = typer.Option(
        None,
        "--help",
        "-h",
        help="Show help message and exit",
        is_flag=True,
    ),
):
    # Show help when --help/-h flag used
    if help_flag:
        click.echo(ctx.get_help())
        raise typer.Exit()

    # Show version when --version/-v flag used
    if version_flag:
        typer.echo(f"vcm version {VCM_VERSION}")
        raise typer.Exit()

    # Show help when no command provided (no args at all)
    if ctx.invoked_subcommand is None:
        click.echo(ctx.get_help())
        raise typer.Exit()