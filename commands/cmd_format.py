import typer
from pathlib import Path
from utils.formatter import format_vba_code
from utils.file_system import FileSystem  # your existing one

app = typer.Typer(help="Format VBA code in src folder")

@app.command()
def format_command(
    all: bool = typer.Option(False, "--all", "-a"),
    onefile: str | None = typer.Option(None, "--onefile", "-o"),
    indent: int = typer.Option(2, "--indent", "-i", help="Indent size"),
):
    """
    Format VBA code.

    Examples:
        vcm format -a
        vcm format -o cls/MyClass
        vcm format -a -i 4
    """
    fs = FileSystem(Path.cwd())
    src_path = fs.src_dir

    if not src_path.exists():
        raise typer.Exit("❌ src folder not found. Run export first.")

    if all and onefile:
        raise typer.Exit("❌ Use either --all or --onefile, not both.")

    if not all and not onefile:
        raise typer.Exit("❌ Specify --all or --onefile.")

    if all:
        count = _format_all(src_path, indent)
    else:
        count = _format_one(src_path, onefile, indent)

    typer.echo("")
    typer.echo("=" * 50)
    typer.echo("Formatting complete")
    typer.echo(f"   Files processed: {count}")
    typer.echo("=" * 50)


def _format_all(src_path: Path, indent: int) -> int:
    files = _get_vba_files(src_path)

    if not files:
        typer.echo("⚠️ No VBA files found")
        return 0

    processed = 0

    for file_path in files:
        try:
            code = file_path.read_text(encoding="utf-8")
            formatted = format_vba_code(code, indent)

            if formatted != code:
                file_path.write_text(formatted, encoding="utf-8")
                typer.echo(f"✓ Formatted: {file_path.name}")
            else:
                typer.echo(f"⊘ Already formatted: {file_path.name}")

            processed += 1

        except Exception as e:
            typer.echo(f"✗ Failed: {file_path.name} | {e}", err=True)

    return processed


def _format_one(src_path: Path, onefile: str, indent: int) -> int:
    try:
        component_type, name = onefile.split("/", 1)
    except ValueError:
        raise typer.Exit("Format must be: cls/MyClass")

    folder = src_path / component_type

    matches = list(folder.glob(f"{name}.*"))

    if not matches:
        raise typer.Exit(f"File not found: {onefile}")

    file_path = matches[0]

    code = file_path.read_text(encoding="utf-8")
    formatted = format_vba_code(code, indent)

    if formatted != code:
        file_path.write_text(formatted, encoding="utf-8")
        typer.echo(f"Formatted: {file_path.name}")
    else:
        typer.echo(f"Already formatted: {file_path.name}")

    return 1


def _get_vba_files(src_path: Path) -> list[Path]:
    exts = {".bas", ".cls", ".frm", ".vbs"}
    return [
        p for p in src_path.rglob("*")
        if p.is_file() and p.suffix.lower() in exts
    ]


if __name__ == "__main__":
    app()