import typer
from typing import Optional, List, Tuple
from utils.workbook import get_active_workbook
from utils.file_system import FileSystem
from constants import ComponentType, EXPORT_EXTENSIONS, EXPORT_FOLDERS

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
        fs = FileSystem()

        if not fs.is_structure_valid():
            typer.echo("❌ No src folder structure found. Run 'vcm export' first.", err=True)
            raise typer.Exit(code=1)

        wb = get_active_workbook()
        vbproj = wb.VBProject

        imported_count = 0
        failed_count = 0
        deleted_count = 0

        target_component = (
            onefile.replace("\\", "/").strip()
            if onefile
            else None
        )

        components = (
            _get_single_component(fs, target_component)
            if target_component
            else _get_all_components(fs)
        )

        if not components:
            typer.echo("⚠️ No components found to import")
            return

        typer.echo(f"[vcm] Importing into: {wb.Name}")

        for folder_name, component_name, file_path in components:
            try:
                component_type = _get_component_type_from_folder(folder_name)

                if component_type in {
                    ComponentType.STANDARD_MODULE,
                    ComponentType.CLASS_MODULE,
                    ComponentType.USER_FORM,
                }:
                    existing = _get_component(vbproj, component_name)
                    if existing:
                        vbproj.VBComponents.Remove(existing)

                    vbproj.VBComponents.Import(str(file_path))
                    typer.echo(f"✅ Imported: {folder_name}/{component_name}")

                elif component_type == ComponentType.DOCUMENT_MODULE:
                    vb_comp = _get_component(vbproj, component_name)

                    if not vb_comp:
                        typer.echo(f"⚠️ Excel object not found: {component_name}", err=True)
                        failed_count += 1
                        continue

                    code = file_path.read_text(encoding="utf-8")

                    cm = vb_comp.CodeModule
                    cm.DeleteLines(1, cm.CountOfLines)
                    cm.AddFromString(code)

                    typer.echo(f"✅ Updated: {folder_name}/{component_name}")

                imported_count += 1

            except Exception as e:
                failed_count += 1
                typer.echo(f"❌ Failed {component_name}: {e}", err=True)

        if force:
            typer.echo("[vcm] Force mode: cleaning orphan components...")

            src_names = _collect_src_component_names(fs)
            snapshot = list(vbproj.VBComponents)

            for comp in snapshot:
                try:
                    if comp.Type != ComponentType.DOCUMENT_MODULE:
                        if comp.Name not in src_names:
                            name = comp.Name
                            vbproj.VBComponents.Remove(comp)
                            typer.echo(f"🗑️ Removed: {name}")
                            deleted_count += 1
                except Exception:
                    continue

        try:
            wb.Save()
            typer.echo("[vcm] Workbook saved.")
        except Exception as e:
            typer.echo(f"❌ Save failed: {e}", err=True)

        typer.echo(
            f"\n📦 Import complete | Imported: {imported_count} | Deleted: {deleted_count} | Failed: {failed_count}\n"
        )

        if failed_count > 0:
            raise typer.Exit(code=1)

    except Exception as e:
        typer.echo(f"❌ Import failed: {e}", err=True)
        raise typer.Exit(code=1)

def _get_single_component(fs: FileSystem, target: str) -> List[Tuple[str, str, object]]:
    parts = target.split("/")
    if len(parts) != 2:
        typer.echo("❌ Invalid format. Use 'cls/MyClass'", err=True)
        raise typer.Exit(code=1)

    folder_name, component_name = parts

    if folder_name not in EXPORT_FOLDERS.values():
        typer.echo(f"❌ Invalid folder: {folder_name}", err=True)
        raise typer.Exit(code=1)

    extension = None
    for ctype, fname in EXPORT_FOLDERS.items():
        if fname == folder_name:
            extension = EXPORT_EXTENSIONS[ctype]
            break

    file_path = fs.src_dir / folder_name / f"{component_name}{extension}"

    if not file_path.exists():
        typer.echo(f"❌ File not found: {file_path}", err=True)
        raise typer.Exit(code=1)

    return [(folder_name, component_name, file_path)]


def _get_all_components(fs: FileSystem) -> List[Tuple[str, str, object]]:
    components = []

    for folder_name in EXPORT_FOLDERS.values():
        folder_path = fs.src_dir / folder_name

        if not folder_path.exists():
            continue

        for file_path in folder_path.iterdir():
            if not file_path.is_file():
                continue

            if file_path.suffix == ".frx":
                continue

            components.append((folder_name, file_path.stem, file_path))

    return components

def _get_component(vbproj, name: str):
    try:
        return vbproj.VBComponents(name)
    except Exception:
        return None

def _get_component_type_from_folder(folder_name: str) -> ComponentType:
    for ctype, fname in EXPORT_FOLDERS.items():
        if fname == folder_name:
            return ctype
    raise ValueError(f"Invalid folder: {folder_name}")

def _collect_src_component_names(fs: FileSystem) -> set:
    names = set()

    for folder_name in EXPORT_FOLDERS.values():
        folder_path = fs.src_dir / folder_name

        if not folder_path.exists():
            continue

        for file_path in folder_path.iterdir():
            if file_path.is_file() and file_path.suffix != ".frx":
                names.add(file_path.stem)

    return names

if __name__ == "__main__":
    app()