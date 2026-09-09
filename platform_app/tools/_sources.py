"""Shared helper for the two tools that read a body of C source.

Both of them can be handed either a pile of loose `.c` / `.h` files or a single
`.zip` of a source tree, and the engine underneath wants one directory to walk.
This flattens whichever it gets into a single working directory.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

# Suffixes the call-tree parser and the function parser understand.
C_SUFFIXES = (".c", ".h", ".cpp", ".hpp", ".cc", ".s", ".asm")


def work_dir(output_dir: Path, name: str = "sources") -> Path:
    """A scratch directory beside the run's output, swept with the run."""
    target = output_dir.parent / "work" / name
    target.mkdir(parents=True, exist_ok=True)
    return target


def gather(
    uploads: Sequence[Path],
    destination: Path,
    suffixes: Iterable[str] = C_SUFFIXES,
) -> Tuple[List[Path], int]:
    """Copy `uploads` into `destination`, expanding any zip archives.

    Returns the source files that landed there and how many archives were
    expanded. Files whose suffix is not in `suffixes` are skipped, so a zip of
    a whole checkout does not drag in its build output.
    """
    allowed = {s.lower() for s in suffixes}
    archives = 0

    for upload in uploads:
        if upload.suffix.lower() == ".zip":
            archives += 1
            _extract(upload, destination, allowed)
        else:
            target = _free_name(destination / upload.name)
            target.write_bytes(upload.read_bytes())

    found = sorted(
        p for p in destination.rglob("*")
        if p.is_file() and p.suffix.lower() in allowed
    )
    return found, archives


def _extract(archive: Path, destination: Path, allowed: set) -> None:
    """Unpack the wanted members of `archive`, refusing paths that escape."""
    with zipfile.ZipFile(archive) as zf:
        for member in zf.infolist():
            if member.is_dir():
                continue
            name = Path(member.filename)
            if name.suffix.lower() not in allowed:
                continue

            # A member may not climb out of the destination directory.
            target = (destination / name).resolve()
            if not str(target).startswith(str(destination.resolve())):
                continue

            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(member) as src, open(_free_name(target), "wb") as dst:
                dst.write(src.read())


def _free_name(target: Path) -> Path:
    """First unused name at `target`, so same-named files do not collide."""
    if not target.exists():
        return target
    stem, suffix = target.stem, target.suffix
    counter = 2
    while True:
        candidate = target.with_name(f"{stem}_{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1
