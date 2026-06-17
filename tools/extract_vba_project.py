#!/usr/bin/env python3
"""Extract the compiled VBA project from an Excel .xlsm template."""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path
from zipfile import BadZipFile, ZipFile


def next_backup_path(path: Path) -> Path:
    candidate = path.with_suffix(path.suffix + ".bak")
    counter = 1
    while candidate.exists():
        candidate = path.with_suffix(path.suffix + f".bak{counter}")
        counter += 1
    return candidate


def extract_vba_project(template_path: Path, output_path: Path, make_backup: bool) -> None:
    if not template_path.exists():
        raise FileNotFoundError(f"Template workbook not found: {template_path}")

    try:
        with ZipFile(template_path, "r") as workbook:
            member_lookup = {name.lower(): name for name in workbook.namelist()}
            member_name = member_lookup.get("xl/vbaproject.bin")
            if member_name is None:
                raise ValueError(f"{template_path} does not contain xl/vbaProject.bin. Save the template as .xlsm with macros enabled.")
            data = workbook.read(member_name)
    except BadZipFile as exc:
        raise ValueError(f"{template_path} is not a readable .xlsm/.xlsx zip package.") from exc

    if len(data) < 1024:
        raise ValueError("Extracted vbaProject.bin is unexpectedly small; the template may not contain a compiled VBA project.")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists() and make_backup:
        backup_path = next_backup_path(output_path)
        shutil.copy2(output_path, backup_path)
        print(f"Backed up existing VBA project to {backup_path}")

    output_path.write_bytes(data)
    print(f"Extracted {len(data):,} bytes to {output_path}")


def main(argv: list[str]) -> int:
    repo_root = Path(__file__).resolve().parents[1]
    default_output = repo_root / "vendor" / "excel" / "vbaProject.bin"

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("template", help="Path to the macro-enabled .xlsm template created in Excel")
    parser.add_argument("--output", default=str(default_output), help="Destination vbaProject.bin path")
    parser.add_argument("--no-backup", action="store_true", help="Do not keep a backup of the existing vbaProject.bin")
    args = parser.parse_args(argv)

    try:
        extract_vba_project(
            Path(args.template).expanduser().resolve(),
            Path(args.output).expanduser().resolve(),
            not args.no_backup,
        )
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
