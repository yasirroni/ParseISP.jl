from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent
SCRIPTS = ROOT.parent

TARGETS = {
    "discovery": SCRIPTS / "eda_isp2026_inputs_workbook_tables.py",
    "parsing": SCRIPTS / "parse_isp2026_inputs_workbook_tables.py",
}


def assemble(kind: str) -> str:
    source_dir = ROOT / kind
    parts = [(source_dir / "_setup.py").read_text()]
    parts.extend(path.read_text() for path in sorted(source_dir.glob("[0-9][0-9][0-9]_*.py")))
    parts.append((source_dir / "_footer.py").read_text())
    return "".join(parts)


def write_notebooks() -> None:
    for kind, target in TARGETS.items():
        target.write_text(assemble(kind))


if __name__ == "__main__":
    write_notebooks()
