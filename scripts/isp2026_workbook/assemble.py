from pathlib import Path

ROOT = Path(__file__).resolve().parent
STREAMS = {
    'discovery': ROOT.parent / 'eda_isp2026_inputs_workbook_tables.py',
    'parsing': ROOT.parent / 'parse_isp2026_inputs_workbook_tables.py',
}
EXPECTED_SHEETS = 84


def sheet_files(directory: Path) -> list[Path]:
    files = sorted(directory.glob('[0-9][0-9][0-9]_*.py'))
    expected = list(range(1, EXPECTED_SHEETS + 1))
    found = [int(path.name[:3]) for path in files]
    if found != expected:
        raise RuntimeError(
            f'{directory}: expected worksheet prefixes 001..{EXPECTED_SHEETS:03d}; found {found}'
        )
    return files


def assemble_stream(name: str, output: Path) -> None:
    directory = ROOT / name
    pieces = [directory / '_setup.py', *sheet_files(directory), directory / '_footer.py']
    missing = [str(path) for path in pieces if not path.is_file()]
    if missing:
        raise FileNotFoundError(f'Missing assembly sources: {missing}')
    output.write_text(''.join(path.read_text() for path in pieces))


if __name__ == '__main__':
    for stream, output in STREAMS.items():
        assemble_stream(stream, output)
