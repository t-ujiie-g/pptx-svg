#!/usr/bin/env python3
"""
gen_string_constants.py

Reads the compiled Wasm binary and extracts all string-constant field names
from the '_' import module (importedStringConstants convention used by
MoonBit's wasm-gc + use-js-builtin-string: true).

Prints a JavaScript array literal suitable for pasting into web/host.js as
the STRING_CONSTANTS array.

Usage:
    python3 scripts/gen_string_constants.py
    python3 scripts/gen_string_constants.py --update   # patch host.js in-place
"""

import sys
import re
from pathlib import Path

WASM_PATH = Path(__file__).parent.parent / "_build/wasm-gc/release/build/main/main.wasm"
HOST_JS   = Path(__file__).parent.parent / "web/host.js"


def read_leb128(data: bytes, pos: int) -> tuple[int, int]:
    val, shift = 0, 0
    while True:
        b = data[pos]; pos += 1
        val |= (b & 0x7F) << shift
        if not (b & 0x80):
            break
        shift += 7
    return val, pos


def read_bytes(data: bytes, pos: int) -> tuple[bytes, int]:
    length, pos = read_leb128(data, pos)
    return data[pos:pos + length], pos + length


def extract_underscore_strings(wasm_path: Path) -> list[str]:
    data = wasm_path.read_bytes()
    offset = 8  # skip magic + version

    while offset < len(data):
        section_id = data[offset]; offset += 1
        size, offset = read_leb128(data, offset)

        if section_id != 2:   # not Import section
            offset += size
            continue

        # Import section found
        count, pos = read_leb128(data, offset)
        strings: list[str] = []

        for _ in range(count):
            mod_bytes, pos  = read_bytes(data, pos)
            field_bytes, pos = read_bytes(data, pos)
            kind = data[pos]; pos += 1

            mod_name   = mod_bytes.decode('utf-8', errors='replace')
            field_name = field_bytes.decode('utf-8', errors='replace')

            if kind == 0:  # func — skip type index
                _, pos = read_leb128(data, pos)
            elif kind == 3:  # global — skip val type + mutability
                vt = data[pos]; pos += 1
                if vt in (0x64, 0x63):   # ref (non-null / null) — heap type follows
                    pos += 1
                pos += 1  # mutability byte

                if mod_name == '_':
                    strings.append(field_name)
            else:
                break  # unexpected kind — stop

        return strings

    raise RuntimeError("Import section not found in Wasm binary")


def js_string_literal(s: str) -> str:
    """Produce a JavaScript string literal for s, using single quotes."""
    escaped = s.replace('\\', '\\\\').replace("'", "\\'")
    # Make control chars visible
    escaped = escaped.replace('\n', '\\n').replace('\t', '\\t').replace('\r', '\\r')
    return f"'{escaped}'"


def format_js_array(strings: list[str], indent: int = 2) -> str:
    pad = ' ' * indent
    lines = []
    row: list[str] = []
    row_len = 0

    for s in strings:
        lit = js_string_literal(s)
        if row and row_len + len(lit) + 2 > 72:
            lines.append(pad + ', '.join(row) + ',')
            row = []
            row_len = 0
        row.append(lit)
        row_len += len(lit) + 2

    if row:
        lines.append(pad + ', '.join(row) + ',')

    return '[\n' + '\n'.join(lines) + '\n]'


def update_host_js(strings: list[str]) -> None:
    source = HOST_JS.read_text(encoding='utf-8')
    array_js = format_js_array(strings, indent=2)

    pattern = re.compile(
        r'(const STRING_CONSTANTS\s*=\s*)\[.*?\];',
        re.DOTALL,
    )
    # re.subn treats backslashes in the replacement string specially,
    # so escape them to avoid \n being turned into a literal newline, etc.
    safe_array = array_js.replace('\\', '\\\\')
    replacement = r'\g<1>' + safe_array + ';'
    new_source, count = pattern.subn(replacement, source)

    if count == 0:
        print("ERROR: Could not find 'const STRING_CONSTANTS = [...]' in host.js", file=sys.stderr)
        sys.exit(1)

    HOST_JS.write_text(new_source, encoding='utf-8')
    print(f"✓ Updated STRING_CONSTANTS in {HOST_JS} ({len(strings)} entries)")


def main() -> None:
    if not WASM_PATH.exists():
        print(f"ERROR: Wasm binary not found: {WASM_PATH}", file=sys.stderr)
        print("Run:  moon build --target wasm-gc --release", file=sys.stderr)
        sys.exit(1)

    strings = extract_underscore_strings(WASM_PATH)
    print(f"Found {len(strings)} string constants in '_' module:\n")

    array_js = format_js_array(strings, indent=2)
    print("const STRING_CONSTANTS =", array_js + ";")

    if '--update' in sys.argv:
        print()
        update_host_js(strings)


if __name__ == '__main__':
    main()
