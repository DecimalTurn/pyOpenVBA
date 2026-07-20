"""Run RunFixture macro in each built Excel workbook and verify output files.

Requires pywin32 (`pip install pywin32`). Must be run on Windows after
build_excel_fixtures.py has produced the .xlsm files.

For each .xlsm in build/fixtures/excel/:
  - Opens the workbook via COM automation
  - Calls Module1.RunFixture
  - Verifies <fixture_name>.txt was written alongside the workbook
  - Reports pass / fail per fixture

Run from the repo root:

    python scripts/run_excel_fixtures.py

Exit code: 0 = all passed, 1 = one or more failed.
"""

from __future__ import annotations

import sys
import os

# Force UTF-8 output on Windows
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

from pathlib import Path

BUILD_DIR = Path(__file__).resolve().parents[1] / "build" / "fixtures" / "excel"


def _run_fixture(xlsm: Path) -> bool:
    """Open *xlsm*, run Module1.RunFixture, verify the sentinel file. Return True on pass."""
    import win32com.client  # type: ignore[import]
    import time
    import os
    import traceback

    xl = None
    wb = None
    try:
        print(f"  Opening {xlsm.name}...", end=" ", flush=True)
        xl = win32com.client.Dispatch("Excel.Application")
        xl.DisplayAlerts = False
        xl.Visible = False

        wb = xl.Workbooks.Open(str(xlsm.absolute()))
        time.sleep(0.5)  # Give Excel time to load
        print("running macro...", end=" ", flush=True)
        try:
            xl.Application.Run("Module1.RunFixture")
        except Exception as macro_err:
            print(f"FAIL (macro error: {macro_err})")
            return False
        time.sleep(0.5)  # Give macro time to complete
        wb.Close(SaveChanges=False)
        wb = None
        print("closed.", end=" ", flush=True)

        expected = xlsm.with_suffix(".txt")
        time.sleep(0.5)  # Extra wait for file I/O
        
        # Debug: list files in directory
        dir_contents = list(xlsm.parent.glob("*"))
        if not expected.exists():
            print(f"FAIL (output file not found)")
            print(f"      Expected: {expected}")
            print(f"      Dir contents: {[f.name for f in dir_contents]}")
            return False

        content = expected.read_text(encoding="utf-8").strip()
        print(f"PASS  -> {expected.name}: {content!r}")
        return True

    except Exception as exc:
        print(f"FAIL  {type(exc).__name__}: {exc}")
        if exc:
            traceback.print_exc()
        return False

    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass


def main() -> None:
    # Accept optional path argument for testing a single file
    if len(sys.argv) > 1:
        xlsm_files = [Path(sys.argv[1])]
    else:
        xlsm_files = sorted(BUILD_DIR.glob("*.xlsm"))
    
    if not xlsm_files:
        print(f"No .xlsm files found")
        sys.exit(1)

    print(f"Running {len(xlsm_files)} fixture(s) from {BUILD_DIR}\n")

    results = [_run_fixture(xlsm) for xlsm in xlsm_files]

    passed = sum(results)
    failed = len(results) - passed
    print(f"\n{'-' * 40}")
    print(f"{passed}/{len(results)} passed" + (f"  ({failed} failed)" if failed else ""))

    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
