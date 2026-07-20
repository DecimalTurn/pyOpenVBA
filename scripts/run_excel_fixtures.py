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
from pathlib import Path

BUILD_DIR = Path(__file__).resolve().parents[1] / "build" / "fixtures" / "excel"


def _run_fixture(xlsm: Path) -> bool:
    """Open *xlsm*, run Module1.RunFixture, verify the sentinel file. Return True on pass."""
    import win32com.client  # type: ignore[import]
    import time

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
        xl.Application.Run("Module1.RunFixture")
        time.sleep(0.5)  # Give macro time to complete
        wb.Close(SaveChanges=False)
        wb = None
        print("closed.", end=" ", flush=True)

        expected = xlsm.with_suffix(".txt")
        if not expected.exists():
            print(f"FAIL  output file not created: {expected.name}")
            return False

        content = expected.read_text(encoding="utf-8").strip()
        print(f"PASS  → {expected.name}: {content!r}")
        return True

    except Exception as exc:
        print(f"FAIL  {exc}")
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
    xlsm_files = sorted(BUILD_DIR.glob("*.xlsm"))
    if not xlsm_files:
        print(f"No .xlsm files found in {BUILD_DIR}")
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
