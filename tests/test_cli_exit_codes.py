"""Unsupported input files must exit nonzero, not just print an error."""

import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
MAIN = REPO / "xlsxgrep" / "xlsxgrep.py"


def run_xlsxgrep(*args):
    return subprocess.run(
        [sys.executable, str(MAIN), *args],
        capture_output=True,
        text=True,
        cwd=REPO,
    )


def test_unsupported_format_exits_nonzero(tmp_path):
    bad = tmp_path / "notes.txt"
    bad.write_text("hello world\n", encoding="utf-8")
    proc = run_xlsxgrep("hello", str(bad))
    assert "Unsupported file format" in proc.stderr
    assert proc.returncode == 1


def test_supported_csv_still_exits_zero(tmp_path):
    good = tmp_path / "data.csv"
    good.write_text("name,note\namy,hello there\n", encoding="utf-8")
    proc = run_xlsxgrep("hello", str(good))
    assert proc.returncode == 0
    assert "hello" in proc.stdout
