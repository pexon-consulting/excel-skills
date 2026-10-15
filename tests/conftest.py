from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from excel_skills import cli

ROOT = Path(__file__).resolve().parents[1]
FIX = ROOT / "fixtures"


def sha(pfad: Path) -> str:
    return hashlib.sha256(pfad.read_bytes()).hexdigest()


@pytest.fixture
def lauf(capsys):
    """Ruft excel-skills auf. Rückgabe: Exit-Code und JSON von stdout oder stderr."""

    def _lauf(*args: str) -> tuple[int, dict]:
        code = cli.main([str(a) for a in args])
        aus = capsys.readouterr()
        text = aus.out if code == 0 else aus.err
        return code, json.loads(text) if text.strip() else {}

    return _lauf
