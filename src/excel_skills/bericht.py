"""`report.md`, `report.json` und `aenderungen.json`. Zellinhalte sind Daten und werden für Markdown entschärft."""

from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path
from typing import Any

_STEUER = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def jsonwert(wert: Any) -> Any:
    if isinstance(wert, dt.datetime):
        return wert.isoformat(timespec="seconds") if wert.time() != dt.time() else wert.date().isoformat()
    if isinstance(wert, (dt.date, dt.time)):
        return wert.isoformat()
    if isinstance(wert, float) and wert.is_integer():
        return int(wert)
    if isinstance(wert, (str, int, float, bool)) or wert is None:
        return wert
    return str(wert)


def zelle_md(wert: Any, laenge: int = 80) -> str:
    """Eine Markdown-Tabellenzelle, die keine Zeile bricht und keine Spalte öffnet."""
    if wert is None:
        return ""
    s = str(jsonwert(wert))
    s = s.replace("\r", " ").replace("\n", " ").replace("\u2028", " ").replace("\u2029", " ")
    s = _STEUER.sub(" ", s)
    s = s.replace("\\", "\\\\").replace("|", "\\|")
    if len(s) > laenge:
        s = s[: laenge - 3] + "..."
    return s


def tabelle_md(kopf: list[str], zeilen: list[list[Any]]) -> list[str]:
    if not zeilen:
        return ["(keine)", ""]
    out = ["| " + " | ".join(kopf) + " |", "|" + "|".join("---" for _ in kopf) + "|"]
    out += ["| " + " | ".join(zelle_md(w) for w in z) + " |" for z in zeilen]
    out.append("")
    return out


def schreibe_json(pfad: Path, daten: dict[str, Any]) -> None:
    pfad.write_text(json.dumps(daten, ensure_ascii=False, indent=2, default=jsonwert) + "\n", encoding="utf-8")


def schreibe_bericht(md_pfad: Path, titel: str, zeilen: list[str], daten: dict[str, Any]) -> Path:
    """Schreibt report.md und daneben report.json mit gleichem Namensstamm."""
    json_pfad = md_pfad.with_suffix(".json")
    md_pfad.parent.mkdir(parents=True, exist_ok=True)
    md_pfad.write_text("\n".join([f"# {titel}", "", *zeilen]).rstrip() + "\n", encoding="utf-8")
    schreibe_json(json_pfad, daten)
    return json_pfad
