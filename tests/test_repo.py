from __future__ import annotations

import importlib.util
import json
import re
import sys

from conftest import ROOT

EINZEILER = (
    "Claude Skills für Excel: sechs Agent Skills für Claude Code und Codex, die deutsche Excel-Dateien prüfen und bereinigen, "
    "zwei Excel-Tabellen vergleichen, mehrere Excel-Dateien zusammenführen und in CSV für DATEV umwandeln, "
    "lokal und ohne das Original zu ändern."
)
SKILLS = ("excel-bereinigen", "excel-csv", "excel-formeln", "excel-profil", "excel-stapeln", "excel-vergleich")


def _skript(name: str):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / f"{name}.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def _frontmatter(text: str) -> dict[str, str]:
    kopf = text.split("---", 2)[1]
    out = {}
    for zeile in kopf.strip().splitlines():
        k, v = zeile.split(":", 1)
        out[k.strip()] = v.strip().strip('"')
    return out


def test_einzeiler_ueberall_gleich():
    assert EINZEILER in (ROOT / "README.md").read_text().splitlines()[2]
    assert EINZEILER in (ROOT / "AGENTS.md").read_text()
    for p in (".claude-plugin/plugin.json", ".codex-plugin/plugin.json"):
        assert json.loads((ROOT / p).read_text())["description"] == EINZEILER
    assert json.loads((ROOT / ".claude-plugin/marketplace.json").read_text())["metadata"]["description"] == EINZEILER
    assert json.loads((ROOT / ".codex-plugin/plugin.json").read_text())["interface"]["longDescription"].startswith(EINZEILER)


def test_skills():
    gefunden = tuple(sorted(p.parent.name for p in (ROOT / "skills").glob("*/SKILL.md")))
    assert gefunden == SKILLS
    for name in SKILLS:
        text = (ROOT / "skills" / name / "SKILL.md").read_text()
        fm = _frontmatter(text)
        assert fm["name"] == name
        assert len(fm["description"]) <= 1024
        assert "<" not in fm["description"] and ">" not in fm["description"]
        assert "Claude Skills für Excel" in fm["description"]
        assert EINZEILER in text
        assert "excel-skills@v0.1.1" in text


def test_versionen_gleich():
    from excel_skills import __version__

    assert __version__ == "0.1.1"
    assert 'version = "0.1.1"' in (ROOT / "pyproject.toml").read_text()
    for p in (".claude-plugin/plugin.json", ".codex-plugin/plugin.json"):
        assert json.loads((ROOT / p).read_text())["version"] == __version__


def test_spiegel_und_beispiele_aktuell():
    assert _skript("sync_skills").main(["--check"]) == 0
    assert _skript("beispiele_bauen").main(["--check"]) == 0


def test_readme_hat_20_beispiele_und_anker():
    readme = (ROOT / "README.md").read_text()
    assert len(re.findall(r"^\| \[\d+\]\(docs/beispiele\.md#beispiel-\d+\)", readme, re.M)) == 20
    regeln = (ROOT / "docs" / "regeln.md").read_text()
    for anker in re.findall(r"docs/regeln\.md#([a-z-]+)", readme + "".join(
            (ROOT / "skills" / n / "SKILL.md").read_text() for n in SKILLS)):
        assert f'<a id="{anker}"></a>' in regeln, anker


def test_kein_gedankenstrich():
    for p in ROOT.rglob("*"):
        if p.suffix not in (".md", ".py", ".json", ".toml", ".yml") or ".venv" in p.parts:
            continue
        text = p.read_text(encoding="utf-8")
        assert "\u2014" not in text and "\u2013" not in text, p
