"""Erzeugt docs/beispiele.md und die Beispieltabelle im README aus echten Läufen gegen fixtures/.

--check vergleicht nur und endet mit 1, wenn eine Datei veraltet ist.
"""

from __future__ import annotations

import contextlib
import io
import json
import os
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from excel_skills import cli  # noqa: E402

FIX = ROOT / "fixtures"
DOC = ROOT / "docs" / "beispiele.md"
README = ROOT / "README.md"
START = "<!-- beispiele:start -->"
ENDE = "<!-- beispiele:ende -->"
ALLE = ("kopfzeile", "verbunden", "leerzeilen", "summen", "zahlen", "datum", "trimmen")


def nur(*regeln: str) -> str:
    return ",".join(r for r in ALLE if r not in regeln)


@dataclass
class Lauf:
    code: int
    erg: dict[str, Any]
    ordner: Path

    def json(self, name: str) -> dict[str, Any]:
        return json.loads((self.ordner / name).read_text(encoding="utf-8"))

    def text(self, name: str, kodierung: str = "utf-8") -> str:
        return (self.ordner / name).read_bytes().decode(kodierung)


@dataclass
class Beispiel:
    satz: str
    args: list[str]
    ergebnis: Callable[[Lauf], str]
    detail: Callable[[Lauf], list[str]]


def code_json(obj: Any) -> list[str]:
    return ["```json", json.dumps(obj, ensure_ascii=False, indent=2), "```"]


def code_text(text: str, sprache: str = "text") -> list[str]:
    return [f"```{sprache}", text.replace("\r\n", "\n").rstrip("\n"), "```"]


def bericht_abschnitt(lauf: Lauf, blatt: str | None = None) -> list[str]:
    zeilen = lauf.text("report.md").splitlines()[1:]
    if blatt:
        start = zeilen.index(f"## Blatt {blatt}")
        ende = next((i for i in range(start + 1, len(zeilen)) if zeilen[i].startswith("## Blatt ")), len(zeilen))
        zeilen = zeilen[start:ende]
    return [("#### " + z[3:]) if z.startswith("## ") else z for z in zeilen]


def aenderungen(lauf: Lauf, regel: str | None = None, felder: tuple[str, ...] = ("aenderungen",)) -> dict[str, Any]:
    d = lauf.json("kopie.aenderungen.json")
    out: dict[str, Any] = {}
    for f in felder:
        v = d[f]
        if regel and isinstance(v, list):
            v = [x for x in v if x.get("regel") == regel]
        out[f] = v
    return out


def kalk(lauf: Lauf) -> dict[str, Any]:
    return lauf.json("report.json")["blaetter"][0]


def fehler_text(lauf: Lauf, fehler: str) -> str:
    z = [f["zelle"] for f in kalk(lauf)["fehler"] if f["fehler"] == fehler]
    return ", ".join(z)


BEISPIELE = [
    Beispiel(
        "Finde die Kopfzeile in dieser Excel, die erst in Zeile 4 anfängt.",
        ["profil", "fixtures/umsatz_bericht.xlsx", "--out", "report.md"],
        lambda l: f"Kopfzeile Zeile {l.json('report.json')['blaetter'][0]['kopfzeile']['zeile']}, Datenbereich {l.json('report.json')['blaetter'][0]['datenbereich']}",
        lambda l: bericht_abschnitt(l, "Umsatz"),
    ),
    Beispiel(
        "Ist dieses Excel-Blatt eine Tabelle oder ein Bericht?",
        ["profil", "fixtures/kunden_liste.xlsx", "--out", "report.md"],
        lambda l: f"Urteil `{l.erg['blaetter']['Kunden']}`, PLZ bleibt Text",
        lambda l: bericht_abschnitt(l, "Kunden"),
    ),
    Beispiel(
        "Verbundene Zellen auflösen und den Wert nach unten füllen.",
        ["bereinigen", "fixtures/umsatz_bericht.xlsx", "--ohne", nur("verbunden"), "--out", "kopie.xlsx"],
        lambda l: ", ".join(f"{a['zelle']} = {a['neu']}" for a in aenderungen(l, "verbunden")["aenderungen"] if a["neu"] != "Umsatzbericht 2026"),
        lambda l: code_json(aenderungen(l, "verbunden")),
    ),
    Beispiel(
        "Leere Zeilen und Zwischensummen aus der Liste löschen.",
        ["bereinigen", "fixtures/umsatz_bericht.xlsx", "--ohne", nur("leerzeilen", "summen"), "--out", "kopie.xlsx"],
        lambda l: "Zeilen " + ", ".join(str(z["zeile"]) for z in aenderungen(l, felder=("geloeschte_zeilen",))["geloeschte_zeilen"]) + " entfernt, Formeln angepasst",
        lambda l: code_json(aenderungen(l, felder=("geloeschte_zeilen", "formeln_angepasst"))),
    ),
    Beispiel(
        "Deutsche Zahlen 1.234,56 in echte Zahlen umwandeln.",
        ["bereinigen", "fixtures/umsatz_bericht.xlsx", "--ohne", nur("zahlen"), "--out", "kopie.xlsx"],
        lambda l: f"{len(aenderungen(l, 'deutsche-zahl')['aenderungen'])} Zellen zu Zahl, `1.234` in E7 bleibt Text",
        lambda l: code_json(aenderungen(l, "deutsche-zahl", ("aenderungen", "hinweise"))),
    ),
    Beispiel(
        "Datum 31.12.2026 als Text in ein Excel-Datum umwandeln.",
        ["bereinigen", "fixtures/umsatz_bericht.xlsx", "--ohne", nur("datum"), "--out", "kopie.xlsx"],
        lambda l: f"{len(aenderungen(l, 'datum')['aenderungen'])} Zellen zu Datum, `01/31/2026` bleibt Text",
        lambda l: code_json(aenderungen(l, "datum", ("aenderungen", "hinweise"))),
    ),
    Beispiel(
        "Alle #BEZUG!-Fehler in der Arbeitsmappe finden.",
        ["formeln", "fixtures/formeln_kaputt.xlsx", "--out", "report.md"],
        lambda l: f"`#BEZUG!` in {fehler_text(l, '#REF!')}",
        lambda l: code_json([f for f in kalk(l)["fehler"] if f["fehler"] == "#REF!"]),
    ),
    Beispiel(
        "Warum zeigt mein SVERWEIS #NV, und in welchen Zellen?",
        ["formeln", "fixtures/formeln_kaputt.xlsx", "--out", "report.md"],
        lambda l: f"{fehler_text(l, '#N/A')}: SVERWEIS findet den Suchwert nicht",
        lambda l: code_json([f for f in kalk(l)["fehler"] if f["fehler"] == "#N/A"]),
    ),
    Beispiel(
        "Harte Zahlen in einer Formelspalte finden.",
        ["formeln", "fixtures/formeln_kaputt.xlsx", "--out", "report.md"],
        lambda l: ", ".join(f"{h['zelle']} = {h['wert']} in Spalte {h['spalte']}" for h in kalk(l)["harte_werte"]),
        lambda l: code_json({"formelspalten": kalk(l)["formelspalten"], "harte_werte": kalk(l)["harte_werte"]}),
    ),
    Beispiel(
        "Externe Verknüpfungen zu anderen Excel-Dateien auflisten.",
        ["formeln", "fixtures/formeln_kaputt.xlsx", "--out", "report.md"],
        lambda l: ", ".join(f"{b['blatt']}!{e['zelle']}" for b in l.json("report.json")["blaetter"] for e in b["externe_verknuepfungen"]),
        lambda l: code_json([{"blatt": b["blatt"], **e} for b in l.json("report.json")["blaetter"] for e in b["externe_verknuepfungen"]]),
    ),
    Beispiel(
        "Formeln finden, die nicht bis zur letzten Zeile mitlaufen.",
        ["formeln", "fixtures/formeln_kaputt.xlsx", "--out", "report.md"],
        lambda l: ", ".join(f"{h['zelle']} in Spalte {h['spalte']}" for h in kalk(l)["formel_fehlt"]),
        lambda l: code_json(kalk(l)["formel_fehlt"]),
    ),
    Beispiel(
        "Zwei Excel-Tabellen vergleichen: alte und neue Preisliste, Schlüssel Artikelnummer.",
        ["vergleich", "fixtures/preise_alt.xlsx", "fixtures/preise_neu.xlsx", "--schluessel", "Artikelnummer", "--out", "report.md"],
        lambda l: f"{l.erg['neu']} neu, {l.erg['geloescht']} gelöscht, {l.erg['geaendert']} geändert",
        lambda l: [z for z in l.text("report.md").splitlines()[1:]],
    ),
    Beispiel(
        "Welche Kunden sind in der neuen Liste dazugekommen oder weggefallen?",
        ["vergleich", "fixtures/kunden_liste.xlsx", "fixtures/kunden_liste_neu.xlsx", "--schluessel", "Kundennummer", "--out", "report.md"],
        lambda l: "neu " + ", ".join(z["schluessel"] for z in l.json("report.json")["neu"]) + ", weg " + ", ".join(z["schluessel"] for z in l.json("report.json")["geloescht"]),
        lambda l: code_json({k: l.json("report.json")[k] for k in ("neu", "geloescht", "geaendert")}),
    ),
    Beispiel(
        "Zwölf Excel-Dateien zusammenführen, eine je Monat.",
        ["stapeln", "fixtures/monate", "--out", "kopie.xlsx"],
        lambda l: f"{l.erg['dateien']} Dateien, {l.erg['zeilen_gesamt']} Zeilen, {l.erg['uebersprungen']} Summenzeile übersprungen",
        lambda l: code_json({k: l.json("kopie.aenderungen.json")[k] for k in ("spalten", "quellen", "uebersprungen")}),
    ),
    Beispiel(
        "Prüfen, ob alle Dateien im Ordner dieselben Spalten haben.",
        ["stapeln", "fixtures/monate_abweichend", "--out", "kopie.xlsx"],
        lambda l: "Abbruch mit Exit-Code 3: " + "; ".join(f"{d['datei']} fehlt {', '.join(d['fehlt'])}, hat {', '.join(d['zusaetzlich'])}" for d in l.erg["details"]),
        lambda l: code_json(l.erg),
    ),
    Beispiel(
        "Excel als CSV für den DATEV-Import exportieren.",
        ["csv", "fixtures/buchungen.xlsx", "--format", "datev", "--out", "kopie.csv"],
        lambda l: "Semikolon, Komma, `TT.MM.JJJJ`, Windows-1252, Text in Anführungszeichen",
        lambda l: code_text(l.text("kopie.csv", "cp1252"), "csv"),
    ),
    Beispiel(
        "CSV mit Semikolon und Komma für deutsches Excel speichern.",
        ["csv", "fixtures/buchungen.xlsx", "--format", "excel-de", "--out", "kopie.csv"],
        lambda l: "Semikolon, Komma, UTF-8 mit BOM, Umlaute bleiben",
        lambda l: code_text(l.text("kopie.csv", "utf-8-sig"), "csv"),
    ),
    Beispiel(
        "Excel als UTF-8-CSV für Python oder eine Datenbank ausgeben.",
        ["csv", "fixtures/buchungen.xlsx", "--format", "utf8", "--out", "kopie.csv"],
        lambda l: "Komma, Punkt, `JJJJ-MM-TT`, UTF-8",
        lambda l: code_text(l.text("kopie.csv"), "csv") + ["", "Hinweise aus `kopie.aenderungen.json`:", ""] + code_json(l.json("kopie.aenderungen.json")["hinweise"]),
    ),
    Beispiel(
        "Diese Excel vor dem Import nach Power BI oder Microsoft Fabric prüfen.",
        ["profil", "fixtures/umsatz_bericht.xlsx", "--out", "report.md"],
        lambda l: f"Urteil `{l.erg['blaetter']['Umsatz']}`, erst bereinigen",
        lambda l: code_json({k: l.json("report.json")["blaetter"][0][k] for k in ("urteil", "gruende", "verbundene_im_datenbereich", "summenzeilen")}),
    ),
    Beispiel(
        "Eine Makro-Datei .xlsm prüfen, ohne Makros auszuführen.",
        ["profil", "fixtures/makro.xlsm", "--out", "report.md"],
        lambda l: f"Urteil `{l.erg['blaetter']['Eingabe']}`, Makros nicht ausgeführt",
        lambda l: bericht_abschnitt(l, "Eingabe"),
    ),
]


def ausfuehren(b: Beispiel, ordner: Path) -> Lauf:
    for p in ordner.iterdir():
        if p.is_file():
            p.unlink()
    argv = [str(FIX / a.removeprefix("fixtures/")) if a.startswith("fixtures/") else a for a in b.args]
    aus, err = io.StringIO(), io.StringIO()
    alt = Path.cwd()
    os.chdir(ordner)
    try:
        with contextlib.redirect_stdout(aus), contextlib.redirect_stderr(err):
            code = cli.main(argv)
    finally:
        os.chdir(alt)
    text = aus.getvalue() if code == 0 else err.getvalue()
    erg = json.loads(text)
    if code != 0 and "details" not in erg:
        raise SystemExit(f"Beispiel scheitert: {b.satz}: {text}")
    if code != 0:
        erg = {"exit_code": code, **erg}
    return Lauf(code, erg, ordner)


def befehl(b: Beispiel) -> str:
    teile = []
    for a in b.args:
        teile.append(a if a.replace("-", "").replace(".", "").replace("/", "").replace("_", "").replace(",", "").isalnum() else f'"{a}"')
    return "excel-skills " + " ".join(teile)


def bauen() -> tuple[str, str]:
    doc = [
        "# 20 Beispiele: Claude Skills für Excel",
        "",
        "Jedes Beispiel ist ein Satz, wie man ihn Claude Code, Codex oder einem anderen Agenten schreibt. Darunter steht der Befehl, den das Skill ausführt, und die echte Ausgabe gegen die Testdateien in `fixtures/`. Diese Datei erzeugt `python3 scripts/beispiele_bauen.py`, ein Test prüft, dass sie aktuell ist.",
        "",
        "Vor jedem Befehl steht im Skill `uvx --from git+https://github.com/pexon-consulting/excel-skills@v0.1.0`.",
        "",
    ]
    tabelle = ["| # | Das schreibst du dem Agenten | Befehl | Ergebnis |", "|---|---|---|---|"]
    with tempfile.TemporaryDirectory() as tmp:
        ordner = Path(tmp)
        for i, b in enumerate(BEISPIELE, start=1):
            lauf = ausfuehren(b, ordner)
            ergebnis = b.ergebnis(lauf)
            cmd = befehl(b)
            tabelle.append(f"| [{i}](docs/beispiele.md#beispiel-{i}) | {b.satz.replace('|', '/')} | `{b.args[0]}` | {ergebnis.replace('|', '/')} |")
            doc += [f'<a id="beispiel-{i}"></a>', f"## {i}. {b.satz}", "", "```bash", cmd, "```", "", f"Ergebnis: {ergebnis}", ""]
            doc += b.detail(lauf)
            doc.append("")
    return "\n".join(doc).rstrip() + "\n", "\n".join(tabelle)


def readme_mit(tabelle: str) -> str:
    text = README.read_text(encoding="utf-8")
    a, b = text.index(START), text.index(ENDE)
    return text[: a + len(START)] + "\n" + tabelle + "\n" + text[b:]


def main(argv: list[str] | None = None) -> int:
    args = argv if argv is not None else sys.argv[1:]
    doc, tabelle = bauen()
    readme = readme_mit(tabelle)
    if "--check" in args:
        ok = True
        if not DOC.is_file() or DOC.read_text(encoding="utf-8") != doc:
            print(f"veraltet: {DOC}", file=sys.stderr)
            ok = False
        if README.read_text(encoding="utf-8") != readme:
            print(f"veraltet: {README}", file=sys.stderr)
            ok = False
        return 0 if ok else 1
    DOC.write_text(doc, encoding="utf-8")
    README.write_text(readme, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
