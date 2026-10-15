"""Konsolenbefehl excel-skills. Jeder Unterbefehl gibt eine Zeile JSON auf stdout aus."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import __version__, bereinigen, csv_export, formeln, laden, profil, regeln, stapeln, vergleich


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="excel-skills",
        description="Claude Skills für Excel: deutsche Excel-Dateien lokal prüfen, bereinigen, vergleichen, zusammenführen und als CSV ausgeben. Das Original wird nie geändert.",
    )
    p.add_argument("--version", action="version", version=f"excel-skills {__version__}")
    sub = p.add_subparsers(dest="befehl", required=True)

    s = sub.add_parser("profil", help="Kopfzeile, Datenbereich, Spaltentypen, Urteil tabelle, bericht oder formular")
    s.add_argument("datei", type=Path)
    s.add_argument("--out", type=Path, default=Path("report.md"))

    s = sub.add_parser("bereinigen", help="bereinigte Kopie schreiben, plus Änderungsprotokoll als JSON")
    s.add_argument("datei", type=Path)
    s.add_argument("--blatt")
    s.add_argument("--out", type=Path)
    s.add_argument("--ohne", action="append", default=[],
                   help=f"Regeln abschalten, kommagetrennt: {','.join(bereinigen.REGELN)}")

    s = sub.add_parser("formeln", help="Fehlerwerte, harte Werte, externe Verknüpfungen, fehlende Formeln")
    s.add_argument("datei", type=Path)
    s.add_argument("--out", type=Path, default=Path("report.md"))

    s = sub.add_parser("vergleich", help="zwei Excel-Tabellen vergleichen")
    s.add_argument("alt", type=Path)
    s.add_argument("neu", type=Path)
    s.add_argument("--blatt")
    s.add_argument("--schluessel", help="Spaltenname, über den Zeilen zugeordnet werden")
    s.add_argument("--out", type=Path, default=Path("report.md"))

    s = sub.add_parser("stapeln", help="Excel-Dateien eines Ordners mit gleichen Spalten zusammenführen")
    s.add_argument("ordner", type=Path)
    s.add_argument("--blatt")
    s.add_argument("--out", type=Path)
    s.add_argument("--werte", action="store_true", help="Formeln als gespeicherte Ergebnisse übernehmen")

    s = sub.add_parser("csv", help="Blatt als CSV: datev, excel-de oder utf8")
    s.add_argument("datei", type=Path)
    s.add_argument("--blatt")
    s.add_argument("--out", type=Path)
    s.add_argument("--format", default="excel-de", choices=sorted(regeln.CSV_FORMATE))

    sub.add_parser("schema", help="die Regeln dieser Version als JSON")
    return p


def _bericht_pfad(out: Path, eingaben: list[Path]) -> Path:
    return laden.pruefe_ausgabe(out, eingaben, (".md",))


def ausfuehren(args: argparse.Namespace) -> dict:
    b = args.befehl
    if b == "schema":
        return {"version": __version__, **regeln.schema()}
    if b == "profil":
        return profil.ausfuehren(args.datei, _bericht_pfad(args.out, [args.datei]))
    if b == "formeln":
        return formeln.ausfuehren(args.datei, _bericht_pfad(args.out, [args.datei]))
    if b == "vergleich":
        return vergleich.ausfuehren(args.alt, args.neu, args.blatt, args.schluessel, _bericht_pfad(args.out, [args.alt, args.neu]))
    if b == "bereinigen":
        ohne = {r.strip() for teil in args.ohne for r in teil.split(",") if r.strip()}
        return bereinigen.ausfuehren(args.datei, args.blatt, args.out, ohne)
    if b == "stapeln":
        return stapeln.ausfuehren(args.ordner, args.blatt, args.out, args.werte)
    if b == "csv":
        return csv_export.ausfuehren(args.datei, args.blatt, args.out, args.format)
    raise laden.Abbruch(f"unbekannter Befehl {b}")


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        erg = ausfuehren(args)
    except laden.Abbruch as exc:
        print(json.dumps({"fehler": exc.meldung, "details": exc.details}, ensure_ascii=False, default=str), file=sys.stderr)
        return exc.code
    print(json.dumps(erg, ensure_ascii=False, indent=None if args.befehl != "schema" else 2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
