"""excel-skills bereinigen: schreibt eine bereinigte Kopie. Das Original bleibt unverändert."""

from __future__ import annotations

import bisect
import re
from pathlib import Path
from typing import Any

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from . import bericht, laden, regeln
from .profil import hinweise_zelle

REGELN = ("kopfzeile", "verbunden", "leerzeilen", "summen", "zahlen", "datum", "trimmen")
_LEER = "\u00a0\u2009\u202f "
_ZELLE = re.compile(r"^(\$?[A-Za-z]{1,3})(\$?)(\d+)$")
_ZEILEN = re.compile(r"^(\$?)(\d+):(\$?)(\d+)$")


class Zeilenplan:
    """Welche Zeilen eines Blatts entfallen, und wohin die übrigen rutschen."""

    def __init__(self, blatt: str, geloescht: set[int]):
        self.blatt = blatt
        self.geloescht = sorted(geloescht)
        self._menge = set(geloescht)

    def neu(self, zeile: int) -> int | None:
        if zeile in self._menge:
            return None
        return zeile - bisect.bisect_left(self.geloescht, zeile)

    def bereich(self, von: int, bis: int) -> tuple[int, int] | None:
        z = [r for r in range(von, bis + 1) if r not in self._menge] if bis - von < 100_000 else None
        if z is None:
            a = next((r for r in range(von, bis + 1) if r not in self._menge), None)
            b = next((r for r in range(bis, von - 1, -1) if r not in self._menge), None)
            if a is None or b is None:
                return None
            return self.neu(a), self.neu(b)  # type: ignore[return-value]
        if not z:
            return None
        return self.neu(z[0]), self.neu(z[-1])  # type: ignore[return-value]


def _blattname(prefix: str) -> str | None:
    if "[" in prefix:
        return None
    if prefix.startswith("'") and prefix.endswith("'"):
        return prefix[1:-1].replace("''", "'")
    return prefix


def _verschiebe_bezug(ref: str, plan: Zeilenplan) -> str | None:
    m = _ZELLE.match(ref)
    if m:
        neu = plan.neu(int(m.group(3)))
        return None if neu is None else f"{m.group(1)}{m.group(2)}{neu}"
    if ":" in ref:
        a, b = ref.split(":", 1)
        ma, mb = _ZELLE.match(a), _ZELLE.match(b)
        if ma and mb:
            r = plan.bereich(int(ma.group(3)), int(mb.group(3)))
            if r is None:
                return None
            return f"{ma.group(1)}{ma.group(2)}{r[0]}:{mb.group(1)}{mb.group(2)}{r[1]}"
        mz = _ZEILEN.match(ref)
        if mz:
            r = plan.bereich(int(mz.group(2)), int(mz.group(4)))
            if r is None:
                return None
            return f"{mz.group(1)}{r[0]}:{mz.group(3)}{r[1]}"
    return ref


def verschiebe_formel(formel: str, auf_blatt: str, plan: Zeilenplan) -> tuple[str, str | None]:
    """Passt Zeilenbezüge auf plan.blatt an. Rückgabe: neue Formel und Konfliktgrund oder None."""
    try:
        tok = Tokenizer(formel)
    except Exception as exc:
        return formel, f"Formel nicht lesbar: {exc}"
    geaendert = False
    for item in tok.items:
        if item.type != "OPERAND" or item.subtype != "RANGE":
            continue
        wert = item.value
        if "!" in wert:
            prefix, ref = wert.rsplit("!", 1)
            if _blattname(prefix) != plan.blatt:
                continue
        else:
            prefix, ref = None, wert
            if auf_blatt != plan.blatt:
                continue
        neu = _verschiebe_bezug(ref, plan)
        if neu is None:
            return formel, f"Bezug {wert} zeigt auf eine Zeile, die entfällt"
        if neu != ref:
            item.value = f"{prefix}!{neu}" if prefix is not None else neu
            geaendert = True
    return (tok.render() if geaendert else formel), None


def _alle_formeln(wb) -> list[tuple[Worksheet, Any]]:
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if laden.ist_formel(c):
                    out.append((ws, c))
    return out


def _zeile_leer(ws: Worksheet, nr: int) -> bool:
    return all(laden.ist_leer(c.value) and not laden.ist_formel(c) for c in ws[nr])


def ausfuehren(datei: Path, blatt_name: str | None, out: Path | None, ohne: set[str]) -> dict[str, Any]:
    unbekannt = ohne - set(REGELN)
    if unbekannt:
        raise laden.Abbruch(f"--ohne kennt nur {', '.join(REGELN)}, nicht {', '.join(sorted(unbekannt))}")
    xlsm = datei.suffix.lower() == ".xlsm"
    out = laden.pruefe_ausgabe(out, [datei], (".xlsm",) if xlsm else (".xlsx",))
    wb = laden.lade(datei, vba=True)
    try:
        return _bereinigen(wb, datei, blatt_name, out, ohne)
    finally:
        laden.schliessen(wb)


def _bereinigen(wb, datei: Path, blatt_name: str | None, out: Path, ohne: set[str]) -> dict[str, Any]:
    ws = laden.blatt(wb, blatt_name)
    t = laden.tabelle(ws)
    if t.kopf_zeile is None:
        raise laden.Abbruch(f"Blatt {ws.title}: keine Kopfzeile gefunden ({t.begruendung}). Nichts geschrieben.", code=3)
    aktiv = [r for r in REGELN if r not in ohne]
    kopf = t.kopf_zeile
    aenderungen: list[dict[str, Any]] = []
    hinweise: list[dict[str, Any]] = []

    geloescht: dict[int, str] = {}
    if "kopfzeile" in aktiv:
        for nr in range(1, kopf):
            geloescht[nr] = "kopfzeile"
    letzte = t.zeilen[-1][0] if t.zeilen else kopf
    if "leerzeilen" in aktiv:
        for nr in range(kopf + 1, letzte):
            if _zeile_leer(ws, nr):
                geloescht[nr] = "leerzeilen"
    if "summen" in aktiv:
        for nr, werte in t.zeilen:
            if regeln.ist_summenzeile(werte):
                geloescht[nr] = "summen"
    plan = Zeilenplan(ws.title, set(geloescht))

    neue_formeln: dict[tuple[str, str], Any] = {}
    konflikte: list[dict[str, Any]] = []
    if geloescht:
        for fws, c in _alle_formeln(wb):
            if fws is ws and c.row in geloescht:
                continue
            text = laden.formeltext(c)
            neu, grund = verschiebe_formel(text, fws.title, plan)
            if grund:
                konflikte.append({"blatt": fws.title, "zelle": c.coordinate, "formel": text, "grund": grund})
            elif neu != text:
                if isinstance(c.value, ArrayFormula):
                    ref = _verschiebe_bezug(c.value.ref, plan) if fws is ws else c.value.ref
                    neue_formeln[(fws.title, c.coordinate)] = ArrayFormula(ref=ref or c.value.ref, text=neu)
                else:
                    neue_formeln[(fws.title, c.coordinate)] = neu
        for name, dn in list(wb.defined_names.items()):
            text = dn.attr_text or ""
            neu, grund = verschiebe_formel("=" + text, "", plan)
            if grund:
                konflikte.append({"blatt": None, "zelle": f"Name {name}", "formel": text, "grund": grund})
            elif neu != "=" + text:
                dn.attr_text = neu[1:]
    if konflikte:
        raise laden.Abbruch(
            "Formeln zeigen auf Zeilen, die entfallen würden. Nichts geschrieben. "
            "Mit --ohne summen oder --ohne leerzeilen erneut versuchen.",
            code=3, details=konflikte,
        )

    # Verbundene Zellen: auflösen und füllen, oder nach dem Löschen verschoben neu verbinden.
    neu_verbinden: list[tuple[int, int, int, int]] = []
    for m in sorted(ws.merged_cells.ranges, key=lambda m: (m.min_row, m.min_col)):
        bereich = str(m)
        r1, c1, r2, c2 = m.min_row, m.min_col, m.max_row, m.max_col
        oben_links = ws.cell(row=r1, column=c1)
        wert = oben_links.value
        ws.unmerge_cells(bereich)
        ganz_weg = all(r in geloescht for r in range(r1, r2 + 1))
        if ganz_weg:
            continue
        if "verbunden" in aktiv and r2 >= kopf:
            if laden.ist_formel(oben_links):
                hinweise.append({"zelle": bereich, "regel": "verbunden", "hinweis": "verbunden mit Formel, nur die Zelle oben links behält die Formel"})
                aenderungen.append({"blatt": ws.title, "zelle": bereich, "alt": "verbunden", "neu": "aufgelöst", "regel": "verbunden"})
                continue
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) != (r1, c1):
                        ws.cell(row=r, column=c).value = wert
            aenderungen.append({"blatt": ws.title, "zelle": bereich, "alt": "verbunden", "neu": wert, "regel": "verbunden"})
        else:
            rb = plan.bereich(r1, r2)
            if rb:
                neu_verbinden.append((rb[0], c1, rb[1], c2))

    for nr in range(kopf, ws.max_row + 1):
        if nr in geloescht:
            continue
        for sp in range(t.erste_spalte, t.letzte_spalte + 1):
            c = ws.cell(row=nr, column=sp)
            v = c.value
            if not isinstance(v, str) or laden.ist_formel(c):
                continue
            ref = c.coordinate
            if "trimmen" in aktiv:
                s = v.strip(_LEER)
                if s != v:
                    c.value = s
                    aenderungen.append({"blatt": ws.title, "zelle": ref, "alt": v, "neu": s, "regel": "trimmen"})
                    v = s
            if nr == kopf:
                continue
            if "zahlen" in aktiv:
                z = regeln.deutsche_zahl(v)
                if z.wert is not None:
                    c.value = z.wert
                    if c.number_format == "@":
                        c.number_format = "General"
                    aenderungen.append({"blatt": ws.title, "zelle": ref, "alt": v, "neu": z.wert, "regel": "deutsche-zahl"})
                    continue
            if "datum" in aktiv:
                d = regeln.deutsches_datum(v)
                if d is not None:
                    c.value = d
                    c.number_format = "DD.MM.YYYY"
                    aenderungen.append({"blatt": ws.title, "zelle": ref, "alt": v, "neu": d, "regel": "datum"})
                    continue
            h = hinweise_zelle(v, ref)
            if h:
                hinweise.append(h)

    geloeschte_zeilen = [
        {"zeile": nr, "regel": geloescht[nr], "werte": [bericht.jsonwert(c.value) for c in ws[nr] if not laden.ist_leer(c.value)]}
        for nr in sorted(geloescht)
    ]
    if geloescht:
        for name, vorhanden in (
            ("Datenüberprüfung", bool(ws.data_validations.dataValidation)),
            ("bedingte Formatierung", bool(len(ws.conditional_formatting))),
            ("intelligente Tabelle", bool(ws.tables)),
            ("Autofilter", bool(ws.auto_filter.ref)),
        ):
            if vorhanden:
                hinweise.append({"zelle": None, "regel": "zeilen", "hinweis": f"{name} vorhanden, deren Bereich wurde nicht verschoben, bitte prüfen"})
        for nr in sorted(geloescht, reverse=True):
            ws.delete_rows(nr, 1)

    for (titel, koord), formel in neue_formeln.items():
        ziel_ws = wb[titel]
        zelle = ziel_ws[koord]
        if ziel_ws is ws:
            neu_zeile = plan.neu(zelle.row)
            zelle = ws.cell(row=neu_zeile, column=zelle.column)  # type: ignore[arg-type]
        zelle.value = formel
    for r1, c1, r2, c2 in neu_verbinden:
        if (r1, c1) != (r2, c2):
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    wb.save(out)
    protokoll = laden.begleitdatei(out, "aenderungen.json")
    daten = {
        "befehl": "bereinigen",
        "datei": datei.name,
        "blatt": ws.title,
        "ausgabe": out.name,
        "zellbezuege": "Zellen und Zeilen in der Zählung des Originals",
        "regeln_aktiv": aktiv,
        "kopfzeile_original": kopf,
        "geloeschte_zeilen": geloeschte_zeilen,
        "formeln_angepasst": len(neue_formeln),
        "aenderungen": aenderungen,
        "hinweise": hinweise,
    }
    bericht.schreibe_json(protokoll, daten)
    return {
        "befehl": "bereinigen",
        "ausgabe": str(out),
        "aenderungen": str(protokoll),
        "zellen_geaendert": len(aenderungen),
        "zeilen_geloescht": len(geloeschte_zeilen),
        "formeln_angepasst": len(neue_formeln),
        "hinweise": len(hinweise),
        "kopfzeile_jetzt": 1 if "kopfzeile" in aktiv else kopf,
        "spalte_letzte": get_column_letter(t.letzte_spalte),
    }
