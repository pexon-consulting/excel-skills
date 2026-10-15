"""Die Entscheidungsregeln. Jede Konstante hier steht wortgleich in docs/regeln.md."""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from typing import Any

KOPFZEILE_TEXTANTEIL = 0.6
KOPFZEILE_SUCHTIEFE = 50
FORMELSPALTE_ANTEIL = 0.8
FORMELSPALTE_MIN = 2
SPALTENTYP_ANTEIL = 0.8
FORMULAR_MAX_SPALTEN = 3
SUMMEN_WOERTER = ("zwischensumme", "summe", "gesamt", "total")

FEHLERWERTE = {
    "#REF!": "#BEZUG!",
    "#N/A": "#NV",
    "#VALUE!": "#WERT!",
    "#DIV/0!": "#DIV/0!",
    "#NAME?": "#NAME?",
    "#NUM!": "#ZAHL!",
    "#NULL!": "#NULL!",
}

CSV_FORMATE = {
    "datev": {
        "trennzeichen": ";",
        "dezimalzeichen": ",",
        "datum": "TT.MM.JJJJ",
        "kodierung": "cp1252",
        "zeilenende": "CRLF",
        "text_in_anfuehrungszeichen": True,
        "wahrheitswerte": ["WAHR", "FALSCH"],
    },
    "excel-de": {
        "trennzeichen": ";",
        "dezimalzeichen": ",",
        "datum": "TT.MM.JJJJ",
        "kodierung": "utf-8-sig",
        "zeilenende": "CRLF",
        "text_in_anfuehrungszeichen": False,
        "wahrheitswerte": ["WAHR", "FALSCH"],
    },
    "utf8": {
        "trennzeichen": ",",
        "dezimalzeichen": ".",
        "datum": "JJJJ-MM-TT",
        "kodierung": "utf-8",
        "zeilenende": "LF",
        "text_in_anfuehrungszeichen": False,
        "wahrheitswerte": ["TRUE", "FALSE"],
    },
}

_LEER = "\u00a0\u2009\u202f "
_ZAHL_PUNKT = re.compile(r"^-?\d{1,3}(?:\.\d{3})+,\d+$")
_ZAHL_LEER = re.compile(r"^-?\d{1,3}(?:[ \u00a0\u2009\u202f]\d{3})+,\d+$")
_ZAHL_KOMMA = re.compile(r"^-?\d+,\d+$")
_ZAHL_MEHRDEUTIG = re.compile(r"^-?\d{1,3}(?:\.\d{3})+$")
_DATUM_DE = re.compile(r"^(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})$")
_DATUM_US = re.compile(r"^\d{1,2}/\d{1,2}/(?:\d{4}|\d{2})$")
_SUMME = re.compile(r"^\s*(?:" + "|".join(SUMMEN_WOERTER) + ")", re.IGNORECASE)


@dataclass(frozen=True)
class Zahl:
    wert: float | None
    mehrdeutig: bool = False


def deutsche_zahl(text: str) -> Zahl:
    """#deutsche-zahl: 1.234,56 und 1 234,56 werden Zahl, 1.234 bleibt Text."""
    s = text.strip(_LEER)
    if _ZAHL_PUNKT.match(s) or _ZAHL_LEER.match(s) or _ZAHL_KOMMA.match(s):
        roh = re.sub(r"[.\u00a0\u2009\u202f ]", "", s).replace(",", ".")
        return Zahl(float(roh))
    if _ZAHL_MEHRDEUTIG.match(s):
        return Zahl(None, mehrdeutig=True)
    return Zahl(None)


def deutsches_datum(text: str) -> dt.datetime | None:
    """#datum: 31.12.2026 und 31.12.26 werden Datum. Zweistellig: 00-29 ist 20xx, 30-99 ist 19xx."""
    m = _DATUM_DE.match(text.strip(_LEER))
    if not m:
        return None
    tag, monat, jahr = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if len(m.group(3)) == 2:
        jahr += 2000 if jahr <= 29 else 1900
    try:
        return dt.datetime(jahr, monat, tag)
    except ValueError:
        return None


def ist_us_datum(text: str) -> bool:
    return bool(_DATUM_US.match(text.strip(_LEER)))


def ist_summenzeile(werte: list[Any]) -> bool:
    """#summenzeile: die erste Textzelle der Zeile beginnt mit Summe, Gesamt, Total oder Zwischensumme."""
    for w in werte:
        if isinstance(w, str) and w.strip() and not w.startswith("="):
            return bool(_SUMME.match(w))
    return False


def zelltyp(wert: Any, formel: bool = False) -> str:
    """Typ einer Zelle für Kopfzeile und Spaltenprofil."""
    if formel:
        return "formel"
    if wert is None or (isinstance(wert, str) and not wert.strip(_LEER)):
        return "leer"
    if isinstance(wert, bool):
        return "wahrheitswert"
    if isinstance(wert, (int, float)):
        return "zahl"
    if isinstance(wert, (dt.datetime, dt.date, dt.time)):
        return "datum"
    if isinstance(wert, str):
        if wert in FEHLERWERTE:
            return "fehler"
        if deutsche_zahl(wert).wert is not None:
            return "zahl_als_text"
        if deutsches_datum(wert) is not None:
            return "datum_als_text"
        return "text"
    return "text"


def finde_kopfzeile(zeilen: list[list[tuple[Any, bool]]], spalten: int) -> tuple[int | None, str]:
    """#kopfzeile: erste Zeile mit mindestens 60 Prozent Text, deren Folgezeile einen anderen Typ hat.

    zeilen: je Zeile eine Liste von (wert, ist_formel). Rückgabe: Index ab 0 und Begründung.
    """
    if spalten <= 0:
        return None, "Blatt ist leer"
    kandidaten: list[int] = []
    for i, zeile in enumerate(zeilen[:KOPFZEILE_SUCHTIEFE]):
        texte = sum(1 for w, f in zeile if zelltyp(w, f) == "text")
        if texte / spalten >= KOPFZEILE_TEXTANTEIL:
            kandidaten.append(i)
    for i in kandidaten:
        if i + 1 >= len(zeilen):
            continue
        darunter = [zelltyp(w, f) for w, f in zeilen[i + 1]]
        if any(t not in ("text", "leer") for t in darunter):
            anteil = sum(1 for w, f in zeilen[i] if zelltyp(w, f) == "text")
            return i, f"{anteil} von {spalten} Spalten Text, Zeile darunter hat Zahl, Datum oder Formel"
    if kandidaten:
        return kandidaten[0], "erste Zeile mit mindestens 60 Prozent Text, darunter nur Text"
    return None, "keine Zeile mit mindestens 60 Prozent Text in den ersten 50 Zeilen"


def schema() -> dict[str, Any]:
    return {
        "kopfzeile": {
            "textanteil_mindestens": KOPFZEILE_TEXTANTEIL,
            "suchtiefe_zeilen": KOPFZEILE_SUCHTIEFE,
            "folgezeile": "mindestens eine Zelle Zahl, Datum, Wahrheitswert oder Formel",
            "ersatz": "erste Zeile mit genug Text, wenn keine Folgezeile passt",
        },
        "tabelle_oder_bericht": {
            "tabelle": "Kopfzeile gefunden, keine verbundenen Zellen im Datenbereich, keine Summenzeile vor der letzten Datenzeile",
            "bericht": "sonst",
            "formular": f"weniger als {FORMULAR_MAX_SPALTEN} Spalten mit Daten",
        },
        "deutsche_zahl": {
            "wird_zahl": ["1.234,56", "-1.234,56", "1 234,56", "12,5"],
            "bleibt_text": ["1.234", "01067", "4711"],
        },
        "datum": {
            "wird_datum": ["31.12.2026", "31.12.26"],
            "zweistellig": "00-29 ist 20xx, 30-99 ist 19xx",
            "bleibt_text_mit_hinweis": ["12/31/2026"],
        },
        "summenzeile": {"erste_textzelle_beginnt_mit": list(SUMMEN_WOERTER)},
        "formelspalte": {"formelanteil_mindestens": FORMELSPALTE_ANTEIL, "formeln_mindestens": FORMELSPALTE_MIN},
        "spaltentyp": {"anteil_mindestens": SPALTENTYP_ANTEIL, "sonst": "gemischt"},
        "fehlerwerte": FEHLERWERTE,
        "csv_formate": CSV_FORMATE,
    }
