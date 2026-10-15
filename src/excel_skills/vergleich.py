"""excel-skills vergleich: zwei Excel-Tabellen vergleichen, über eine Schlüsselspalte oder Zelle für Zelle."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl.utils import get_column_letter

from . import bericht, laden


def _gleich(a: Any, b: Any) -> bool:
    if laden.ist_leer(a) and laden.ist_leer(b):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
        return abs(float(a) - float(b)) <= 1e-9 * max(1.0, abs(float(a)), abs(float(b)))
    return a == b


def _schluessel(wert: Any) -> str:
    v = bericht.jsonwert(wert)
    return "" if v is None else str(v).strip()


def nach_schluessel(alt: laden.Tabelle, neu: laden.Tabelle, spalte: str) -> dict[str, Any]:
    ia, ineu = alt.spalte(spalte), neu.spalte(spalte)
    gemeinsam = [k for k in alt.kopf if k in neu.kopf and k != spalte]
    hinweise: list[str] = []

    def index(t: laden.Tabelle, i: int, name: str) -> dict[str, tuple[int, list[Any]]]:
        d: dict[str, tuple[int, list[Any]]] = {}
        for nr, werte in t.zeilen:
            k = _schluessel(werte[i])
            if not k:
                continue
            if k in d:
                hinweise.append(f"{name}: Schlüssel {k} doppelt in Zeile {d[k][0]} und {nr}, verglichen wird die erste")
                continue
            d[k] = (nr, werte)
        return d

    da, dn = index(alt, ia, "alt"), index(neu, ineu, "neu")
    neu_zeilen = [{"schluessel": k, "zeile": dn[k][0], "werte": dict(zip(neu.kopf, map(bericht.jsonwert, dn[k][1])))} for k in dn if k not in da]
    weg = [{"schluessel": k, "zeile": da[k][0], "werte": dict(zip(alt.kopf, map(bericht.jsonwert, da[k][1])))} for k in da if k not in dn]
    geaendert = []
    for k, (nra, wa) in da.items():
        if k not in dn:
            continue
        nrn, wn = dn[k]
        for name in gemeinsam:
            a, b = wa[alt.kopf.index(name)], wn[neu.kopf.index(name)]
            if not _gleich(a, b):
                geaendert.append({"schluessel": k, "spalte": name, "alt": bericht.jsonwert(a), "neu": bericht.jsonwert(b), "zeile_alt": nra, "zeile_neu": nrn})
    return {"neu": neu_zeilen, "geloescht": weg, "geaendert": geaendert, "hinweise": hinweise}


def nach_zellen(alt: laden.Tabelle, neu: laden.Tabelle) -> dict[str, Any]:
    """Ohne Schlüssel: Zeile n der alten Tabelle gegen Zeile n der neuen, Spalten über den Namen."""
    gemeinsam = [k for k in alt.kopf if k in neu.kopf]
    geaendert = []
    for i in range(max(len(alt.zeilen), len(neu.zeilen))):
        za = alt.zeilen[i] if i < len(alt.zeilen) else None
        zn = neu.zeilen[i] if i < len(neu.zeilen) else None
        for name in gemeinsam:
            a = za[1][alt.kopf.index(name)] if za else None
            b = zn[1][neu.kopf.index(name)] if zn else None
            if not _gleich(a, b):
                sp = get_column_letter(neu.erste_spalte + neu.kopf.index(name))
                geaendert.append({"datenzeile": i + 1, "zelle": f"{sp}{zn[0]}" if zn else None, "spalte": name, "alt": bericht.jsonwert(a), "neu": bericht.jsonwert(b)})
    return {"geaendert": geaendert, "zeilen_alt": len(alt.zeilen), "zeilen_neu": len(neu.zeilen)}


def ausfuehren(alt_pfad: Path, neu_pfad: Path, blatt: str | None, schluessel: str | None, out: Path) -> dict[str, Any]:
    tabellen = []
    for p in (alt_pfad, neu_pfad):
        wb, wv = laden.lade_beide(p)
        ws = laden.blatt(wb, blatt)
        tabellen.append(laden.tabelle_pflicht(ws, wv[ws.title], datei=p))
    alt, neu = tabellen
    spalten = {"nur_alt": [k for k in alt.kopf if k not in neu.kopf], "nur_neu": [k for k in neu.kopf if k not in alt.kopf]}
    if schluessel:
        erg = nach_schluessel(alt, neu, schluessel)
        modus = "schluessel"
    else:
        erg = nach_zellen(alt, neu)
        modus = "zellen"
    daten = {"befehl": "vergleich", "alt": alt_pfad.name, "neu": neu_pfad.name, "modus": modus,
             "schluessel": schluessel, "spalten": spalten, **erg}

    zeilen = [f"Alt: `{bericht.zelle_md(alt_pfad.name)}`, Blatt {bericht.zelle_md(alt.blatt)}. "
              f"Neu: `{bericht.zelle_md(neu_pfad.name)}`, Blatt {bericht.zelle_md(neu.blatt)}.", ""]
    zeilen.append(f"Vergleich über Schlüsselspalte **{bericht.zelle_md(schluessel)}**." if schluessel
                  else "Vergleich Zeile für Zeile, Spalten über den Namen.")
    zeilen.append("")
    if spalten["nur_alt"] or spalten["nur_neu"]:
        zeilen += [f"Spalten nur alt: {', '.join(spalten['nur_alt']) or 'keine'}. Spalten nur neu: {', '.join(spalten['nur_neu']) or 'keine'}.", ""]
    if schluessel:
        zeilen += [f"## Neu ({len(erg['neu'])})", ""]
        zeilen += bericht.tabelle_md([schluessel, "Zeile"], [[z["schluessel"], z["zeile"]] for z in erg["neu"]])
        zeilen += [f"## Gelöscht ({len(erg['geloescht'])})", ""]
        zeilen += bericht.tabelle_md([schluessel, "Zeile"], [[z["schluessel"], z["zeile"]] for z in erg["geloescht"]])
        zeilen += [f"## Geändert ({len(erg['geaendert'])})", ""]
        zeilen += bericht.tabelle_md([schluessel, "Spalte", "alt", "neu"], [[g["schluessel"], g["spalte"], g["alt"], g["neu"]] for g in erg["geaendert"]])
        for h in erg["hinweise"]:
            zeilen.append(f"- {bericht.zelle_md(h, 200)}")
    else:
        zeilen += [f"## Geändert ({len(erg['geaendert'])})", ""]
        zeilen += bericht.tabelle_md(["Datenzeile", "Zelle neu", "Spalte", "alt", "neu"],
                                     [[g["datenzeile"], g["zelle"], g["spalte"], g["alt"], g["neu"]] for g in erg["geaendert"]])
    json_pfad = bericht.schreibe_bericht(out, "Excel-Vergleich", zeilen, daten)
    zusammen: dict[str, Any] = {"befehl": "vergleich", "bericht": str(out), "json": str(json_pfad), "modus": modus,
                                "geaendert": len(erg["geaendert"])}
    if schluessel:
        zusammen.update({"neu": len(erg["neu"]), "geloescht": len(erg["geloescht"])})
    return zusammen
