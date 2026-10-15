"""Erzeugt die synthetischen Testdateien in fixtures/. Keine Kundendaten, alle Namen erfunden."""

from __future__ import annotations

import datetime as dt
import re
import zipfile
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
FIX = ROOT / "fixtures"
FEST = dt.datetime(2026, 10, 15, 12, 0, 0)


def _speichern(wb: Workbook, pfad: Path, cache: dict[str, dict[str, tuple[str, str]]] | None = None) -> None:
    """Speichert und setzt gespeicherte Formel-Ergebnisse, wie Excel sie hinterlegt."""
    wb.properties.creator = "excel-skills fixtures"
    wb.properties.created = FEST
    wb.properties.modified = FEST
    pfad.parent.mkdir(parents=True, exist_ok=True)
    wb.save(pfad)
    if not cache:
        return
    namen = {ws.title: f"xl/worksheets/sheet{i}.xml" for i, ws in enumerate(wb.worksheets, start=1)}
    teile = {}
    with zipfile.ZipFile(pfad) as z:
        infos = z.infolist()
        for i in infos:
            teile[i.filename] = z.read(i.filename)
    for blatt, zellen in cache.items():
        xml = teile[namen[blatt]].decode("utf-8")
        for ref, (typ, wert) in zellen.items():
            muster = re.compile(rf'<c r="{ref}"([^>]*)><f>(.*?)</f><v\s*/?>(?:</v>)?</c>')
            attr = {"n": "", "e": ' t="e"', "str": ' t="str"'}[typ]

            def ersetzen(m: re.Match[str], ref: str = ref, attr: str = attr, wert: str = wert) -> str:
                rest = re.sub(r' t="[^"]*"', "", m.group(1))
                return f'<c r="{ref}"{rest}{attr}><f>{m.group(2)}</f><v>{wert}</v></c>'

            xml, n = muster.subn(ersetzen, xml)
            if n != 1:
                raise SystemExit(f"Cache für {blatt}!{ref} nicht gesetzt")
        teile[namen[blatt]] = xml.encode("utf-8")
    with zipfile.ZipFile(pfad, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i, teile[i.filename])


def umsatz_bericht() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Umsatz"
    ws["A1"] = "Umsatzbericht 2026"
    ws.merge_cells("A1:F1")
    ws["A2"] = "Stand: 31.12.2026, Vertrieb"
    ws.append([])
    ws.append(["Region", "Kunde", "Monat", "Menge", "Umsatz", "Menge doppelt"])
    zeilen = [
        ("Nord", " Muster Maschinenbau GmbH ", "31.01.2026", 120, "1.234,56"),
        (None, "Beispiel Logistik AG", "28.02.2026", 80, "987,10"),
        (None, "Muster Handel KG", "31.03.2026", 45, "1.234"),
    ]
    for i, z in enumerate(zeilen, start=5):
        ws.append([*z, f"=D{i}*2"])
    ws.merge_cells("A5:A7")
    ws.append(["Zwischensumme Nord", None, None, "=SUM(D5:D7)", None, None])
    ws.append([])
    zeilen = [
        ("Süd", "Beispiel Werkzeuge GmbH", "01/31/2026", 200, "12.500,00"),
        (None, "Muster Software GmbH", "30.04.26", 15, "-250,75"),
    ]
    for i, z in enumerate(zeilen, start=10):
        ws.append([*z, f"=D{i}*2"])
    ws.merge_cells("A10:A11")
    ws.append(["Zwischensumme Süd", None, None, "=SUM(D10:D11)", None, None])
    ws.append(["Gesamt", None, None, "=D8+D12", None, None])
    notiz = wb.create_sheet("Notizen")
    notiz["A1"] = "Erste Menge Nord"
    notiz["B1"] = "=Umsatz!D5"
    notiz["A2"] = "Menge Süd, zweite Zeile"
    notiz["B2"] = "=Umsatz!D11"
    _speichern(wb, FIX / "umsatz_bericht.xlsx")


def kunden_liste(name: str, zeilen: list[tuple]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Kunden"
    ws.append(["Kundennummer", "Name", "Ort", "PLZ", "Umsatz 2026"])
    for z in zeilen:
        ws.append(list(z))
    _speichern(wb, FIX / name)


KUNDEN = [
    ("K-1001", "Muster Maschinenbau GmbH", "Dresden", "01067", 48200.0),
    ("K-1002", "Beispiel Logistik AG", "Hamburg", "20095", 12750.5),
    ("K-1003", "Muster Handel KG", "Kiel", "24103", 9100.0),
    ("K-1004", "Beispiel Werkzeuge GmbH", "München", "80331", 66300.0),
]


def formeln_kaputt() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Kalkulation"
    ws.append(["Artikel", "Menge", "Preis", "Summe", "Lieferant", "Anteil"])
    daten = [("A-100", 10, 2.5), ("A-200", 4, 12.0), ("A-300", 7, 3.2), ("A-999", 3, 8.0), ("A-400", 0, 5.0), ("A-500", 6, 1.5)]
    for i, (a, m, p) in enumerate(daten, start=2):
        ws.append([a, m, p, f"=B{i}*C{i}", f"=VLOOKUP(A{i},Lieferanten!A:B,2,FALSE)", f"=D{i}/B{i}"])
    ws["D4"] = 99
    ws["D5"] = "=B5*#REF!"
    ws["D7"] = None
    lief = wb.create_sheet("Lieferanten")
    lief.append(["Artikel", "Lieferant", "Listenpreis extern"])
    for a, l in (("A-100", "Muster Metall GmbH"), ("A-200", "Beispiel Kunststoff AG"), ("A-300", "Muster Metall GmbH"),
                 ("A-400", "Muster Handel KG"), ("A-500", "Beispiel Werkzeuge GmbH")):
        lief.append([a, l, None])
    lief["C2"] = "='[1]Preisliste'!B2"
    cache = {
        "Kalkulation": {
            "D2": ("n", "25"), "D3": ("n", "48"), "D5": ("e", "#REF!"), "D6": ("n", "0"),
            "E2": ("str", "Muster Metall GmbH"), "E3": ("str", "Beispiel Kunststoff AG"), "E4": ("str", "Muster Metall GmbH"),
            "E5": ("e", "#N/A"), "E6": ("str", "Muster Handel KG"), "E7": ("str", "Beispiel Werkzeuge GmbH"),
            "F2": ("n", "2.5"), "F3": ("n", "12"), "F4": ("n", "33"), "F5": ("e", "#REF!"), "F6": ("e", "#DIV/0!"),
        }
    }
    _speichern(wb, FIX / "formeln_kaputt.xlsx", cache)


def preise(name: str, zeilen: list[tuple]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Preise"
    ws.append(["Artikelnummer", "Bezeichnung", "Preis", "Einheit"])
    for z in zeilen:
        ws.append(list(z))
    _speichern(wb, FIX / name)


def monat(ordner: str, name: str, kopf: list[str], zeilen: list[tuple], summe: bool = False) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Buchungen"
    ws.append(kopf)
    for z in zeilen:
        ws.append(list(z))
    if summe:
        n = len(zeilen) + 1
        ws.append(["Summe", None, f"=SUM(C2:C{n})"])
    for row in ws.iter_rows(min_row=2, max_col=1):
        for c in row:
            if isinstance(c.value, dt.datetime):
                c.number_format = "DD.MM.YYYY"
    _speichern(wb, FIX / ordner / name)


def buchungen() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Buchungen"
    ws.append(["Belegdatum", "Konto", "Gegenkonto", "Betrag netto", "Betrag brutto", "Buchungstext"])
    daten = [
        (dt.datetime(2026, 1, 15), 8400, 10001, 1000.0, "Rechnung 2026-001 Müller; Wartung"),
        (dt.datetime(2026, 1, 31), 8400, 10002, 2499.99, 'Projekt "Fabric" Phase 1'),
        (dt.datetime(2026, 2, 3), 4930, 70001, 89.9, "Bürobedarf"),
    ]
    for i, (d, k, g, n, t) in enumerate(daten, start=2):
        ws.append([d, k, g, n, f"=ROUND(D{i}*1.19,2)", t])
        ws.cell(row=i, column=1).number_format = "DD.MM.YYYY"
    cache = {"Buchungen": {"E2": ("n", "1190"), "E3": ("n", "2974.99")}}
    _speichern(wb, FIX / "buchungen.xlsx", cache)


def makro() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Eingabe"
    ws.append(["Artikel", "Menge", "Preis"])
    ws.append(["A-100", "10", "2,50"])
    ws.append(["A-200", "4", "12,00"])
    tmp = FIX / "_makro.xlsx"
    _speichern(wb, tmp)
    ziel = FIX / "makro.xlsm"
    with zipfile.ZipFile(tmp) as zin, zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as zout:
        for i in zin.infolist():
            d = zin.read(i.filename)
            if i.filename == "[Content_Types].xml":
                s = d.decode().replace(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                ).replace("</Types>", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
                d = s.encode()
            if i.filename == "xl/_rels/workbook.xml.rels":
                d = d.decode().replace(
                    "</Relationships>",
                    '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
                ).encode()
            zout.writestr(i, d)
        # Platzhalter statt echtem VBA-Projekt: Bytes, die kein Programm ausführen kann.
        zout.writestr("xl/vbaProject.bin", b"EXCEL-SKILLS-FIXTURE: kein ausfuehrbarer Code\n")
    tmp.unlink()


def main() -> None:
    umsatz_bericht()
    kunden_liste("kunden_liste.xlsx", KUNDEN)
    kunden_liste("kunden_liste_neu.xlsx", [
        KUNDEN[0],
        ("K-1002", "Beispiel Logistik AG", "Bremen", "28195", 12750.5),
        KUNDEN[3],
        ("K-1005", "Muster Software GmbH", "München", "80333", 18400.0),
    ])
    formeln_kaputt()
    preise("preise_alt.xlsx", [
        ("4711", "Schraube M6 verzinkt", 0.12, "Stück"),
        ("4712", "Mutter M6", 0.05, "Stück"),
        ("4713", "Unterlegscheibe 6 mm", 0.03, "Stück"),
        ("4714", "Gewindestange M6 1 m", 2.4, "Stück"),
    ])
    preise("preise_neu.xlsx", [
        ("4711", "Schraube M6 verzinkt", 0.14, "Stück"),
        ("4712", "Mutter M6 verzinkt", 0.05, "Stück"),
        ("4714", "Gewindestange M6 1 m", 2.4, "Stück"),
        ("4715", "Hutmutter M6", 0.22, "Stück"),
    ])
    kopf = ["Datum", "Kunde", "Betrag"]
    monat("monate", "2026-01.xlsx", kopf, [(dt.datetime(2026, 1, 5), "Muster Maschinenbau GmbH", 1200.0), (dt.datetime(2026, 1, 19), "Beispiel Logistik AG", 640.5)], summe=True)
    monat("monate", "2026-02.xlsx", kopf, [(dt.datetime(2026, 2, 2), "Muster Handel KG", 310.0)])
    monat("monate", "2026-03.xlsx", kopf, [(dt.datetime(2026, 3, 9), "Beispiel Werkzeuge GmbH", 2150.0), (dt.datetime(2026, 3, 23), "Muster Maschinenbau GmbH", 980.0)])
    monat("monate_abweichend", "2026-01.xlsx", kopf, [(dt.datetime(2026, 1, 5), "Muster Maschinenbau GmbH", 1200.0)])
    monat("monate_abweichend", "2026-02.xlsx", ["Datum", "Kunde", "Netto"], [(dt.datetime(2026, 2, 2), "Muster Handel KG", 310.0)])
    buchungen()
    makro()


if __name__ == "__main__":
    main()
