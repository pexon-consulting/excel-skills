# Claude Skills für Excel: 20 Beispiele mit echter Ausgabe

Jedes Beispiel ist ein Satz, wie man ihn Claude Code, Codex oder einem anderen Agenten schreibt. Darunter steht der Befehl, den das Skill ausführt, und die echte Ausgabe gegen die Testdateien in `fixtures/`. Diese Datei erzeugt `python3 scripts/beispiele_bauen.py`, ein Test prüft, dass sie aktuell ist.

Vor jedem Befehl steht im Skill `uvx --from git+https://github.com/pexon-consulting/excel-skills@v0.1.1`.

<a id="beispiel-1"></a>
## 1. Finde die Kopfzeile in dieser Excel, die erst in Zeile 4 anfängt.

```bash
excel-skills profil fixtures/umsatz_bericht.xlsx --out report.md
```

Ergebnis: Kopfzeile Zeile 4, Datenbereich A5:F13

#### Blatt Umsatz

- Urteil: **bericht**
- Kopfzeile: 4 (6 von 6 Spalten Text, Zeile darunter hat Zahl, Datum oder Formel)
- Datenbereich: A5:F13
- Verbundene Zellen: 3, davon im Datenbereich 2
- Summenzeilen: 8, 12, 13
- Gründe: 2 verbundene Bereiche im Datenbereich; Summenzeile vor der letzten Datenzeile: 8, 12; Kopfzeile erst in Zeile 4

| Spalte | Name | Typ | leer |
|---|---|---|---|
| A | Region | text | 60 % |
| B | Kunde | text | 0 % |
| C | Monat | datum_als_text | 0 % |
| D | Menge | zahl | 0 % |
| E | Umsatz | zahl_als_text | 0 % |
| F | Menge doppelt | formel | 0 % |

Hinweise:

| Zelle | Wert | Hinweis |
|---|---|---|
| C10 | 01/31/2026 | Datum mit Schrägstrich, Reihenfolge Tag und Monat unklar, bleibt Text |
| E7 | 1.234 | ohne Komma mehrdeutig, bleibt Text |


<a id="beispiel-2"></a>
## 2. Ist dieses Excel-Blatt eine Tabelle oder ein Bericht?

```bash
excel-skills profil fixtures/kunden_liste.xlsx --out report.md
```

Ergebnis: Urteil `tabelle`, PLZ bleibt Text

#### Blatt Kunden

- Urteil: **tabelle**
- Kopfzeile: 1 (5 von 5 Spalten Text, Zeile darunter hat Zahl, Datum oder Formel)
- Datenbereich: A2:E5
- Verbundene Zellen: 0, davon im Datenbereich 0
- Summenzeilen: keine
- Gründe: Kopfzeile in Zeile 1, keine verbundenen Zellen, keine Summenzeile mitten drin

| Spalte | Name | Typ | leer |
|---|---|---|---|
| A | Kundennummer | text | 0 % |
| B | Name | text | 0 % |
| C | Ort | text | 0 % |
| D | PLZ | text | 0 % |
| E | Umsatz 2026 | zahl | 0 % |

<a id="beispiel-3"></a>
## 3. Verbundene Zellen auflösen und den Wert nach unten füllen.

```bash
excel-skills bereinigen fixtures/umsatz_bericht.xlsx --ohne kopfzeile,leerzeilen,summen,zahlen,datum,trimmen --out kopie.xlsx
```

Ergebnis: A5:A7 = Nord, A10:A11 = Süd

```json
{
  "aenderungen": [
    {
      "blatt": "Umsatz",
      "zelle": "A5:A7",
      "alt": "verbunden",
      "neu": "Nord",
      "regel": "verbunden"
    },
    {
      "blatt": "Umsatz",
      "zelle": "A10:A11",
      "alt": "verbunden",
      "neu": "Süd",
      "regel": "verbunden"
    }
  ]
}
```

<a id="beispiel-4"></a>
## 4. Leere Zeilen und Zwischensummen aus der Excel-Liste löschen.

```bash
excel-skills bereinigen fixtures/umsatz_bericht.xlsx --ohne kopfzeile,verbunden,zahlen,datum,trimmen --out kopie.xlsx
```

Ergebnis: Zeilen 8, 9, 12, 13 entfernt, Formeln angepasst

```json
{
  "geloeschte_zeilen": [
    {
      "zeile": 8,
      "regel": "summen",
      "werte": [
        "Zwischensumme Nord",
        "=SUM(D5:D7)"
      ]
    },
    {
      "zeile": 9,
      "regel": "leerzeilen",
      "werte": []
    },
    {
      "zeile": 12,
      "regel": "summen",
      "werte": [
        "Zwischensumme Süd",
        "=SUM(D10:D11)"
      ]
    },
    {
      "zeile": 13,
      "regel": "summen",
      "werte": [
        "Gesamt",
        "=D8+D12"
      ]
    }
  ],
  "formeln_angepasst": 3
}
```

<a id="beispiel-5"></a>
## 5. Text in Zahl umwandeln: deutsche Zahlen 1.234,56 in Excel.

```bash
excel-skills bereinigen fixtures/umsatz_bericht.xlsx --ohne kopfzeile,verbunden,leerzeilen,summen,datum,trimmen --out kopie.xlsx
```

Ergebnis: 4 Zellen zu Zahl, `1.234` in E7 bleibt Text

```json
{
  "aenderungen": [
    {
      "blatt": "Umsatz",
      "zelle": "E5",
      "alt": "1.234,56",
      "neu": 1234.56,
      "regel": "deutsche-zahl"
    },
    {
      "blatt": "Umsatz",
      "zelle": "E6",
      "alt": "987,10",
      "neu": 987.1,
      "regel": "deutsche-zahl"
    },
    {
      "blatt": "Umsatz",
      "zelle": "E10",
      "alt": "12.500,00",
      "neu": 12500.0,
      "regel": "deutsche-zahl"
    },
    {
      "blatt": "Umsatz",
      "zelle": "E11",
      "alt": "-250,75",
      "neu": -250.75,
      "regel": "deutsche-zahl"
    }
  ],
  "hinweise": [
    {
      "zelle": "E7",
      "regel": "deutsche-zahl",
      "wert": "1.234",
      "hinweis": "ohne Komma mehrdeutig, bleibt Text"
    }
  ]
}
```

<a id="beispiel-6"></a>
## 6. Datum 31.12.2026 als Text in ein Excel-Datum umwandeln.

```bash
excel-skills bereinigen fixtures/umsatz_bericht.xlsx --ohne kopfzeile,verbunden,leerzeilen,summen,zahlen,trimmen --out kopie.xlsx
```

Ergebnis: 4 Zellen zu Datum, `01/31/2026` bleibt Text

```json
{
  "aenderungen": [
    {
      "blatt": "Umsatz",
      "zelle": "C5",
      "alt": "31.01.2026",
      "neu": "2026-01-31",
      "regel": "datum"
    },
    {
      "blatt": "Umsatz",
      "zelle": "C6",
      "alt": "28.02.2026",
      "neu": "2026-02-28",
      "regel": "datum"
    },
    {
      "blatt": "Umsatz",
      "zelle": "C7",
      "alt": "31.03.2026",
      "neu": "2026-03-31",
      "regel": "datum"
    },
    {
      "blatt": "Umsatz",
      "zelle": "C11",
      "alt": "30.04.26",
      "neu": "2026-04-30",
      "regel": "datum"
    }
  ],
  "hinweise": [
    {
      "zelle": "C10",
      "regel": "datum",
      "wert": "01/31/2026",
      "hinweis": "Datum mit Schrägstrich, Reihenfolge Tag und Monat unklar, bleibt Text"
    }
  ]
}
```

<a id="beispiel-7"></a>
## 7. Alle #BEZUG!-Fehler in der Arbeitsmappe finden.

```bash
excel-skills formeln fixtures/formeln_kaputt.xlsx --out report.md
```

Ergebnis: `#BEZUG!` in D5, F5

```json
[
  {
    "zelle": "D5",
    "fehler": "#REF!",
    "deutsch": "#BEZUG!",
    "formel": "=B5*#REF!",
    "hinweis": "Bezug zeigt auf eine gelöschte Zeile, Spalte oder ein gelöschtes Blatt"
  },
  {
    "zelle": "F5",
    "fehler": "#REF!",
    "deutsch": "#BEZUG!",
    "formel": "=D5/B5",
    "hinweis": "Bezug zeigt auf eine gelöschte Zeile, Spalte oder ein gelöschtes Blatt"
  }
]
```

<a id="beispiel-8"></a>
## 8. Warum zeigt mein SVERWEIS #NV, und in welchen Zellen?

```bash
excel-skills formeln fixtures/formeln_kaputt.xlsx --out report.md
```

Ergebnis: E5: SVERWEIS findet den Suchwert nicht

```json
[
  {
    "zelle": "E5",
    "fehler": "#N/A",
    "deutsch": "#NV",
    "formel": "=VLOOKUP(A5,Lieferanten!A:B,2,FALSE)",
    "hinweis": "SVERWEIS findet den Suchwert nicht im Suchbereich"
  }
]
```

<a id="beispiel-9"></a>
## 9. Harte Zahlen in einer Formelspalte finden.

```bash
excel-skills formeln fixtures/formeln_kaputt.xlsx --out report.md
```

Ergebnis: D4 = 99 in Spalte Summe

```json
{
  "formelspalten": [
    {
      "spalte": "D",
      "name": "Summe",
      "formeln": 4,
      "gefuellt": 5
    },
    {
      "spalte": "E",
      "name": "Lieferant",
      "formeln": 6,
      "gefuellt": 6
    },
    {
      "spalte": "F",
      "name": "Anteil",
      "formeln": 6,
      "gefuellt": 6
    }
  ],
  "harte_werte": [
    {
      "zelle": "D4",
      "spalte": "Summe",
      "wert": 99,
      "hinweis": "harter Wert in einer Formelspalte"
    }
  ]
}
```

<a id="beispiel-10"></a>
## 10. Externe Verknüpfungen zu anderen Excel-Dateien auflisten.

```bash
excel-skills formeln fixtures/formeln_kaputt.xlsx --out report.md
```

Ergebnis: Lieferanten!C2

```json
[
  {
    "blatt": "Lieferanten",
    "zelle": "C2",
    "formel": "='[1]Preisliste'!B2"
  }
]
```

<a id="beispiel-11"></a>
## 11. Formeln finden, die nicht bis zur letzten Zeile mitlaufen.

```bash
excel-skills formeln fixtures/formeln_kaputt.xlsx --out report.md
```

Ergebnis: D7 in Spalte Summe

```json
[
  {
    "zelle": "D7",
    "spalte": "Summe",
    "hinweis": "Zeile hat Daten, aber die Formel fehlt"
  }
]
```

<a id="beispiel-12"></a>
## 12. Zwei Excel-Tabellen vergleichen: alte und neue Preisliste, Schlüssel Artikelnummer.

```bash
excel-skills vergleich fixtures/preise_alt.xlsx fixtures/preise_neu.xlsx --schluessel Artikelnummer --out report.md
```

Ergebnis: 1 neu, 1 gelöscht, 2 geändert


Alt: `preise_alt.xlsx`, Blatt Preise. Neu: `preise_neu.xlsx`, Blatt Preise.

Vergleich über Schlüsselspalte **Artikelnummer**.

## Neu (1)

| Artikelnummer | Zeile |
|---|---|
| 4715 | 5 |

## Gelöscht (1)

| Artikelnummer | Zeile |
|---|---|
| 4713 | 4 |

## Geändert (2)

| Artikelnummer | Spalte | alt | neu |
|---|---|---|---|
| 4711 | Preis | 0.12 | 0.14 |
| 4712 | Bezeichnung | Mutter M6 | Mutter M6 verzinkt |

<a id="beispiel-13"></a>
## 13. Welche Kunden sind in der neuen Liste dazugekommen oder weggefallen?

```bash
excel-skills vergleich fixtures/kunden_liste.xlsx fixtures/kunden_liste_neu.xlsx --schluessel Kundennummer --out report.md
```

Ergebnis: neu K-1005, weg K-1003

```json
{
  "neu": [
    {
      "schluessel": "K-1005",
      "zeile": 5,
      "werte": {
        "Kundennummer": "K-1005",
        "Name": "Muster Software GmbH",
        "Ort": "München",
        "PLZ": "80333",
        "Umsatz 2026": 18400
      }
    }
  ],
  "geloescht": [
    {
      "schluessel": "K-1003",
      "zeile": 4,
      "werte": {
        "Kundennummer": "K-1003",
        "Name": "Muster Handel KG",
        "Ort": "Kiel",
        "PLZ": "24103",
        "Umsatz 2026": 9100
      }
    }
  ],
  "geaendert": [
    {
      "schluessel": "K-1002",
      "spalte": "Ort",
      "alt": "Hamburg",
      "neu": "Bremen",
      "zeile_alt": 3,
      "zeile_neu": 3
    },
    {
      "schluessel": "K-1002",
      "spalte": "PLZ",
      "alt": "20095",
      "neu": "28195",
      "zeile_alt": 3,
      "zeile_neu": 3
    }
  ]
}
```

<a id="beispiel-14"></a>
## 14. Mehrere Excel-Dateien zusammenführen, eine je Monat.

```bash
excel-skills stapeln fixtures/monate --out kopie.xlsx
```

Ergebnis: 3 Dateien, 5 Zeilen, 1 Summenzeile übersprungen

```json
{
  "spalten": [
    "Datum",
    "Kunde",
    "Betrag",
    "quelle_datei"
  ],
  "quellen": [
    {
      "datei": "2026-01.xlsx",
      "blatt": "Buchungen",
      "kopfzeile": 1,
      "zeilen": 2
    },
    {
      "datei": "2026-02.xlsx",
      "blatt": "Buchungen",
      "kopfzeile": 1,
      "zeilen": 1
    },
    {
      "datei": "2026-03.xlsx",
      "blatt": "Buchungen",
      "kopfzeile": 1,
      "zeilen": 2
    }
  ],
  "uebersprungen": [
    {
      "datei": "2026-01.xlsx",
      "zeile": 4,
      "regel": "summen",
      "werte": [
        "Summe",
        "=SUM(C2:C3)"
      ]
    }
  ]
}
```

<a id="beispiel-15"></a>
## 15. Prüfen, ob alle Dateien im Ordner dieselben Spalten haben.

```bash
excel-skills stapeln fixtures/monate_abweichend --out kopie.xlsx
```

Ergebnis: Abbruch mit Exit-Code 3: 2026-02.xlsx fehlt Betrag, hat Netto

```json
{
  "exit_code": 3,
  "fehler": "Spalten passen nicht zu 2026-01.xlsx. Nichts geschrieben.",
  "details": [
    {
      "datei": "2026-02.xlsx",
      "fehlt": [
        "Betrag"
      ],
      "zusaetzlich": [
        "Netto"
      ],
      "doppelt": []
    }
  ]
}
```

<a id="beispiel-16"></a>
## 16. Excel in CSV umwandeln für den DATEV-Import.

```bash
excel-skills csv fixtures/buchungen.xlsx --format datev --out kopie.csv
```

Ergebnis: Semikolon, Komma, `TT.MM.JJJJ`, Windows-1252, Text in Anführungszeichen

```csv
"Belegdatum";"Konto";"Gegenkonto";"Betrag netto";"Betrag brutto";"Buchungstext"
15.01.2026;8400;10001;1000;1190;"Rechnung 2026-001 Müller; Wartung"
31.01.2026;8400;10002;2499,99;2974,99;"Projekt ""Fabric"" Phase 1"
03.02.2026;4930;70001;89,9;;"Bürobedarf"
```

<a id="beispiel-17"></a>
## 17. CSV mit Semikolon und Komma für deutsches Excel speichern.

```bash
excel-skills csv fixtures/buchungen.xlsx --format excel-de --out kopie.csv
```

Ergebnis: Semikolon, Komma, UTF-8 mit BOM, Umlaute bleiben

```csv
Belegdatum;Konto;Gegenkonto;Betrag netto;Betrag brutto;Buchungstext
15.01.2026;8400;10001;1000;1190;"Rechnung 2026-001 Müller; Wartung"
31.01.2026;8400;10002;2499,99;2974,99;"Projekt ""Fabric"" Phase 1"
03.02.2026;4930;70001;89,9;;Bürobedarf
```

<a id="beispiel-18"></a>
## 18. Excel als UTF-8-CSV für Python oder eine Datenbank ausgeben.

```bash
excel-skills csv fixtures/buchungen.xlsx --format utf8 --out kopie.csv
```

Ergebnis: Komma, Punkt, `JJJJ-MM-TT`, UTF-8

```csv
Belegdatum,Konto,Gegenkonto,Betrag netto,Betrag brutto,Buchungstext
2026-01-15,8400,10001,1000,1190,Rechnung 2026-001 Müller; Wartung
2026-01-31,8400,10002,2499.99,2974.99,"Projekt ""Fabric"" Phase 1"
2026-02-03,4930,70001,89.9,,Bürobedarf
```

Hinweise aus `kopie.aenderungen.json`:

```json
[
  {
    "zelle": null,
    "hinweis": "3 Formeln als gespeicherte Werte ausgegeben, CSV kennt keine Formeln"
  },
  {
    "zelle": "E4",
    "hinweis": "Formel ohne gespeichertes Ergebnis, Feld bleibt leer. Datei einmal in Excel öffnen und speichern."
  }
]
```

<a id="beispiel-19"></a>
## 19. Diese Excel vor dem Import nach Power BI oder Microsoft Fabric prüfen.

```bash
excel-skills profil fixtures/umsatz_bericht.xlsx --out report.md
```

Ergebnis: Urteil `bericht`, erst bereinigen

```json
{
  "urteil": "bericht",
  "gruende": [
    "2 verbundene Bereiche im Datenbereich",
    "Summenzeile vor der letzten Datenzeile: 8, 12",
    "Kopfzeile erst in Zeile 4"
  ],
  "verbundene_im_datenbereich": [
    "A10:A11",
    "A5:A7"
  ],
  "summenzeilen": [
    8,
    12,
    13
  ]
}
```

<a id="beispiel-20"></a>
## 20. Eine Makro-Datei .xlsm prüfen, ohne Makros auszuführen.

```bash
excel-skills profil fixtures/makro.xlsm --out report.md
```

Ergebnis: Urteil `tabelle`, Makros nicht ausgeführt

#### Blatt Eingabe

- Urteil: **tabelle**
- Kopfzeile: 1 (3 von 3 Spalten Text, Zeile darunter hat Zahl, Datum oder Formel)
- Datenbereich: A2:C3
- Verbundene Zellen: 0, davon im Datenbereich 0
- Summenzeilen: keine
- Gründe: Kopfzeile in Zeile 1, keine verbundenen Zellen, keine Summenzeile mitten drin

| Spalte | Name | Typ | leer |
|---|---|---|---|
| A | Artikel | text | 0 % |
| B | Menge | text | 0 % |
| C | Preis | zahl_als_text | 0 % |
