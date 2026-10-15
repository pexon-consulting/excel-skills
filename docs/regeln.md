# Regeln von Claude Skills für Excel

Jede Regel hat einen festen Anker. Die Werte stehen identisch in `src/excel_skills/regeln.py` und kommen maschinenlesbar aus `excel-skills schema`. Kein Sprachmodell trifft diese Entscheidungen, das Programm tut es, jedes Mal gleich.

<a id="kopfzeile"></a>
## Kopfzeile finden

Die Kopfzeile ist die erste Zeile in den ersten 50 Zeilen, in der mindestens 60 Prozent der genutzten Spalten Text enthalten und deren Folgezeile mindestens eine Zahl, ein Datum, einen Wahrheitswert oder eine Formel hat. Passt keine Folgezeile, gilt die erste Zeile mit genug Text. Ein Titel wie „Umsatzbericht 2026“ in Zeile 1 füllt nur eine Spalte und ist deshalb keine Kopfzeile.

| Prüfung | Wert |
|---|---|
| Textanteil der Zeile | mindestens 60 Prozent der genutzten Spalten |
| Suchtiefe | erste 50 Zeilen |
| Folgezeile | mindestens eine Zelle Zahl, Datum, Wahrheitswert oder Formel |
| Ersatz | erste Zeile mit genug Text |
| Zahl als Text, `1.234,56` | zählt als Zahl, nicht als Text |

<a id="tabelle-oder-bericht"></a>
## Tabelle, Bericht oder Formular

| Urteil | Bedingung |
|---|---|
| `formular` | weniger als 3 Spalten mit Daten |
| `bericht` | keine Kopfzeile gefunden, oder verbundene Zellen unterhalb der Kopfzeile, oder eine Summenzeile vor der letzten Datenzeile |
| `tabelle` | sonst |

Eine `tabelle` kann direkt nach Power BI, Microsoft Fabric oder in eine Datenbank. Ein `bericht` braucht vorher `excel-skills bereinigen`.

<a id="deutsche-zahl"></a>
## Deutsche Zahl, Text in Zahl umwandeln

| Text in der Zelle | Ergebnis | Grund |
|---|---|---|
| `1.234,56` | 1234,56 als Zahl | Punkt als Tausender, Komma als Dezimalzeichen |
| `-1.234,56` | -1234,56 als Zahl | Minus vorne |
| `1 234,56` | 1234,56 als Zahl | Leerzeichen, geschütztes oder schmales Leerzeichen als Tausender |
| `12,5` | 12,5 als Zahl | Komma als Dezimalzeichen |
| `1.234` | bleibt Text, Hinweis | kann 1234 oder 1,234 sein |
| `01067`, `4711` | bleibt Text | kann Postleitzahl, Artikel- oder Kundennummer sein |

Nur Zellen unterhalb der Kopfzeile werden umgewandelt. Formeln werden nie umgewandelt.

<a id="datum"></a>
## Datum als Text in Excel-Datum umwandeln

| Text in der Zelle | Ergebnis |
|---|---|
| `31.12.2026` | Datum 31.12.2026, Format `TT.MM.JJJJ` |
| `31.12.26` | Datum 31.12.2026. Zweistellige Jahre 00-29 sind 20xx, 30-99 sind 19xx, wie in Excel |
| `31.02.2026` | bleibt Text, kein gültiges Datum |
| `12/31/2026` | bleibt Text, Hinweis: Reihenfolge Tag und Monat unklar |

<a id="summenzeile"></a>
## Summenzeile und Zwischensumme

Eine Zeile ist Summenzeile, wenn ihre erste Textzelle mit `Summe`, `Gesamt`, `Total` oder `Zwischensumme` beginnt, ohne Rücksicht auf Groß- und Kleinschreibung. Formeltexte zählen nicht als Textzelle. `bereinigen` entfernt Summenzeilen, `stapeln` überspringt sie, `csv` gibt sie mit Hinweis aus.

<a id="formelspalte"></a>
## Formelspalte, harter Wert, Formel fehlt

Eine Spalte ist Formelspalte, wenn mindestens 80 Prozent ihrer gefüllten Datenzellen Formeln sind und es mindestens 2 Formeln gibt. Summenzeilen zählen nicht mit.

| Befund | Bedingung |
|---|---|
| harter Wert | eine Zahl oder ein Text in einer Formelspalte |
| Formel fehlt | leere Zelle in einer Formelspalte, obwohl die Zeile Daten hat |
| Fehlerwert | gespeichertes Ergebnis ist ein Excel-Fehler, oder die Formel enthält `#REF!` |
| externe Verknüpfung | Formel mit `[1]`, `[2]` und so weiter, das ist ein Bezug auf eine andere Datei |

Fehlerwerte stehen in der Datei englisch. Der Bericht nennt beide Namen:

| Datei | Deutsches Excel | Häufiger Grund |
|---|---|---|
| `#REF!` | `#BEZUG!` | Bezug auf gelöschte Zeile, Spalte oder Blatt |
| `#N/A` | `#NV` | SVERWEIS, XVERWEIS oder VERGLEICH findet den Suchwert nicht |
| `#VALUE!` | `#WERT!` | Text statt Zahl |
| `#DIV/0!` | `#DIV/0!` | Teilung durch null oder leere Zelle |
| `#NAME?` | `#NAME?` | unbekannte Funktion oder unbekannter Name |
| `#NUM!` | `#ZAHL!` | ungültige Zahl |
| `#NULL!` | `#NULL!` | Bereiche ohne Schnittmenge |

<a id="zeilen-loeschen"></a>
## Zeilen löschen, ohne Formeln zu brechen

Wenn `bereinigen` Zeilen löscht, schreibt es jeden Zeilenbezug im ganzen Arbeitsbuch um, auch auf anderen Blättern und in benannten Bereichen. Aus `=D5*2` wird `=D2*2`, aus `=SUM(D5:D7)` wird `=SUM(D2:D4)`. Zeigt eine Formel, die bleibt, auf eine Zeile, die entfällt, bricht der Befehl mit Exit-Code 3 ab und schreibt nichts.

<a id="csv-formate"></a>
## CSV-Formate

| Format | Trennzeichen | Dezimalzeichen | Datum | Kodierung | Zeilenende | Text in Anführungszeichen |
|---|---|---|---|---|---|---|
| `datev` | `;` | `,` | `TT.MM.JJJJ` | Windows-1252 (ANSI) | CRLF | immer |
| `excel-de` | `;` | `,` | `TT.MM.JJJJ` | UTF-8 mit BOM | CRLF | nur wenn nötig |
| `utf8` | `,` | `.` | `JJJJ-MM-TT` | UTF-8 ohne BOM | LF | nur wenn nötig |

Zahlen haben kein Tausendertrennzeichen. Wahrheitswerte sind `WAHR` und `FALSCH`, bei `utf8` `TRUE` und `FALSE`. Formeln werden als gespeicherte Ergebnisse ausgegeben. Passt ein Zeichen bei `datev` nicht in Windows-1252, bricht der Befehl ab. `datev` schreibt eine Tabelle für den ASCII-Import, keinen EXTF-Buchungsstapel mit Kopfsatz.

<a id="grenzen"></a>
## Grenzen

| Grenze | Wert |
|---|---|
| Dateitypen | `.xlsx`, `.xlsm`. `.xls` bitte als `.xlsx` speichern |
| Dateigröße | 50 MB |
| Entpackt | 500 MB |
| Zellen je Blatt | 1.000.000 |
| Makros | werden nie ausgeführt, das VBA-Projekt bleibt in `.xlsm`-Kopien erhalten |
| Netz | kein Zugriff, kein Upload |
