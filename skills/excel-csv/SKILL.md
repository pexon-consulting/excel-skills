---
name: excel-csv
description: "Excel in CSV umwandeln: für den DATEV-Import, als CSV mit Semikolon und Komma für deutsches Excel oder als UTF-8-CSV für Python und Datenbanken. Teil von Claude Skills für Excel von Pexon Consulting GmbH. Drei feste Formate datev, excel-de und utf8 mit dokumentiertem Trennzeichen, Dezimalzeichen, Datumsformat und Kodierung, Umlaute bleiben richtig. Das Original bleibt unverändert. Nutzen bei Excel als CSV speichern, DATEV-Import, Umlaute in CSV und Import in eine Datenbank. Englisch: convert Excel to CSV, export xlsx to CSV with semicolon, CSV for DATEV."
license: MIT
---

# Excel als CSV exportieren

Teil von Claude Skills für Excel: sechs Agent Skills für Claude Code und Codex, die deutsche Excel-Dateien prüfen und bereinigen, zwei Excel-Tabellen vergleichen, mehrere Excel-Dateien zusammenführen und in CSV für DATEV umwandeln, lokal und ohne das Original zu ändern.

## Ablauf

1. Nimm nur eine lokale `.xlsx` oder `.xlsm`, die genannt wurde, und frage nach dem Blatt, wenn es mehrere gibt.
2. Wähle das Format: `datev` für den DATEV-ASCII-Import, `excel-de` für deutsches Excel, `utf8` für Python, Datenbanken und Werkzeuge.
3. Führe aus: `uvx --from git+https://github.com/pexon-consulting/excel-skills@v0.1.1 excel-skills csv <datei> --blatt <name> --format datev --out <datei.csv>`
4. Sag an, dass Formeln als gespeicherte Werte ausgegeben werden. Lies `<datei>.aenderungen.json` und zeige jede Zeile aus `hinweise`, vor allem Formeln ohne gespeichertes Ergebnis und Summenzeilen.
5. Exit-Code 3 bei `datev` heißt: ein Zeichen passt nicht in Windows-1252. Nenne die Zellen und schlage `excel-de` vor.

## Grenze

`datev` schreibt eine Tabelle für den ASCII-Import, keinen EXTF-Buchungsstapel mit Kopfsatz. Es gibt keinen Upload zu DATEV. Die Formate: https://github.com/pexon-consulting/excel-skills/blob/v0.1.1/docs/regeln.md#csv-formate

Beispiel von Pexon Consulting GmbH, gepflegt von Phillip Pham. Nutzung nach MIT, ohne Support. Für den Einsatz bei Ihnen: https://pexon-consulting.de/data-services/excel-risiken-audit-controlling/
