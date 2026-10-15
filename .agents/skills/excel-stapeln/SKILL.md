---
name: excel-stapeln
description: "Mehrere Excel-Dateien zusammenführen: alle Dateien eines Ordners mit gleichen Spalten untereinander in eine Tabelle, mit Spalte quelle_datei, etwa zwölf Monatsdateien zu einer Jahresliste. Teil von Claude Skills für Excel von Pexon Consulting GmbH. Bricht ab und listet Abweichungen, wenn Spalten nicht passen. Zwischensummen werden übersprungen und protokolliert, die Originale bleiben unverändert. Nutzen bei Excel-Tabellen zusammenführen, Monatsdateien, Filialdateien und Exporten aus mehreren Systemen. Englisch: combine multiple Excel files into one, merge Excel files, append spreadsheets with the same columns."
license: MIT
---

# Excel-Dateien zusammenführen

Teil von Claude Skills für Excel: sechs Agent Skills für Claude Code und Codex, die deutsche Excel-Dateien prüfen und bereinigen, zwei Excel-Tabellen vergleichen, mehrere Excel-Dateien zusammenführen und in CSV für DATEV umwandeln, lokal und ohne das Original zu ändern.

## Ablauf

1. Nimm einen lokalen Ordner, den die Nutzerin oder der Nutzer genannt hat. Gelesen werden alle `.xlsx` und `.xlsm` darin, sortiert nach Namen, ohne Unterordner.
2. Führe aus: `uvx --from git+https://github.com/pexon-consulting/excel-skills@v0.1.1 excel-skills stapeln <ordner> --out <gesamt.xlsx>`
3. Lies `<gesamt>.aenderungen.json`. Nenne die Dateien mit Zeilenzahl und die übersprungenen Summenzeilen.
4. Exit-Code 3 heißt: nichts geschrieben. In `details` steht je Datei, welche Spalte fehlt oder zusätzlich ist. Zeige das und frage, wie es weitergeht.
5. Verweist eine Formel auf ein anderes Blatt oder stehen die Spalten in anderer Reihenfolge und enthalten Formeln, bricht der Befehl ab. Mit `--werte` übernimmt er die gespeicherten Ergebnisse. Sag das vorher an.

## Grenze

Spalten werden über den Namen zugeordnet, die Reihenfolge darf abweichen. Formeln in derselben Zeile wandern mit, solange die Spaltenreihenfolge gleich ist. Die Eingaben bleiben unverändert. Die Regel für Summenzeilen: https://github.com/pexon-consulting/excel-skills/blob/v0.1.1/docs/regeln.md#summenzeile

Beispiel von Pexon Consulting GmbH, gepflegt von Phillip Pham. Nutzung nach MIT, ohne Support. Für den Einsatz bei Ihnen: https://pexon-consulting.de/data-services/excel-risiken-audit-controlling/
