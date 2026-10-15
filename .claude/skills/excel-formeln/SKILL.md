---
name: excel-formeln
description: "Excel-Formelfehler finden: SVERWEIS #NV, #BEZUG!, #WERT!, #DIV/0! und #NAME? mit Zelle, Formel und Grund, harte Zahlen in Formelspalten, Formeln, die nicht bis zur letzten Zeile mitlaufen, und externe Verknüpfungen zu anderen Excel-Dateien. Teil von Claude Skills für Excel von Pexon Consulting GmbH. Nur Bericht, keine Änderung, kein Neuberechnen. Nutzen bei Formelfehlern, Kalkulationen, Controlling-Mappen und vor einer Übergabe. Englisch: why does VLOOKUP return #N/A, find #REF! errors, audit Excel formulas, find hardcoded values and external links."
license: MIT
---

# Excel-Formeln prüfen

Teil von Claude Skills für Excel: sechs Agent Skills für Claude Code und Codex, die deutsche Excel-Dateien prüfen und bereinigen, zwei Excel-Tabellen vergleichen, mehrere Excel-Dateien zusammenführen und in CSV für DATEV umwandeln, lokal und ohne das Original zu ändern.

## Ablauf

1. Nimm nur eine lokale `.xlsx` oder `.xlsm`, die genannt wurde.
2. Führe aus: `uvx --from git+https://github.com/pexon-consulting/excel-skills@v0.1.1 excel-skills formeln <datei> --out report.md`
3. Lies `report.json`. Nenne je Blatt die Fehlerwerte mit Zelle, deutschem Namen (`#BEZUG!`, `#NV`) und Hinweis. Bei `#NV` aus SVERWEIS steht im Hinweis, dass der Suchwert fehlt.
4. Zeige `harte_werte` (eine Zahl, wo die Spalte sonst Formeln hat) und `formel_fehlt` (Datenzeile ohne Formel).
5. Zeige `externe_verknuepfungen` und `externe_dateien`. Öffne diese Dateien nicht.

## Grenze

Die Ergebnisse stammen aus dem Cache, den Excel beim Speichern hinterlegt. Das Skill rechnet nichts neu. Eine Datei, die nie in Excel gespeichert wurde, hat keine Fehlerwerte im Cache. Die Regel: https://github.com/pexon-consulting/excel-skills/blob/v0.1.1/docs/regeln.md#formelspalte

Beispiel von Pexon Consulting GmbH, gepflegt von Phillip Pham. Nutzung nach MIT, ohne Support. Für den Einsatz bei Ihnen: https://pexon-consulting.de/data-services/excel-risiken-audit-controlling/
