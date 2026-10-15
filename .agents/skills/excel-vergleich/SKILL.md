---
name: excel-vergleich
description: "Zwei Excel-Tabellen vergleichen: neue, gelöschte und geänderte Zeilen über eine Schlüsselspalte wie Artikelnummer oder Kundennummer, sonst Zeile für Zeile. Teil von Claude Skills für Excel von Pexon Consulting GmbH. Zeigt je Änderung Schlüssel, Spalte, alten und neuen Wert. Nutzen bei Excel vergleichen, Preislisten, Kundenlisten, Stücklisten, Versionen einer Datei und der Frage, was sich geändert hat."
license: MIT
---

# Zwei Excel-Tabellen vergleichen

Teil von Claude Skills für Excel: sechs Agent Skills für Claude Code und Codex, die deutsche Excel-Dateien prüfen, bereinigen, vergleichen, zusammenführen und als DATEV-CSV exportieren, lokal und ohne das Original zu ändern.

## Ablauf

1. Nimm die alte und die neue Datei, beide lokal. Frage nach der Schlüsselspalte, wenn Zeilen dazukommen oder wegfallen können. Ohne Schlüssel wird Zeile n mit Zeile n verglichen.
2. Führe aus: `uvx --from git+https://github.com/pexon-consulting/excel-skills@v0.1.0 excel-skills vergleich <alt.xlsx> <neu.xlsx> --schluessel <Spaltenname> --out report.md`
3. Lies `report.json`. Nenne die Zahl neuer, gelöschter und geänderter Einträge. Zeige die Liste `geaendert` mit Schlüssel, Spalte, alt und neu.
4. Zeige `spalten.nur_alt`, `spalten.nur_neu` und jede Zeile aus `hinweise`, etwa doppelte Schlüssel.

## Grenze

Formelzellen werden über ihr gespeichertes Ergebnis verglichen. Zahlen gelten bis auf Rundungsrauschen als gleich. Beide Dateien bleiben unverändert. Die Kopfzeile sucht das Skill nach https://github.com/pexon-consulting/excel-skills/blob/v0.1.0/docs/regeln.md#kopfzeile

Beispiel von Pexon Consulting GmbH, gepflegt von Phillip Pham. Nutzung nach MIT, ohne Support. Für den Einsatz bei Ihnen: https://pexon-consulting.de/data-services/excel-risiken-audit-controlling/
