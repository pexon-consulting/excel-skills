# Änderungen an Claude Skills für Excel

## 0.1.1 - 2026-10-15

Korrekturen:

- `--out` wird auch dann abgelehnt, wenn es die Eingabedatei in anderer Groß- und Kleinschreibung oder als Hardlink ist. Auf macOS konnte das Original sonst überschrieben werden.
- `bereinigen` verschiebt beim Löschen von Zeilen jetzt auch intelligente Tabellen, Autofilter, Hyperlinks, benannte Bereiche einzelner Blätter, den Druckbereich und den Bereich von Matrixformeln. Fällt die Kopfzeile einer intelligenten Tabelle weg oder ist eine Datentabelle der Was-wäre-wenn-Analyse betroffen, bricht der Befehl mit Exit-Code 3 ab.
- `stapeln` bricht mit Exit-Code 3 ab, wenn die Spalten in anderer Reihenfolge stehen und Formeln enthalten. Vorher konnten die verschobenen Formeln auf falsche Spalten zeigen. Mit `--werte` kommen die gespeicherten Ergebnisse.
- Summenzeilen werden nur am ganzen Wort erkannt. „Gesamtschule Nord“ oder „Totalbau GmbH“ bleiben stehen, „Gesamtsumme“ zählt neu dazu.
- `--blatt` mit einem Diagrammblatt gibt einen klaren Fehler statt eines Absturzes.

Sonst:

- Tests prüfen für `profil`, `formeln`, `csv`, `vergleich` und `stapeln`, dass das Original byte-gleich bleibt.
- README, Skill-Beschreibungen und `docs/regeln.md` überarbeitet, mit Vergleichen zum xlsx-Skill von Anthropic, Claude for Excel, Copilot, Power Query und Excel MCP Server und einem englischen Teil.

## 0.1.0 - 2026-10-15

Erste Version mit sechs Skills: `excel-profil`, `excel-bereinigen`, `excel-formeln`, `excel-vergleich`, `excel-stapeln` und `excel-csv`.
