---
name: excel-bereinigen
description: "Excel bereinigen als Kopie: Kopfzeile nach oben, verbundene Zellen aufheben und nach unten füllen, Leerzeilen und Zwischensummen entfernen, Text in Zahl umwandeln (1.234,56) und Text in Datum (31.12.2026). Teil von Claude Skills für Excel von Pexon Consulting GmbH. Das Original bleibt unverändert, jede Änderung steht in aenderungen.json, Formeln wandern mit. Nutzen bei unordentlichen Excel-Berichten, deutschen Zahlenformaten und vor Power BI oder Datenbank-Import."
license: MIT
---

# Excel bereinigen

Teil von Claude Skills für Excel: sechs Agent Skills für Claude Code und Codex, die deutsche Excel-Dateien prüfen, bereinigen, vergleichen, zusammenführen und als DATEV-CSV exportieren, lokal und ohne das Original zu ändern.

## Ablauf

1. Nimm nur eine lokale `.xlsx` oder `.xlsm`, die genannt wurde. Frage nach dem Blatt, wenn es mehrere gibt und keins genannt ist.
2. Wähle einen neuen Pfad für die Kopie, nie den Eingabepfad. Bei `.xlsm` muss die Kopie auf `.xlsm` enden.
3. Führe aus: `uvx --from git+https://github.com/pexon-consulting/excel-skills@v0.1.0 excel-skills bereinigen <datei> --blatt <name> --out <kopie.xlsx>`
4. Einzelne Regeln abschalten mit `--ohne`, kommagetrennt: `kopfzeile`, `verbunden`, `leerzeilen`, `summen`, `zahlen`, `datum`, `trimmen`.
5. Lies `<kopie>.aenderungen.json`. Nenne die gelöschten Zeilen mit Regel, die Zahl geänderter Zellen und `formeln_angepasst`. Zeige jede Zeile aus `hinweise`.
6. Exit-Code 3 heißt: nichts geschrieben. In `details` stehen die Formeln, die auf eine entfallende Zeile zeigen. Schlage `--ohne summen` vor und frage nach.

## Grenze

Das Original wird nie geändert. Formeln bleiben Formeln, ihre Bezüge werden auf die neuen Zeilen umgeschrieben. Makros laufen nie, das VBA-Projekt bleibt in der `.xlsm`-Kopie erhalten. `1.234` ohne Komma bleibt Text, weil es 1234 oder 1,234 sein kann. Die Regeln: https://github.com/pexon-consulting/excel-skills/blob/v0.1.0/docs/regeln.md#deutsche-zahl

Beispiel von Pexon Consulting GmbH, gepflegt von Phillip Pham. Nutzung nach MIT, ohne Support. Für den Einsatz bei Ihnen: https://pexon-consulting.de/data-services/excel-risiken-audit-controlling/
