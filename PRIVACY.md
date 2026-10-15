# Datenschutz

Claude Skills für Excel lesen nur die lokalen Dateien und Ordner, die du angibst. Es gibt keinen Upload, keine Telemetrie und keinen Netzabruf. Einzige Ausnahme ist `uvx`, das beim ersten Aufruf dieses Paket und `openpyxl` herunterlädt.

Alle Regeln laufen auf diesem Rechner. Kopien, Berichte und `aenderungen.json` legt das Programm nur dorthin, wo `--out` hinzeigt. Das Original wird nie geändert. Makros in `.xlsm` werden nie ausgeführt.

Was dein Agent, also Claude, Codex oder ein anderes Modell, mit dem Inhalt der Berichte macht, regelt dessen Anbieter.
