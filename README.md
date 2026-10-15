# Claude Skills für Excel (und Codex)

Claude Skills für Excel: sechs Agent Skills für Claude Code und Codex, die deutsche Excel-Dateien prüfen und bereinigen, zwei Excel-Tabellen vergleichen, mehrere Excel-Dateien zusammenführen und in CSV für DATEV umwandeln, lokal und ohne das Original zu ändern.

Beispiel von Pexon Consulting GmbH, gepflegt von Phillip Pham, Lizenz MIT. Agent Skills plus ein lokales Kommandozeilenprogramm `excel-skills`, gebaut für deutsche Excel-Dateien.

[![tests](https://github.com/pexon-consulting/excel-skills/actions/workflows/tests.yml/badge.svg)](https://github.com/pexon-consulting/excel-skills/actions/workflows/tests.yml) ![Lizenz MIT](https://img.shields.io/badge/Lizenz-MIT-blue) ![Python 3.11+](https://img.shields.io/badge/Python-3.11%2B-blue)

Stand: Release [v0.1.1](https://github.com/pexon-consulting/excel-skills/releases/tag/v0.1.1) vom 15.10.2026. Die Tests laufen bei jedem Push auf Python 3.11 und 3.13, die CI prüft die 20 Beispiele gegen die Testdateien in `fixtures/`.

```bash
npx skills add pexon-consulting/excel-skills
```

Danach schreibst du deinem Agenten einen normalen Satz, zum Beispiel „Vergleiche preise_alt.xlsx und preise_neu.xlsx über die Artikelnummer“. Das passende Skill ruft ein lokales Programm auf, das Programm entscheidet nach festen Regeln, der Agent erklärt das Ergebnis. Kein Modell rechnet eine Zahl aus, keine Datei verlässt den Rechner.

Gebaut für die Dateien, an denen allgemeine Excel-Werkzeuge scheitern: `1.234,56` als Text, `31.12.2026` als Text, Kopfzeile erst in Zeile 4, verbundene Zellen, Zwischensummen mitten in der Liste, `#BEZUG!` und SVERWEIS `#NV`, CSV mit Semikolon für DATEV.

*In English:* agent skills for German Excel files, for Claude Code and Codex. See [English](#english).

## Die sechs Claude Skills für Excel

| Skill | Wofür | Befehl |
|---|---|---|
| [excel-profil](skills/excel-profil/SKILL.md) | Excel-Kopfzeile finden, Tabelle oder Bericht, Spaltentypen | `excel-skills profil <datei>` |
| [excel-bereinigen](skills/excel-bereinigen/SKILL.md) | Kopie ohne verbundene Zellen, leere Zeilen und Zwischensummen, mit echten Zahlen und Daten | `excel-skills bereinigen <datei> --out <kopie.xlsx>` |
| [excel-formeln](skills/excel-formeln/SKILL.md) | `#BEZUG!`, SVERWEIS `#NV`, harte Werte, fehlende Formeln, externe Verknüpfungen | `excel-skills formeln <datei>` |
| [excel-vergleich](skills/excel-vergleich/SKILL.md) | zwei Excel-Tabellen vergleichen: neu, gelöscht, geändert | `excel-skills vergleich <alt> <neu> --schluessel <Spalte>` |
| [excel-stapeln](skills/excel-stapeln/SKILL.md) | mehrere Excel-Dateien mit gleichen Spalten zusammenführen | `excel-skills stapeln <ordner> --out <gesamt.xlsx>` |
| [excel-csv](skills/excel-csv/SKILL.md) | Excel in CSV umwandeln: DATEV, deutsches Excel oder UTF-8 | `excel-skills csv <datei> --format datev --out <datei.csv>` |

Im Skill steht vor jedem Befehl `uvx --from git+https://github.com/pexon-consulting/excel-skills@v0.1.1`. So läuft immer genau dieses Release, ohne Installation und ohne `pip`.

## 20 Beispiele: Claude Skills für Excel im Einsatz

Jede Zeile ist ein Satz an Claude oder Codex, der Befehl, den das Skill wählt, und das echte Ergebnis gegen die Testdateien in [`fixtures/`](fixtures/). Die volle Ausgabe je Beispiel steht in [docs/beispiele.md](docs/beispiele.md).

<!-- beispiele:start -->
| # | Das schreibst du dem Agenten | Befehl | Ergebnis |
|---|---|---|---|
| [1](docs/beispiele.md#beispiel-1) | Finde die Kopfzeile in dieser Excel, die erst in Zeile 4 anfängt. | `profil` | Kopfzeile Zeile 4, Datenbereich A5:F13 |
| [2](docs/beispiele.md#beispiel-2) | Ist dieses Excel-Blatt eine Tabelle oder ein Bericht? | `profil` | Urteil `tabelle`, PLZ bleibt Text |
| [3](docs/beispiele.md#beispiel-3) | Verbundene Zellen auflösen und den Wert nach unten füllen. | `bereinigen` | A5:A7 = Nord, A10:A11 = Süd |
| [4](docs/beispiele.md#beispiel-4) | Leere Zeilen und Zwischensummen aus der Excel-Liste löschen. | `bereinigen` | Zeilen 8, 9, 12, 13 entfernt, Formeln angepasst |
| [5](docs/beispiele.md#beispiel-5) | Text in Zahl umwandeln: deutsche Zahlen 1.234,56 in Excel. | `bereinigen` | 4 Zellen zu Zahl, `1.234` in E7 bleibt Text |
| [6](docs/beispiele.md#beispiel-6) | Datum 31.12.2026 als Text in ein Excel-Datum umwandeln. | `bereinigen` | 4 Zellen zu Datum, `01/31/2026` bleibt Text |
| [7](docs/beispiele.md#beispiel-7) | Alle #BEZUG!-Fehler in der Arbeitsmappe finden. | `formeln` | `#BEZUG!` in D5, F5 |
| [8](docs/beispiele.md#beispiel-8) | Warum zeigt mein SVERWEIS #NV, und in welchen Zellen? | `formeln` | E5: SVERWEIS findet den Suchwert nicht |
| [9](docs/beispiele.md#beispiel-9) | Harte Zahlen in einer Formelspalte finden. | `formeln` | D4 = 99 in Spalte Summe |
| [10](docs/beispiele.md#beispiel-10) | Externe Verknüpfungen zu anderen Excel-Dateien auflisten. | `formeln` | Lieferanten!C2 |
| [11](docs/beispiele.md#beispiel-11) | Formeln finden, die nicht bis zur letzten Zeile mitlaufen. | `formeln` | D7 in Spalte Summe |
| [12](docs/beispiele.md#beispiel-12) | Zwei Excel-Tabellen vergleichen: alte und neue Preisliste, Schlüssel Artikelnummer. | `vergleich` | 1 neu, 1 gelöscht, 2 geändert |
| [13](docs/beispiele.md#beispiel-13) | Welche Kunden sind in der neuen Liste dazugekommen oder weggefallen? | `vergleich` | neu K-1005, weg K-1003 |
| [14](docs/beispiele.md#beispiel-14) | Mehrere Excel-Dateien zusammenführen, eine je Monat. | `stapeln` | 3 Dateien, 5 Zeilen, 1 Summenzeile übersprungen |
| [15](docs/beispiele.md#beispiel-15) | Prüfen, ob alle Dateien im Ordner dieselben Spalten haben. | `stapeln` | Abbruch mit Exit-Code 3: 2026-02.xlsx fehlt Betrag, hat Netto |
| [16](docs/beispiele.md#beispiel-16) | Excel in CSV umwandeln für den DATEV-Import. | `csv` | Semikolon, Komma, `TT.MM.JJJJ`, Windows-1252, Text in Anführungszeichen |
| [17](docs/beispiele.md#beispiel-17) | CSV mit Semikolon und Komma für deutsches Excel speichern. | `csv` | Semikolon, Komma, UTF-8 mit BOM, Umlaute bleiben |
| [18](docs/beispiele.md#beispiel-18) | Excel als UTF-8-CSV für Python oder eine Datenbank ausgeben. | `csv` | Komma, Punkt, `JJJJ-MM-TT`, UTF-8 |
| [19](docs/beispiele.md#beispiel-19) | Diese Excel vor dem Import nach Power BI oder Microsoft Fabric prüfen. | `profil` | Urteil `bericht`, erst bereinigen |
| [20](docs/beispiele.md#beispiel-20) | Eine Makro-Datei .xlsm prüfen, ohne Makros auszuführen. | `profil` | Urteil `tabelle`, Makros nicht ausgeführt |
<!-- beispiele:ende -->

## Excel-Tabellen vergleichen: zwei Dateien über eine Schlüsselspalte

Mit `excel-vergleich` aus Claude Skills für Excel vergleicht Claude zwei Versionen einer Preisliste, Kundenliste oder Stückliste: Zeilen werden über eine Schlüsselspalte zugeordnet, der Bericht nennt neue, gelöschte und geänderte Zeilen mit altem und neuem Wert.

```bash
excel-skills vergleich preise_alt.xlsx preise_neu.xlsx --schluessel Artikelnummer --out report.md
```

| Artikelnummer | Spalte | alt | neu |
|---|---|---|---|
| 4711 | Preis | 0.12 | 0.14 |
| 4712 | Bezeichnung | Mutter M6 | Mutter M6 verzinkt |

Dazu kommen `4715` als neu und `4713` als gelöscht. Ohne `--schluessel` vergleicht das Skill Zeile 1 mit Zeile 1, Zeile 2 mit Zeile 2. Das passt nur, wenn keine Zeile dazukam oder wegfiel. Doppelte Schlüssel stehen als Hinweis im Bericht, verglichen wird der erste Treffer.

## Mehrere Excel-Dateien zusammenführen

Zwölf Monatsdateien, vier Filialen, drei Exporte: `excel-stapeln` aus Claude Skills für Excel legt alle `.xlsx` eines Ordners untereinander und schreibt in die Spalte `quelle_datei`, woher jede Zeile kommt.

```bash
excel-skills stapeln ./monate --out gesamt.xlsx
```

Die Spalten werden über den Namen zugeordnet, die Reihenfolge darf abweichen. Fehlt in einer Datei eine Spalte oder heißt sie anders, bricht der Befehl ab, schreibt nichts und nennt die Abweichung: `2026-02.xlsx fehlt Betrag, hat Netto`. Stehen die Spalten in anderer Reihenfolge und enthalten Formeln, bricht der Befehl ebenfalls ab, weil die Bezüge sonst falsch würden. Mit `--werte` kommen dann die gespeicherten Ergebnisse. Summenzeilen wie „Summe Januar“ werden übersprungen und in `gesamt.aenderungen.json` protokolliert, damit die Jahressumme nicht doppelt zählt.

## Excel-Kopfzeile finden, wenn sie nicht in Zeile 1 steht

Viele Exporte haben oben einen Titel, ein Datum und eine Leerzeile. `excel-profil` sucht die Kopfzeile nach einer festen Regel: die erste Zeile, in der mindestens 60 Prozent der Spalten Text haben und deren Folgezeile eine Zahl, ein Datum oder eine Formel enthält ([Regel](docs/regeln.md#kopfzeile)).

```text
Kopfzeile: 4 (6 von 6 Spalten Text, Zeile darunter hat Zahl, Datum oder Formel)
Urteil: bericht
Gründe: 2 verbundene Bereiche im Datenbereich; Summenzeile vor der letzten Datenzeile: 8, 12; Kopfzeile erst in Zeile 4
```

Das Urteil ist `tabelle`, `bericht` oder `formular` ([Regel](docs/regeln.md#tabelle-oder-bericht)). Eine `tabelle` kann direkt nach Power BI, Microsoft Fabric oder in eine Datenbank. Ein `bericht` braucht vorher `excel-bereinigen`.

## Excel bereinigen: leere Zeilen löschen, verbundene Zellen aufheben, Zwischensummen entfernen

`excel-bereinigen` schreibt eine Kopie. Es zieht die Kopfzeile nach oben, löst verbundene Zellen auf und füllt den Wert nach unten, löscht leere Zeilen und Summenzeilen, trimmt Leerzeichen. Jede Regel lässt sich mit `--ohne` abschalten.

```bash
excel-skills bereinigen umsatz_bericht.xlsx --out umsatz_sauber.xlsx
```

Formeln wandern mit. Wenn drei Zeilen über der Kopfzeile entfallen, wird aus `=D5*2` die Formel `=D2*2`, und ein Verweis von einem anderen Blatt `=Umsatz!D5` wird zu `=Umsatz!D2`. Dasselbe gilt für benannte Bereiche, den Druckbereich, intelligente Tabellen, Autofilter, Matrixformeln und Hyperlinks ([Regel](docs/regeln.md#zeilen-loeschen)). Zeigt eine Formel, die bleibt, auf eine Zeile, die entfallen würde, bricht der Befehl ab und schreibt nichts. Jede Änderung steht mit Zelle, altem Wert, neuem Wert und Regel in `umsatz_sauber.aenderungen.json`.

## Excel: Text in Zahl umwandeln (1.234,56) und Text in Datum

Deutsche Zahlen, die als Text in der Zelle stehen, wandelt `excel-bereinigen` in echte Zahlen um ([Regel](docs/regeln.md#deutsche-zahl)).

| Text | wird | Grund |
|---|---|---|
| `1.234,56` | 1234,56 | Punkt als Tausender, Komma als Dezimalzeichen |
| `-250,75` | -250,75 | Minus vorne |
| `1 234,56` | 1234,56 | Leerzeichen als Tausender |
| `1.234` | bleibt Text | kann 1234 oder 1,234 sein, steht als Hinweis im Bericht |
| `01067` | bleibt Text | kann eine Postleitzahl oder Nummer sein |

Text in Datum genauso: `31.12.2026` und `31.12.26` werden ein Excel-Datum, `12/31/2026` bleibt Text mit Hinweis, weil Tag und Monat vertauscht sein können ([Regel](docs/regeln.md#datum)).

## Excel-Formelfehler finden: SVERWEIS #NV und #BEZUG!

`excel-formeln` liest die Ergebnisse, die Excel beim Speichern hinterlegt hat, und nennt jeden Fehlerwert mit Zelle, Formel, deutschem Namen und Grund.

| Zelle | Fehler | Deutsch | Formel | Hinweis |
|---|---|---|---|---|
| D5 | #REF! | #BEZUG! | =B5*#REF! | Bezug zeigt auf eine gelöschte Zeile, Spalte oder ein gelöschtes Blatt |
| E5 | #N/A | #NV | =VLOOKUP(A5,Lieferanten!A:B,2,FALSE) | SVERWEIS findet den Suchwert nicht im Suchbereich |

Dazu kommen harte Zahlen in einer Formelspalte (eine `99`, wo sonst `=B4*C4` steht), Datenzeilen, in denen die Formel fehlt, und externe Verknüpfungen wie `='[1]Preisliste'!B2` ([Regel](docs/regeln.md#formelspalte)). Das Skill ändert nichts und rechnet nichts neu.

## Excel in CSV umwandeln: DATEV, Semikolon, UTF-8

`excel-csv` kennt drei feste Formate ([Regel](docs/regeln.md#csv-formate)):

| Format | Trennzeichen | Dezimal | Datum | Kodierung |
|---|---|---|---|---|
| `datev` | `;` | `,` | `TT.MM.JJJJ` | Windows-1252, Text in Anführungszeichen |
| `excel-de` | `;` | `,` | `TT.MM.JJJJ` | UTF-8 mit BOM, Umlaute öffnen richtig in Excel |
| `utf8` | `,` | `.` | `JJJJ-MM-TT` | UTF-8, für Python, pandas, Datenbanken |

```csv
"Belegdatum";"Konto";"Gegenkonto";"Betrag netto";"Betrag brutto";"Buchungstext"
15.01.2026;8400;10001;1000;1190;"Rechnung 2026-001 Müller; Wartung"
```

Formeln werden als gespeicherte Ergebnisse ausgegeben, das sagt der Befehl an. Eine Formel ohne gespeichertes Ergebnis bleibt leer und steht als Hinweis im Protokoll. Passt ein Zeichen nicht in Windows-1252, bricht `datev` ab und nennt die Zelle. `datev` schreibt eine Tabelle für den ASCII-Import, keinen EXTF-Buchungsstapel mit Kopfsatz.

## Was das Original schützt

| Grenze | So ist es |
|---|---|
| Original | wird nie geschrieben. `--out` ist Pflicht und darf nicht die Eingabe sein, auch nicht in anderer Groß- und Kleinschreibung oder als Hardlink |
| Protokoll | neben jeder Ausgabe `<name>.aenderungen.json` mit Zelle, alt, neu, Regel |
| Formeln | bleiben Formeln. Nur `csv` gibt Werte aus und sagt das an |
| Makros | laufen nie. Eine `.xlsm`-Kopie behält das VBA-Projekt, eine `.xlsx`-Kopie einer `.xlsm` wird abgelehnt |
| Netz | kein Upload, kein Schreibweg nach DATEV, SAP, SharePoint oder Fabric |
| Größe | 50 MB je Datei, 1.000.000 Zellen je Blatt |
| Exit-Codes | 0 fertig, 2 Aufruf oder Datei falsch, 3 Regel verbietet das Schreiben, nichts geschrieben |

## Claude Skills für Excel im Vergleich

<a id="vergleich-anthropic-xlsx"></a>
### Claude Skills für Excel oder der xlsx-Skill von Anthropic?

Anthropic veröffentlicht in [anthropics/skills](https://github.com/anthropics/skills) einen allgemeinen xlsx-Skill zum Erstellen, Bearbeiten und Analysieren von Tabellen, unter einer proprietären Lizenz von Anthropic. Claude Skills für Excel sind enger: sechs Befehle für vorhandene deutsche Dateien, mit festen Regeln für `1.234,56`, `31.12.2026`, `#NV` und DATEV-CSV, unter MIT-Lizenz und mit einem Protokoll je Änderung. Wer eine neue Mappe bauen lässt, nimmt den allgemeinen Skill, wer eine bestehende Datei prüfen, bereinigen oder exportieren will, diese Skills.

<a id="vergleich-claude-for-excel"></a>
### Claude Skills für Excel oder Claude for Excel?

Claude for Excel ist ein Add-in von Anthropic, das Claude in der geöffneten Mappe arbeiten lässt. Claude Skills für Excel sind kein Add-in, sondern Agent Skills für Claude Code und Codex mit einem lokalen Programm. Sie arbeiten auf Dateien im Ordner, schreiben Kopien und ein Protokoll und entscheiden nach festen Regeln. Beides schließt sich nicht aus.

<a id="vergleich-copilot"></a>
### Claude Skills für Excel oder Microsoft Copilot in Excel?

Copilot in Excel läuft in der geöffneten Mappe, braucht eine Microsoft-365-Lizenz mit Copilot und verarbeitet die Anfrage in der Microsoft-Cloud. Claude Skills für Excel sind ein lokales Programm, das ein Agent wie Claude Code oder Codex aufruft: Die Datei verlässt den Rechner nicht, die Entscheidungen trifft das Programm nach festen Regeln, nicht ein Modell. Was der Agent aus dem Bericht liest, geht an dessen Anbieter.

<a id="vergleich-power-query"></a>
### Claude Skills für Excel oder Power Query?

Power Query ist in Excel eingebaut, kann alle Dateien eines Ordners zusammenführen, Spalten umwandeln und den Ablauf per Aktualisieren wiederholen. Das passt für eine Abfrage, die jeden Monat gleich in derselben Mappe läuft. `excel-stapeln` und `excel-bereinigen` passen für eine Aufgabe per Satz an den Agenten: Abweichende Spalten brechen den Lauf mit Exit-Code 3 ab und werden je Datei genannt, übersprungene Summenzeilen stehen im Änderungsprotokoll.

<a id="vergleich-mcp"></a>
### Claude Skills für Excel oder ein Excel MCP Server?

Ein Excel MCP Server wie [haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) gibt dem Agenten Werkzeuge, um Zellen, Blätter und Diagramme live zu lesen und zu schreiben. Das passt, wenn der Agent eine Mappe aufbauen oder Werte eintragen soll. Claude Skills für Excel passen, wenn eine vorhandene deutsche Datei geprüft, als Kopie bereinigt, verglichen oder exportiert werden soll, nach festen Regeln und mit Protokoll. Beides lässt sich kombinieren.

## Claude Code Skills installieren (auch Codex und Cursor)

Für Claude Code, Codex, Cursor und andere Agenten reicht eine Zeile:

```bash
npx skills add pexon-consulting/excel-skills
```

Wer die Skills im Produkt aktuell halten will, nutzt den Marketplace. Die Kopie von `npx skills add` zieht spätere Änderungen nicht von selbst.

```bash
claude plugin marketplace add pexon-consulting/excel-skills
codex plugin marketplace add pexon-consulting/excel-skills
```

Dauerhaft auf den PATH, dieselbe Version wie der Agenten-Befehl:

```bash
uv tool install git+https://github.com/pexon-consulting/excel-skills@v0.1.1
excel-skills profil umsatz.xlsx --out report.md
excel-skills schema
```

Entwicklung in diesem Checkout:

```bash
python3 -m venv .venv
.venv/bin/pip install -e ".[dev]"
.venv/bin/pytest
python3 scripts/fixtures_bauen.py     # Testdateien neu erzeugen
python3 scripts/beispiele_bauen.py    # README-Tabelle und docs/beispiele.md neu erzeugen
python3 scripts/sync_skills.py        # skills/ nach .claude/skills/ und .agents/skills/ spiegeln
```

Python 3.11 oder neuer. Einzige Laufzeitabhängigkeit ist [openpyxl](https://openpyxl.readthedocs.io/). Kein pandas, kein Excel, kein LibreOffice. Was sich je Version geändert hat, steht in [CHANGELOG.md](CHANGELOG.md).

## Fragen und Antworten

### Gibt es einen Claude Skill für Excel?

Ja. Claude Skills für Excel sind sechs Agent Skills für Claude Code und Codex: `excel-profil`, `excel-bereinigen`, `excel-formeln`, `excel-vergleich`, `excel-stapeln` und `excel-csv`. Installiert wird mit `npx skills add pexon-consulting/excel-skills`, Lizenz MIT, von Pexon Consulting GmbH. Für allgemeine Tabellenarbeit gibt es zusätzlich den xlsx-Skill von Anthropic, siehe [Vergleich](#vergleich-anthropic-xlsx).

### Wie räume ich eine unordentliche Excel-Datei mit Claude auf?

Mit `excel-bereinigen` aus Claude Skills für Excel. Ein Satz wie „Bereinige umsatz_bericht.xlsx“ reicht: Das Skill schreibt eine Kopie ohne verbundene Zellen, leere Zeilen und Zwischensummen, mit echten Zahlen und Datumswerten, und schreibt Formelbezüge um. Jede Änderung steht mit Zelle, altem und neuem Wert im Änderungsprotokoll, das Original bleibt unverändert.

### Was sind Claude Skills, und wie installiere ich sie?

Ein Claude Skill ist ein Ordner mit einer `SKILL.md`, die dem Agenten sagt, wann und wie er eine Aufgabe erledigt. Claude Code und Codex laden Skills nur, wenn die Beschreibung zur Anfrage passt. Installiert wird mit `npx skills add pexon-consulting/excel-skills` oder über den Plugin-Marketplace.

### Ist das dasselbe wie Claude for Excel?

Nein. Claude for Excel ist ein Add-in von Anthropic für die geöffnete Mappe. Claude Skills für Excel sind Agent Skills mit einem lokalen Programm, das Dateien prüft und Kopien schreibt. Siehe [Vergleich](#vergleich-claude-for-excel).

### Überschreibt das Skill meine Excel-Datei?

Nein. Jeder schreibende Befehl braucht `--out` mit einem neuen Pfad. Ist `--out` die Eingabedatei, auch in anderer Groß- und Kleinschreibung oder als Hardlink, bricht der Befehl mit Exit-Code 2 ab. Ein Test prüft für jeden Befehl, dass das Original byte-gleich bleibt.

### Laufen Makros aus einer .xlsm?

Nein. Die Datei wird gelesen, VBA wird nie ausgeführt. Beim Bereinigen bleibt das VBA-Projekt in der `.xlsm`-Kopie unverändert erhalten.

### Rechnet das Skill Formeln neu?

Nein. Es liest die Ergebnisse, die Excel beim letzten Speichern hinterlegt hat. Eine Datei, die nie in Excel gespeichert wurde, hat für neue Formeln kein Ergebnis. Das steht dann als Hinweis im Bericht.

### Warum zeigt SVERWEIS #NV?

`#NV` (englisch `#N/A`) heißt, SVERWEIS findet den Suchwert nicht in der ersten Spalte des Suchbereichs. Häufige Gründe sind ein Suchwert, der als Text statt als Zahl in der Zelle steht, Leerzeichen am Anfang oder Ende und ein Suchbereich, der nicht bis zur letzten Zeile reicht. `excel-formeln` listet jede `#NV`-Zelle mit Formel, `excel-profil` zeigt Spalten vom Typ `zahl_als_text`, `excel-bereinigen` trimmt Leerzeichen und wandelt `1.234,56` in einer Kopie in eine Zahl um. Ganzzahlen wie `4711` bleiben bewusst Text, weil sie Artikelnummern sein können.

### Braucht es Excel, LibreOffice oder eine Cloud?

Nein. Das Programm liest `.xlsx` und `.xlsm` direkt mit openpyxl, auf macOS, Linux und Windows. Kein API-Schlüssel, kein Konto, kein Netz.

### Warum bleibt 1.234 Text?

Weil `1.234` ohne Komma 1234 oder 1,234 heißen kann. Das Skill rät nicht. Die Zelle steht als Hinweis im Bericht, damit ein Mensch entscheidet.

### Wie sieht der DATEV-Export aus?

Semikolon als Trennzeichen, Komma als Dezimalzeichen, Datum `TT.MM.JJJJ`, Kodierung Windows-1252, Text in Anführungszeichen, Zeilenende CRLF. Das ist eine Tabelle für den DATEV-ASCII-Import, kein EXTF-Buchungsstapel.

### Geht das auch mit ChatGPT, Copilot oder Cursor?

Mit jedem Agenten, der lokale Shell-Befehle ausführen darf, etwa Claude Code, Codex, Cursor, Gemini CLI oder GitHub Copilot im Agent-Modus. Der Befehl ist immer `excel-skills`. Im ChatGPT-Chat und in Microsoft Copilot in Excel geht es nicht, dort startet kein lokales Programm. Den Unterschied zu Copilot in Excel beschreibt der [Vergleich](#vergleich-copilot).

### Funktioniert es mit Google Sheets?

Nach dem Export als `.xlsx` ja. Eine direkte Verbindung zur Google-Sheets-API gibt es nicht.

### Wohin gehen meine Daten?

Nirgendwohin. Das Programm läuft lokal und schreibt nur dorthin, wo `--out` hinzeigt. Was dein Agent mit dem Bericht macht, regelt dessen Anbieter. Details in [PRIVACY.md](PRIVACY.md).

### Was ist der Unterschied zum Excel MCP Server?

Ein MCP Server gibt dem Agenten Werkzeuge zum Lesen und Schreiben einzelner Zellen. Diese Skills sind feste Prüf- und Bereinigungsregeln für deutsche Dateien, mit Protokoll und ohne Änderung am Original. Siehe [Vergleich](#vergleich-mcp).

### Welche Version soll ein Agent aufrufen?

`v0.1.1`, angehängt an die Git-URL: `git+https://github.com/pexon-consulting/excel-skills@v0.1.1`. Damit läuft dieses Release und nicht ein späterer Stand von `main`. Die Regeln dieser Version liefert `excel-skills schema`.

<a id="english"></a>
## English: Agent skills for German Excel files

Claude Skills for Excel are six agent skills for Claude Code and Codex that check, clean, compare, merge and export German Excel files, locally and without touching the original. Built by Pexon Consulting GmbH, maintained by Phillip Pham.

German spreadsheets break general tools: numbers like `1.234,56` and dates like `31.12.2026` stored as text, the header row in row 4, merged cells, subtotals in the middle of a list, `#BEZUG!` instead of `#REF!`, and CSV with semicolons for DATEV. These skills handle exactly that, with fixed rules instead of model guesses.

- **Find the header row** when it is not in row 1, and tell a table from a report: `excel-skills profil file.xlsx`.
- **Clean an Excel file** into a copy: unmerge cells and fill down, remove blank rows and subtotals, convert the German number format `1.234,56` and dates like `31.12.2026`. Formula references, named ranges, tables and hyperlinks are rewritten when rows move: `excel-skills bereinigen file.xlsx --out clean.xlsx`.
- **Find formula errors**: `#REF!` (`#BEZUG!`), VLOOKUP `#N/A` (`#NV`), hard-coded numbers in formula columns, external links: `excel-skills formeln file.xlsx`.
- **Compare two Excel files** by a key column: added, removed and changed rows: `excel-skills vergleich old.xlsx new.xlsx --schluessel SKU`.
- **Combine multiple Excel files into one** table with the same columns, with a `quelle_datei` column: `excel-skills stapeln ./folder --out all.xlsx`.
- **Convert Excel to CSV** for DATEV (semicolon, decimal comma, Windows-1252), German Excel or UTF-8: `excel-skills csv file.xlsx --format datev --out file.csv`.

```bash
npx skills add pexon-consulting/excel-skills
uvx --from git+https://github.com/pexon-consulting/excel-skills@v0.1.1 excel-skills schema
```

Macros never run. Nothing is uploaded. Every change is logged in `<name>.aenderungen.json`. MIT licensed.

### Which Claude skill should I use for Excel?

Anthropic publishes a general xlsx skill for creating, editing and analysing spreadsheets under a proprietary license. Claude Skills for Excel are narrower: six MIT-licensed commands for existing German Excel files, with fixed rules for `1.234,56`, `31.12.2026`, `#NV` and DATEV CSV, and a log of every change.

### How do I clean a messy Excel file with Claude Code?

Run `npx skills add pexon-consulting/excel-skills` and ask Claude Code to clean the file. `excel-bereinigen` writes a copy without merged cells, blank rows and subtotals, converts German numbers and dates, rewrites formula references and logs every change in `<name>.aenderungen.json`.

### Why does VLOOKUP return #N/A?

`#N/A` (`#NV` in German Excel) means VLOOKUP did not find the lookup value. `excel-formeln` lists every `#N/A` cell with its formula, read from the values Excel saved, without recalculating.

### How do I compare two Excel files with Claude?

`excel-vergleich` matches rows by a key column such as an SKU and reports added, removed and changed rows with old and new values: `excel-skills vergleich old.xlsx new.xlsx --schluessel SKU`.

## Danksagung

Lesen und Schreiben übernimmt [openpyxl](https://openpyxl.readthedocs.io/). Das Plugin-Schema folgt [cathrynlavery/diagram-design](https://github.com/cathrynlavery/diagram-design), die Installation läuft über [vercel-labs/skills](https://github.com/vercel-labs/skills). Mehr in [docs/danksagung.md](docs/danksagung.md). Aus derselben Werkstatt: [mailbox-triage](https://github.com/pexon-consulting/mailbox-triage) sortiert ein Postfach in Rechnung, Mahnung, Angebot und Auftrag.

## Herkunft

Beispiel von Pexon Consulting GmbH, gepflegt von Phillip Pham. Nutzung nach MIT, ohne Support. Für den Einsatz bei Ihnen: https://pexon-consulting.de/data-services/excel-risiken-audit-controlling/
