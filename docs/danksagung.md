# Danksagung

Claude Skills für Excel bauen auf offener Arbeit auf. Kein Code wurde kopiert, Form und Schema schon.

- [openpyxl](https://openpyxl.readthedocs.io/) (MIT) liest und schreibt die Excel-Dateien. Es ist die einzige Laufzeitabhängigkeit.
- [cathrynlavery/diagram-design](https://github.com/cathrynlavery/diagram-design) zeigt, wie ein Skill-Paket zugleich als Claude-Plugin und Codex-Plugin funktioniert. Die Dateien in `.claude-plugin/`, `.codex-plugin/` und `.agents/plugins/` folgen diesem Schema.
- [vercel-labs/skills](https://github.com/vercel-labs/skills) liefert `npx skills add`, über das die Skills in Claude Code, Codex, Cursor und weitere Agenten kommen.
- [haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server) (MIT) ist der richtige Weg, wenn ein Agent Zellen live über MCP schreiben soll. Die Abgrenzung steht im README.
