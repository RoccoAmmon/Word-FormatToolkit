# Changelog

Alle wichtigen Änderungen an diesem Projekt werden hier dokumentiert.

Das Format basiert auf [Keep a Changelog](https://keepachangelog.com/de/1.1.0/),
und dieses Projekt folgt [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.7.0] – 2026-09-30

### Added
- **Manuell nummerierte Standard-Absätze umwandeln** (`Repair-ManualNumberingHeadings`) – Absätze im Stil "Standard"/"Normal" mit manuell getippter Nummerierung (z.B. "2 Startinvestition", "2.1 Ziele") werden automatisch in "Überschrift N" umgewandelt. Ebene aus Gliederungstiefe ("2"→1, "2.1"→2). Schutz vor Fehltreffern: Tabellen, echte Listen/Auto-Nummerierung, Texte >150 Zeichen sowie Jahreszahlen/große Zahlen (>99) werden ignoriert.
- 🇬🇧 *Convert manually numbered Standard paragraphs into headings (`Repair-ManualNumberingHeadings`) – e.g. "2 Project Start" or "2.1 Goals" become "Heading N", level derived from depth ("2"→1, "2.1"→2). Skips tables, real lists/auto-numbering, texts >150 chars and year-like/large numbers (>99).*
- **Verarbeitungsreihenfolge** – "Manuelle Nummerierung" läuft jetzt vor "Doppelte Nummern entfernen", damit frisch umgewandelte Überschriften direkt von der Duplikat-Bereinigung profitieren.
- 🇬🇧 *Processing order – "Manual numbering" now runs before "Remove duplicate numbers" so freshly converted headings benefit from duplicate cleanup.*

### Changed
- **GUI** – Checkbox umbenannt in "Manuelle Nummerierung" und in der Aktionen-Liste vor "Doppelte Nummern entfernen" sortiert.
- **Script-Header** – Beschreibung um Umwandlung manueller Nummerierungen erweitert.
- 🇬🇧 *GUI checkbox renamed to "Manuelle Nummerierung" and moved above "Remove duplicate numbers"; script header description extended.*

---

## [1.6.0] – 2026-07-30

### Added
- **Fehlende Überschriften erkennen und korrigieren** (`Repair-MissingHeadings`) – Absätze mit Stil "Standard" (Normal), die eine Kapitelnummerierung (Gliederungsebene 1–9) haben, werden automatisch auf "Überschrift N" (bzw. "Heading N") gesetzt. Die Nummerierung bleibt erhalten, nur der Stil wird korrigiert.
- **Kopf-/Fußzeilen aus Vorlage klonen** (`Update-HeaderFooter`) – Alle Kopf- und Fußzeilen (Primär, Erste Seite, Gerade Seiten) werden aus der Vorlage in alle Abschnitte des Dokuments kopiert. Die Einstellungen "DifferentFirstPage" und "OddAndEvenPages" werden von der Vorlage übernommen und auf jeden Abschnitt angewendet. Inklusive frei positionierter Grafiken/Shapes.
- **HTML-Report** – Neue Spalte "Kopf/Fußz." im Vergleichsbericht

### Fixed
- **Kopf-/Fußzeilen-Klon robust gemacht** – Zwei-Pass-Verfahren (erst entkoppeln, dann kopieren) eliminiert "Befehl nicht verfügbar"-Fehler bei verlinkten Abschnitten. Fallback auf `FormattedText` bei fehlgeschlagenem `Range.Paste`.
- **Frei positionierte Grafiken in Kopf-/Fußzeilen** – Werden jetzt per `Range.Copy`/`Range.Paste` statt `FormattedText` übertragen, wodurch auch Shapes korrekt kopiert werden.
- 🇬🇧 *Robust header/footer cloning – two-pass approach (unlink, then copy) fixes "command unavailable" errors; floating shapes are transferred via Range.Copy/Paste with FormattedText fallback.*

---

## [1.5.0] – 2026-07-30

🇬🇧 *Clone headers/footers from the template (primary, first page, even pages) incl. DifferentFirstPage/OddAndEvenPages settings; new "Header/Footer" column in the HTML report.*

### Added
- **Kopf-/Fußzeilen aus Vorlage übernehmen** – Alle Kopf- und Fußzeilen (Primär, Erste Seite, Gerade Seiten) werden aus der Vorlage in alle Abschnitte des Dokuments kopiert. Die Einstellungen "DifferentFirstPage" und "OddAndEvenPages" werden von der Vorlage übernommen.

### Changed
- **HTML-Report** – Neue Spalte "Kopf/Fußz." im Vergleichsbericht
- **Script-Header** – Beschreibung um Kopf-/Fußzeilen-Feature erweitert

---

## [1.4.0] – 2026-07-29

🇬🇧 *Apply the "Standard" (Normal) style from the template (title page protected), splash screen with live status, GUI responsiveness fixes.*

### Added
- **„Standard" aus Vorlage übernehmen** – Die Formatvorlage "Standard" (Normal) wird aus der gewählten .dotx/.dotm-Vorlage in jedes Dokument kopiert. Schriftart, -größe, Absatzformatierung, Sprache u.v.m. werden übernommen
- **Schutz der ersten Seite** – Beim Standard-Style-Update bleibt die Titelseite (Seite 1) unverändert; ihre originale Formatierung wird gesichert und nach dem Style-Update wiederhergestellt
- **Splash-Screen mit Live-Status** – Der Start-Splash zeigt jetzt jede Phase an (Vorlagensuche → Styles laden → Bereit), der Titel wechselt dynamisch
- **Dynamischer Splash-Titel** – Der Fortschrittsbalken animiert dank verbessertem `Do-WpfEvents` (DispatcherTimer 30 ms) jetzt auch während langer Suchläufe

### Fixed
- **Fenster bleibt im Vordergrund** – `Topmost`-Flag vom Hauptfenster entfernt; das Toolkit verhält sich wie ein normales Fenster
- **GUI reagiert nicht während Batch** – Zahlreiche `DoEvents()`-Aufrufe in `Invoke-ProcessDocument` eingebaut, damit das Fenster auch während der Verarbeitung minimiert, verschoben oder in den Hintergrund geklickt werden kann
- **MessageBox geht verloren** – Alle `MessageBox.Show()`-Aufrufe erhalten jetzt das Hauptfenster als Owner, sodass sie nicht hinter anderen Fenstern verschwinden; nach der Batch-Meldung wird das Hauptfenster wieder aktiviert
- **Alte `Update-StandardStyle`-Funktion** – Doppelte, einfachere Version ohne Seitenschutz entfernt

---

## [1.3.0] – 2026-07-28

🇬🇧 *Page break before tables-of-tables/figures via PageBreakBefore; existing directories are rebuilt; figure captions live in their own paragraph; directory detection via field-code regex.*

### Fixed
- **Seitenumbruch vor Tabellen-/Abbildungsverzeichnis** – `PageBreakBefore` auf der Überschrift statt Formfeed-Zeichen
- **Bestehende Verzeichnisse** werden jetzt gelöscht und neu erstellt (mit korrektem Seitenumbruch)
- **Abbildungsbeschriftung im eigenen Absatz** – Bild erscheint nicht mehr im Abbildungsverzeichnis
- **Einheitliche Schriftgröße** bei Abbildungsbeschriftungen (Stil nur auf Beschriftungs-Absatz angewandt)
- **Keine Leerzeilen** mehr nach Abbildungsbeschriftungen

### Changed
- Erkennung bestehender Verzeichnisse per Feldcode-Regex (`\c "Tabelle"` / `\c "Abbildung"`)
- 🇬🇧 *Page breaks before tables-of-tables/figures via PageBreakBefore; existing directories are rebuilt; figure captions live in their own paragraph.*

---

## [1.2.0] – 2026-07-25

🇬🇧 *Automatic table/figure captions with SEQ fields plus tables of tables/figures; template font size is preserved.*

### Added
- **Tabellenbeschriftung** – Automatische Beschriftung unter jeder Tabelle
  - Format: „Tabelle N: Kapitelüberschrift" mit SEQ-Feldfunktion
  - Kapiteltext aus der letzten Überschrift vor der Tabelle
  - Vorhandene Beschriftungen werden vor Neuerstellung entfernt
- **Abbildungsbeschriftung** – Automatische Beschriftung unter jeder Inline-Abbildung
  - Format: „Abbildung N: Kapitelüberschrift" mit SEQ-Feldfunktion
  - Vorhandene Beschriftungen werden erkannt und entfernt
- **Tabellenverzeichnis** erstellen/aktualisieren (nach dem Inhaltsverzeichnis, neue Seite)
- **Abbildungsverzeichnis** erstellen/aktualisieren (nach dem Tabellenverzeichnis, neue Seite)
- Null-Prüfung nach `Documents.Open()` – klarere Fehlermeldung bei nicht öffenbaren Dokumenten

### Fixed
- Schriftgröße wird bei Template-Anwendung bewahrt (Standard-Style Font Name/Size sichern + wiederherstellen)

---

## [1.1.0] – 2026-07-11

🇬🇧 *Visual table-style preview (Word CopyAsPicture + Win32 clipboard EMF extraction) with caching and a persistent Word instance.*

### Added
- **Visuelle Tabellen-Style-Vorschau** in der GUI
  - Zeigt eine Echtzeit-Vorschau des ausgewählten Tabellen-Styles
  - Verwendet Word-CopyAsPicture + Win32-API (GetClipboardData/CopyEnhMetaFile)
  - Extrahiert Enhanced Metafile direkt aus der Zwischenablage (kein HTML-Umweg)
  - Mit Retry-Logik bei Timing-Problemen
  - Gecacht für sofortigen Wechsel zwischen bereits geladenen Styles
- Persistente Word-Instanz für die Vorschau (startet nur einmal, wird für alle Previews wiederverwendet)

### Fixed
- Pipeline-Leaks in `Write-Log` und `Show-TableStylePreview` unterdrückt (`| Out-Null`)
- `SelectionChanged`-Event bei editable ComboBox durch `DropDownClosed` ergänzt
- Word-Speicherdialog bei `Quit()` unterdrückt (`wdDoNotSaveChanges = 0`)

### Changed
- Timing optimiert: kürzere Sleeps (150–200ms statt 300–400ms)
- Log-Fenster in die linke Spalte verschoben, nimmt gesamte Resthöhe ein
- `Gdi32Helper` als separaten `Add-Type` für `DeleteObject`

---

## [1.0.0] – 2026-07-10

🇬🇧 *First release: WPF GUI with file management, template search, heading repair, level-jump fix, duplicate-number removal, table formatting, TOC update, link checker, manual-numbering detection, HTML comparison report, cleanup routine, live log and backups.*

### Added
- **Erstveröffentlichung** des Word-Format-Toolkit
- WPF-GUI mit Datei-Management (Einzeldaten, Ordner-Import, Drag & Drop)
- Vorlagen-Suche in Registry (`PersonalTemplates`), `%APPDATA%` und bekannten Pfaden
- Dynamisches Auslesen von Tabellen-Styles aus der gewählten `.dotx`/`.dotm`
- **Überschriften-Reparatur** – Style-Neuaufsetzung für Überschrift 1–9
- **Levelsprung-Korrektur** – Automatische Angleichung von Hierarchiesprüngen (bis zu 3 Durchgänge)
- **Doppelte Kapitelnummern entfernen** – Erkennt und löscht manuelle Nummern bei aktiver Auto-Nummerierung
- **Tabellen-Formatierung** – Einheitlicher Tabellen-Style aus Vorlage, entfernt direkte Formatierung
- **Inhaltsverzeichnis aktualisieren** – TOC, Abbildungsverzeichnisse, Felder, Kopf-/Fußzeilen
- **Link-Prüfer** – Erkennt tote Hyperlinks, fehlende Lesezeichen und Querverweise
- **Manuelle Nummerierung erkennen** – Listet manuell nummerierte Überschriften
- **HTML-Vergleichsbericht** – Vorher/Nachher-Dashboard mit Statistiken
- **Aufräum-Routine** – Löscht alte Logs, Reports und Backups (konfigurierbar)
- Live-Log in der GUI (farbcodiert, RichTextBox)
- Fortschrittsbalken mit Status-Text
- Abbrechen-Button für laufende Batch-Verarbeitung
- Word-Turbo-Modus (deaktiviert ScreenUpdating, Pagination, Rechtschreibprüfung)
- Backups vor jeder Verarbeitung (`_Backup_*`)
- Splash-Screen beim Start
- Konfigurierbare Extra-Template-Verzeichnisse

[1.0.0]: https://github.com/RoccoAmmon/Word-FormatToolkit/releases/tag/v1.0
