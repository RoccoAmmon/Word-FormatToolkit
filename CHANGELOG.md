# Changelog

Alle wichtigen Änderungen an diesem Projekt werden hier dokumentiert.

Das Format basiert auf [Keep a Changelog](https://keepachangelog.com/de/1.1.0/),
und dieses Projekt folgt [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.3.0] – 2026-07-28

### Fixed
- **Seitenumbruch vor Tabellen-/Abbildungsverzeichnis** – `PageBreakBefore` auf der Überschrift statt Formfeed-Zeichen
- **Bestehende Verzeichnisse** werden jetzt gelöscht und neu erstellt (mit korrektem Seitenumbruch)
- **Abbildungsbeschriftung im eigenen Absatz** – Bild erscheint nicht mehr im Abbildungsverzeichnis
- **Einheitliche Schriftgröße** bei Abbildungsbeschriftungen (Stil nur auf Beschriftungs-Absatz angewandt)
- **Keine Leerzeilen** mehr nach Abbildungsbeschriftungen

### Changed
- Erkennung bestehender Verzeichnisse per Feldcode-Regex (`\c "Tabelle"` / `\c "Abbildung"`)

---

## [1.2.0] – 2026-07-25

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
