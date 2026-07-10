# 🏗️ Architektur

## Überblick

Das Word-Format-Toolkit ist als einzelnes PowerShell-Skript (`WordFormatTool-GUI.ps1`) realisiert. Es kombiniert eine WPF-GUI mit der Microsoft Word COM-Schnittstelle zur automatisierten Dokumenten-Verarbeitung.

## Komponenten

```
WordFormatTool-GUI.ps1
├── 1. Initialisierung
│   ├── Konfiguration ($Global:Config)
│   ├── Word-Konstanten ($Global:WdConst)
│   ├── Synchronisierter Zustand ($Global:Sync)
│   └── Assemblys laden (WPF, Forms)
│
├── 2. Hilfsfunktionen
│   ├── Write-Log            – Logging in Datei + GUI
│   ├── Show/Update/Close-SplashScreen – Splash-Overlay
│   ├── Get-WordTemplatesFolder – Registry-Lese für Vorlagenordner
│   ├── Find-WordTemplates   – Vorlagen-Suche im Dateisystem
│   ├── Get-TemplateStyles   – Tabellen-Styles aus .dotx/.dotm auslesen
│   ├── Enable/Disable-WordTurboMode – Performance-Optimierung
│   └── Get-DocumentStats    – Vorher/Nachher-Statistik
│
├── 3. Dokument-Funktionen
│   ├── Repair-Headings      – Überschriften neu aufsetzen
│   ├── Repair-HeadingLevels – Levelsprünge korrigieren
│   ├── Remove-DuplicateHeadingNumbers – Doppelte Nummern entfernen
│   ├── Format-Tables        – Tabellen nach Vorlagen-Style formatieren
│   ├── Update-DocumentTOC   – Inhaltsverzeichnisse aktualisieren
│   ├── Test-DeadLinks       – Tote Hyperlinks & Querverweise prüfen
│   └── Test-ManualNumbering – Manuelle Nummerierung erkennen
│
├── 4. Batch-Verarbeitung
│   ├── Invoke-ProcessDocument – Einzeldokument mit Backup + Word-Handling
│   ├── New-ComparisonReport    – HTML-Vergleichsbericht generieren
│   └── Invoke-Cleanup          – Aufräum-Routine
│
└── 5. GUI (Show-MainGUI)
    ├── XAML-Definition         – WPF-Layout
    ├── Event-Handler           – Button-Clicks, SelectionChanged
    ├── fillTemplates           – Template-ComboBox befüllen
    ├── loadTemplateStyles      – Styles aus Template laden
    └── btnStart                – Batch-Verarbeitung starten
```

## Datenfluss

```
Benutzer wählt Dokumente + Vorlage
        │
        ▼
Template-Styles werden geladen (Get-TemplateStyles)
        │
        ▼
Benutzer wählt Tabellen-Style + Aktionen
        │
        ▼
Batch-Start: Für jedes Dokument:
        │
        ├── Backup erstellen
        ├── Word im Hintergrund starten
        ├── Turbo-Modus aktivieren
        ├── Vorher-Statistik (Get-DocumentStats)
        ├── Gewählte Aktionen ausführen
        │   ├── Repair-Headings
        │   ├── Repair-HeadingLevels
        │   ├── Remove-DuplicateHeadingNumbers
        │   ├── Format-Tables
        │   ├── Update-DocumentTOC
        │   ├── Test-DeadLinks
        │   └── Test-ManualNumbering
        ├── Nachher-Statistik (Get-DocumentStats)
        ├── Dokument speichern + schließen
        ├── Word beenden + COM-Ressourcen freigeben
        │
        ▼
Vergleichsbericht erstellen (optional)
```

## Word-COM-Integration

- Das Toolkit startet eine unsichtbare Word-Instanz (`$word.Visible = $false`)
- Alle Meldungen werden unterdrückt (`$word.DisplayAlerts = 0`)
- Der Turbo-Modus deaktiviert:
  - Bildschirmaktualisierung (`ScreenUpdating = $false`)
  - Seitenumbrüche (`Pagination = $false`)
  - Rechtschreib- und Grammatikprüfung
  - Auto-Speichern
- Nach der Verarbeitung wird Word sauber geschlossen und alle COM-Ressourcen werden freigegeben

## WPF-GUI

- Die GUI wird als XAML-String definiert und mit `XamlReader.Load()` geladen
- Die Steuerelemente werden über `FindName()` referenziert und im `$Global:UI`-Hash gespeichert
- Das Live-Log verwendet eine `RichTextBox` mit farbigen `Paragraph`-Elementen
- Der Fortschritt wird über `ProgressBar` und `TextBlock` dargestellt
- `[System.Windows.Forms.Application]::DoEvents()` hält die GUI während der Batch-Verarbeitung reaktionsfähig

## Threading

- Die gesamte Verarbeitung läuft im GUI-Thread (vereinfacht das Logging und die Fortschrittsanzeige)
- `DoEvents()` sorgt für GUI-Reaktionsfähigkeit
- Der Abbrechen-Button setzt ein Flag (`$Global:Sync.CancelRequested`), das vor jedem Dokument geprüft wird
