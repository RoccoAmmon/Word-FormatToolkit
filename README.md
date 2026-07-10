# 📄 Word-Format-Toolkit v1.0

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)](https://docs.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)](https://www.microsoft.com/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Version](https://img.shields.io/badge/Version-1.0.0-purple)](CHANGELOG.md)

> **Batch-Formatierung und Qualitätssicherung für Microsoft Word-Dokumente**  
> PowerShell-WPF-Tool zur automatisierten Reparatur von Überschriften, Tabellen, Nummerierungen und Links in Word-Dokumenten (.docx/.doc).

---

## 📑 Inhaltsverzeichnis

- [Features](#-features)
- [Systemanforderungen](#-systemanforderungen)
- [Installation](#-installation)
- [Verwendung](#-verwendung)
- [Screenshots](#-screenshots)
- [Konfiguration](#-konfiguration)
- [Vergleichsbericht](#-vergleichsbericht)
- [Architektur](#-architektur)
- [Troubleshooting](#-troubleshooting)
- [Changelog](#-changelog)
- [License](#-license)

---

## ✨ Features

### 📝 Überschriften-Reparatur
- Setzt Überschriften-Styles neu auf (Style-Neuaufsetzung)
- Behebt Formatierungs-Müll durch direkte Formatierung
- Erhält die bestehenden Style-Namen (Überschrift 1–9 / Heading 1–9)

### 🔢 Levelsprung-Korrektur
- Automatische Erkennung von Hierarchiesprüngen (z. B. Überschrift 1 → Überschrift 3)
- Mehrstufige Korrektur (bis zu 3 Durchgänge)
- Nachfolgende Nummerierung wird aktualisiert

### 🔁 Doppelte Kapitelnummern entfernen
- Erkennt manuell eingegebene Nummern am Überschrift-Anfang
- Entfernt diese, wenn automatische Nummerierung aktiv ist
- Verhindert doppelte Darstellung („1.1 1.1 Einleitung")

### 📊 Tabellen-Formatierung
- Wähle einen beliebigen Tabellen-Style aus der gewählten Vorlage (.dotx/.dotm)
- Alle Tabellen im Dokument erhalten einheitlich diesen Style
- Automatische Aktivierung von Heading Rows, First Column, Row Bands
- Entfernt vorherige direkte Tabellen-Formatierungen
- Setzt Tabellen auf Seitenbreite (100 %)

### 📑 Inhaltsverzeichnis aktualisieren
- Aktualisiert alle TablesOfContents, TablesOfFigures
- Aktualisiert Kopf- und Fußzeilen-Felder
- Seitenumbrüche werden neu berechnet (Repaginate)

### 🔗 Link-Prüfer
- Prüft alle Hyperlinks auf gültige Ziele
- Erkennt fehlende Lesezeichen (interne Sprungziele)
- Prüft Datei-Verknüpfungen auf Existenz
- Findet tote Querverweise (REF-Felder ohne Ziel)

### ✏️ Manuelle Nummerierung erkennen
- Findet Überschriften, die manuell nummeriert wurden (z. B. „3.2.1 Einleitung")
- Listet diese im Log auf – hilfreich für die Migration auf automatische Nummerierung

### 📈 HTML-Vergleichsbericht
- Automatische Erstellung nach der Batch-Verarbeitung
- Vorher/Nachher-Vergleich (Levelsprünge, Duplikate, Tabellen)
- Übersichtliches Dashboard mit Erfolgs-/Fehlerzähler
- Wird nach Abschluss automatisch zur Ansicht angeboten

### 🧹 Aufräum-Routine
- Löscht alte Logs (Standard: >30 Tage)
- Löscht alte Vergleichsberichte (Standard: >90 Tage)
- Löscht alte _Backup-Dateien (Standard: >14 Tage)
- Konfigurierbar über die GUI

---

## 💻 Systemanforderungen

- **Windows** 10/11 oder Windows Server 2016+
- **PowerShell** 5.1+ (mit STA-Modus)
- **Microsoft Word** (Office 2013/2016/2019/2021/365) – COM-Komponente
- **.NET Framework** 4.6.1+ (für WPF)
- **Ausführungsrichtlinie**: `Bypass` oder `RemoteSigned`

---

## 🚀 Installation

```powershell
# 1. Repository klonen
git clone https://github.com/RoccoAmmon/Word-FormatToolkit.git
cd Word-FormatToolkit

# 2. Ausführungsrichtlinie (falls nötig)
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser

# 3. Skript starten
.\WordFormatTool-GUI.ps1
```

> **Hinweis:** Das Skript muss mit der `-STA`-Flag gestartet werden (Single-Threaded Apartment).  
> `powershell.exe -ExecutionPolicy Bypass -STA -File "WordFormatTool-GUI.ps1"`  
> Die PowerShell ISE startet standardmäßig im STA-Modus. Die neue PS Console (pwsh.exe) **nicht** – dort ggf. mit `-STA` starten oder in der ISE öffnen.

---

## 🎮 Verwendung

### Workflow

1. **📁 Dokumente hinzufügen** – Einzelne Dateien oder ganze Ordner mit `.docx`/`.doc`
2. **📋 Vorlage auswählen** – Automatische Suche nach `.dotx`/`.dotm` in bekannten Pfaden
3. **📊 Tabellen-Style wählen** – Die verfügbaren Styles werden aus der Vorlage geladen
4. **⚙️ Aktionen konfigurieren** – Checkboxen für gewünschte Reparaturen
5. **🚀 Verarbeitung starten** – Batch läuft mit Live-Log und Fortschrittsanzeige
6. **📈 Bericht prüfen** – Optionaler HTML-Vergleichsbericht

### GUI im Detail

| Bereich | Beschreibung |
|---------|-------------|
| **📁 Dokumente** | Liste der zu verarbeitenden Dateien – via ➕/📂/➖ verwalten |
| **📋 Vorlage** | Auswahl einer `.dotx`/`.dotm` – sucht automatisch in Registry + bekannten Pfaden |
| **📊 Tabellen-Style** | Alle verfügbaren Tabellen-Styles aus der gewählten Vorlage |
| **⚙️ Aktionen** | Checkboxen zum Aktivieren/Deaktivieren einzelner Reparaturen |
| **📋 Live-Log** | Farbcodiertes Log (INFO/WARN/ERROR/SUCCESS/STEP) |
| **📈 Fortschritt** | Fortschrittsbalken + Status-Text |

---

## 📸 Screenshots

*(Screenshots folgen in einer späteren Version)*

---

## ⚙️ Konfiguration

Die wichtigsten Einstellungen werden über die GUI gesteuert. Darüber hinaus können folgende Variablen im Skript-Kopf angepasst werden (`$Global:Config`):

| Variable | Standard | Beschreibung |
|----------|----------|-------------|
| `TableStyleName` | `"Gitternetztabelle 4 – Akzent 1"` | Standard-Tabellen-Style (wird aus GUI überschrieben) |
| `LogFolder` | `"C:\ScriptLog"` | Ausgabeverzeichnis für Logs |
| `ReportFolder` | `"C:\ScriptLog\Reports"` | Ausgabeverzeichnis für Vergleichsberichte |
| `VerboseSteps` | `$true` | Detaillierte Schritt-für-Schritt-Ausgabe im Log |
| `ExtraTemplateDirs` | `@("E:\", "\\Server\Vorlagen")` | Zusätzliche Verzeichnisse für die Vorlagen-Suche |

---

## 📄 Vergleichsbericht

Der HTML-Vergleichsbericht zeigt pro Dokument:

- **Status** (OK/FEHLER)
- **Überschriften** vor/nach
- **Levelsprünge** vor → nach
- **Doppelte Nummern** vor → nach
- **Manuelle Nummerierung** (gefunden)
- **Tote Links** (gefunden)
- **Formatierte Tabellen** & **TOC-Updates**
- **Dauer** pro Dokument

Der Bericht wird unter `C:\ScriptLog\Reports\Vergleichsbericht_*.html` gespeichert.

---

## 🏗️ Architektur

```
WordFormatTool-GUI.ps1
├── Konfiguration ($Global:Config, $Global:WdConst)
├── Splash-Screen (WPF-Overlay)
├── Vorlagen-Suche (Find-WordTemplates, Get-WordTemplatesFolder)
├── Style-Auslesen (Get-TemplateStyles – liest Tabellen-Styles aus .dotx/.dotm)
├── Dokument-Funktionen
│   ├── Repair-Headings          – Überschriften neu aufsetzen
│   ├── Repair-HeadingLevels     – Levelsprünge korrigieren
│   ├── Remove-DuplicateHeadingNumbers – Doppelte Nummern entfernen
│   ├── Format-Tables            – Tabellen nach Vorlagen-Style formatieren
│   ├── Update-DocumentTOC       – Inhaltsverzeichnisse aktualisieren
│   ├── Test-DeadLinks           – Tote Links prüfen
│   ├── Test-ManualNumbering     – Manuelle Nummerierung erkennen
│   └── Get-DocumentStats        – Vorher/Nachher-Statistik
├── Batch-Verarbeitung
│   ├── Invoke-ProcessDocument   – Einzeldokument verarbeiten
│   ├── New-ComparisonReport     – HTML-Vergleichsbericht
│   └── Invoke-Cleanup           – Aufräum-Routine
├── Word-Turbo-Modus (Enable/Disable-WordTurboMode)
├── Logging (Write-Log)
└── WPF-GUI (Show-MainGUI)
    ├── Datei-Management (➕📂➖)
    ├── Template-Auswahl + Style-Preview
    ├── Aktionen-Checkboxen
    ├── Live-Log (RichTextBox)
    ├── Progress-Bar
    └── Start/Cancel/Report-Buttons
```

---

## 🔧 Troubleshooting

| Problem | Lösung |
|---------|--------|
| **Word wird nicht gefunden** | Office-Installation prüfen. Nur 32-Bit Word wird unterstützt. |
| **Speichern-Dialog erscheint** | `$doc.Saved = $true` wird jetzt direkt nach dem Öffnen gesetzt. |
| **Keine Tabellen-Styles geladen** | Template (.dotx/.dotm) muss Tabellen-Styles enthalten. Neues Template wählen. |
| **Skript startet nicht** | Mit `-STA`-Flag starten: `powershell.exe -STA -File "WordFormatTool-GUI.ps1"` |
| **PowerShell 7 (pwsh.exe)** | STA wird nicht unterstützt. In der PowerShell ISE oder Windows PowerShell öffnen. |
| **GUI bleibt leer** | .NET Framework 4.6.1+ erforderlich für WPF. |
| **Backup-Dateien sammeln sich** | Aufräum-Routine verwenden (🧹 Button) oder `$BackupDays` anpassen. |

---

## 📋 Changelog

Siehe [CHANGELOG.md](CHANGELOG.md) für die vollständige Versionshistorie.

---

## 📜 License

Dieses Projekt ist unter der **MIT License** lizenziert – siehe [LICENSE](LICENSE) für Details.

---

## 👤 Autor

**Rocco Ammon**

- 🔗 GitHub: [@RoccoAmmon](https://github.com/RoccoAmmon)
- 📧 Kontakt via [GitHub Issues](https://github.com/RoccoAmmon/Word-FormatToolkit/issues)
- 🌍 Repository: [Word-FormatToolkit](https://github.com/RoccoAmmon/Word-FormatToolkit)

---

*Erstellt mit ❤️ und PowerShell*
