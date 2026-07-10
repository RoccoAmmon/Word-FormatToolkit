# 🎮 Bedienung

## GUI-Übersicht

Die Benutzeroberfläche ist in mehrere Bereiche unterteilt:

### 1. 📁 Dokumente (linke Seite)

Hier verwalten Sie die zu verarbeitenden Word-Dokumente.

| Button | Aktion |
|--------|--------|
| **➕ Dateien** | Öffnet einen Datei-Dialog zur Auswahl einzelner `.docx`/`.doc`-Dateien (Mehrfachauswahl möglich) |
| **📂 Ordner** | Wählt einen kompletten Ordner aus – alle `.docx`-Dateien werden rekursiv eingelesen |
| **➖ Entfernen** | Entfernt die markierten Einträge aus der Liste |
| **🗑️ Leeren** | Leert die gesamte Liste |

### 2. 📋 Vorlage & Tabellen-Style (rechte Seite)

**Vorlage (.dotx/.dotm)**
- Dropdown mit allen gefundenen Word-Vorlagen
- Die Suche erfolgt automatisch in Registry + bekannten Pfaden
- **🔄 Neu suchen** – Durchsucht die Pfade erneut
- **📂 Datei...** – Manuelle Auswahl einer Vorlagendatei
- Der Pfad kann auch direkt eingetippt werden (IsEditable)

**Tabellen-Style**
- Sobald eine Vorlage ausgewählt ist, werden alle verfügbaren Tabellen-Styles geladen
- Wählen Sie den gewünschten Style für die Tabellen-Formatierung

### 3. ⚙️ Aktionen (rechte Seite)

| Checkbox | Funktion |
|----------|----------|
| **📝 Überschriften reparieren** | Setzt Überschriften-Styles neu auf |
| **🔢 Levelsprünge korrigieren** | Gleicht Hierarchiesprünge an |
| **🔁 Doppelte Nummern entfernen** | Löscht manuelle Nummern bei Auto-Nummerierung |
| **✏️ Manuelle Nummerierung finden** | Listet manuell nummerierte Überschriften (keine Korrektur) |
| **🔗 Tote Links prüfen** | Prüft Hyperlinks und Querverweise |
| **📊 Tabellen formatieren** | Wendet den gewählten Tabellen-Style an |
| **📑 Inhaltsverzeichnis updaten** | Aktualisiert TOC, Abbildungsverzeichnisse, Felder |

### 4. 📋 Live-Log (unten)

Das Log zeigt farbcodierte Meldungen während der Verarbeitung:

| Farbe | Level | Bedeutung |
|-------|-------|-----------|
| 🔵 Blau (`#0078D4`) | STEP | Einzelner Verarbeitungsschritt |
| 🟢 Grün (`#107C10`) | SUCCESS | Erfolgreich abgeschlossen |
| 🟠 Orange (`#CA5010`) | WARN | Warnung (z. B. toter Link gefunden) |
| 🔴 Rot (`#D13438`) | ERROR | Fehler bei der Verarbeitung |
| ⚫ Dunkelgrau | INFO | Allgemeine Information |

### 5. 🗂️ Aufräumen (rechte Seite)

| Feld | Standard | Beschreibung |
|------|----------|-------------|
| **Logs** | 30 Tage | Log-Dateien älter als X Tage löschen |
| **Reports** | 90 Tage | Vergleichsberichte älter als X Tage löschen |
| **Backups** | 14 Tage | _Backup-Dateien älter als X Tage löschen |

**🧹 Jetzt aufräumen** – Führt die Bereinigung sofort aus.

### 6. Fortschritt & Steuerung (unten)

- **Fortschrittsbalken** zeigt den aktuellen Status der Batch-Verarbeitung
- **Status-Text** zeigt die aktuelle Datei und den Gesamtfortschritt
- **⏹️ Abbrechen** – Bricht die Verarbeitung nach dem aktuellen Dokument ab
- **🚀 Verarbeitung starten** – Startet die Batch-Verarbeitung
- **📊 Letzten Bericht öffnen** – Öffnet den zuletzt erstellten Vergleichsbericht

## Workflow

1. **Dokumente hinzufügen** – Einzeldateien oder Ordner
2. **Vorlage auswählen** – Automatische Suche oder manuell
3. **Tabellen-Style wählen** – Wird dynamisch aus der Vorlage geladen
4. **Aktionen konfigurieren** – Checkboxen setzen
5. **Verarbeitung starten** – 🚀 Button klicken
6. **Live-Log verfolgen** – Status live im Log
7. **Bericht öffnen** – Nach Abschluss erscheint ein Dialog mit dem Bericht

## Batch-Verarbeitung

- Alle Dokumente werden nacheinander verarbeitet
- Vor jedem Durchlauf wird automatisch ein Backup erstellt (`_Backup_JJJJMMTT_HHMMSS.docx`)
- Der Abbrechen-Button stoppt die Verarbeitung nach dem aktuellen Dokument
- Nach Abschluss erscheint eine Zusammenfassung mit Erfolgs-/Fehlerzähler
- Optional: HTML-Vergleichsbericht mit Vorher/Nachher-Statistiken
