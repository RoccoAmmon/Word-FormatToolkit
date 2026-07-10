# 📊 Vergleichsbericht

Nach Abschluss einer Batch-Verarbeitung kann optional ein HTML-Vergleichsbericht erstellt werden. Dieser zeigt für jedes verarbeitete Dokument die wichtigsten Kennzahlen im Vorher/Nachher-Vergleich.

## Bericht-Aufbau

### 1. Dashboard (Zusammenfassung)

Oben im Bericht befinden sich drei Kacheln:

| Kachel | Farbe | Inhalt |
|--------|-------|--------|
| **Dokumente gesamt** | Blau | Anzahl aller verarbeiteten Dokumente |
| **Erfolgreich** | Grün | Anzahl erfolgreich verarbeiteter Dokumente |
| **Fehler** | Rot | Anzahl der Dokumente, bei denen ein Fehler aufgetreten ist |

### 2. Detailtabelle

Pro Dokument werden folgende Spalten angezeigt:

| Spalte | Beschreibung |
|--------|-------------|
| **Datei** | Dateiname des Dokuments |
| **Status** | OK (grün) oder FEHLER (rot) |
| **Überschr.** | Anzahl der gefundenen Überschriften |
| **Tabellen** | Anzahl der Tabellen im Dokument |
| **Levelsprünge (vor→nach)** | Hierarchiesprünge vor/nach der Korrektur – gelb wenn vorher > 0 |
| **Duplikate (vor→nach)** | Doppelte Kapitelnummern vor/nach – gelb wenn vorher > 0 |
| **Manuell nummeriert** | Gefundene manuelle Nummerierungen – gelb wenn > 0 |
| **Tote Links** | Gefundene defekte Verweise – rot wenn > 0 |
| **Tab. format.** | Anzahl formatierter Tabellen |
| **TOC** | Anzahl aktualisierter Verzeichnisse |
| **Dauer** | Verarbeitungszeit pro Dokument |

### 3. Fußzeile

Enthält den Namen des Toolkits und den Pfad zur Log-Datei.

## Interpretation

- **Levelsprünge**: Sollten nach der Korrektur (→) bei 0 sein. Falls nicht, konnte die Hierarchie nicht vollständig repariert werden.
- **Duplikate**: Nach der Korrektur sollten keine doppelten Nummern mehr vorhanden sein.
- **Manuelle Nummerierung**: Diese werden nur gefunden, nicht korrigiert – als Hinweis für den Benutzer.
- **Tote Links**: Sollten manuell überprüft und korrigiert werden.

## Speicherort

Die Berichte werden unter `C:\ScriptLog\Reports\Vergleichsbericht_JJJJMMTT_HHMMSS.html` gespeichert.

## Automatische Anzeige

Nach der Batch-Verarbeitung fragt das Toolkit, ob der Bericht sofort im Browser geöffnet werden soll.
