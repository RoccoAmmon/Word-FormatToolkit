# ⚙️ Konfiguration

## Globale Konfiguration ($Global:Config)

Die Konfiguration befindet sich am Anfang des Skripts (ca. Zeile 27) in der `$Global:Config`-Hashtable.

```powershell
$Global:Config = @{
    TemplatePath       = ""                                    # Wird dynamisch ermittelt
    TableStyleName     = "Gitternetztabelle 4 – Akzent 1"       # Fallback
    LogFolder          = "C:\ScriptLog"                         # Log-Verzeichnis
    ReportFolder       = "C:\ScriptLog\Reports"                 # Bericht-Verzeichnis
    VerboseSteps       = $true                                  # Detailliertes Log
    ExtraTemplateDirs  = @("E:\", "\\Server\Vorlagen")          # Zusätzliche Suchpfade
}
```

### Einzelne Einstellungen

| Variable | Typ | Standard | Beschreibung |
|----------|-----|----------|-------------|
| `TemplatePath` | `string` | `""` | Pfad zur Word-Vorlage (.dotx/.dotm). Wird dynamisch aus der GUI befüllt. |
| `TableStyleName` | `string` | `"Gitternetztabelle 4 – Akzent 1"` | Name des Tabellen-Styles, der auf alle Tabellen angewendet wird. Wird aus der GUI-ComboBox überschrieben. |
| `LogFolder` | `string` | `"C:\ScriptLog"` | Verzeichnis für Log-Dateien. Wird automatisch erstellt, falls nicht vorhanden. |
| `ReportFolder` | `string` | `"C:\ScriptLog\Reports"` | Verzeichnis für HTML-Vergleichsberichte. |
| `VerboseSteps` | `bool` | `$true` | Wenn `$true`, werden einzelne Schritte detailliert im Log ausgegeben. |
| `ExtraTemplateDirs` | `string[]` | `@("E:\", "\\Server\Vorlagen")` | Zusätzliche Verzeichnisse, die nach Word-Vorlagen durchsucht werden. |

## Word-Konstanten ($Global:WdConst)

Die `$Global:WdConst`-Hashtable enthält wichtige Word-Enum-Werte für die COM-Interop:

| Konstante | Wert | Bedeutung |
|-----------|------|-----------|
| `LineStyleNone` | 0 | Kein Linienstil |
| `LineStyleSingle` | 1 | Einfache Linie |
| `LineWidth050pt` | 4 | Linienstärke 0,5 pt |
| `AutoFitWindow` | 2 | Tabelle an Fensterbreite anpassen |
| `PreferredWidthPercent` | 2 | Breitenangabe in Prozent |
| `FindStop` | 0 | Suche am Ende stoppen |
| `Paragraph` | 4 | Einheit: Absatz |
| `CollapseEnd` | 0 | Auswahl ans Ende einklappen |
| `BorderTop` | -1 | Oberer Rand |
| `BorderLeft` | -2 | Linker Rand |
| `BorderBottom` | -3 | Unterer Rand |
| `BorderRight` | -4 | Rechter Rand |

## Logging

- Log-Dateien werden unter `$LogFolder\WordFormatGUI_JJJJMMTT_HHMMSS.log` gespeichert
- Das Log-Format ist: `[HH:mm:ss] [LEVEL] Nachricht`
- Im GUI wird das Log farbcodiert in einer RichTextBox dargestellt

## Vergleichsbericht

- HTML-Berichte werden unter `$ReportFolder\Vergleichsbericht_JJJJMMTT_HHMMSS.html` gespeichert
- Der Bericht enthält eine Zusammenfassung (Dashboard) und eine Detailtabelle pro Dokument

## Aufräum-Routine

Die Aufräum-Routine (`Invoke-Cleanup`) löscht automatisch:

- **Logs** älter als X Tage (Standard: 30)
- **Vergleichsberichte** älter als X Tage (Standard: 90)
- **_Backup-Dateien** älter als X Tage (Standard: 14) – sucht in den Verzeichnissen der verarbeiteten Dokumente

Die Werte können in der GUI unter **🗂️ Aufräumen (Tage)** angepasst werden.
