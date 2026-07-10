# 🔧 Troubleshooting

## Häufige Probleme

### Word wird nicht gefunden / COM-Fehler

**Problem:** Beim Start erscheint ein COM-Fehler „Class not registered" oder „New-Object -ComObject Word.Application" schlägt fehl.

**Lösungen:**
- Prüfen Sie, ob Microsoft Word installiert ist
- Die Office-Installation muss die COM-Komponente enthalten (Standard bei jeder Office-Installation)
- Bei side-by-side Installationen (32-Bit + 64-Bit) muss die zum PowerShell passende Word-Version verwendet werden
- PowerShell 5.1 64-Bit → Word 64-Bit
- PowerShell 5.1 32-Bit → Word 32-Bit

### Speichern-Dialog erscheint beim Template-Laden

**Problem:** Beim Auswählen einer Vorlage (.dotx/.dotm) öffnet Word einen Dialog „Möchten Sie die Änderungen speichern?".

**Lösung:** Dieses Problem wurde in v1.0 behoben durch:
- `$doc.Saved = $true` direkt nach dem Öffnen
- `$doc.Close(0)` mit `wdDoNotSaveChanges`

Falls das Problem weiterhin auftritt, stellen Sie sicher, dass Sie die aktuelle Version verwenden.

### Keine Tabellen-Styles werden geladen

**Problem:** In der ComboBox „Tabellen-Style" steht nur der Fallback-Name, aber keine Styles aus der Vorlage.

**Lösungen:**
- Die ausgewählte Vorlage (.dotx/.dotm) muss Tabellen-Styles enthalten
- Öffnen Sie die Vorlage in Word und prüfen Sie unter `Start → Formatvorlagen → Formatvorlagen verwalten`, ob Tabellen-Styles vorhanden sind
- Wählen Sie eine andere Vorlage mit **📂 Datei...**
- Führen Sie **🔄 Neu suchen** aus, um die Vorlagen-Liste zu aktualisieren

### Skript startet nicht / PowerShell-Fehler

**Problem:** Das Skript startet nicht oder es erscheint ein Syntax-Fehler.

**Lösungen:**
- **STA-Modus**: PowerShell 7 (pwsh.exe) unterstützt keinen STA-Modus. Verwenden Sie:
  ```powershell
  powershell.exe -ExecutionPolicy Bypass -STA -File "WordFormatTool-GUI.ps1"
  ```
- **Ausführungsrichtlinie**:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
  ```
- **Leerzeichen im Pfad**: Wenn der Skript-Pfad Leerzeichen enthält, in Anführungszeichen setzen:
  ```powershell
  powershell.exe -STA -File "C:\Mein Ordner\WordFormatTool-GUI.ps1"
  ```

### GUI bleibt leer / WPF-Fehler

**Problem:** Das Skript startet, aber die GUI bleibt leer oder es erscheint ein XAML-Fehler.

**Lösungen:**
- .NET Framework 4.6.1+ ist für WPF erforderlich
- Prüfen Sie mit `[System.Environment]::Version`
- Bei Windows 7/8/8.1 ggf. .NET Framework manuell aktualisieren
- Windows 10/11 haben .NET Framework 4.8 standardmäßig installiert

### Backup-Dateien sammeln sich an

**Problem:** In den Dokumenten-Ordnern sammeln sich viele `_Backup_*`-Dateien.

**Lösungen:**
- Nutzen Sie die **🧹 Jetzt aufräumen**-Funktion in der GUI
- Passen Sie die Aufbewahrungsdauer für Backups an (Standard: 14 Tage)
- Die Aufräum-Routine durchsucht die Verzeichnisse der verarbeiteten Dokumente

### Word-Prozess bleibt hängen

**Problem:** Nach einem Abbruch oder Fehler bleibt ein Winword.exe-Prozess im Task-Manager.

**Lösungen:**
- Das Skript räumt COM-Ressourcen im `finally`-Block auf
- Bei hartem Abbruch (Taskkill / PowerShell-Fenster schließen) kann ein Word-Prozess übrig bleiben
- Manuell beenden:
  ```powershell
  Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force
  ```

### HTML-Vergleichsbericht wird nicht geöffnet

**Problem:** Nach der Verarbeitung erscheint kein Dialog zum Öffnen des Berichts.

**Lösungen:**
- Prüfen Sie, ob die Checkbox **📈 Vergleichsbericht** aktiviert ist
- Der Bericht wird unter `C:\ScriptLog\Reports\Vergleichsbericht_*.html` gespeichert
- Öffnen Sie ihn manuell über **📊 Letzten Bericht öffnen**

## Log-Datei

Bei Problemen finden Sie detaillierte Informationen in der Log-Datei:

```
C:\ScriptLog\WordFormatGUI_JJJJMMTT_HHMMSS.log
```

Das Log enthält:
- Zeitstempel jedes Schritts
- Erfolgs-/Fehlermeldungen pro Dokument
- Technische Details bei COM-Fehlern

## Support

- **GitHub Issues**: [https://github.com/RoccoAmmon/Word-FormatToolkit/issues](https://github.com/RoccoAmmon/Word-FormatToolkit/issues)
- **Wiki**: [https://github.com/RoccoAmmon/Word-FormatToolkit/wiki](https://github.com/RoccoAmmon/Word-FormatToolkit/wiki)
- **Repository**: [https://github.com/RoccoAmmon/Word-FormatToolkit](https://github.com/RoccoAmmon/Word-FormatToolkit)
