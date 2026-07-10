# 📥 Installation

## Systemanforderungen

| Komponente | Anforderung |
|------------|-------------|
| **Betriebssystem** | Windows 10/11 oder Windows Server 2016+ |
| **PowerShell** | 5.1+ (Windows PowerShell, **nicht** pwsh.exe) |
| **Microsoft Word** | Office 2013/2016/2019/2021/365 (COM-Komponente) |
| **.NET Framework** | 4.6.1+ (für WPF) |
| **Ausführungsrichtlinie** | `Bypass` oder `RemoteSigned` |

> **Wichtig:** PowerShell 7 (pwsh.exe) unterstützt **keinen STA-Modus** (Single-Threaded Apartment), der für WPF benötigt wird. Verwenden Sie die Windows PowerShell (powershell.exe) oder die PowerShell ISE.

## Setup

### 1. Repository klonen

```powershell
git clone https://github.com/RoccoAmmon/Word-FormatToolkit.git
cd Word-FormatToolkit
```

### 2. Ausführungsrichtlinie anpassen (falls nötig)

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
```

### 3. Skript starten

```powershell
powershell.exe -ExecutionPolicy Bypass -STA -File "WordFormatTool-GUI.ps1"
```

Alternativ: Rechtsklick auf `WordFormatTool-GUI.ps1` → „Mit PowerShell ausführen"

## Microsoft Word-Konfiguration

Das Toolkit greift über die COM-Schnittstelle auf Microsoft Word zu. Folgende Voraussetzungen müssen erfüllt sein:

- Microsoft Word muss installiert sein (jede Edition ab 2013)
- Es wird automatisch eine unsichtbare Word-Instanz im Hintergrund gestartet
- Während der Verarbeitung wird die Bildschirmaktualisierung deaktiviert (Turbo-Modus)
- Nach der Verarbeitung wird Word sauber geschlossen

## Vorlagen (.dotx/.dotm)

Für die Tabellen-Formatierung wird eine Word-Vorlage benötigt. Das Toolkit sucht automatisch in diesen Pfaden:

1. **Registry**: `HKCU:\Software\Microsoft\Office\<Version>\Word\Options\PersonalTemplates`
2. **AppData**: `%APPDATA%\Microsoft\Templates`
3. **Eigene Dokumente**: `%USERPROFILE%\Documents\Benutzerdefinierte Office-Vorlagen`
4. **Eigene Dokumente**: `%USERPROFILE%\Documents\Custom Office Templates`
5. **Program Files**: `%ProgramFiles%\Microsoft Office\Templates`
6. **Program Files (x86)**: `%ProgramFiles(x86)%\Microsoft Office\Templates`

Zusätzlich können eigene Verzeichnisse in der Konfiguration angegeben werden (`ExtraTemplateDirs`).
