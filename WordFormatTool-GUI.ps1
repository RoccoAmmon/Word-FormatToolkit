<#
.SYNOPSIS
    Word-Format-Toolkit – Batch-Formatierung und Qualitätssicherung für Microsoft Word-Dokumente.

.DESCRIPTION
    Ein leistungsstarkes PowerShell-WPF-Tool zur automatisierten Batch-Verarbeitung
    von Word-Dokumenten (.docx/.doc). Es repariert Überschriften, korrigiert
    Levelsprünge, entfernt doppelte Kapitelnummern, formatiert Tabellen nach
    Vorlagen-Styles, aktualisiert Inhaltsverzeichnisse, prüft auf tote Links und
    erkennt manuelle Nummerierungen – alles mit Live-Log und Vergleichsbericht.

    Features:
    - WPF-GUI mit Drag & Drop, Dateiauswahl und Ordner-Import
    - Hintergrundverarbeitung mit Live-Progress und Abbrechen-Button
    - Überschriften-Reparatur (Style-Neuaufsetzung)
    - Levelsprung-Korrektur (automatische Hierarchie-Fix)
    - Doppelte Kapitelnummern entfernen
    - Tabellen-Formatierung nach wählbarem Vorlagen-Style
    - Inhaltsverzeichnis-Update (TOC, Abbildungsverzeichnisse, Felder)
    - Link-Prüfer (tote Hyperlinks & Querverweise)
    - Manuelle Nummerierung erkennen
    - HTML-Vergleichsbericht (Vorher/Nachher)
    - Automatische Vorlagen-Suche in Registry + bekannten Pfaden
    - Aufräum-Routine für Logs, Reports und Backups
    - Detailliertes Live-Log im GUI-RichTextBox

.PARAMETER FunctionsOnly
    Lädt nur die Funktionen (für den Worker-Runspace), startet keine GUI.
.PARAMETER PreloadFile
    Datei, die beim Start in die Liste geladen wird (Kontextmenü).
.PARAMETER PreloadFolder
    Ordner, dessen .docx beim Start geladen werden (Kontextmenü).

.NOTES
    Autor   : Rocco Ammon
    Erstellt: 2026-06-10
    Version : 1.0.0
    Lizenz  : MIT
    Aufruf  : .\WordFormatTool-GUI.ps1
    Wiki    : https://github.com/RoccoAmmon/Word-FormatToolkit/wiki
#>

param(
    [switch]$FunctionsOnly,
    [string]$PreloadFile,
    [string]$PreloadFolder
)

# ============================================================================
# KONFIGURATION
# ============================================================================
$Global:Config = @{
    TemplatePath    = ""     # Wird dynamisch aus Find-WordTemplates ermittelt
    TableStyleName  = "Gitternetztabelle 4 – Akzent 1"
    LogFolder       = "C:\ScriptLog"
    ReportFolder    = "C:\ScriptLog\Reports"
    VerboseSteps    = $true
    ExtraTemplateDirs = @("E:\", "\\Server\Vorlagen")
}
$Global:LogFile = Join-Path $Global:Config.LogFolder ("WordFormatGUI_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

$Global:WdConst = @{
    LineStyleNone=0; LineStyleSingle=1; LineWidth050pt=4; AutoFitWindow=2
    StyleTypeTable=3; PreferredWidthPercent=2; FindStop=0; Paragraph=4; CollapseEnd=0
    BorderTop=-1; BorderLeft=-2; BorderBottom=-3; BorderRight=-4
    BorderHorizontal=-5; BorderVertical=-6; BorderDiagonalDown=-7; BorderDiagonalUp=-8
}

# Synchronisierter Zustand für Worker <-> GUI
$Global:Sync = [hashtable]::Synchronized(@{
    CancelRequested = $false
    Running         = $false
    Results         = (New-Object System.Collections.ArrayList)
    LogQueue        = (New-Object System.Collections.Concurrent.ConcurrentQueue[string])
    Progress        = 0
    ProgressText    = ""
})

# Globale Timer-/Worker-Referenzen (WICHTIG: verhindert GC des Timers!)
$Global:BatchTimer = $null
$Global:Worker     = $null

# ============================================================================
# ASSEMBLIES
# ============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ============================================================================
# LOGGING (GUI direkt ODER über Queue im Worker)
# ============================================================================
function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS","STEP")][string]$Level="INFO"
    )
    $ts = (Get-Date).ToString("HH:mm:ss")
    $entry = "[$ts] [$Level] $Message"
    try { Add-Content -Path $Global:LogFile -Value $entry -Encoding UTF8 } catch { }

    if ($null -ne $Global:UI -and $null -ne $Global:UI.LogBox) {
        try {
            $color = switch ($Level) {
                "ERROR" {"#D13438"} "WARN" {"#CA5010"} "SUCCESS" {"#107C10"}
                "STEP" {"#0078D4"} default {"#333333"}
            }
            $para = New-Object System.Windows.Documents.Paragraph
            $para.Margin = "0"
            if ($Level -eq "STEP") { $para.Margin = "15,0,0,0" }
            $run = New-Object System.Windows.Documents.Run($entry)
            $run.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($color)
            $para.Inlines.Add($run)
            $Global:UI.LogBox.Document.Blocks.Add($para)
            $Global:UI.LogBox.ScrollToEnd()
            [System.Windows.Forms.Application]::DoEvents()
        } catch { }
    }
}

# ============================================================================
# SPLASH-SCREEN
# ============================================================================
$Global:Splash = @{ Window=$null; StatusText=$null }
function Show-SplashScreen {
    param([string]$InitialText="Initialisiere...")
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Height="220" Width="460" WindowStartupLocation="CenterScreen"
        WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True"
        Background="Transparent" Topmost="True">
    <Border CornerRadius="12" Background="#FFFFFF" BorderBrush="#0078D4" BorderThickness="2">
        <Border.Effect><DropShadowEffect Color="#888888" BlurRadius="20" ShadowDepth="0" Opacity="0.5"/></Border.Effect>
        <Grid Margin="25">
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <StackPanel Grid.Row="0" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,8">
                <TextBlock Text="📄" FontSize="38" Margin="0,0,12,0"/>
                <TextBlock Text="Word-Format-Toolkit" FontSize="24" FontWeight="Bold" Foreground="#0078D4" VerticalAlignment="Center"/>
            </StackPanel>
            <TextBlock Grid.Row="1" Text="🔍 Suche nach Word-Vorlagen..." FontSize="14" FontWeight="Bold" Foreground="#333333" HorizontalAlignment="Center" Margin="0,5,0,10"/>
            <ProgressBar Grid.Row="2" Height="8" IsIndeterminate="True" Foreground="#0078D4" VerticalAlignment="Center"/>
            <TextBlock Grid.Row="3" x:Name="StatusText" Text="$InitialText" FontSize="11" Foreground="Gray" HorizontalAlignment="Center" TextWrapping="Wrap" TextAlignment="Center" Margin="0,12,0,0"/>
        </Grid>
    </Border>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $Global:Splash.Window = $window
    $Global:Splash.StatusText = $window.FindName("StatusText")
    $window.Show()
    $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
    [System.Windows.Forms.Application]::DoEvents()
}
function Update-SplashScreen {
    param([string]$Text)
    if ($null -ne $Global:Splash.Window) {
        try {
            $Global:Splash.Window.Dispatcher.Invoke([Action]{ $Global:Splash.StatusText.Text = $Text }, [System.Windows.Threading.DispatcherPriority]::Render)
            [System.Windows.Forms.Application]::DoEvents()
        } catch { }
    }
}
function Close-SplashScreen {
    if ($null -ne $Global:Splash.Window) {
        try { $Global:Splash.Window.Close() } catch { }
        $Global:Splash.Window = $null
    }
}

# ============================================================================
# STANDARD-WORD-VORLAGENORDNER ERMITTELN
# ============================================================================
function Get-WordTemplatesFolder {
    <#
    .SYNOPSIS
        Ermittelt den persönlichen Word-Vorlagenordner (PersonalTemplates)
        aus der Registry, ohne eine bestimmte Vorlage vorauszusetzen.
    .DESCRIPTION
        Liest den Pfad aus HKCU:\Software\Microsoft\Office\<Version>\Word\Options\PersonalTemplates.
        Fallback: %APPDATA%\Microsoft\Templates.
    .OUTPUTS
        string[] – Ein oder mehrere bekannte Verzeichnisse für Word-Vorlagen.
    #>
    $folders = New-Object System.Collections.ArrayList
    # 1. Persönlicher Vorlagenordner aus Registry (bevorzugt)
    try {
        foreach ($ver in @("16.0","15.0","14.0")) {
            $rp = "HKCU:\Software\Microsoft\Office\$ver\Word\Options"
            if (Test-Path $rp) {
                $pt = (Get-ItemProperty $rp -ErrorAction SilentlyContinue).PersonalTemplates
                if ($pt -and (Test-Path $pt)) { [void]$folders.Add($pt) }
            }
        }
    } catch { }
    # 2. Benutzervorlagen aus %APPDATA%
    $appDataTpl = Join-Path $env:APPDATA "Microsoft\Templates"
    if (Test-Path $appDataTpl) { [void]$folders.Add($appDataTpl) }
    # 3. "Benutzerdefinierte Office-Vorlagen" unter Eigene Dokumente
    $myDocs = [Environment]::GetFolderPath("MyDocuments")
    foreach ($sub in @("Benutzerdefinierte Office-Vorlagen","Custom Office Templates")) {
        $p = Join-Path $myDocs $sub
        if (Test-Path $p) { [void]$folders.Add($p) }
    }
    # 4. Fallback: %APPDATA%\Microsoft\Templates auch wenn nicht existent (als Text-Vorschlag)
    if ($folders.Count -eq 0) { [void]$folders.Add($appDataTpl) }
    return @($folders | Select-Object -Unique)
}

# ============================================================================
# STYLES AUS TEMPLATE AUSLESEN
# ============================================================================
function Get-TemplateStyles {
    <#
    .SYNOPSIS
        Öffnet eine Word-Vorlage und liest verfügbare Tabellen-Styles
        und Überschriften-Styles aus.
    .PARAMETER TemplatePath
        Pfad zur .dotx/.dotm-Datei.
    .OUTPUTS
        Hashtable mit Keys 'TableStyles' und 'HeadingStyles' (jeweils string[]).
    #>
    param([string]$TemplatePath)

    $result = @{
        TableStyles  = New-Object System.Collections.ArrayList
        HeadingStyles = New-Object System.Collections.ArrayList
    }
    if ([string]::IsNullOrWhiteSpace($TemplatePath) -or -not (Test-Path $TemplatePath)) {
        return $result
    }

    $word = $null; $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        # Template als Dokument öffnen (schreibgeschützt, keine Dialoge)
        $doc = $word.Documents.Open($TemplatePath, $false, $true)
        # Als "nicht geändert" markieren, damit Word kein Speichern-Dialog öffnet
        $doc.Saved = $true

        # Tabellen-Styles (StyleType = 3) – über numerischen Index iterieren (COM-sicher)
        $styleCount = $doc.Styles.Count
        Write-Log "Style-Count im Template: $styleCount" -Level STEP
        for ($i = 1; $i -le $styleCount; $i++) {
            try {
                $style = $doc.Styles.Item($i)
                if ($style.Type -eq 3) {  # wdStyleTypeTable = 3
                    [void]$result.TableStyles.Add($style.NameLocal)
                }
            } catch { }
        }

        # Überschriften-Styles: Paragraph-Styles mit OutlineLevel 1-9
        for ($i = 1; $i -le $styleCount; $i++) {
            try {
                $style = $doc.Styles.Item($i)
                if ($style.Type -eq 1) { # Paragraph-Style
                    $ol = $style.ParagraphFormat.OutlineLevel
                    if ($null -ne $ol -and [int]$ol -ge 1 -and [int]$ol -le 9) {
                        [void]$result.HeadingStyles.Add($style.NameLocal)
                    }
                }
            } catch { }
        }

        # Fallback: Falls keine gefunden, Standard-Heading-Styles eintragen
        if ($result.HeadingStyles.Count -eq 0) {
            for ($n = 1; $n -le 9; $n++) {
                foreach ($base in @("Überschrift $n","Heading $n")) {
                    try { $s = $doc.Styles.Item($base); if ($s) { [void]$result.HeadingStyles.Add($s.NameLocal); break } } catch { }
                }
            }
        }
    }
    catch {
        Write-Log "Fehler beim Auslesen der Styles aus '$TemplatePath': $($_.Exception.Message)" -Level WARN
    }
    finally {
        # 0 = wdDoNotSaveChanges – kein Speichern-Dialog, auch nicht bei Vorlagen
        try { if ($doc) { $doc.Close(0) | Out-Null } } catch { }
        try { if ($word) { $word.Quit() | Out-Null } } catch { }
        try { if ($doc) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } } catch { }
        try { if ($word) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } } catch { }
        [GC]::Collect(); [GC]::WaitForFinalizers()
    }

    $result.TableStyles  = @($result.TableStyles | Sort-Object -Unique)
    $result.HeadingStyles = @($result.HeadingStyles | Sort-Object -Unique)
    return $result
}

# ============================================================================
# VORLAGEN SUCHEN
# ============================================================================
function Find-WordTemplates {
    $searchDirs = New-Object System.Collections.ArrayList
    $knownPaths = @(
        (Join-Path $env:APPDATA "Microsoft\Templates"),
        (Join-Path $env:APPDATA "Microsoft\Word\STARTUP"),
        (Join-Path ([Environment]::GetFolderPath("MyDocuments")) "Benutzerdefinierte Office-Vorlagen"),
        (Join-Path ([Environment]::GetFolderPath("MyDocuments")) "Custom Office Templates"),
        (Join-Path ${env:ProgramFiles} "Microsoft Office\Templates"),
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft Office\Templates")
    )
    foreach ($p in $knownPaths) { if (-not [string]::IsNullOrWhiteSpace($p)) { [void]$searchDirs.Add($p) } }
    try {
        foreach ($ver in @("16.0","15.0","14.0")) {
            $rp = "HKCU:\Software\Microsoft\Office\$ver\Common\General"
            if (Test-Path $rp) { $wg=(Get-ItemProperty $rp -ErrorAction SilentlyContinue).SharedTemplates; if($wg){[void]$searchDirs.Add($wg)} }
            $rp2 = "HKCU:\Software\Microsoft\Office\$ver\Word\Options"
            if (Test-Path $rp2) { $u=(Get-ItemProperty $rp2 -ErrorAction SilentlyContinue).PersonalTemplates; if($u){[void]$searchDirs.Add($u)} }
        }
    } catch { }
    foreach ($extra in $Global:Config.ExtraTemplateDirs) { if (-not [string]::IsNullOrWhiteSpace($extra)) { [void]$searchDirs.Add($extra) } }
    $uniqueDirs = $searchDirs | Select-Object -Unique
    $found = New-Object System.Collections.ArrayList; $seenFiles = @{}
    foreach ($dir in $uniqueDirs) {
        try {
            if (Test-Path $dir) {
                Update-SplashScreen -Text "Durchsuche: $dir"
                $depth = if ($Global:Config.ExtraTemplateDirs -contains $dir) {2} else {3}
                $files = Get-ChildItem -Path $dir -Include "*.dotx","*.dotm" -File -Recurse -Depth $depth -ErrorAction SilentlyContinue |
                         Where-Object { $_.Name -notlike "~*" -and $_.Name -notlike "Normal.dot*" }
                foreach ($f in $files) {
                    if (-not $seenFiles.ContainsKey($f.FullName)) {
                        $seenFiles[$f.FullName]=$true
                        [void]$found.Add([PSCustomObject]@{Name=$f.Name;FullPath=$f.FullName;Folder=$f.DirectoryName;Modified=$f.LastWriteTime})
                    }
                }
            }
        } catch { Write-Log "Verzeichnis nicht durchsuchbar: $dir" -Level WARN }
    }
    # Stelle sicher, dass der persönliche Vorlagenordner aus der Registry durchsucht wird
    foreach ($tplDir in (Get-WordTemplatesFolder)) {
        if (-not $searchDirs.Contains($tplDir) -and (Test-Path $tplDir)) {
            try {
                $files = Get-ChildItem -Path $tplDir -Include "*.dotx","*.dotm" -File -ErrorAction SilentlyContinue |
                         Where-Object { $_.Name -notlike "~*" -and $_.Name -notlike "Normal.dot*" }
                foreach ($f in $files) {
                    if (-not $seenFiles.ContainsKey($f.FullName)) {
                        $seenFiles[$f.FullName]=$true
                        [void]$found.Add([PSCustomObject]@{Name=$f.Name;FullPath=$f.FullName;Folder=$f.DirectoryName;Modified=$f.LastWriteTime})
                    }
                }
            } catch { }
        }
    }
    $result = @($found | Sort-Object Name)
    Write-Log "Vorlagen-Suche: $($result.Count) gefunden." -Level INFO
    return $result
}

# ============================================================================
# TABELLEN-HILFSFUNKTIONEN
# ============================================================================
function Get-StyleBorderInfo {
    param($Style)
    $c = $Global:WdConst
    foreach ($id in @($c.BorderTop,$c.BorderLeft,$c.BorderBottom,$c.BorderRight,$c.BorderHorizontal,$c.BorderVertical)) {
        try { $b = $Style.Table.Borders.Item($id); if ($b.LineStyle -ne $c.LineStyleNone) { return @{Color=$b.Color;LineStyle=$b.LineStyle;LineWidth=$b.LineWidth} } } catch { }
    }
    return @{Color=-16777216;LineStyle=$c.LineStyleSingle;LineWidth=$c.LineWidth050pt}
}
function Set-TableBordersFromStyle {
    param($Table,[hashtable]$BorderInfo)
    $c = $Global:WdConst
    foreach ($id in @($c.BorderTop,$c.BorderLeft,$c.BorderBottom,$c.BorderRight,$c.BorderHorizontal,$c.BorderVertical)) {
        try { $b=$Table.Borders.Item($id); $b.LineStyle=$BorderInfo.LineStyle; $b.LineWidth=$BorderInfo.LineWidth; $b.Color=$BorderInfo.Color } catch { }
    }
}
function Clear-TableDirectFormatting {
    param($Table)
    $c = $Global:WdConst
    try { $Table.Range.Font.Reset() } catch { }
    try { foreach ($bId in @($c.BorderTop,$c.BorderLeft,$c.BorderBottom,$c.BorderRight)) { $Table.Range.Cells.Borders.Item($bId).LineStyle=$c.LineStyleNone } } catch { }
}
function Set-TableToPageWidth {
    param($Table)
    $c = $Global:WdConst
    try { $Table.AutoFitBehavior($c.AutoFitWindow); $Table.PreferredWidthType=$c.PreferredWidthPercent; $Table.PreferredWidth=100 } catch { }
}

# ============================================================================
# TURBO-MODUS
# ============================================================================
function Enable-WordTurboMode {
    param($WordApp,$Doc,[ref]$SavedOptions)
    $SavedOptions.Value=@{}
    try {
        $SavedOptions.Value["ScreenUpdating"]=$WordApp.ScreenUpdating
        $SavedOptions.Value["Pagination"]=$WordApp.Options.Pagination
        $SavedOptions.Value["CheckSpellingAsYouType"]=$WordApp.Options.CheckSpellingAsYouType
        $SavedOptions.Value["CheckGrammarAsYouType"]=$WordApp.Options.CheckGrammarAsYouType
        $SavedOptions.Value["SaveInterval"]=$WordApp.Options.SaveInterval
    } catch { }
    try { $WordApp.ScreenUpdating=$false } catch { }
    try { $WordApp.Options.Pagination=$false } catch { }
    try { $WordApp.Options.CheckSpellingAsYouType=$false } catch { }
    try { $WordApp.Options.CheckGrammarAsYouType=$false } catch { }
    try { $WordApp.Options.SaveInterval=0 } catch { }
    try { $Doc.ShowSpellingErrors=$false } catch { }
    try { $Doc.ShowGrammaticalErrors=$false } catch { }
}
function Disable-WordTurboMode {
    param($WordApp,[hashtable]$SavedOptions)
    foreach ($key in $SavedOptions.Keys) {
        try {
            switch ($key) {
                "ScreenUpdating" { $WordApp.ScreenUpdating=$SavedOptions[$key] }
                "Pagination" { $WordApp.Options.Pagination=$SavedOptions[$key] }
                "CheckSpellingAsYouType" { $WordApp.Options.CheckSpellingAsYouType=$SavedOptions[$key] }
                "CheckGrammarAsYouType" { $WordApp.Options.CheckGrammarAsYouType=$SavedOptions[$key] }
                "SaveInterval" { $WordApp.Options.SaveInterval=$SavedOptions[$key] }
            }
        } catch { }
    }
}

# ============================================================================
# DOKUMENT-STATISTIK
# ============================================================================
function Get-DocumentStats {
    param($Document)
    $c = $Global:WdConst
    $stats = [ordered]@{ Paragraphs=0;Headings=0;Tables=0;LevelJumps=0;DuplicateNum=0;ManualNum=0;DeadLinks=0;Words=0;Pages=0 }
    try { $stats.Words=$Document.Words.Count } catch { }
    try { $stats.Tables=$Document.Tables.Count } catch { }
    try { $stats.Paragraphs=$Document.Paragraphs.Count } catch { }
    try { $stats.Pages=$Document.ComputeStatistics(2) } catch { }
    $headings = New-Object System.Collections.ArrayList; $seen=@{}
    for ($n=1;$n -le 9;$n++) {
        $styleObj=$null
        foreach ($base in @("Überschrift $n","Heading $n")) { try { $s=$Document.Styles.Item($base); if($s){$styleObj=$s;break} } catch { } }
        if ($null -eq $styleObj) { continue }
        $find=$Document.Content.Find; $find.ClearFormatting(); $find.Style=$styleObj; $find.Text=""; $find.Forward=$true; $find.Wrap=$c.FindStop; $find.Format=$true
        $cont=$find.Execute(); $loop=0
        while ($cont -and $loop -lt 10000) {
            $loop++
            try {
                $r=$find.Parent.Duplicate; $r.Expand($c.Paragraph)
                if (-not $seen.ContainsKey($r.Start)) {
                    $seen[$r.Start]=$true
                    $an=""; try { $an=($r.ListFormat.ListString -replace '\s','').Trim() } catch { }
                    $headings.Add([PSCustomObject]@{Level=$n;Position=$r.Start;Text=($r.Text -replace '\r|\n','').Trim();AutoNum=$an}) | Out-Null
                }
            } catch { }
            $find.Parent.Collapse($c.CollapseEnd); $cont=$find.Execute()
        }
    }
    $stats.Headings=$headings.Count
    $prev=0; $rx='^(\d+(?:\.\d+)*\.?)[\s\t\u00A0]+'
    foreach ($h in @($headings | Sort-Object Position)) {
        if ($prev -ne 0 -and $h.Level -gt ($prev+1)) { $stats.LevelJumps++ }
        $prev=$h.Level
        if ($h.Text -match $rx) { if (-not [string]::IsNullOrWhiteSpace($h.AutoNum)) { $stats.DuplicateNum++ } else { $stats.ManualNum++ } }
    }
    return $stats
}

# ============================================================================
# ÜBERSCHRIFTEN REPARIEREN
# ============================================================================
function Repair-Headings {
    param($Document)
    $c=$Global:WdConst
    Write-Log "Überschriften-Reparatur gestartet..." -Level INFO
    $hCount=0; $names=@()
    for ($n=1;$n -le 9;$n++) { foreach ($base in @("Überschrift $n","Heading $n")) { try { $s=$Document.Styles.Item($base); if($s){$names+=$s.NameLocal;break} } catch { } } }
    $names=$names | Select-Object -Unique
    Write-Log "Suche nach Style-Vorkommen..." -Level STEP
    $toFix=New-Object System.Collections.ArrayList; $seen=@{}
    foreach ($styleName in $names) {
        try { $styleObj=$Document.Styles.Item($styleName) } catch { continue }
        $find=$Document.Content.Find; $find.ClearFormatting(); $find.Style=$styleObj; $find.Text=""; $find.Forward=$true; $find.Wrap=$c.FindStop; $find.Format=$true
        $cont=$find.Execute(); $loop=0
        while ($cont -and $loop -lt 10000) {
            $loop++
            try { $r=$find.Parent.Duplicate; $r.Expand($c.Paragraph); if(-not $seen.ContainsKey($r.Start)){$seen[$r.Start]=$true; $toFix.Add(@{Range=$r;StyleName=$styleName})|Out-Null} } catch { }
            $find.Parent.Collapse($c.CollapseEnd); $cont=$find.Execute()
        }
    }
    Write-Log "$($toFix.Count) Überschriften gefunden - beginne Reparatur..." -Level STEP
    $standard=$Document.Styles.Item("Standard"); $total=$toFix.Count
    foreach ($e in $toFix) {
        try {
            $hCount++; $r=$e.Range
            if ($Global:Config.VerboseSteps) { $p=($r.Text -replace '\r|\n','').Trim(); if($p.Length -gt 45){$p=$p.Substring(0,45)+"..."}; Write-Log "[$hCount/$total] Repariere: $p" -Level STEP }
            $orig=$Document.Styles.Item($e.StyleName)
            $r.Style=$standard
            try { $r.Font.Reset(); $r.ParagraphFormat.Reset() } catch { }
            $r.Style=$orig
        } catch { }
    }
    Write-Log "Überschriften repariert: $hCount" -Level SUCCESS
    return $hCount
}

# ============================================================================
# LEVELSPRÜNGE KORRIGIEREN
# ============================================================================
function Repair-HeadingLevels {
    param($Document,[int]$MaxPasses=3)
    $c=$Global:WdConst
    Write-Log "Levelsprung-Prüfung gestartet..." -Level INFO
    $totalFixed=0
    $getHeadings={
        param($Doc)
        $list=New-Object System.Collections.ArrayList; $seen=@{}; $styles=@()
        for ($n=1;$n -le 9;$n++) { foreach ($base in @("Überschrift $n","Heading $n")) { try { $s=$Doc.Styles.Item($base); if($s){$styles+=[PSCustomObject]@{Level=$n;StyleObj=$s};break} } catch { } } }
        foreach ($entry in $styles) {
            $find=$Doc.Content.Find; $find.ClearFormatting(); $find.Style=$entry.StyleObj; $find.Text=""; $find.Forward=$true; $find.Wrap=$c.FindStop; $find.Format=$true
            $cont=$find.Execute(); $loop=0
            while ($cont -and $loop -lt 10000) {
                $loop++
                try { $r=$find.Parent.Duplicate; $r.Expand($c.Paragraph); if(-not $seen.ContainsKey($r.Start)){$seen[$r.Start]=$true; $list.Add([PSCustomObject]@{Range=$r;Level=$entry.Level;Position=$r.Start;StyleNameLocal=$entry.StyleObj.NameLocal})|Out-Null} } catch { }
                $find.Parent.Collapse($c.CollapseEnd); $cont=$find.Execute()
            }
        }
        return @($list | Sort-Object Position)
    }
    for ($pass=1;$pass -le $MaxPasses;$pass++) {
        Write-Log "Durchgang $pass von $MaxPasses..." -Level STEP
        $headings=& $getHeadings $Document
        if ($headings.Count -eq 0) { return 0 }
        $prev=0; $passFixed=0
        foreach ($h in $headings) {
            $cur=$h.Level
            if ($prev -eq 0) { $prev=$cur; continue }
            $maxAllowed=$prev+1
            if ($cur -gt $maxAllowed) {
                $newLevel=$maxAllowed
                $newName=if($h.StyleNameLocal -like "Heading*"){"Heading $newLevel"}else{"Überschrift $newLevel"}
                try {
                    $p=($h.Range.Text -replace '\r|\n','').Trim(); if($p.Length -gt 40){$p=$p.Substring(0,40)+"..."}
                    $h.Range.Style=$Document.Styles.Item($newName)
                    $passFixed++; $totalFixed++; $prev=$newLevel
                    if ($Global:Config.VerboseSteps) { Write-Log "Korrigiert L$cur->L$newLevel : $p" -Level STEP }
                } catch { $prev=$cur }
            } else { $prev=$cur }
        }
        Write-Log "Durchgang $pass : $passFixed Korrektur(en)" -Level STEP
        if ($passFixed -eq 0) { Write-Log "Hierarchie stabil." -Level STEP; break }
    }
    Write-Log "Aktualisiere Listen-Nummerierung..." -Level STEP
    try { $Document.Fields.Update()|Out-Null; foreach ($l in $Document.Lists){ try{$l.Range.Fields.Update()|Out-Null}catch{} } } catch { }
    Write-Log "Levelsprünge korrigiert: $totalFixed" -Level SUCCESS
    return $totalFixed
}

# ============================================================================
# DOPPELTE KAPITELNUMMERN ENTFERNEN
# ============================================================================
function Remove-DuplicateHeadingNumbers {
    param($Document,[bool]$OnlyIfAutoNumbered=$true)
    $c=$Global:WdConst
    Write-Log "Suche nach doppelten Kapitelnummern..." -Level INFO
    $headings=New-Object System.Collections.ArrayList; $seen=@{}
    for ($n=1;$n -le 9;$n++) {
        $styleObj=$null
        foreach ($base in @("Überschrift $n","Heading $n")) { try { $s=$Document.Styles.Item($base); if($s){$styleObj=$s;break} } catch { } }
        if ($null -eq $styleObj) { continue }
        $find=$Document.Content.Find; $find.ClearFormatting(); $find.Style=$styleObj; $find.Text=""; $find.Forward=$true; $find.Wrap=$c.FindStop; $find.Format=$true
        $cont=$find.Execute(); $loop=0
        while ($cont -and $loop -lt 10000) {
            $loop++
            try { $r=$find.Parent.Duplicate; $r.Expand($c.Paragraph); if(-not $seen.ContainsKey($r.Start)){$seen[$r.Start]=$true; $headings.Add([PSCustomObject]@{Range=$r;Level=$n})|Out-Null} } catch { }
            $find.Parent.Collapse($c.CollapseEnd); $cont=$find.Execute()
        }
    }
    if ($headings.Count -eq 0) { return 0 }
    Write-Log "$($headings.Count) Überschriften werden geprüft..." -Level STEP
    $sortedDesc=@($headings | Sort-Object -Property @{Expression={$_.Range.Start};Descending=$true})
    $rx='^(\d+(?:\.\d+)*\.?)[\s\t\u00A0]+(.+)$'; $removed=0
    foreach ($h in $sortedDesc) {
        try {
            $r=$h.Range; $text=$r.Text -replace '\r|\n',''
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            $an=""; try { $an=($r.ListFormat.ListString -replace '\s','').Trim() } catch { }
            if ($text -notmatch $rx) { continue }
            if ($OnlyIfAutoNumbered -and [string]::IsNullOrWhiteSpace($an)) { continue }
            $tn=$matches[1]
            $pm=[regex]::Match($text,'^(\d+(?:\.\d+)*\.?)[\s\t\u00A0]+'); if(-not $pm.Success){continue}
            $endPos=$r.Start+$pm.Length; if($endPos -gt $r.End){continue}
            ($Document.Range($r.Start,$endPos)).Text=""
            $removed++
            if ($Global:Config.VerboseSteps) { $rest=($text -replace $rx,'$2'); if($rest.Length -gt 35){$rest=$rest.Substring(0,35)+"..."}; Write-Log "Entfernt '$tn' aus: $rest" -Level STEP }
        } catch { }
    }
    Write-Log "Doppelte Kapitelnummern entfernt: $removed" -Level SUCCESS
    return $removed
}

# ============================================================================
# TABELLEN FORMATIEREN
# ============================================================================
function Format-Tables {
    param($Document)
    $c=$Global:WdConst
    Write-Log "Tabellen-Formatierung gestartet..." -Level INFO
    if (-not [string]::IsNullOrWhiteSpace($Global:Config.TemplatePath) -and (Test-Path $Global:Config.TemplatePath)) {
        Write-Log "Hänge Vorlage an und aktualisiere Styles..." -Level STEP
        $Document.AttachedTemplate=$Global:Config.TemplatePath
        $Document.UpdateStyles()
    } else {
        Write-Log "Keine Vorlage gesetzt – verwende dokumenteigene Styles." -Level INFO
    }
    $tableStyle=$null
    try { $tableStyle=$Document.Styles.Item($Global:Config.TableStyleName) } catch { throw "Tabellenstyle '$($Global:Config.TableStyleName)' fehlt in der Vorlage." }
    $borderInfo=Get-StyleBorderInfo -Style $tableStyle
    $tableCount=$Document.Tables.Count
    Write-Log "$tableCount Tabellen gefunden - beginne Formatierung..." -Level STEP
    $tCount=0
    for ($i=1;$i -le $tableCount;$i++) {
        try {
            $table=$Document.Tables.Item($i); $tCount++
            if ($Global:Config.VerboseSteps) { Write-Log "[$tCount/$tableCount] Formatiere Tabelle ($($table.Rows.Count)x$($table.Columns.Count))..." -Level STEP }
            try { foreach ($bId in @($c.BorderTop,$c.BorderLeft,$c.BorderBottom,$c.BorderRight,$c.BorderHorizontal,$c.BorderVertical,$c.BorderDiagonalDown,$c.BorderDiagonalUp)){$table.Borders.Item($bId).LineStyle=$c.LineStyleNone} } catch { }
            Clear-TableDirectFormatting -Table $table
            $table.Style=$Global:Config.TableStyleName
            $table.ApplyStyleHeadingRows=$true; $table.ApplyStyleLastRow=$false; $table.ApplyStyleFirstColumn=$true
            $table.ApplyStyleLastColumn=$false; $table.ApplyStyleRowBands=$true; $table.ApplyStyleColumnBands=$false
            Set-TableBordersFromStyle -Table $table -BorderInfo $borderInfo
            Set-TableToPageWidth -Table $table
        } catch { Write-Log "Fehler bei Tabelle #$i : $($_.Exception.Message)" -Level WARN }
    }
    Write-Log "Tabellen formatiert: $tCount" -Level SUCCESS
    return $tCount
}

# ============================================================================
# INHALTSVERZEICHNIS AKTUALISIEREN
# ============================================================================
function Update-DocumentTOC {
    param($Document)
    Write-Log "Aktualisiere Verzeichnisse..." -Level INFO
    $count=0
    Write-Log "Berechne Seitenumbrüche neu (Repaginate)..." -Level STEP
    try { $Document.Application.Options.Pagination=$true } catch { }
    try { $Document.Repaginate() } catch { }
    Write-Log "Aktualisiere Inhaltsverzeichnis(se)..." -Level STEP
    try { foreach ($toc in $Document.TablesOfContents){ try{$toc.Update();$toc.UpdatePageNumbers();$count++}catch{} } } catch { }
    Write-Log "Aktualisiere Abbildungs-/Tabellenverzeichnisse..." -Level STEP
    try { foreach ($tof in $Document.TablesOfFigures){ try{$tof.Update();$count++}catch{} } } catch { }
    Write-Log "Aktualisiere alle Felder + Kopf-/Fußzeilen..." -Level STEP
    try { $Document.Fields.Update()|Out-Null } catch { }
    foreach ($section in $Document.Sections) {
        try {
            foreach ($h in $section.Headers){ try{$h.Range.Fields.Update()|Out-Null}catch{} }
            foreach ($f in $section.Footers){ try{$f.Range.Fields.Update()|Out-Null}catch{} }
        } catch { }
    }
    try { $Document.Application.Options.Pagination=$false } catch { }
    Write-Log "Verzeichnisse aktualisiert: $count" -Level SUCCESS
    return $count
}

# ============================================================================
# TOTE LINKS PRÜFEN
# ============================================================================
function Test-DeadLinks {
    param($Document)
    Write-Log "Prüfe Hyperlinks und Querverweise..." -Level INFO
    $dead=New-Object System.Collections.ArrayList
    $bookmarks=@{}
    foreach ($bm in $Document.Bookmarks) { try { $bookmarks[$bm.Name]=$true } catch { } }
    $hCount=0
    foreach ($link in $Document.Hyperlinks) {
        $hCount++
        try {
            $sub=$link.SubAddress; $addr=$link.Address; $ttd=""
            try { $ttd=$link.TextToDisplay } catch { }
            if (-not [string]::IsNullOrWhiteSpace($sub)) {
                if (-not $bookmarks.ContainsKey($sub)) {
                    [void]$dead.Add([PSCustomObject]@{Typ="Internes Lesezeichen fehlt";Ziel=$sub;Text=$ttd})
                    if ($Global:Config.VerboseSteps) { Write-Log "TOT: Lesezeichen '$sub' fehlt (Text: $ttd)" -Level WARN }
                }
            } elseif (-not [string]::IsNullOrWhiteSpace($addr) -and $addr -notmatch '^(https?|mailto|ftp):' -and $addr -notmatch '^#') {
                if ($addr -match '^[a-zA-Z]:\\' -or $addr -match '^\\\\') {
                    if (-not (Test-Path $addr)) {
                        [void]$dead.Add([PSCustomObject]@{Typ="Verknüpfte Datei fehlt";Ziel=$addr;Text=$ttd})
                        if ($Global:Config.VerboseSteps) { Write-Log "TOT: Datei '$addr' fehlt" -Level WARN }
                    }
                }
            }
        } catch { }
    }
    foreach ($field in $Document.Fields) {
        try {
            if ($field.Type -eq 3) {
                $code=($field.Code.Text).Trim()
                if ($code -match 'REF\s+(\S+)') {
                    $rt=$matches[1]
                    if (-not $bookmarks.ContainsKey($rt)) {
                        [void]$dead.Add([PSCustomObject]@{Typ="Querverweis-Ziel fehlt";Ziel=$rt;Text=$code})
                        if ($Global:Config.VerboseSteps) { Write-Log "TOT: Querverweis auf '$rt' fehlt" -Level WARN }
                    }
                }
            }
        } catch { }
    }
    Write-Log "Link-Prüfung: $hCount Hyperlinks, $($dead.Count) Problem(e)." -Level SUCCESS
    return $dead
}

# ============================================================================
# MANUELLE NUMMERIERUNG ERKENNEN
# ============================================================================
function Test-ManualNumbering {
    param($Document)
    $c=$Global:WdConst
    Write-Log "Prüfe auf manuelle Nummerierung..." -Level INFO
    $manual=New-Object System.Collections.ArrayList; $seen=@{}; $rx='^(\d+(?:\.\d+)*\.?)[\s\t\u00A0]+'
    for ($n=1;$n -le 9;$n++) {
        $styleObj=$null
        foreach ($base in @("Überschrift $n","Heading $n")) { try { $s=$Document.Styles.Item($base); if($s){$styleObj=$s;break} } catch { } }
        if ($null -eq $styleObj) { continue }
        $find=$Document.Content.Find; $find.ClearFormatting(); $find.Style=$styleObj; $find.Text=""; $find.Forward=$true; $find.Wrap=$c.FindStop; $find.Format=$true
        $cont=$find.Execute(); $loop=0
        while ($cont -and $loop -lt 10000) {
            $loop++
            try {
                $r=$find.Parent.Duplicate; $r.Expand($c.Paragraph)
                if (-not $seen.ContainsKey($r.Start)) {
                    $seen[$r.Start]=$true
                    $text=($r.Text -replace '\r|\n','').Trim()
                    $an=""; try { $an=($r.ListFormat.ListString -replace '\s','').Trim() } catch { }
                    if ($text -match $rx -and [string]::IsNullOrWhiteSpace($an)) {
                        $p=$text; if($p.Length -gt 55){$p=$p.Substring(0,55)+"..."}
                        [void]$manual.Add([PSCustomObject]@{Level=$n;TextNummer=$matches[1];Text=$p})
                        if ($Global:Config.VerboseSteps) { Write-Log "MANUELL (L$n): $p" -Level WARN }
                    }
                }
            } catch { }
            $find.Parent.Collapse($c.CollapseEnd); $cont=$find.Execute()
        }
    }
    Write-Log "Manuelle Nummerierung: $($manual.Count) gefunden." -Level SUCCESS
    return $manual
}

# ============================================================================
# AUFRÄUM-ROUTINE
# ============================================================================
function Invoke-Cleanup {
    param([int]$LogDays=30,[int]$ReportDays=90,[int]$BackupDays=14,[string[]]$BackupSearchDirs=@())
    $deleted=0
    Write-Log "Aufräumen (Logs>$LogDays Tg, Reports>$ReportDays Tg, Backups>$BackupDays Tg)..." -Level INFO
    try { Get-ChildItem -Path $Global:Config.LogFolder -Filter "WordFormat*.log" -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogDays) } | ForEach-Object { Remove-Item $_.FullName -Force; $deleted++ } } catch { }
    try { if (Test-Path $Global:Config.ReportFolder) { Get-ChildItem -Path $Global:Config.ReportFolder -Filter "Vergleichsbericht_*.html" -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$ReportDays) } | ForEach-Object { Remove-Item $_.FullName -Force; $deleted++ } } } catch { }
    foreach ($dir in $BackupSearchDirs) {
        try { if (Test-Path $dir) { Get-ChildItem -Path $dir -Filter "*_Backup_*.doc*" -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$BackupDays) } | ForEach-Object { Remove-Item $_.FullName -Force; $deleted++ } } } catch { }
    }
    Write-Log "Aufräumen abgeschlossen: $deleted Datei(en) gelöscht." -Level SUCCESS
    return $deleted
}

# ============================================================================
# EINZELDOKUMENT VERARBEITEN
# ============================================================================
function Invoke-ProcessDocument {
    param([string]$DocPath,[hashtable]$Actions)
    $result=[ordered]@{
        File=$DocPath; FileName=[System.IO.Path]::GetFileName($DocPath)
        Success=$false; Error=""; BackupPath=""
        StatsBefore=$null; StatsAfter=$null
        Headings=0; Levels=0; Duplicates=0; Tables=0; TOC=0
        DeadLinks=0; ManualNum=0; DeadLinkList=@(); ManualNumList=@(); Duration=$null
    }
    $word=$null; $document=$null; $savedOptions=@{}; $startTime=Get-Date
    try {
        Write-Log "===== Verarbeite: $($result.FileName) =====" -Level INFO
        if (-not (Test-Path $DocPath)) { throw "Datei nicht gefunden." }
        $suffix="_Backup_{0:yyyyMMdd_HHmmss}" -f (Get-Date)
        $backup=[System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($DocPath),([System.IO.Path]::GetFileNameWithoutExtension($DocPath)+$suffix+[System.IO.Path]::GetExtension($DocPath)))
        Copy-Item -Path $DocPath -Destination $backup -Force
        $result.BackupPath=$backup
        Write-Log "Backup erstellt: $([System.IO.Path]::GetFileName($backup))" -Level SUCCESS
        Write-Log "Starte Word im Hintergrund..." -Level STEP
        $word=New-Object -ComObject Word.Application; $word.Visible=$false; $word.DisplayAlerts=0
        Write-Log "Öffne Dokument..." -Level STEP
        $document=$word.Documents.Open($DocPath)
        Enable-WordTurboMode -WordApp $word -Doc $document -SavedOptions ([ref]$savedOptions)
        Write-Log "Analysiere Dokument (vorher)..." -Level STEP
        $result.StatsBefore=Get-DocumentStats -Document $document
        Write-Log ("Vorher: Headings={0}, Levelsprünge={1}, Duplikate={2}, Manuell={3}, Tabellen={4}" -f $result.StatsBefore.Headings,$result.StatsBefore.LevelJumps,$result.StatsBefore.DuplicateNum,$result.StatsBefore.ManualNum,$result.StatsBefore.Tables) -Level INFO
        if ($Actions.Headings)   { $result.Headings   = Repair-Headings -Document $document }
        if ($Actions.Levels)     { $result.Levels     = Repair-HeadingLevels -Document $document }
        if ($Actions.Duplicates) { $result.Duplicates = Remove-DuplicateHeadingNumbers -Document $document }
        if ($Actions.ManualNum)  { $mn = Test-ManualNumbering -Document $document; $result.ManualNum = $mn.Count; $result.ManualNumList = @($mn) }
        if ($Actions.DeadLinks)  { $dl = Test-DeadLinks -Document $document; $result.DeadLinks = $dl.Count; $result.DeadLinkList = @($dl) }
        if ($Actions.Tables)     { $result.Tables     = Format-Tables -Document $document }
        if ($Actions.TOC)        { $result.TOC        = Update-DocumentTOC -Document $document }
        Write-Log "Analysiere Dokument (nachher)..." -Level STEP
        $result.StatsAfter=Get-DocumentStats -Document $document
        Write-Log "Speichere Dokument..." -Level STEP
        $document.Save()
        $result.Success=$true
        Write-Log "Fertig: $($result.FileName)" -Level SUCCESS
    }
    catch {
        $result.Error=$_.Exception.Message
        Write-Log "FEHLER bei $($result.FileName): $($_.Exception.Message)" -Level ERROR
    }
    finally {
        if ($null -ne $word) { try { Disable-WordTurboMode -WordApp $word -SavedOptions $savedOptions } catch { } }
        try { if ($null -ne $document) { $document.Close()|Out-Null } } catch { }
        try { if ($null -ne $word) { $word.Quit()|Out-Null } } catch { }
        if ($null -ne $document) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)|Out-Null }
        if ($null -ne $word)     { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        $result.Duration=(Get-Date)-$startTime
    }
    Write-Output ([PSCustomObject]$result)
}

# ============================================================================
# VERGLEICHSBERICHT (HTML)
# ============================================================================
function New-ComparisonReport {
    param([array]$Results)
    if (-not (Test-Path $Global:Config.ReportFolder)) { New-Item -Path $Global:Config.ReportFolder -ItemType Directory -Force | Out-Null }
    $reportPath=Join-Path $Global:Config.ReportFolder ("Vergleichsbericht_{0:yyyyMMdd_HHmmss}.html" -f (Get-Date))
    $validResults=@($Results | Where-Object { $_ -is [System.Management.Automation.PSCustomObject] -and $_.PSObject.Properties.Name -contains 'Success' })
    $rows=""
    $totalDocs=$validResults.Count
    $okDocs=@($validResults | Where-Object { $_.Success -eq $true }).Count
    $errDocs=@($validResults | Where-Object { $_.Success -ne $true }).Count
    foreach ($r in $validResults) {
        $badge=if($r.Success){"<span class='badge ok'>OK</span>"}else{"<span class='badge err'>FEHLER</span>"}
        if ($r.StatsBefore -and $r.StatsAfter) {
            $b=$r.StatsBefore; $a=$r.StatsAfter
            $rows += @"
        <tr>
            <td>$($r.FileName)</td><td>$badge</td>
            <td class='num'>$($b.Headings)</td><td class='num'>$($b.Tables)</td>
            <td class='num'><span class='$( if($b.LevelJumps -gt 0){"warn-text"} )'>$($b.LevelJumps)</span> &rarr; <span class='ok-text'>$($a.LevelJumps)</span></td>
            <td class='num'><span class='$( if($b.DuplicateNum -gt 0){"warn-text"} )'>$($b.DuplicateNum)</span> &rarr; <span class='ok-text'>$($a.DuplicateNum)</span></td>
            <td class='num'><span class='$( if($r.ManualNum -gt 0){"warn-text"} )'>$($r.ManualNum)</span></td>
            <td class='num'><span class='$( if($r.DeadLinks -gt 0){"err-text"} )'>$($r.DeadLinks)</span></td>
            <td class='num'>$($r.Tables)</td><td class='num'>$($r.TOC)</td>
            <td>$("{0:hh\:mm\:ss}" -f $r.Duration)</td>
        </tr>
"@
        } else {
            $rows += @"
        <tr><td>$($r.FileName)</td><td>$badge</td><td colspan='8' class='err-text'>$($r.Error)</td><td>$("{0:hh\:mm\:ss}" -f $r.Duration)</td></tr>
"@
        }
    }
    $html = @"
<!DOCTYPE html><html lang='de'><head><meta charset='UTF-8'><title>Word-Format-Vergleichsbericht</title>
<style>
body{font-family:'Segoe UI',Tahoma,sans-serif;margin:30px;color:#1a1a1a;background:#f5f5f5}
h1{color:#0078D4}
.summary{display:flex;gap:20px;margin:20px 0}
.card{background:white;border-radius:8px;padding:20px;box-shadow:0 2px 6px rgba(0,0,0,0.1);flex:1}
.card .big{font-size:32px;font-weight:bold}
.card.total .big{color:#0078D4}.card.ok .big{color:#107C10}.card.err .big{color:#D13438}
table{width:100%;border-collapse:collapse;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.1)}
th{background:#0078D4;color:white;padding:10px 6px;text-align:left;font-size:12px}
td{padding:9px 6px;border-bottom:1px solid #eee;font-size:12px}
tr:hover{background:#f0f8ff}
.num{text-align:center}
.badge{padding:3px 10px;border-radius:12px;color:white;font-size:11px;font-weight:bold}
.badge.ok{background:#107C10}.badge.err{background:#D13438}
.ok-text{color:#107C10;font-weight:bold}.warn-text{color:#CA5010;font-weight:bold}.err-text{color:#D13438;font-weight:bold}
.footer{margin-top:20px;color:#888;font-size:12px}
</style></head><body>
<h1>📊 Word-Format-Vergleichsbericht</h1>
<p>Erstellt am $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')</p>
<p>Verwendete Vorlage: <b>$($Global:Config.TemplatePath)</b></p>
<div class='summary'>
<div class='card total'><div>Dokumente gesamt</div><div class='big'>$totalDocs</div></div>
<div class='card ok'><div>Erfolgreich</div><div class='big'>$okDocs</div></div>
<div class='card err'><div>Fehler</div><div class='big'>$errDocs</div></div>
</div>
<table><thead><tr>
<th>Datei</th><th>Status</th><th>Überschr.</th><th>Tabellen</th>
<th>Levelsprünge<br>(vor&rarr;nach)</th><th>Duplikate<br>(vor&rarr;nach)</th>
<th>Manuell<br>nummeriert</th><th>Tote<br>Links</th>
<th>Tab.<br>format.</th><th>TOC</th><th>Dauer</th>
</tr></thead><tbody>
$rows
</tbody></table>
<div class='footer'>Word-Format-Toolkit | Logfile: $Global:LogFile</div>
</body></html>
"@
    Set-Content -Path $reportPath -Value $html -Encoding UTF8
    Write-Log "Vergleichsbericht erstellt: $reportPath" -Level SUCCESS
    return $reportPath
}

# ============================================================================
# BATCH IM HINTERGRUND-RUNSPACE STARTEN
# ============================================================================

# ============================================================================
# GUI
# ============================================================================
function Show-MainGUI {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Word-Format-Toolkit" Height="860" Width="940"
        WindowStartupLocation="CenterScreen" Background="#F3F3F3">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,15">
            <TextBlock Text="📄" FontSize="32" Margin="0,0,10,0"/>
            <StackPanel>
                <TextBlock Text="Word-Format-Toolkit" FontSize="24" FontWeight="Bold" Foreground="#0078D4"/>
                <TextBlock Text="Batch + Runspace + Link-Prüfer + Aufräumen + Vergleichsbericht" FontSize="13" Foreground="Gray"/>
            </StackPanel>
        </StackPanel>
        <Grid Grid.Row="1" Margin="0,0,0,10">
            <Grid.ColumnDefinitions><ColumnDefinition Width="2*"/><ColumnDefinition Width="1*"/></Grid.ColumnDefinitions>
            <Border Grid.Column="0" Background="White" CornerRadius="8" Padding="12" Margin="0,0,8,0">
                <Grid>
                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="180"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Text="📁 Zu verarbeitende Dokumente" FontWeight="Bold" Margin="0,0,0,8"/>
                    <ListBox Grid.Row="1" x:Name="FileList" SelectionMode="Extended" BorderBrush="#DDD" BorderThickness="1"/>
                    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,0">
                        <Button x:Name="btnAddFiles" Content="➕ Dateien" Width="90" Height="30" Margin="0,0,5,0"/>
                        <Button x:Name="btnAddFolder" Content="📂 Ordner" Width="90" Height="30" Margin="0,0,5,0"/>
                        <Button x:Name="btnRemove" Content="➖ Entfernen" Width="90" Height="30" Margin="0,0,5,0"/>
                        <Button x:Name="btnClear" Content="🗑️ Leeren" Width="80" Height="30"/>
                    </StackPanel>
                </Grid>
            </Border>
            <Border Grid.Column="1" Background="White" CornerRadius="8" Padding="12" Margin="8,0,0,0">
                <StackPanel>
                    <TextBlock Text="📋 Vorlage (.dotx/.dotm)" FontWeight="Bold" Margin="0,0,0,4"/>
                    <ComboBox x:Name="cmbTemplate" Height="28" Margin="0,0,0,4" IsEditable="True"/>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
                        <Button x:Name="btnRescan" Content="🔄 Neu suchen" Width="105" Height="26" Margin="0,0,5,0"/>
                        <Button x:Name="btnBrowseTemplate" Content="📂 Datei..." Width="80" Height="26"/>
                    </StackPanel>
                    <Separator Margin="0,0,0,6"/>
                    <TextBlock Text="📊 Tabellen-Style" FontWeight="Bold" Margin="0,0,0,2"/>
                    <ComboBox x:Name="cmbTableStyle" Height="26" Margin="0,0,0,6" IsEditable="True"/>
                    <Separator Margin="0,0,0,6"/>
                    <TextBlock Text="⚙️ Aktionen" FontWeight="Bold" Margin="0,0,0,6"/>
                    <CheckBox x:Name="chkHeadings"   Content="📝 Überschriften reparieren" IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="chkLevels"     Content="🔢 Levelsprünge korrigieren" IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="chkDuplicates" Content="🔁 Doppelte Nummern entfernen" IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="chkManualNum"  Content="✏️ Manuelle Nummerierung finden" Margin="0,3"/>
                    <CheckBox x:Name="chkDeadLinks"  Content="🔗 Tote Links prüfen" Margin="0,3"/>
                    <CheckBox x:Name="chkTables"     Content="📊 Tabellen formatieren" IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="chkTOC"        Content="📑 Inhaltsverzeichnis updaten" IsChecked="True" Margin="0,3"/>
                    <Separator Margin="0,6"/>
                    <CheckBox x:Name="chkReport"     Content="📈 Vergleichsbericht" IsChecked="True" Margin="0,3"/>
                    <CheckBox x:Name="chkVerbose"    Content="🔍 Detaillierte Schritte" IsChecked="True" Margin="0,3"/>
                    <Separator Margin="0,6"/>
                    <TextBlock Text="🗂️ Aufräumen (Tage)" FontWeight="Bold" Margin="0,0,0,4"/>
                    <Grid>
                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0"><TextBlock Text="Logs" FontSize="10" Foreground="Gray"/><TextBox x:Name="txtCleanLogs" Text="30" Height="24" HorizontalContentAlignment="Center"/></StackPanel>
                        <StackPanel Grid.Column="1"><TextBlock Text="Reports" FontSize="10" Foreground="Gray"/><TextBox x:Name="txtCleanReports" Text="90" Height="24" HorizontalContentAlignment="Center"/></StackPanel>
                        <StackPanel Grid.Column="2"><TextBlock Text="Backups" FontSize="10" Foreground="Gray"/><TextBox x:Name="txtCleanBackups" Text="14" Height="24" HorizontalContentAlignment="Center"/></StackPanel>
                    </Grid>
                    <Button x:Name="btnCleanup" Content="🧹 Jetzt aufräumen" Height="28" Margin="0,6,0,0"/>
                </StackPanel>
            </Border>
        </Grid>
        <Border Grid.Row="2" Background="White" CornerRadius="8" Padding="12">
            <Grid>
                <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                <TextBlock Grid.Row="0" Text="📋 Verarbeitungs-Log (live)" FontWeight="Bold" Margin="0,0,0,8"/>
                <RichTextBox Grid.Row="1" x:Name="LogBox" IsReadOnly="True" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="11" BorderBrush="#DDD" BorderThickness="1"/>
            </Grid>
        </Border>
        <Grid Grid.Row="3" Margin="0,10,0,0">
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <TextBlock Grid.Row="0" x:Name="ProgressText" Text="Bereit." Margin="0,0,0,4" Foreground="#555"/>
            <ProgressBar Grid.Row="1" x:Name="ProgressBar" Height="20" Minimum="0" Maximum="100" Value="0"/>
        </Grid>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnReport" Content="📊 Letzten Bericht öffnen" Width="170" Height="36" Margin="0,0,10,0" IsEnabled="False"/>
            <Button x:Name="btnCancel" Content="⏹️ Abbrechen" Width="120" Height="36" Margin="0,0,10,0" Background="#D13438" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
            <Button x:Name="btnStart" Content="🚀 Verarbeitung starten" Width="180" Height="36" Background="#0078D4" Foreground="White" FontWeight="Bold" FontSize="14"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader=New-Object System.Xml.XmlNodeReader $xaml
    $window=[Windows.Markup.XamlReader]::Load($reader)
    $Global:UI=@{
        Window=$window; FileList=$window.FindName("FileList"); LogBox=$window.FindName("LogBox")
        ProgressBar=$window.FindName("ProgressBar"); ProgressText=$window.FindName("ProgressText")
        cmbTemplate=$window.FindName("cmbTemplate")
        cmbTableStyle=$window.FindName("cmbTableStyle")
        chkHeadings=$window.FindName("chkHeadings"); chkLevels=$window.FindName("chkLevels")
        chkDuplicates=$window.FindName("chkDuplicates"); chkManualNum=$window.FindName("chkManualNum")
        chkDeadLinks=$window.FindName("chkDeadLinks"); chkTables=$window.FindName("chkTables")
        chkTOC=$window.FindName("chkTOC"); chkReport=$window.FindName("chkReport"); chkVerbose=$window.FindName("chkVerbose")
        txtCleanLogs=$window.FindName("txtCleanLogs"); txtCleanReports=$window.FindName("txtCleanReports"); txtCleanBackups=$window.FindName("txtCleanBackups")
        btnStart=$window.FindName("btnStart"); btnReport=$window.FindName("btnReport"); btnCancel=$window.FindName("btnCancel")
    }
    $btnAddFiles=$window.FindName("btnAddFiles"); $btnAddFolder=$window.FindName("btnAddFolder")
    $btnRemove=$window.FindName("btnRemove"); $btnClear=$window.FindName("btnClear")
    $btnRescan=$window.FindName("btnRescan"); $btnBrowseTemplate=$window.FindName("btnBrowseTemplate")
    $btnCleanup=$window.FindName("btnCleanup")
    $Global:LastReportPath=$null

    $fillTemplates={
        $Global:UI.cmbTemplate.Items.Clear()
        foreach ($t in (Find-WordTemplates)) {
            $item=New-Object System.Windows.Controls.ComboBoxItem
            $item.Content="$($t.Name)   —   $($t.Folder)"; $item.Tag=$t.FullPath; $item.ToolTip=$t.FullPath
            [void]$Global:UI.cmbTemplate.Items.Add($item)
        }
        if ($Global:UI.cmbTemplate.Items.Count -gt 0) {
            # Vorherige Auswahl merken (z.B. nach Rescan)
            $pre=$null
            if (-not [string]::IsNullOrWhiteSpace($Global:Config.TemplatePath)) {
                foreach ($it in $Global:UI.cmbTemplate.Items) {
                    if ($it.Tag -ieq $Global:Config.TemplatePath) { $pre=$it; break }
                }
            }
            if ($pre) { $Global:UI.cmbTemplate.SelectedItem=$pre }
            else { $Global:UI.cmbTemplate.SelectedIndex=0 }
        } else {
            # Keine Vorlage gefunden – Standard-Ordner als Text vorschlagen
            $defaultFolders = Get-WordTemplatesFolder
            $Global:UI.cmbTemplate.Text = if ($defaultFolders.Count -gt 0) { $defaultFolders[0] } else { "" }
            $Global:UI.cmbTemplate.ToolTip = "Keine Vorlagen gefunden. Bitte Pfad eingeben oder durchsuchen."
        }
    }

    # Styles aus dem aktuell ausgewählten Template laden
    $getSelectedTemplate={
        $sel=$Global:UI.cmbTemplate.SelectedItem
        if ($null -ne $sel -and $null -ne $sel.Tag) { return [string]$sel.Tag }
        return ([string]$Global:UI.cmbTemplate.Text).Trim()
    }

    $loadTemplateStyles = {
        $tplPath = & $getSelectedTemplate
        if ([string]::IsNullOrWhiteSpace($tplPath) -or -not (Test-Path $tplPath)) {
            $Global:UI.cmbTableStyle.Items.Clear()
            return
        }
        Write-Log "Lade Styles aus: $([System.IO.Path]::GetFileName($tplPath))" -Level STEP
        $styles = Get-TemplateStyles -TemplatePath $tplPath

        # Tabellen-Styles befüllen
        $Global:UI.cmbTableStyle.Items.Clear()
        foreach ($s in $styles.TableStyles) { [void]$Global:UI.cmbTableStyle.Items.Add($s) }
        # Vorherige Auswahl merken oder ersten passenden Eintrag
        $preTbl = $null
        foreach ($item in $Global:UI.cmbTableStyle.Items) {
            if ($item -ieq $Global:Config.TableStyleName) { $preTbl = $item; break }
        }
        if ($preTbl) { $Global:UI.cmbTableStyle.Text = $preTbl }
        elseif ($Global:UI.cmbTableStyle.Items.Count -gt 0) { $Global:UI.cmbTableStyle.SelectedIndex = 0 }
        else { $Global:UI.cmbTableStyle.Text = $Global:Config.TableStyleName }


    }

    & $fillTemplates
    & $loadTemplateStyles   # Styles für initial ausgewähltes Template laden

    # Template-SelectionChanged: Styles nachladen
    $Global:UI.cmbTemplate.Add_SelectionChanged({
        # Vermeide Neu-Laden während des Befüllens
        if ($Global:UI.cmbTemplate.Items.Count -eq 0) { return }
        & $loadTemplateStyles
    })

    if (-not [string]::IsNullOrWhiteSpace($PreloadFile) -and (Test-Path $PreloadFile)) { $Global:UI.FileList.Items.Add($PreloadFile) | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace($PreloadFolder) -and (Test-Path $PreloadFolder)) {
        $pf=Get-ChildItem -Path $PreloadFolder -Filter "*.docx" -File -Recurse | Where-Object { $_.Name -notlike "*_Backup_*" }
        foreach ($f in $pf) { $Global:UI.FileList.Items.Add($f.FullName) | Out-Null }
    }

    $btnAddFiles.Add_Click({
        $dlg=New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter="Word (*.docx;*.doc)|*.docx;*.doc"; $dlg.Multiselect=$true
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { foreach ($f in $dlg.FileNames){ if(-not $Global:UI.FileList.Items.Contains($f)){$Global:UI.FileList.Items.Add($f)|Out-Null} } }
    })
    $btnAddFolder.Add_Click({
        $dlg=New-Object System.Windows.Forms.FolderBrowserDialog; $dlg.Description="Ordner wählen"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $files=Get-ChildItem -Path $dlg.SelectedPath -Filter "*.docx" -File -Recurse | Where-Object { $_.Name -notlike "*_Backup_*" }
            foreach ($f in $files){ if(-not $Global:UI.FileList.Items.Contains($f.FullName)){$Global:UI.FileList.Items.Add($f.FullName)|Out-Null} }
            Write-Log "Ordner eingelesen: $($files.Count) Dateien." -Level INFO
        }
    })
    $btnRemove.Add_Click({ foreach ($item in @($Global:UI.FileList.SelectedItems)){ $Global:UI.FileList.Items.Remove($item) } })
    $btnClear.Add_Click({ $Global:UI.FileList.Items.Clear() })
    $btnRescan.Add_Click({ Write-Log "Suche Vorlagen neu..." -Level INFO; & $fillTemplates })  # SelectionChanged lädt Styles automatisch nach
    $btnBrowseTemplate.Add_Click({
        $dlg=New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter="Word-Vorlagen (*.dotx;*.dotm;*.dot)|*.dotx;*.dotm;*.dot|Alle (*.*)|*.*"
        $cur=& $getSelectedTemplate; if($cur -and (Test-Path $cur)){$dlg.InitialDirectory=[System.IO.Path]::GetDirectoryName($cur)}
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $item=New-Object System.Windows.Controls.ComboBoxItem
            $item.Content="$([System.IO.Path]::GetFileName($dlg.FileName))   —   $([System.IO.Path]::GetDirectoryName($dlg.FileName))"; $item.Tag=$dlg.FileName; $item.ToolTip=$dlg.FileName
            [void]$Global:UI.cmbTemplate.Items.Insert(0,$item); $Global:UI.cmbTemplate.SelectedItem=$item
        }
    })
    $btnCleanup.Add_Click({
        $logD=[int]($Global:UI.txtCleanLogs.Text); $repD=[int]($Global:UI.txtCleanReports.Text); $bakD=[int]($Global:UI.txtCleanBackups.Text)
        $dirs=@($Global:UI.FileList.Items | ForEach-Object { [System.IO.Path]::GetDirectoryName($_) } | Select-Object -Unique)
        $n=Invoke-Cleanup -LogDays $logD -ReportDays $repD -BackupDays $bakD -BackupSearchDirs $dirs
        [System.Windows.MessageBox]::Show("$n Datei(en) gelöscht.","Aufräumen","OK","Information") | Out-Null
    })
    $Global:UI.btnReport.Add_Click({ if ($Global:LastReportPath -and (Test-Path $Global:LastReportPath)) { Start-Process $Global:LastReportPath } })
    $Global:UI.btnCancel.Add_Click({ $Global:Sync.CancelRequested=$true; $Global:UI.btnCancel.IsEnabled=$false; Write-Log "Abbruch angefordert - stoppe nach aktuellem Dokument..." -Level WARN })

    # ----- START (Runspace + Timer) -----
        # ----- START (Runspace + Polling-Schleife statt Timer) -----


        # ----- START (direkte Verarbeitung im GUI-Thread mit DoEvents) -----
    $Global:UI.btnStart.Add_Click({
        $files=@($Global:UI.FileList.Items)
        if ($files.Count -eq 0) { [System.Windows.MessageBox]::Show("Bitte zuerst Dokumente hinzufügen!","Hinweis","OK","Warning")|Out-Null; return }
        $Global:Config.TemplatePath=& $getSelectedTemplate
        $Global:Config.TableStyleName = $Global:UI.cmbTableStyle.Text
        $Global:Config.VerboseSteps=$Global:UI.chkVerbose.IsChecked
        $actions=@{
            Headings=$Global:UI.chkHeadings.IsChecked; Levels=$Global:UI.chkLevels.IsChecked
            Duplicates=$Global:UI.chkDuplicates.IsChecked; ManualNum=$Global:UI.chkManualNum.IsChecked
            DeadLinks=$Global:UI.chkDeadLinks.IsChecked; Tables=$Global:UI.chkTables.IsChecked; TOC=$Global:UI.chkTOC.IsChecked
        }
        if (-not ($actions.Values -contains $true)) { [System.Windows.MessageBox]::Show("Bitte mindestens eine Aktion wählen!","Hinweis","OK","Warning")|Out-Null; return }
        if ($actions.Tables -and (-not (Test-Path $Global:Config.TemplatePath))) { [System.Windows.MessageBox]::Show("Vorlage nicht gefunden!","Vorlage fehlt","OK","Warning")|Out-Null; return }

        # Cancel-Flag zurücksetzen
        $Global:Sync.CancelRequested = $false

        $Global:UI.btnStart.IsEnabled=$false; $Global:UI.btnCancel.IsEnabled=$true
        $Global:UI.LogBox.Document.Blocks.Clear()
        Write-Log "===== Batch gestartet: $($files.Count) Dokumente =====" -Level INFO

        $results = New-Object System.Collections.ArrayList
        $total = $files.Count
        $idx = 0

        foreach ($file in $files) {
            # Abbruch-Prüfung VOR jedem Dokument
            if ($Global:Sync.CancelRequested) {
                Write-Log "Abbruch durch Benutzer - Verarbeitung gestoppt." -Level WARN
                break
            }

            $idx++
            $pct = [int](($idx / $total) * 100)
            $Global:UI.ProgressBar.Value = $pct
            $Global:UI.ProgressText.Text = "Verarbeite $idx von $total : $([System.IO.Path]::GetFileName($file))"
            [System.Windows.Forms.Application]::DoEvents()

            # Dokument verarbeiten (läuft im GUI-Thread, Write-Log zeigt live)
            $res = Invoke-ProcessDocument -DocPath $file -Actions $actions

            # Defensiv: nur echtes Ergebnis-Objekt behalten
            $res = @($res) | Where-Object {
                $_ -is [System.Management.Automation.PSCustomObject] -and
                $_.PSObject.Properties.Name -contains 'Success'
            } | Select-Object -Last 1
            if ($null -ne $res) { [void]$results.Add($res) }

            [System.Windows.Forms.Application]::DoEvents()
        }

        Write-Log "===== Batch-Verarbeitung abgeschlossen =====" -Level SUCCESS

        # Abschluss
        $Global:UI.ProgressBar.Value=100
        $Global:UI.btnStart.IsEnabled=$true; $Global:UI.btnCancel.IsEnabled=$false
        $resultArr=@($results)
        $okCount=@($resultArr | Where-Object { $_.Success -eq $true }).Count
        $errCount=@($resultArr | Where-Object { $_.Success -ne $true }).Count
        $Global:UI.ProgressText.Text="Fertig: $okCount erfolgreich, $errCount Fehler."

        if ($Global:UI.chkReport.IsChecked -and $resultArr.Count -gt 0) {
            $Global:UI.ProgressText.Text="Erstelle Vergleichsbericht..."
            [System.Windows.Forms.Application]::DoEvents()
            $Global:LastReportPath=New-ComparisonReport -Results $resultArr
            $Global:UI.btnReport.IsEnabled=$true
        }

        $cn=if($Global:Sync.CancelRequested){"`n(Abgebrochen durch Benutzer)"}else{""}
        $msg="Verarbeitung beendet!`n`nErfolgreich: $okCount`nFehler: $errCount$cn"
        if ($Global:LastReportPath) {
            $msg+="`n`nVergleichsbericht erstellt. Jetzt öffnen?"
            if ([System.Windows.MessageBox]::Show($msg,"Fertig","YesNo","Information") -eq "Yes") { Start-Process $Global:LastReportPath }
        } else { [System.Windows.MessageBox]::Show($msg,"Fertig","OK","Information")|Out-Null }
    })

    Close-SplashScreen
    $window.ShowDialog() | Out-Null
}

# ============================================================================
# START
# ============================================================================
if ($FunctionsOnly) { return }   # Für Worker-Runspace: nur Funktionen laden

if (-not (Test-Path $Global:Config.LogFolder)) { New-Item -Path $Global:Config.LogFolder -ItemType Directory -Force | Out-Null }
Write-Log "===== GUI gestartet =====" -Level INFO
Show-SplashScreen -InitialText "Starte Word-Format-Toolkit..."
Show-MainGUI
Write-Log "===== GUI beendet =====" -Level INFO
