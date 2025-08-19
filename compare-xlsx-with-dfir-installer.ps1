# Pfade setzen
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$logFile = Join-Path $scriptDir "$scriptName.log"

$excelFile = Join-Path $scriptDir "dfir-installer-selector.xlsx"
$tempFile  = Join-Path $scriptDir "tool.temp"
$listFile  = Join-Path $scriptDir "..\dfir-installer\list-of-tools.txt"

# --- Excel einlesen ---
# Hinweis: Benötigt ImportExcel Modul -> Install-Module -Name ImportExcel -Scope CurrentUser
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Error "Das Modul 'ImportExcel' ist nicht installiert. Bitte mit 'Install-Module -Name ImportExcel -Scope CurrentUser' installieren."
    exit
}

try {
    $excelData = Import-Excel -Path $excelFile -ErrorAction Stop
} catch {
    Write-Error "Fehler beim Einlesen der Excel-Datei: $_"
    exit
}

# Spalte "Tool" (ab Zeile 2, Kopfzeile wird automatisch von ImportExcel behandelt)
$tools = $excelData | Select-Object -ExpandProperty Tool

# Inhalte in tool.temp schreiben
$tools | Out-File -FilePath $tempFile -Encoding UTF8
"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] tool.temp erstellt: $tempFile" | Tee-Object -FilePath $logFile -Append

# --- Vergleich ---
if (-not (Test-Path $listFile)) {
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] FEHLER: Datei list-of-tools.txt wurde nicht gefunden: $listFile" | Tee-Object -FilePath $logFile -Append
    exit
}

$tempTools = Get-Content $tempFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
$listTools = Get-Content $listFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

# Unterschiede
$inListNotInTemp = $listTools | Where-Object { $_ -notin $tempTools }
$inTempNotInList = $tempTools | Where-Object { $_ -notin $listTools }

"`n--- Neue Tools in DFIR Installer ---" | Tee-Object -FilePath $logFile -Append
if ($inListNotInTemp) {
    $inListNotInTemp | Tee-Object -FilePath $logFile -Append
} else {
    "Keine neuen Tools gefunden." | Tee-Object -FilePath $logFile -Append
}

"`n--- Entfernte Tools in DFIR Installer ---" | Tee-Object -FilePath $logFile -Append
if ($inTempNotInList) {
    $inTempNotInList | Tee-Object -FilePath $logFile -Append
} else {
    "Keine entfernten Tools gefunden." | Tee-Object -FilePath $logFile -Append
}
