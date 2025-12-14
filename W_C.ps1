<#
.SYNOPSIS
Erstellt einen detaillierten Treiberinventarbericht im strikten CSV-Format mit Gerätetyp (Type), Vendor ID und Device ID als separate Felder.

.DESCRIPTION Die Ausgabe verwendet Kommas als strikte Trennzeichen. Die Chrome-Suche wird nur bei einem positiven Treffer ("Updates:") im Katalog geöffnet.
Das Skript stellt sicher, dass der CSV-Header und die Datenzeilen exakt dem gewünschten Format mit 6 Spalten entsprechen.
#>

Function Get-DriverReport {
    # --- Konfiguration ---
    $OSPrefix = "windows 11 client, version 25H2"
    $OutputPath = "$([Environment]::GetFolderPath('Desktop'))\Driver_Report_Detailliert.csv"
    $BaseUrl = "https://www.catalog.update.microsoft.com/Search.aspx?q="
    
    # Muster für einen positiven Treffer: Link zum Herunterladen/Download muss existieren
    $SuccessPattern = 'Updates:' 
    
    # KORRIGIERTER HEADER: Strikte CSV-Formatierung mit 6 Spalten
    $ReportHeader = "Hersteller,Gerätename,Type,Vendor ID,Device ID,Installierte_Version"

    Write-Host "Starte Treiberdaten-Erfassung..."

    # --- Teil 1: Datenbeschaffung und Formatierung (Striktes CSV) ---
    $PnpDevices = Get-PnpDevice -Status OK | Select-Object FriendlyName, InstanceId
    $AllDrivers = Get-CimInstance -ClassName Win32_PnPSignedDriver
    
    $FullOutputLines = @()
    $SearchUrlsToOpen = @()

    foreach ($Device in $PnpDevices) {
        $DriverInfo = $AllDrivers | Where-Object { $_.DeviceID -eq $Device.InstanceId }
        
        if ($DriverInfo) {
            $DriverInfo = $DriverInfo | Select-Object -First 1
            $RawHardwareId = ($DriverInfo.HardWareID | Select-Object -First 1)

            # Standardwerte fÃ¼r IDs (für den Report)
            $VendorID = "N/A"
            $DeviceID = "N/A"
            $DeviceType = "ROOT/SW"
            
            # RegulÃ¤re Ausdrücke zur ID-Extraktion
            
            # PCI Typ (VEN/DEV)
            if ($RawHardwareId -match '^PCI\\(VEN|VID)_([0-9A-F]{4}).*?(DEV|PID)_([0-9A-F]{4})') {
                $DeviceType = "PCI"
                $VendorID = $matches[2]
                $DeviceID = $matches[4]
            }
            # USB Typ (VID/PID)
            elseif ($RawHardwareId -match '^USB\\(VEN|VID)_([0-9A-F]{4}).*?(DEV|PID)_([0-9A-F]{4})') {
                $DeviceType = "USB"
                $VendorID = $matches[2]
                $DeviceID = $matches[4]
            }
            # ACPI/HDAUDIO Typ (VEN/DEV)
            elseif ($RawHardwareId -match '^(ACPI|HDAUDIO)\\(.+?)(VEN|VID)_([0-9A-F]{4}).*?(DEV|PID)_([0-9A-F]{4})') {
                $DeviceType = $matches[1]
                $VendorID = $matches[4]
                $DeviceID = $matches[6]
            }
            # ACPI PnP Typen ohne klare VEN/DEV-Struktur
            elseif ($RawHardwareId -match '^ACPI\\(.+)') {
                $DeviceType = "ACPI"
            }
            
            # 1. Zeile fÃ¼r das Textfile (STRIKTES CSV-Format mit KLARTEXT-Spalten)
            # Hersteller | GerÃ¤tename | Type=XXX | Vendor ID=XXX | Device ID=XXX | Version
            $TypeColumn = "Type=$DeviceType"
            $VendorColumn = "Vendor ID=$VendorID"
            $DeviceColumn = "Device ID=$DeviceID"
            
            $OutputLine = "$($DriverInfo.Manufacturer),$($Device.FriendlyName),$TypeColumn,$VendorColumn,$DeviceColumn,$($DriverInfo.DriverVersion)"
            $FullOutputLines += $OutputLine

            # 2. Suchanfrage fÃ¼r die URL (NUR GerÃ¤tename und OS-PrÃ¤fix)
            $SearchQuery = "$($OSPrefix)$($Device.FriendlyName)"

            # 3. Relevante GerÃ¤te fÃ¼r die Suche vormerken (Nur GerÃ¤te mit klarer ID-Struktur)
            if ($DeviceType -ne "ROOT/SW" -and $VendorID -ne "N/A") {
                $EncodedQuery = [uri]::EscapeDataString($SearchQuery)
                $FullUrl = $BaseUrl + $EncodedQuery
                
                $SearchUrlsToOpen += $FullUrl
            }
        }
    }

    # --- Teil 2: Speichern und Ã–ffnen in Notepad ---
    
    # 1. Header schreiben
    $ReportHeader | Out-File -FilePath $OutputPath -Encoding UTF8
    
    # 2. Daten sortieren und anhÃ¤ngen
    $FullOutputLines | Sort-Object | Out-File -FilePath $OutputPath -Encoding UTF8 -Append

    Write-Host "`n Detaillierter Inventarbericht im CSV-Format (6 Spalten) erfolgreich erstellt unter: $OutputPath"
    
    # Öffnen in Notepad (erster Schritt)
    Start-Process notepad.exe -ArgumentList $OutputPath
    Write-Host "Die vollstÃ¤ndige Liste wurde in Notepad geöffnet."

    # --- Teil 3: Gefilterte Suche im Chrome Browser ---
    Write-Host "`n Starte gefilterte Suche im Microsoft Update Catalog..."

    $ChromeTabsToOpen = @()
    $SkippedCount = 0

    foreach ($Url in $SearchUrlsToOpen) {
        try {
            # HTML-Code abrufen
            $Response = Invoke-WebRequest -Uri $Url -UseBasicParsing -ErrorAction SilentlyContinue

            # PRÜFUNG: Nur Öffnen, wenn der Muster-String "Updates:" im Inhalt existiert
            if ($Response.Content -match $SuccessPattern) {
                $ChromeTabsToOpen += $Url
            } else {
                $SkippedCount++
            }
        }
        catch {
            # String-Concatenation zur Vermeidung von Parser-Fehlern
            $ErrorMessage = $_.Exception.Message
            Write-Warning ("Fehler beim Abruf von " + $Url + ": " + $ErrorMessage)
            $SkippedCount++
        }

        # Kurze Pause, um die Katalog-Seite nicht zu blockieren
        Start-Sleep -Milliseconds 100
    }

    # --- Teil 4: Öffnen der gefilterten Tabs ---
    
    if ($ChromeTabsToOpen.Count -gt 0) {
        Write-Host "`n Starte $($ChromeTabsToOpen.Count) relevante Suchen..."
        foreach ($FinalUrl in $ChromeTabsToOpen) {
            Start-Process "chrome.exe" -ArgumentList $FinalUrl
            Start-Sleep -Milliseconds 200
        }
        Write-Host "`n Filterung abgeschlossen. $SkippedCount Suchen wurden übersprungen."
    } else {
        Write-Host "`n Filterung abgeschlossen. Es wurden keine relevanten Updates gefunden oder alle ($SkippedCount) wurden übersprungen."
    }
}

# Skript ausführen
Get-DriverReport
