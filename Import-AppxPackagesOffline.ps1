<#
.SYNOPSIS
    Installiert zuvor exportierte APPX/MSIX-Pakete inkl. Abhaengigkeiten
    systemweit (fuer bestehende und zukuenftige Benutzer) auf einem
    Air-Gapped Zielsystem.

.DESCRIPTION
    Erwartet eine Ordnerstruktur wie von Export-AppxPackages.ps1 erzeugt:
    pro Paket ein Unterordner mit genau einer Hauptpaketdatei
    (.appx/.appxbundle/.msix/.msixbundle) und einem "Dependencies"-
    Unterordner. Optional wird vor der Installation die SHA256-Pruefsumme
    gegen das mitgelieferte Manifest (export-manifest.csv) verglichen.

    WICHTIG: Add-AppxPackage installiert grundsaetzlich immer nur fuer das
    aktuell angemeldete Benutzerkonto und besitzt keinen All-Users-Schalter.
    Fuer eine echte systemweite Bereitstellung nutzt dieses Skript daher
    Add-AppxProvisionedPackage -Online (DISM-Modul). Provisionierte Pakete
    werden fuer bestehende Benutzer bei der naechsten Anmeldung automatisch
    registriert, und fuer neu angelegte Benutzer ebenfalls automatisch.
    Zusaetzlich installiert das Skript per einfachem Add-AppxPackage auch
    sofort fuer die aktuell laufende (administrative) Sitzung, damit das
    Ergebnis ohne Ab-/Anmelden ueberprueft werden kann.

.PARAMETER SourceRoot
    Ordner mit den exportierten Paket-Unterordnern (auf das Air-Gapped
    System uebertragen, z. B. per USB).

.PARAMETER VerifyHash
    Prueft vor der Installation die SHA256-Pruefsumme jeder Hauptpaketdatei
    gegen export-manifest.csv im SourceRoot. Empfohlen.

.PARAMETER SkipCurrentUserInstall
    Ueberspringt die zusaetzliche sofortige Installation fuer die aktuell
    angemeldete (administrative) Sitzung. Ohne diesen Schalter wird nach
    der Provisionierung zusaetzlich ein einfacher Add-AppxPackage-Aufruf
    fuer den aktuellen Benutzer ausgefuehrt.

.PARAMETER Force
    Ueberspringt die Pruefung auf Windows PowerShell 5.1 (siehe Hinweis
    unten). Nur verwenden, wenn Sie das auf Ihrem PowerShell-7-Build
    getestet haben.

.PARAMETER ForceApplicationShutdown
    Beendet bei der Sofort-Installation fuer die aktuelle Sitzung laufende
    Prozesse des betroffenen Pakets und erzwingt so das Update. Behebt den
    Fehler HRESULT 0x80073D02 ("Das Paket konnte nicht installiert werden,
    da die davon geaenderten Ressourcen derzeit verwendet werden" bzw.
    "folgende Apps muessen geschlossen werden"), der auftritt, wenn eine
    aeltere Version der App oder eine gemeinsam genutzte Framework-
    Abhaengigkeit gerade laeuft. Achtung: laufende Apps werden ohne
    Rueckfrage geschlossen.

    Auch ohne diesen Schalter wiederholt das Skript die Sofort-Installation
    bei genau diesem Fehler automatisch einmal mit -ForceApplicationShutdown.
    Der Schalter erzwingt das Verhalten bereits beim ersten Versuch.

.EXAMPLE
    .\Import-AppxPackagesOffline.ps1 -SourceRoot "D:\AppxImport" -VerifyHash

.NOTES
    Muss elevated (als Administrator) ausgefuehrt werden, da
    Add-AppxProvisionedPackage -Online administrative Rechte erfordert.

    Muss in Windows PowerShell 5.1 laufen (powershell.exe), nicht in
    PowerShell 7 (pwsh.exe): Es gibt einen bekannten, gemeldeten Fehler, bei
    dem -DependencyPath als Array in PowerShell 7 fehlerhaft zu einem
    einzigen, ungueltigen String zusammengefasst wird, wodurch die
    Installation der Abhaengigkeiten mit einem "Cannot find path" Fehler
    fehlschlaegt. Das Skript prueft dies und bricht standardmaessig ab,
    falls es unter PowerShell 7 gestartet wird.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceRoot,

    [Parameter(Mandatory = $false)]
    [switch]$VerifyHash,

    [Parameter(Mandatory = $false)]
    [switch]$SkipCurrentUserInstall,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$ForceApplicationShutdown
)

Import-Module Dism -ErrorAction SilentlyContinue

# --- Vorpruefung 1: Administratorrechte ---
$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$currentPrincipal = New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList $currentIdentity
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Error "[FAIL] Dieses Skript muss als Administrator ausgefuehrt werden (Add-AppxProvisionedPackage -Online erfordert Elevation)."
    exit 1
}

# --- Vorpruefung 2: PowerShell-Edition ---
if ($PSVersionTable.PSEdition -eq "Core" -and -not $Force) {
    Write-Error "[FAIL] Dieses Skript sollte in Windows PowerShell 5.1 laufen (powershell.exe), nicht in PowerShell 7/Core (pwsh.exe)."
    Write-Error "Grund: bekannter Fehler, der -DependencyPath als Array in PowerShell 7 zu einem einzigen ungueltigen String zusammenfasst."
    Write-Error "Mit -Force kann diese Pruefung uebersprungen werden, auf eigenes Risiko."
    exit 1
}

if (-not (Test-Path -Path $SourceRoot)) {
    Write-Error "[FAIL] SourceRoot nicht gefunden: $SourceRoot"
    exit 1
}

# --- Optional: Manifest fuer Hash-Pruefung laden ---
$manifest = $null
if ($VerifyHash) {
    $manifestPath = Join-Path -Path $SourceRoot -ChildPath "export-manifest.csv"
    if (Test-Path -Path $manifestPath) {
        $manifest = Import-Csv -Path $manifestPath
    }
    else {
        Write-Warning "[WARN] -VerifyHash gesetzt, aber export-manifest.csv nicht gefunden unter: $manifestPath"
        Write-Warning "Fahre ohne Hash-Pruefung fort."
    }
}

$results = @()
$appFolders = Get-ChildItem -Path $SourceRoot -Directory

if ($appFolders.Count -eq 0) {
    Write-Warning "[WARN] Keine Unterordner unter $SourceRoot gefunden. Nichts zu installieren."
    exit 0
}

foreach ($folder in $appFolders) {

    Write-Host "---> Verarbeite Ordner: $($folder.Name)"

    # Hauptpaketdatei suchen (liegt direkt im Ordner, nicht im Dependencies-Unterordner,
    # da Get-ChildItem hier ohne -Recurse nur die unmittelbaren Kindelemente listet)
    $mainFile = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in ".msixbundle", ".appxbundle" } |
        Select-Object -First 1

    if (-not $mainFile) {
        $mainFile = Get-ChildItem -Path $folder.FullName -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in ".msix", ".appx" } |
            Select-Object -First 1
    }

    if (-not $mainFile) {
        Write-Warning "[SKIP] Keine .msixbundle/.appxbundle/.msix/.appx Datei in $($folder.FullName) gefunden."
        $results += [PSCustomObject]@{ Folder = $folder.Name; Status = "SKIP"; Detail = "Keine Hauptpaketdatei gefunden" }
        continue
    }

    # Optional: Hash gegen Manifest pruefen
    if ($VerifyHash -and $manifest) {
        $entry = $manifest | Where-Object { $_.MainFile -eq $mainFile.Name -and $_.Status -eq "OK" }
        if ($entry) {
            $actualHash = (Get-FileHash -Path $mainFile.FullName -Algorithm SHA256).Hash
            if ($actualHash -ne $entry.SHA256) {
                Write-Warning "[FAIL] SHA256 stimmt nicht mit dem Manifest ueberein: $($mainFile.Name)"
                Write-Warning "       Erwartet: $($entry.SHA256)"
                Write-Warning "       Ist:      $actualHash"
                $results += [PSCustomObject]@{ Folder = $folder.Name; Status = "FAIL"; Detail = "Hash-Abweichung, Installation uebersprungen" }
                continue
            }
            else {
                Write-Host "[OK] Hash bestaetigt fuer $($mainFile.Name)"
            }
        }
        else {
            Write-Warning "[WARN] Kein Manifest-Eintrag fuer $($mainFile.Name), Hash-Pruefung nicht moeglich."
        }
    }

    # Abhaengigkeiten sammeln (mehrere Wildcard-Pfade statt -Include, da -Include
    # nur zuverlaessig greift, wenn Path selbst bereits ein Wildcard-Element enthaelt)
    $depFolder = Join-Path -Path $folder.FullName -ChildPath "Dependencies"
    $depFiles = @()
    if (Test-Path -Path $depFolder) {
        $depPatterns = @(
            (Join-Path $depFolder "*.appx"),
            (Join-Path $depFolder "*.msix"),
            (Join-Path $depFolder "*.msixbundle"),
            (Join-Path $depFolder "*.appxbundle")
        )
        $depFiles = @(Get-ChildItem -Path $depPatterns -File -ErrorAction SilentlyContinue)
    }

    try {
        # Primaerer Weg fuer "systemweit": DISM-Provisionierung. Wirkt fuer
        # bestehende Benutzer ab der naechsten Anmeldung und automatisch
        # fuer jeden neu angelegten Benutzer.
        if ($depFiles.Count -gt 0) {
            Add-AppxProvisionedPackage -Online -PackagePath $mainFile.FullName -DependencyPackagePath $depFiles.FullName -SkipLicense -ErrorAction Stop | Out-Null
        }
        else {
            Add-AppxProvisionedPackage -Online -PackagePath $mainFile.FullName -SkipLicense -ErrorAction Stop | Out-Null
        }

        Write-Host "[OK] Systemweit provisioniert: $($mainFile.Name) (bestehende Benutzer ab naechster Anmeldung, neue Benutzer automatisch)"
        $status = "OK"
        $detail = "Provisioniert per Add-AppxProvisionedPackage"

        # Zusaetzlich sofort fuer die aktuelle (administrative) Sitzung installieren,
        # damit das Ergebnis ohne Ab-/Anmelden getestet werden kann
        if (-not $SkipCurrentUserInstall) {
            # Lokale Funktion fuer den eigentlichen Installationsversuch. -force
            # steuert, ob -ForceApplicationShutdown gesetzt wird (beendet laufende
            # Prozesse des Pakets, behebt HRESULT 0x80073D02).
            $installCurrentSession = {
                param([bool]$UseForceShutdown)
                $addAppxParams = @{ ErrorAction = "Stop" }
                if ($UseForceShutdown) {
                    $addAppxParams["ForceApplicationShutdown"] = $true
                }
                if ($depFiles.Count -gt 0) {
                    Add-AppxPackage -Path $mainFile.FullName -DependencyPath $depFiles.FullName @addAppxParams
                }
                else {
                    Add-AppxPackage -Path $mainFile.FullName @addAppxParams
                }
            }

            try {
                & $installCurrentSession $ForceApplicationShutdown.IsPresent
                Write-Host "[OK] Zusaetzlich sofort fuer aktuelle Sitzung installiert: $($mainFile.Name)"
                $detail += "; sofort fuer aktuelle Sitzung installiert"
            }
            catch {
                # 0x80073D02: Ressourcen/Prozesse des Pakets in Verwendung. Wurde noch
                # nicht mit erzwungenem Shutdown versucht -> genau einmal wiederholen.
                $isInUseError = $_.Exception.Message -match "0x80073D02"
                if ($isInUseError -and -not $ForceApplicationShutdown) {
                    Write-Warning "[WARN] $($mainFile.Name): App/Ressourcen in Verwendung (0x80073D02). Wiederhole mit erzwungenem Beenden laufender Prozesse..."
                    try {
                        & $installCurrentSession $true
                        Write-Host "[OK] Nach erzwungenem Beenden sofort fuer aktuelle Sitzung installiert: $($mainFile.Name)"
                        $detail += "; sofort fuer aktuelle Sitzung installiert (nach ForceApplicationShutdown)"
                    }
                    catch {
                        Write-Warning "[WARN] Sofort-Installation auch mit ForceApplicationShutdown fehlgeschlagen (Provisionierung ist trotzdem aktiv): $($_.Exception.Message)"
                        $detail += "; Sofort-Installation fehlgeschlagen (auch mit ForceApplicationShutdown): $($_.Exception.Message)"
                    }
                }
                else {
                    Write-Warning "[WARN] Sofort-Installation fuer aktuelle Sitzung fehlgeschlagen (Provisionierung ist trotzdem aktiv): $($_.Exception.Message)"
                    $detail += "; Sofort-Installation fuer aktuelle Sitzung fehlgeschlagen: $($_.Exception.Message)"
                }
            }
        }

        $results += [PSCustomObject]@{ Folder = $folder.Name; Status = $status; Detail = $detail }
    }
    catch {
        Write-Warning "[FAIL] Provisionierung fehlgeschlagen fuer $($mainFile.Name): $($_.Exception.Message)"
        $results += [PSCustomObject]@{ Folder = $folder.Name; Status = "FAIL"; Detail = $_.Exception.Message }
    }
}

Write-Host ""
Write-Host "---> Zusammenfassung:"
$results | Format-Table -AutoSize

$failCount = ($results | Where-Object Status -eq "FAIL").Count
if ($failCount -gt 0) {
    Write-Warning "$failCount Paket(e) konnten nicht installiert werden. Details siehe obige Tabelle."
    exit 1
}