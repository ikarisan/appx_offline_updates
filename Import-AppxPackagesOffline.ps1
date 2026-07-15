<#
.SYNOPSIS
    Installiert zuvor exportierte APPX/MSIX-Pakete inkl. Abhaengigkeiten
    systemweit (fuer bestehende und zukuenftige Benutzer) auf einem
    Air-Gapped Zielsystem.

.DESCRIPTION
    Erwartet eine Ordnerstruktur wie von Export-AppxPackages.ps1 erzeugt:
    pro Paket ein Unterordner mit genau einer Hauptpaketdatei
    (.appx/.appxbundle/.msix/.msixbundle) sowie ein gemeinsamer
    "Dependencies"-Ordner direkt unter -SourceRoot. Welche Abhaengigkeits-
    Dateien ein Paket tatsaechlich benoetigt, wird aus der Spalte
    "Dependencies" von export-manifest.csv gelesen. Fehlt das Manifest oder
    die Spalte, wird aus Rueckwaerts-Kompatibilitaet auf einen
    paket-eigenen "Dependencies"-Unterordner zurueckgegriffen (altes
    Layout). Optional wird vor der Installation zusaetzlich die
    SHA256-Pruefsumme gegen das Manifest verglichen.

    WICHTIG: Add-AppxPackage installiert grundsaetzlich immer nur fuer das
    aktuell angemeldete Benutzerkonto und besitzt keinen All-Users-Schalter.
    Fuer eine echte systemweite Bereitstellung nutzt dieses Skript daher
    Add-AppxProvisionedPackage -Online (DISM-Modul). Provisionierte Pakete
    werden fuer bestehende Benutzer bei der naechsten Anmeldung automatisch
    registriert, und fuer neu angelegte Benutzer ebenfalls automatisch.
    Zusaetzlich installiert das Skript per einfachem Add-AppxPackage auch
    sofort fuer die aktuell laufende (administrative) Sitzung, damit das
    Ergebnis ohne Ab-/Anmelden ueberprueft werden kann.

    Sowohl Add-AppxProvisionedPackage als auch Add-AppxPackage koennen ohne
    Fehler zurueckkehren, ohne dass das Paket fuer die aktuelle Sitzung
    tatsaechlich nutzbar ist (z. B. weil es fuer den aktuellen Benutzer nur
    als "Staged" statt "Installed" registriert wurde und erst bei der
    naechsten Anmeldung aktiv wird). Das Skript verifiziert deshalb nach
    beiden Schritten per Get-AppxProvisionedPackage bzw. Get-AppxPackage, ob
    das Paket wirklich (a) provisioniert und (b) fuer die aktuelle Sitzung
    installiert ist. Stellt die Verifikation den Zustand "Staged" fest, wird
    das Paket per Add-AppxPackage -RegisterByFamilyName direkt fuer die
    aktuelle Sitzung registriert (ohne erneutes Staging); nur wenn auch das
    fehlschlaegt, wird der Status "WARN" statt "OK" gemeldet (siehe
    Detail-Spalte fuer den genauen Grund).

    Vor jeder Installation prueft das Skript ausserdem, ob das Paket bereits
    in gleicher oder neuerer Version auf dem System vorhanden ist. Name und
    Version werden dazu direkt aus der Paketdatei gelesen (AppxManifest bzw.
    AppxBundleManifest), da Ordner- und Dateinamen nicht garantiert dem
    echten Appx-Identity-Namen entsprechen. Ist bereits eine gleiche oder
    neuere Version vorhanden, wird nichts installiert und der Status "SKIP"
    gemeldet (eine aeltere vorhandene Version wird normal aktualisiert).

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

# --- Manifest laden (fuer Hash-Pruefung UND Abhaengigkeits-Aufloesung) ---
$manifestPath = Join-Path -Path $SourceRoot -ChildPath "export-manifest.csv"
$manifest = $null
if (Test-Path -Path $manifestPath) {
    $manifest = Import-Csv -Path $manifestPath
}
elseif ($VerifyHash) {
    Write-Warning "[WARN] -VerifyHash gesetzt, aber export-manifest.csv nicht gefunden unter: $manifestPath"
    Write-Warning "Fahre ohne Hash-Pruefung fort."
}

function Get-PackageIdentityFromFile {
    <#
        Liest den Appx-Identity-Namen und die Version direkt aus der
        Paketdatei (Bundle: AppxMetadata/AppxBundleManifest.xml, Einzelpaket:
        AppxManifest.xml). Zuverlaessiger als Ordner-/Dateiname, die (v. a.
        beim winget-Pfad) nicht garantiert dem Identity-Namen entsprechen.
        Gibt $null zurueck, wenn die Datei nicht lesbar ist.
    #>
    param([string]$PackagePath)

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
        try {
            $entry = $zip.Entries | Where-Object { $_.FullName -eq "AppxMetadata/AppxBundleManifest.xml" }
            if (-not $entry) {
                $entry = $zip.Entries | Where-Object { $_.FullName -eq "AppxManifest.xml" }
            }
            if (-not $entry) { return $null }

            $reader = New-Object System.IO.StreamReader($entry.Open())
            $manifestText = $reader.ReadToEnd()
            $reader.Close()

            if ($manifestText -notmatch '<Identity[^>]*\sName="([^"]+)"') { return $null }
            $name = $Matches[1]
            $version = $null
            if ($manifestText -match '<Identity[^>]*\sVersion="([^"]+)"') { $version = $Matches[1] }

            return [PSCustomObject]@{ Name = $name; Version = $version }
        }
        finally {
            $zip.Dispose()
        }
    }
    catch {
        return $null
    }
}

function Compare-AppxVersion {
    param([string]$VersionA, [string]$VersionB)
    $partsA = $VersionA -split "\." | ForEach-Object { [int]$_ }
    $partsB = $VersionB -split "\." | ForEach-Object { [int]$_ }
    for ($i = 0; $i -lt [Math]::Max($partsA.Count, $partsB.Count); $i++) {
        $a = if ($i -lt $partsA.Count) { $partsA[$i] } else { 0 }
        $b = if ($i -lt $partsB.Count) { $partsB[$i] } else { 0 }
        if ($a -ne $b) { return $a - $b }
    }
    return 0
}

function Register-StagedPackage {
    <#
        Registriert ein bereits provisioniertes/gestagtes Paket sofort fuer
        den aktuellen Benutzer, ohne es erneut zu stagen. Gibt $true bei
        Erfolg zurueck.
    #>
    param([string]$PackageFamilyName)

    try {
        Add-AppxPackage -RegisterByFamilyName -MainPackage $PackageFamilyName -ErrorAction Stop
        return $true
    }
    catch {
        Write-Warning "[WARN] Registrierung des gestagten Pakets '$PackageFamilyName' fehlgeschlagen: $($_.Exception.Message)"
        return $false
    }
}

function Get-InstalledPackageHighestVersion {
    # Liefert das systemweit registrierte Paket mit der hoechsten Version
    # (Get-AppxPackage -AllUsers kann mehrere Versionen zurueckgeben).
    param([string]$Name)

    return Get-AppxPackage -AllUsers -Name $Name -ErrorAction SilentlyContinue |
        Sort-Object -Property { [version]$_.Version } -Descending |
        Select-Object -First 1
}

$results = @()

# Den gemeinsamen "Dependencies"-Ordner ueberspringen - er enthaelt nur
# Framework-Abhaengigkeiten und ist kein App-Ordner (Framework-Pakete lassen
# sich nicht provisionieren).
$appFolders = @(Get-ChildItem -Path $SourceRoot -Directory |
    Where-Object { $_.Name -ne "Dependencies" })

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

    # Echten Appx-Identity-Namen (+ Version) aus der Paketdatei lesen; der
    # Ordnername dient nur als Fallback, falls die Datei nicht lesbar ist.
    $identity = Get-PackageIdentityFromFile -PackagePath $mainFile.FullName
    $pkgName = if ($identity) { $identity.Name } else { $folder.Name }
    if (-not $identity) {
        Write-Warning "[WARN] Identity aus '$($mainFile.Name)' nicht lesbar, verwende Ordnernamen '$($folder.Name)' fuer die Verifikation."
    }

    # Manifest-Eintrag fuer dieses Paket ermitteln (wird sowohl fuer -VerifyHash
    # als auch fuer die Abhaengigkeits-Aufloesung ueber die "Dependencies"-Spalte
    # benoetigt).
    $manifestEntry = $null
    if ($manifest) {
        $manifestEntry = $manifest | Where-Object { $_.MainFile -eq $mainFile.Name -and $_.Status -eq "OK" } | Select-Object -First 1
    }

    # Optional: Hash gegen Manifest pruefen
    if ($VerifyHash -and $manifest) {
        if ($manifestEntry) {
            $actualHash = (Get-FileHash -Path $mainFile.FullName -Algorithm SHA256).Hash
            if ($actualHash -ne $manifestEntry.SHA256) {
                Write-Warning "[FAIL] SHA256 stimmt nicht mit dem Manifest ueberein: $($mainFile.Name)"
                Write-Warning "       Erwartet: $($manifestEntry.SHA256)"
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

    # --- Vorab-Check: Paket bereits in gleicher oder neuerer Version vorhanden? ---
    # (Haeufig bei unter Windows vorinstallierten Apps. Ohne diesen Check laeuft
    # die Sofort-Installation in einen Downgrade-Fehler; eine AELTERE vorhandene
    # Version wird dagegen normal aktualisiert.)
    if ($identity -and $identity.Version) {
        $already = Get-InstalledPackageHighestVersion -Name $pkgName
        if ($already -and (Compare-AppxVersion $already.Version $identity.Version) -ge 0) {
            Write-Host "[SKIP] '$pkgName' ist bereits in Version $($already.Version) vorhanden (Import-Version: $($identity.Version))."
            $detail = "Bereits vorhanden in Version $($already.Version), Import-Version $($identity.Version)"

            # Ist das vorhandene Paket fuer die aktuelle Sitzung nur "Staged",
            # jetzt direkt registrieren statt auf die naechste Anmeldung zu warten.
            if (-not $SkipCurrentUserInstall) {
                $currentSid = $currentIdentity.User.Value
                $stagedHere = $already.PackageUserInformation |
                    Where-Object { $_.UserSecurityId.ToString() -like "$currentSid*" -and $_.InstallState -eq "Staged" } |
                    Select-Object -First 1
                if ($stagedHere) {
                    Write-Host "       Fuer die aktuelle Sitzung nur 'Staged' - registriere jetzt..."
                    if (Register-StagedPackage -PackageFamilyName $already.PackageFamilyName) {
                        Write-Host "[OK] '$pkgName' fuer die aktuelle Sitzung registriert."
                        $detail += "; war 'Staged', jetzt fuer aktuelle Sitzung registriert"
                    }
                    else {
                        $detail += "; 'Staged' fuer aktuelle Sitzung, Registrierung fehlgeschlagen (Ab-/Neuanmeldung noetig)"
                    }
                }
            }

            $results += [PSCustomObject]@{ Folder = $folder.Name; Status = "SKIP"; Detail = $detail }
            continue
        }
    }

    # Abhaengigkeiten aufloesen: bevorzugt ueber die "Dependencies"-Spalte des
    # Manifests (verweist auf Dateien im gemeinsamen SourceRoot\Dependencies-
    # Ordner, aktuelles Export-Layout). Fallback auf einen paket-eigenen
    # <Paketordner>\Dependencies-Unterordner (altes Layout / kein Manifest).
    $sharedDepFolder = Join-Path -Path $SourceRoot -ChildPath "Dependencies"
    $depFiles = @()
    if ($manifestEntry -and $manifestEntry.Dependencies) {
        $depFileNames = $manifestEntry.Dependencies -split ";" | Where-Object { $_ }
        foreach ($depFileName in $depFileNames) {
            $depPath = Join-Path -Path $sharedDepFolder -ChildPath $depFileName
            if (Test-Path -Path $depPath) {
                $depFiles += Get-Item -Path $depPath
            }
            else {
                Write-Warning "[WARN] Im Manifest referenzierte Abhaengigkeit nicht gefunden: $depPath"
            }
        }
    }
    else {
        # Mehrere Wildcard-Pfade statt -Include, da -Include nur zuverlaessig
        # greift, wenn Path selbst bereits ein Wildcard-Element enthaelt.
        $depFolder = Join-Path -Path $folder.FullName -ChildPath "Dependencies"
        if (Test-Path -Path $depFolder) {
            $depPatterns = @(
                (Join-Path $depFolder "*.appx"),
                (Join-Path $depFolder "*.msix"),
                (Join-Path $depFolder "*.msixbundle"),
                (Join-Path $depFolder "*.appxbundle")
            )
            $depFiles = @(Get-ChildItem -Path $depPatterns -File -ErrorAction SilentlyContinue)
        }
    }
    $depFiles = @($depFiles)

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

        # Verifikation: DISM kann Erfolg melden, ohne dass ein Eintrag in
        # Get-AppxProvisionedPackage auftaucht (seltener Randfall) - lieber
        # ehrlich als "WARN" statt ungeprueft "OK" melden.
        $provisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
            Where-Object { $_.PackageName -like "$($pkgName)_*" }

        if ($provisioned) {
            Write-Host "[OK] Systemweit provisioniert (verifiziert): $($mainFile.Name) -> $($provisioned.PackageName) (bestehende Benutzer ab naechster Anmeldung, neue Benutzer automatisch)"
            $status = "OK"
            $detail = "Provisioniert und verifiziert per Get-AppxProvisionedPackage"
        }
        else {
            Write-Warning "[WARN] Add-AppxProvisionedPackage meldete Erfolg, aber '$pkgName' erscheint nicht in Get-AppxProvisionedPackage."
            $status = "WARN"
            $detail = "Provisionierung meldete Erfolg, konnte aber nicht per Get-AppxProvisionedPackage verifiziert werden"
        }

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
                # (HResult-Vergleich per Hex-String, da die Fehlermeldung lokalisiert
                # ist und den Code nicht in jedem Fall enthaelt.)
                $isInUseError = (("0x{0:X8}" -f $_.Exception.HResult) -eq "0x80073D02") -or ($_.Exception.Message -match "0x80073D02")
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

            # Verifikation: Add-AppxPackage kann ohne Fehler zurueckkehren, obwohl
            # das Paket fuer den aktuellen Benutzer nur als "Staged" (statt
            # "Installed") registriert wurde und erst bei der naechsten Anmeldung
            # wirklich verfuegbar ist. Deshalb den tatsaechlichen Zustand ueber
            # PackageUserInformation der aktuell angemeldeten SID pruefen.
            $currentSid = $currentIdentity.User.Value
            $installedPkg = Get-InstalledPackageHighestVersion -Name $pkgName
            $currentUserEntry = $null
            if ($installedPkg) {
                $currentUserEntry = $installedPkg.PackageUserInformation |
                    Where-Object { $_.UserSecurityId.ToString() -like "$currentSid*" } |
                    Select-Object -First 1
            }

            if ($currentUserEntry -and $currentUserEntry.InstallState -eq "Installed") {
                Write-Host "[OK] Verifiziert: '$pkgName' ist fuer die aktuelle Sitzung installiert."
                $detail += "; verifiziert: fuer aktuelle Sitzung installiert"
            }
            elseif ($currentUserEntry -and $currentUserEntry.InstallState -eq "Staged") {
                # Direkte Abhilfe: das bereits gestagte Paket sofort fuer den
                # aktuellen Benutzer registrieren (kein erneutes Staging noetig).
                Write-Warning "[WARN] '$pkgName' ist fuer die aktuelle Sitzung nur 'Staged' - versuche direkte Registrierung..."
                $nowInstalled = $false
                if (Register-StagedPackage -PackageFamilyName $installedPkg.PackageFamilyName) {
                    $recheck = Get-InstalledPackageHighestVersion -Name $pkgName
                    if ($recheck) {
                        $recheckEntry = $recheck.PackageUserInformation |
                            Where-Object { $_.UserSecurityId.ToString() -like "$currentSid*" } |
                            Select-Object -First 1
                        $nowInstalled = ($recheckEntry -and $recheckEntry.InstallState -eq "Installed")
                    }
                }

                if ($nowInstalled) {
                    Write-Host "[OK] '$pkgName' nach direkter Registrierung fuer die aktuelle Sitzung installiert."
                    $detail += "; war 'Staged', nach Registrierung fuer aktuelle Sitzung installiert"
                }
                else {
                    $status = "WARN"
                    $detail += "; nur 'Staged' fuer aktuelle Sitzung, Ab-/Neuanmeldung noetig"
                }
            }
            else {
                Write-Warning "[WARN] '$pkgName' konnte fuer die aktuelle Sitzung nicht verifiziert werden (kein passender Eintrag in Get-AppxPackage)."
                $status = "WARN"
                $detail += "; nicht verifizierbar fuer aktuelle Sitzung"
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

$warnCount = @($results | Where-Object Status -eq "WARN").Count
if ($warnCount -gt 0) {
    Write-Warning "$warnCount Paket(e) wurden provisioniert, konnten aber nicht als sofort einsatzbereit verifiziert werden (Details siehe Detail-Spalte; ggf. Ab-/Neuanmeldung noetig)."
}

$failCount = @($results | Where-Object Status -eq "FAIL").Count
if ($failCount -gt 0) {
    Write-Warning "$failCount Paket(e) konnten nicht installiert werden. Details siehe obige Tabelle."
    exit 1
}