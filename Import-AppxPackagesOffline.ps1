<#
.SYNOPSIS
    Installiert zuvor exportierte APPX/MSIX-Pakete inkl. Abhaengigkeiten
    systemweit (fuer alle bestehenden Benutzer) auf einem Air-Gapped
    Zielsystem.

.DESCRIPTION
    Erwartet eine Ordnerstruktur wie von Export-AppxPackages.ps1 erzeugt:
    pro Paket ein Unterordner mit genau einer Hauptpaketdatei
    (.appx/.appxbundle/.msix/.msixbundle) und einem "Dependencies"-
    Unterordner. Optional wird vor der Installation die SHA256-Pruefsumme
    gegen das mitgelieferte Manifest (export-manifest.csv) verglichen.

.PARAMETER SourceRoot
    Ordner mit den exportierten Paket-Unterordnern (auf das Air-Gapped
    System uebertragen, z. B. per USB).

.PARAMETER VerifyHash
    Prueft vor der Installation die SHA256-Pruefsumme jeder Hauptpaketdatei
    gegen export-manifest.csv im SourceRoot. Empfohlen.

.PARAMETER ProvisionForFutureUsers
    Registriert das Paket zusaetzlich per Add-AppxProvisionedPackage, damit
    es auch fuer zukuenftig neu angelegte lokale Benutzer verfuegbar ist.
    Wirkt nur auf ein lebendes System (kein Sysprep-Generalize-Kontext).

.PARAMETER Force
    Ueberspringt die Pruefung auf Windows PowerShell 5.1 (siehe Hinweis
    unten). Nur verwenden, wenn Sie das auf Ihrem PowerShell-7-Build
    getestet haben.

.EXAMPLE
    .\Import-AppxPackagesOffline.ps1 -SourceRoot "D:\AppxImport" -VerifyHash

.NOTES
    Muss elevated (als Administrator) ausgefuehrt werden, da -AllUsers
    verwendet wird.

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
    [switch]$ProvisionForFutureUsers,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# --- Vorpruefung 1: Administratorrechte ---
$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$currentPrincipal = New-Object -TypeName Security.Principal.WindowsPrincipal -ArgumentList $currentIdentity
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Error "[FAIL] Dieses Skript muss als Administrator ausgefuehrt werden (-AllUsers erfordert Elevation)."
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
        if ($depFiles.Count -gt 0) {
            Add-AppxPackage -Path $mainFile.FullName -DependencyPath $depFiles.FullName -AllUsers -ErrorAction Stop
        }
        else {
            Add-AppxPackage -Path $mainFile.FullName -AllUsers -ErrorAction Stop
        }

        Write-Host "[OK] Installiert (bestehende Benutzer): $($mainFile.Name)"
        $results += [PSCustomObject]@{ Folder = $folder.Name; Status = "OK"; Detail = "Installiert fuer bestehende Benutzer" }

        if ($ProvisionForFutureUsers) {
            try {
                if ($depFiles.Count -gt 0) {
                    Add-AppxProvisionedPackage -Online -PackagePath $mainFile.FullName -DependencyPackagePath $depFiles.FullName -SkipLicense -ErrorAction Stop | Out-Null
                }
                else {
                    Add-AppxProvisionedPackage -Online -PackagePath $mainFile.FullName -SkipLicense -ErrorAction Stop | Out-Null
                }
                Write-Host "[OK] Zusaetzlich fuer zukuenftige Benutzer provisioniert: $($mainFile.Name)"
            }
            catch {
                Write-Warning "[WARN] Provisionierung fuer zukuenftige Benutzer fehlgeschlagen: $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-Warning "[FAIL] Installation fehlgeschlagen fuer $($mainFile.Name): $($_.Exception.Message)"
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
