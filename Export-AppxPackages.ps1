<#
.SYNOPSIS
    Exportiert ausgewaehlte, bereits installierte APPX/MSIX-Pakete inkl. aller
    Abhaengigkeiten von diesem (Online-)Geraet zur spaeteren Offline-Installation
    auf einem Air-Gapped Zielsystem.

.DESCRIPTION
    Nutzt Save-AppxPackage, um pro Paket einen eigenen Unterordner mit dem
    Hauptpaket (.appx/.appxbundle/.msix/.msixbundle) und einem automatisch
    erzeugten "Dependencies"-Unterordner anzulegen.
    Zusaetzlich wird ein Manifest (export-manifest.csv) mit SHA256-Hashes
    geschrieben, damit die Paketintegritaet nach dem Transfer auf dem
    Zielsystem geprueft werden kann.

.PARAMETER PackageNames
    Ein oder mehrere Appx-Paketnamen (Get-AppxPackage -Name Wert), z. B.
    "Microsoft.MicrosoftStickyNotes". Wenn dieser Parameter angegeben wird,
    hat er Vorrang: eine eventuell vorhandene packages.txt (Default oder
    ueber -PackageListFile angegeben) wird dann komplett ignoriert.

.PARAMETER PackageListFile
    Pfad zu einer Textdatei mit einem Appx-Paketnamen pro Zeile. Leere
    Zeilen und Zeilen, die mit "#" beginnen, werden ignoriert (Kommentare).
    Wird nur verwendet, wenn -PackageNames NICHT angegeben ist. Ohne
    -PackageListFile wird in diesem Fall standardmaessig nach einer Datei
    "packages.txt" im Skriptverzeichnis gesucht. Existiert weder
    -PackageNames noch eine gueltige Paketliste, bricht das Skript mit
    einem Fehler ab.

.PARAMETER ExportRoot
    Zielordner fuer den Export. Wird angelegt, falls nicht vorhanden.

.PARAMETER AllUsers
    Ruft die Paketinfo mit Get-AppxPackage -AllUsers ab (erfordert
    Administratorrechte). Ohne diesen Schalter wird nur im Kontext des
    aktuell angemeldeten Benutzers gesucht (kein Admin noetig).

.EXAMPLE
    .\Export-AppxPackages.ps1 -PackageNames "Microsoft.MicrosoftStickyNotes" -ExportRoot "D:\AppxExport"

.EXAMPLE
    .\Export-AppxPackages.ps1 -PackageNames "Microsoft.MicrosoftStickyNotes","Microsoft.WindowsCalculator" -ExportRoot "D:\AppxExport" -AllUsers

.EXAMPLE
    # Nutzt automatisch packages.txt im Skriptverzeichnis
    .\Export-AppxPackages.ps1 -ExportRoot "D:\AppxExport"

.EXAMPLE
    .\Export-AppxPackages.ps1 -PackageListFile "D:\pakete.txt" -ExportRoot "D:\AppxExport"

.NOTES
    Voraussetzung: Die angegebenen Pakete muessen auf diesem Geraet installiert
    sein (fuer den aktuellen Benutzer oder, mit -AllUsers, fuer irgendeinen
    Benutzer). Fehlt ein Paket, zuerst ueber Store/winget installieren und
    einmal starten, dann das Skript erneut ausfuehren.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$PackageNames,

    [Parameter(Mandatory = $false)]
    [string]$PackageListFile,

    [Parameter(Mandatory = $false)]
    [string]$ExportRoot = ".\AppxExport",

    [Parameter(Mandatory = $false)]
    [switch]$AllUsers
)

if (-not $PackageNames) {
    $listPath = if ($PackageListFile) { $PackageListFile } else { Join-Path -Path $PSScriptRoot -ChildPath "packages.txt" }

    if (-not (Test-Path -Path $listPath)) {
        Write-Error "[FAIL] Weder -PackageNames angegeben noch Paketliste gefunden: $listPath"
        exit 1
    }

    $PackageNames = Get-Content -Path $listPath |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and -not $_.StartsWith("#") }

    if (-not $PackageNames -or $PackageNames.Count -eq 0) {
        Write-Error "[FAIL] Paketliste '$listPath' enthaelt keine gueltigen Eintraege."
        exit 1
    }

    Write-Host "---> Paketliste geladen: $listPath ($($PackageNames.Count) Paket(e))"
}

$manifestEntries = @()

if (-not (Test-Path -Path $ExportRoot)) {
    New-Item -Path $ExportRoot -ItemType Directory -Force | Out-Null
}

foreach ($name in $PackageNames) {

    Write-Host "---> Verarbeite Paket: $name"

    $pkg = $null
    try {
        if ($AllUsers) {
            $pkg = Get-AppxPackage -AllUsers -Name $name -ErrorAction Stop
        }
        else {
            $pkg = Get-AppxPackage -Name $name -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "[FAIL] Get-AppxPackage Fehler fuer '$name': $($_.Exception.Message)"
        $manifestEntries += [PSCustomObject]@{
            PackageName   = $name
            Status        = "FAIL"
            Detail        = "Get-AppxPackage Fehler: $($_.Exception.Message)"
            Version       = $null
            ExportPath    = $null
            MainFile      = $null
            SHA256        = $null
            ExportedAtUtc = (Get-Date).ToUniversalTime().ToString("s")
        }
        continue
    }

    if (-not $pkg) {
        Write-Warning "[SKIP] Paket nicht installiert: $name (zuerst ueber Store/winget installieren und einmal starten)"
        $manifestEntries += [PSCustomObject]@{
            PackageName   = $name
            Status        = "SKIP"
            Detail        = "Paket ist auf diesem Geraet nicht installiert"
            Version       = $null
            ExportPath    = $null
            MainFile      = $null
            SHA256        = $null
            ExportedAtUtc = (Get-Date).ToUniversalTime().ToString("s")
        }
        continue
    }

    if ($pkg -is [array]) {
        Write-Warning "[WARN] Mehrere Treffer fuer '$name', verwende den ersten: $($pkg[0].PackageFullName)"
        $pkg = $pkg[0]
    }

    $destFolder = Join-Path -Path $ExportRoot -ChildPath $pkg.Name

    try {
        if (-not (Test-Path -Path $destFolder)) {
            New-Item -Path $destFolder -ItemType Directory -Force | Out-Null
        }

        Save-AppxPackage -Package $pkg -Path $destFolder -IncludeDependencies -ErrorAction Stop

        # Hauptpaketdatei ermitteln (liegt direkt im Zielordner, nicht im
        # automatisch erzeugten Dependencies-Unterordner)
        $mainFile = Get-ChildItem -Path $destFolder -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in ".msixbundle", ".appxbundle" } |
            Select-Object -First 1

        if (-not $mainFile) {
            $mainFile = Get-ChildItem -Path $destFolder -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -in ".msix", ".appx" } |
                Select-Object -First 1
        }

        if (-not $mainFile) {
            throw "Keine .msixbundle/.appxbundle/.msix/.appx Datei nach dem Export gefunden."
        }

        $hash = Get-FileHash -Path $mainFile.FullName -Algorithm SHA256

        Write-Host "[OK] Exportiert: $($mainFile.Name) (Version $($pkg.Version))"

        $manifestEntries += [PSCustomObject]@{
            PackageName   = $name
            Status        = "OK"
            Detail        = ""
            Version       = $pkg.Version
            ExportPath    = $destFolder
            MainFile      = $mainFile.Name
            SHA256        = $hash.Hash
            ExportedAtUtc = (Get-Date).ToUniversalTime().ToString("s")
        }
    }
    catch {
        Write-Warning "[FAIL] Export fehlgeschlagen fuer '$name': $($_.Exception.Message)"
        $manifestEntries += [PSCustomObject]@{
            PackageName   = $name
            Status        = "FAIL"
            Detail        = "Save-AppxPackage Fehler: $($_.Exception.Message)"
            Version       = $pkg.Version
            ExportPath    = $destFolder
            MainFile      = $null
            SHA256        = $null
            ExportedAtUtc = (Get-Date).ToUniversalTime().ToString("s")
        }
    }
}

$manifestPath = Join-Path -Path $ExportRoot -ChildPath "export-manifest.csv"
$manifestEntries | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "---> Manifest geschrieben: $manifestPath"

$failCount = ($manifestEntries | Where-Object Status -ne "OK").Count
if ($failCount -gt 0) {
    Write-Warning "$failCount von $($manifestEntries.Count) Paket(en) konnten nicht vollstaendig exportiert werden. Details siehe Manifest."
}
else {
    Write-Host "Alle $($manifestEntries.Count) Paket(e) erfolgreich exportiert."
}
