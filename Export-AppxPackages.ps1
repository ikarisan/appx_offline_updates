<#
.SYNOPSIS
    Laedt ausgewaehlte APPX/MSIX-Pakete inkl. aller Abhaengigkeiten ueber
    "winget download" herunter, zur spaeteren Offline-Installation auf einem
    Air-Gapped Zielsystem.

.DESCRIPTION
    Nutzt "winget download", um pro Paket einen eigenen Unterordner mit dem
    Hauptpaket (.appx/.appxbundle/.msix/.msixbundle) und einem automatisch
    von winget erzeugten "Dependencies"-Unterordner anzulegen. Das Paket muss
    dafuer NICHT auf diesem Geraet installiert sein.

    Hintergrund: Das fruehere Save-AppxPackage-Cmdlet (Appx-Modul) wurde von
    Microsoft in aktuellen Windows-Versionen (24H2 und neuer) entfernt und
    steht nicht mehr zur Verfuegung. "winget download" ist der offizielle
    Ersatz.

    Wichtig: Fuer Pakete, die ueber den Microsoft Store vertrieben werden
    (die meisten vorinstallierten Windows-Apps wie Rechner, Sticky Notes,
    Paint, Fotos, Notepad, To Do, ...), verlangt winget beim Download eine
    interaktive Microsoft-Anmeldung (Microsoft Entra ID). Es kann sich dabei
    ein Anmeldefenster oeffnen, das manuell ausgefuellt werden muss - das
    Skript wartet in diesem Fall, bis der Download abgeschlossen ist.
    Pakete aus dem offenen Winget-Community-Repository (z. B.
    Microsoft.WindowsTerminal) benoetigen dagegen keine Anmeldung.

    Zusaetzlich wird ein Manifest (export-manifest.csv) mit SHA256-Hashes
    geschrieben, damit die Paketintegritaet nach dem Transfer auf dem
    Zielsystem geprueft werden kann (kompatibel mit
    Import-AppxPackagesOffline.ps1 -VerifyHash).

.PARAMETER PackageNames
    Ein oder mehrere Paket-Bezeichner. Erlaubt sind:
      - Microsoft-Store-Produkt-IDs, z. B. "9WZDNCRFHVN5"
      - Winget-Paket-IDs aus dem Community-Repo, z. B. "Microsoft.WindowsTerminal"
      - Vollstaendige Microsoft-Store-Links, z. B.
        "https://apps.microsoft.com/detail/9NBLGGH4QGHW?hl=de&gl=DE"
        (die Produkt-ID wird automatisch aus dem Link extrahiert)
    Wenn dieser Parameter angegeben wird, hat er Vorrang: eine eventuell
    vorhandene packages.txt (Default oder ueber -PackageListFile angegeben)
    wird dann komplett ignoriert.

.PARAMETER PackageListFile
    Pfad zu einer Textdatei mit einem Paket-Bezeichner pro Zeile (siehe
    -PackageNames fuer die erlaubten Formate). Leere Zeilen und Zeilen, die
    mit "#" beginnen, werden ignoriert (Kommentare). Wird nur verwendet,
    wenn -PackageNames NICHT angegeben ist. Ohne -PackageListFile wird in
    diesem Fall standardmaessig nach einer Datei "packages.txt" im
    Skriptverzeichnis gesucht. Existiert weder -PackageNames noch eine
    gueltige Paketliste, bricht das Skript mit einem Fehler ab.

.PARAMETER ExportRoot
    Zielordner fuer den Export. Wird angelegt, falls nicht vorhanden.

.PARAMETER IncludeLicense
    Laedt zusaetzlich die Offline-Lizenzdatei fuer Microsoft-Store-Pakete
    herunter. Erfordert ein Microsoft-Entra-ID-Konto mit der Rolle "Globaler
    Administrator", "Benutzeradministrator" oder "Lizenzadministrator".
    Ohne diesen Schalter wird die Lizenz uebersprungen (--skip-license) -
    ausreichend fuer kostenlose Standard-Apps.

.EXAMPLE
    .\Export-AppxPackages.ps1 -PackageNames "9WZDNCRFHVN5" -ExportRoot "D:\AppxExport"

.EXAMPLE
    .\Export-AppxPackages.ps1 -PackageNames "Microsoft.WindowsTerminal","9WZDNCRFHVN5" -ExportRoot "D:\AppxExport"

.EXAMPLE
    # Nutzt automatisch packages.txt im Skriptverzeichnis
    .\Export-AppxPackages.ps1 -ExportRoot "D:\AppxExport"

.EXAMPLE
    .\Export-AppxPackages.ps1 -PackageListFile "D:\pakete.txt" -ExportRoot "D:\AppxExport"

.NOTES
    Voraussetzung: winget (App Installer) muss installiert sein. Bei
    Microsoft-Store-Paketen kann waehrend des Downloads ein interaktives
    Anmeldefenster erscheinen - das Skript ist in diesem Fall nicht
    vollstaendig unbeaufsichtigt lauffaehig.
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
    [switch]$IncludeLicense
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

function Resolve-WingetPackageId {
    param([string]$Entry)

    # Vollstaendigen Store-Link (z. B. https://apps.microsoft.com/detail/9NBLGGH4QGHW?hl=de)
    # auf die reine Produkt-ID reduzieren.
    if ($Entry -match "apps\.microsoft\.com/detail/([^/?#]+)") {
        return $Matches[1]
    }

    return $Entry
}

$manifestEntries = @()

if (-not (Test-Path -Path $ExportRoot)) {
    New-Item -Path $ExportRoot -ItemType Directory -Force | Out-Null
}

foreach ($entry in $PackageNames) {

    $id = Resolve-WingetPackageId -Entry $entry
    $folderSafeId = ($id -replace '[<>:"/\\|?*]', "_")

    Write-Host "---> Verarbeite Paket: $entry$(if ($id -ne $entry) { " (Store-ID: $id)" })"

    $destFolder = Join-Path -Path $ExportRoot -ChildPath $folderSafeId

    try {
        if (-not (Test-Path -Path $destFolder)) {
            New-Item -Path $destFolder -ItemType Directory -Force | Out-Null
        }

        $wingetArgs = @(
            "download",
            "--id", $id,
            "--exact",
            "--accept-package-agreements",
            "--accept-source-agreements",
            "--download-directory", $destFolder
        )
        if (-not $IncludeLicense) {
            $wingetArgs += "--skip-license"
        }

        $output = & winget @wingetArgs 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
        Write-Host $output.Trim()

        if ($exitCode -ne 0) {
            throw "winget download beendete sich mit Exit-Code $exitCode"
        }

        # Hauptpaketdatei ermitteln (liegt direkt im Zielordner, nicht im
        # von winget automatisch erzeugten Dependencies-Unterordner)
        $mainFile = Get-ChildItem -Path $destFolder -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in ".msixbundle", ".appxbundle" } |
            Select-Object -First 1

        if (-not $mainFile) {
            $mainFile = Get-ChildItem -Path $destFolder -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -in ".msix", ".appx" } |
                Select-Object -First 1
        }

        if (-not $mainFile) {
            throw "Keine .msixbundle/.appxbundle/.msix/.appx Datei nach dem Download gefunden."
        }

        $hash = Get-FileHash -Path $mainFile.FullName -Algorithm SHA256

        $version = $null
        if ($mainFile.BaseName -match "_(\d+(?:\.\d+){1,3})_") {
            $version = $Matches[1]
        }

        Write-Host "[OK] Heruntergeladen: $($mainFile.Name)"

        $manifestEntries += [PSCustomObject]@{
            PackageName   = $entry
            Status        = "OK"
            Detail        = ""
            Version       = $version
            ExportPath    = $destFolder
            MainFile      = $mainFile.Name
            SHA256        = $hash.Hash
            ExportedAtUtc = (Get-Date).ToUniversalTime().ToString("s")
        }
    }
    catch {
        Write-Warning "[FAIL] Download fehlgeschlagen fuer '$entry': $($_.Exception.Message)"
        $manifestEntries += [PSCustomObject]@{
            PackageName   = $entry
            Status        = "FAIL"
            Detail        = "winget download Fehler: $($_.Exception.Message)"
            Version       = $null
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

$failCount = @($manifestEntries | Where-Object Status -ne "OK").Count
if ($failCount -gt 0) {
    Write-Warning "$failCount von $($manifestEntries.Count) Paket(en) konnten nicht vollstaendig heruntergeladen werden. Details siehe Manifest."
}
else {
    Write-Host "Alle $($manifestEntries.Count) Paket(e) erfolgreich heruntergeladen."
}
