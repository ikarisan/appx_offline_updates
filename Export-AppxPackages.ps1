<#
.SYNOPSIS
    Laedt ausgewaehlte APPX/MSIX-Pakete inkl. aller Abhaengigkeiten herunter,
    zur spaeteren Offline-Installation auf einem Air-Gapped Zielsystem.

.DESCRIPTION
    Unterstuetzt zwei Paketquellen, automatisch anhand des angegebenen
    Bezeichners erkannt:

    1. Microsoft-Store-Produkt-IDs (z. B. "9WZDNCRFHVN5", oder ein
       kompletter "https://apps.microsoft.com/detail/..."-Link). Diese
       werden DIREKT ueber die oeffentliche Microsoft-Store-Katalog- und
       Windows-Update-Auslieferungsschnittstelle heruntergeladen (dasselbe
       Prinzip wie store.rg-adguard.net) - OHNE Anmeldung. Hintergrund:
       Save-AppxPackage (Appx-Modul) wurde von Microsoft in aktuellen
       Windows-Versionen (24H2+) entfernt, und "winget download" verlangt
       fuer Microsoft-Store-Pakete zwingend ein organisatorisches
       Microsoft-Entra-ID-Konto (persoenliche Konten werden abgelehnt,
       AADSTS500200). Diese Schnittstelle basiert auf einem oeffentlich von
       Microsoft dokumentierten Protokoll (MS-WUSP), wird aber inoffiziell
       fuer diesen Zweck genutzt und kann sich jederzeit aendern.
       Nur fuer kostenlose Pakete geeignet (keine Lizenz-/DRM-Behandlung).

    2. Winget-Paket-IDs aus dem offenen Community-Repository (z. B.
       "Microsoft.WindowsTerminal", erkennbar am Punkt im Namen). Diese
       werden ueber "winget download" geholt (benoetigt keine Anmeldung,
       da nicht ueber den Microsoft Store lizenziert).

    Pro Paket wird ein eigener Unterordner mit der Hauptpaketdatei
    (.msix/.msixbundle) angelegt und nach dem Appx-Paketnamen benannt
    (z. B. "Microsoft.WindowsCalculator"), nicht nach der Store-Produkt-ID
    oder dem urspruenglich angegebenen Bezeichner. Alle Abhaengigkeiten
    werden - ueber alle Pakete hinweg gemeinsam und ohne Duplikate - in
    einem einzigen "Dependencies"-Ordner direkt unterhalb von -ExportRoot
    abgelegt. Zusaetzlich wird ein Manifest (export-manifest.csv) mit
    SHA256-Hashes und - je Paket - der Liste der tatsaechlich benoetigten
    Abhaengigkeitsdateien (Spalte "Dependencies", mehrere durch ";"
    getrennt) geschrieben. Dadurch installiert
    Import-AppxPackagesOffline.ps1 pro Paket ausschliesslich die wirklich
    benoetigten Abhaengigkeiten.

.PARAMETER PackageNames
    Ein oder mehrere Paket-Bezeichner (Store-Produkt-ID, Store-Link oder
    winget-Community-ID, siehe .DESCRIPTION). Wenn angegeben, hat dieser
    Parameter Vorrang: eine eventuell vorhandene packages.txt (Default oder
    ueber -PackageListFile angegeben) wird dann komplett ignoriert.

.PARAMETER PackageListFile
    Pfad zu einer Textdatei mit einem Paket-Bezeichner pro Zeile. Leere
    Zeilen und Zeilen, die mit "#" beginnen, werden ignoriert (Kommentare).
    Wird nur verwendet, wenn -PackageNames NICHT angegeben ist. Ohne
    -PackageListFile wird in diesem Fall standardmaessig nach einer Datei
    "packages.txt" im Skriptverzeichnis gesucht. Existiert weder
    -PackageNames noch eine gueltige Paketliste, bricht das Skript mit
    einem Fehler ab.

.PARAMETER ExportRoot
    Zielordner fuer den Export. Wird angelegt, falls nicht vorhanden.

.PARAMETER SkipExisting
    Prueft vor jedem Download, ob im Zielordner bereits eine Datei mit der
    gleichen Version vorliegt (erkennbar am Dateinamen). Ist das der Fall,
    wird bei Store-Produkt-IDs zusaetzlich die tatsaechliche Dateigroesse
    per HTTP-HEAD-Anfrage mit der bereits vorhandenen Datei verglichen -
    nur bei uebereinstimmender Groesse wird der Download uebersprungen
    (schuetzt vor einer unbemerkt unvollstaendigen/beschaedigten alten
    Datei). Bei winget-Community-Paketen wird nur die Version geprueft
    (kein Groessenvergleich moeglich, da winget die Dateigroesse vorab
    nicht preisgibt). Ohne diesen Schalter wird immer neu heruntergeladen.

.EXAMPLE
    .\Export-AppxPackages.ps1 -PackageNames "9WZDNCRFHVN5" -ExportRoot "D:\AppxExport"

.EXAMPLE
    .\Export-AppxPackages.ps1 -PackageNames "Microsoft.WindowsTerminal","9WZDNCRFHVN5" -ExportRoot "D:\AppxExport"

.EXAMPLE
    # Nutzt automatisch packages.txt im Skriptverzeichnis
    .\Export-AppxPackages.ps1 -ExportRoot "D:\AppxExport"

.EXAMPLE
    # Wiederholter Lauf: bereits vorhandene, unveraenderte Pakete werden uebersprungen
    .\Export-AppxPackages.ps1 -ExportRoot "D:\AppxExport" -SkipExisting

.NOTES
    Fuer winget-Community-Pakete muss winget (App Installer) installiert
    sein. Fuer Store-Produkt-IDs wird nur eine Internetverbindung benoetigt.
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
    [switch]$SkipExisting
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

$script:HttpHeaders = @{ "User-Agent" = "StoreLib" }
$script:MachineArch = switch ($env:PROCESSOR_ARCHITECTURE) {
    "AMD64" { "x64" }
    "ARM64" { "arm64" }
    "x86" { "x86" }
    "ARM" { "arm" }
    default { "x64" }
}
$script:WuCookie = $null
$script:ProcessedDepMonikers = @{}

function Resolve-PackageIdentifier {
    param([string]$Entry)

    # Vollstaendigen Store-Link (z. B. https://apps.microsoft.com/detail/9NBLGGH4QGHW?hl=de)
    # auf die reine Produkt-ID reduzieren.
    if ($Entry -match "apps\.microsoft\.com/detail/([^/?#]+)") {
        return $Matches[1]
    }

    return $Entry
}

function New-ManifestEntry {
    param($PackageName, $Status, $Detail, $Version, $ExportPath, $MainFile, $SHA256, $Dependencies)
    [PSCustomObject]@{
        PackageName   = $PackageName
        Status        = $Status
        Detail        = $Detail
        Version       = $Version
        ExportPath    = $ExportPath
        MainFile      = $MainFile
        SHA256        = $SHA256
        Dependencies  = if ($Dependencies) { (@($Dependencies) -join ";") } else { "" }
        ExportedAtUtc = (Get-Date).ToUniversalTime().ToString("s")
    }
}

function Move-DependenciesToShared {
    <#
        Verschiebt Abhaengigkeits-Paketdateien in den gemeinsamen
        "Dependencies"-Ordner unterhalb von -ExportRoot und gibt die Liste
        der Dateinamen zurueck. Ist eine Datei dort bereits vorhanden (von
        einem anderen Paket), wird das Duplikat verworfen statt erneut
        abgelegt.
    #>
    param($CandidateFiles, [string]$ExportRoot)

    $collected = @()
    $sharedDepFolder = Join-Path -Path $ExportRoot -ChildPath "Dependencies"
    foreach ($file in @($CandidateFiles)) {
        if (-not (Test-Path -Path $sharedDepFolder)) {
            New-Item -Path $sharedDepFolder -ItemType Directory -Force | Out-Null
        }
        $target = Join-Path -Path $sharedDepFolder -ChildPath $file.Name
        if (Test-Path -Path $target) {
            Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
        }
        else {
            Move-Item -Path $file.FullName -Destination $target -Force
        }
        $collected += $file.Name
    }
    return @($collected | Select-Object -Unique)
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

#region winget-Pfad (Community-Pakete)

function Get-ExistingMainFile {
    param([string]$DestFolder)

    $existing = Get-ChildItem -Path $DestFolder -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in ".msixbundle", ".appxbundle" } |
        Select-Object -First 1
    if (-not $existing) {
        $existing = Get-ChildItem -Path $DestFolder -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in ".msix", ".appx" } |
            Select-Object -First 1
    }
    return $existing
}

function Export-PackageViaWinget {
    param([string]$Entry, [string]$Id, [string]$ExportRoot, [switch]$SkipExisting)

    # Bei winget-Community-Paketen ist die ID selbst schon der Paketname
    # (z. B. "Microsoft.WindowsTerminal") - passt direkt als Ordnername.
    $folderSafeName = ($Id -replace '[<>:"/\\|?*]', "_")
    $DestFolder = Join-Path -Path $ExportRoot -ChildPath $folderSafeName
    if (-not (Test-Path -Path $DestFolder)) {
        New-Item -Path $DestFolder -ItemType Directory -Force | Out-Null
    }

    try {
        if ($SkipExisting) {
            $showOutput = & winget show --id $Id --exact --accept-source-agreements 2>&1 | Out-String
            if ($showOutput -match "Version:\s*(\S+)") {
                $availableVersion = $Matches[1]
                if ($availableVersion -ne "Unknown") {
                    $existing = Get-ExistingMainFile -DestFolder $DestFolder
                    if ($existing -and $existing.BaseName -like "*$availableVersion*") {
                        $hash = Get-FileHash -Path $existing.FullName -Algorithm SHA256
                        Write-Host "[SKIP] Bereits vorhanden (Version $availableVersion): $($existing.Name)"

                        # Dependencies-Spalte aus dem vorhandenen Manifest uebernehmen,
                        # damit sie beim Ueberspringen nicht verloren geht (die Dateien
                        # liegen aus dem frueheren Lauf bereits im gemeinsamen Ordner).
                        $prevDeps = $null
                        $prevRow = $script:PreviousManifest |
                            Where-Object { $_.PackageName -eq $Entry -and $_.Dependencies } |
                            Select-Object -First 1
                        if ($prevRow) { $prevDeps = @($prevRow.Dependencies -split ";" | Where-Object { $_ }) }

                        return New-ManifestEntry -PackageName $Entry -Status "OK" -Detail "Uebersprungen (bereits vorhanden, gleiche Version)" -Version $availableVersion -ExportPath $DestFolder -MainFile $existing.Name -SHA256 $hash.Hash -Dependencies $prevDeps
                    }
                }
            }
        }

        $wingetArgs = @(
            "download",
            "--id", $Id,
            "--exact",
            "--accept-package-agreements",
            "--accept-source-agreements",
            "--skip-license",
            "--download-directory", $DestFolder
        )

        $output = & winget @wingetArgs 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
        Write-Host $output.Trim()

        if ($exitCode -ne 0) {
            throw "winget download beendete sich mit Exit-Code $exitCode"
        }

        $mainFile = Get-ExistingMainFile -DestFolder $DestFolder
        if (-not $mainFile) {
            throw "Keine .msixbundle/.appxbundle/.msix/.appx Datei nach dem Download gefunden."
        }

        # Von winget zusaetzlich heruntergeladene Paketdateien sind Abhaengigkeiten.
        # Diese in den gemeinsamen Dependencies-Ordner auf ExportRoot-Ebene verschieben,
        # damit sie nicht pro Paket redundant abgelegt werden.
        $depCandidateFiles = Get-ChildItem -Path $DestFolder -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -ne $mainFile.FullName -and $_.Extension -in ".appx", ".msix", ".appxbundle", ".msixbundle" }
        $collectedDeps = Move-DependenciesToShared -CandidateFiles $depCandidateFiles -ExportRoot $ExportRoot

        $hash = Get-FileHash -Path $mainFile.FullName -Algorithm SHA256
        $version = $null
        if ($mainFile.BaseName -match "_(\d+(?:\.\d+){1,3})_") { $version = $Matches[1] }

        Write-Host "[OK] Heruntergeladen (winget): $($mainFile.Name)"
        return New-ManifestEntry -PackageName $Entry -Status "OK" -Detail "" -Version $version -ExportPath $DestFolder -MainFile $mainFile.Name -SHA256 $hash.Hash -Dependencies $collectedDeps
    }
    catch {
        Write-Warning "[FAIL] winget download fehlgeschlagen fuer '$Entry': $($_.Exception.Message)"
        return New-ManifestEntry -PackageName $Entry -Status "FAIL" -Detail "winget download Fehler: $($_.Exception.Message)" -Version $null -ExportPath $DestFolder -MainFile $null -SHA256 $null
    }
}

#endregion

#region Store-API-Pfad (Produkt-IDs, ohne Anmeldung)

function Get-WuCookie {
    if ($script:WuCookie) { return $script:WuCookie }

    $body = @'
<Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="http://www.w3.org/2003/05/soap-envelope">
  <Header>
    <Action d3p1:mustUnderstand="1" xmlns:d3p1="http://www.w3.org/2003/05/soap-envelope" xmlns="http://www.w3.org/2005/08/addressing">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/GetCookie</Action>
    <MessageID xmlns="http://www.w3.org/2005/08/addressing">urn:uuid:b9b43757-2247-4d7b-ae8f-a71ba8a22386</MessageID>
    <To d3p1:mustUnderstand="1" xmlns:d3p1="http://www.w3.org/2003/05/soap-envelope" xmlns="http://www.w3.org/2005/08/addressing">https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx</To>
    <Security d3p1:mustUnderstand="1" xmlns:d3p1="http://www.w3.org/2003/05/soap-envelope" xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
      <Timestamp xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
        <Created>2017-12-02T00:16:15.210Z</Created>
        <Expires>2017-12-29T06:25:43.943Z</Expires>
      </Timestamp>
      <WindowsUpdateTicketsToken d4p1:id="ClientMSA" xmlns:d4p1="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd" xmlns="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization">
        <TicketType Name="MSA" Version="1.0" Policy="MBI_SSL">
          <User />
        </TicketType>
      </WindowsUpdateTicketsToken>
    </Security>
  </Header>
  <Body>
    <GetCookie xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService">
      <oldCookie>
      </oldCookie>
      <lastChange>2015-10-21T17:01:07.1472913Z</lastChange>
      <currentTime>2017-12-02T00:16:15.217Z</currentTime>
      <protocolVersion>1.40</protocolVersion>
    </GetCookie>
  </Body>
</Envelope>
'@

    $resp = Invoke-WebRequest -Uri "https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx" -Method Post -Body $body -ContentType "application/soap+xml" -Headers $script:HttpHeaders -UseBasicParsing
    [xml]$doc = $resp.Content
    $script:WuCookie = $doc.GetElementsByTagName("EncryptedData")[0].InnerText
    return $script:WuCookie
}

function Get-StoreCatalogInfo {
    param([string]$ProductId)

    $uri = "https://displaycatalog.mp.microsoft.com/v7.0/products/${ProductId}?market=US&languages=en-US,en,neutral"
    $dcat = Invoke-RestMethod -Uri $uri -Method Get -Headers $script:HttpHeaders

    if (-not $dcat.Product) {
        throw "Produkt-ID '$ProductId' wurde im Microsoft Store Katalog nicht gefunden."
    }

    $fulfillment = $dcat.Product.DisplaySkuAvailabilities[0].Sku.Properties.FulfillmentData
    if (-not $fulfillment.WuCategoryId) {
        throw "Produkt-ID '$ProductId' hat keine WuCategoryId (evtl. kein Appx/MSIX-Paket, sondern z. B. ein Add-on)."
    }

    [PSCustomObject]@{
        Title              = $dcat.Product.LocalizedProperties[0].ProductTitle
        WuCategoryId       = $fulfillment.WuCategoryId
        PackageFamilyName  = $fulfillment.PackageFamilyName
        MainName           = ($fulfillment.PackageFamilyName -replace "_[^_]+$", "")
    }
}

function Get-WuSyncCandidates {
    param([string]$WuCategoryId)

    $cookie = Get-WuCookie
    $template = $script:SyncUpdatesTemplate
    $body = $template.Replace("__COOKIE__", $cookie).Replace("__CATEGORYID__", $WuCategoryId)

    $resp = Invoke-WebRequest -Uri "https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx" -Method Post -Body $body -ContentType "application/soap+xml" -Headers $script:HttpHeaders -UseBasicParsing
    $decoded = [System.Net.WebUtility]::HtmlDecode($resp.Content)
    [xml]$doc = $decoded

    $candidates = @()
    foreach ($updateInfo in $doc.GetElementsByTagName("UpdateInfo")) {
        $xmlNode = $updateInfo.GetElementsByTagName("Xml")[0]
        if (-not $xmlNode) { continue }
        $appxMeta = $xmlNode.GetElementsByTagName("AppxMetadata")[0]
        $secFrag = $xmlNode.GetElementsByTagName("SecuredFragment")[0]
        if (-not $appxMeta -or -not $secFrag) { continue }
        $updateIdentity = $xmlNode.GetElementsByTagName("UpdateIdentity")[0]
        $moniker = $appxMeta.Attributes["PackageMoniker"].Value
        $parts = $moniker -split "_"
        if ($parts.Count -lt 3) { continue }
        $candidates += [PSCustomObject]@{
            UpdateId   = $updateIdentity.Attributes["UpdateID"].Value
            RevisionId = $updateIdentity.Attributes["RevisionNumber"].Value
            Moniker    = $moniker
            Name       = $parts[0]
            Version    = $parts[1]
            Arch       = $parts[2]
        }
    }
    return $candidates
}

function Get-WuFileUrls {
    param([string]$UpdateId, [string]$RevisionId)

    $body = $script:FileUrlTemplate.Replace("__UPDATEID__", $UpdateId).Replace("__REVISIONID__", $RevisionId)
    $resp = Invoke-WebRequest -Uri "https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx/secured" -Method Post -Body $body -ContentType "application/soap+xml" -Headers $script:HttpHeaders -UseBasicParsing
    [xml]$doc = $resp.Content

    $urls = @()
    foreach ($fileLocation in $doc.GetElementsByTagName("FileLocation")) {
        $urlNode = $fileLocation.GetElementsByTagName("Url")[0]
        # Die BlockMap-Datei (Datei-Integritaetsmetadaten) hat immer eine 99 Zeichen
        # lange URL - die wollen wir nicht, nur die eigentliche Paketdatei.
        if ($urlNode -and $urlNode.InnerText.Length -ne 99) {
            $urls += $urlNode.InnerText
        }
    }
    return $urls
}

function Select-BestCandidate {
    param([array]$Candidates, [string]$Name)

    $matching = $Candidates | Where-Object { $_.Name -eq $Name }
    if (-not $matching) { return $null }

    $preferred = $matching | Where-Object { $_.Arch -eq "neutral" }
    if (-not $preferred) { $preferred = $matching | Where-Object { $_.Arch -eq $script:MachineArch } }
    if (-not $preferred) { $preferred = $matching }

    $best = $null
    foreach ($c in $preferred) {
        if (-not $best -or (Compare-AppxVersion $c.Version $best.Version) -gt 0) { $best = $c }
    }
    return $best
}

function Get-PackageInfoFromZip {
    <#
        Oeffnet die heruntergeladene Hauptpaketdatei und ermittelt:
        - die korrekte Dateiendung (.msixbundle bei Bundles, sonst .msix)
        - die exakt im Manifest deklarierten Abhaengigkeiten (PackageDependency-Namen)
    #>
    param([string]$ZipPath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $bundleManifestEntry = $zip.Entries | Where-Object { $_.FullName -eq "AppxMetadata/AppxBundleManifest.xml" }

        if ($bundleManifestEntry) {
            $reader = New-Object System.IO.StreamReader($bundleManifestEntry.Open())
            $bundleManifestXml = $reader.ReadToEnd()
            $reader.Close()

            [xml]$bundleDoc = $bundleManifestXml
            $appPackages = $bundleDoc.GetElementsByTagName("Package") | Where-Object { $_.Attributes["Type"].Value -eq "application" }
            $chosen = $appPackages | Where-Object { $_.Attributes["Architecture"] -and $_.Attributes["Architecture"].Value -eq $script:MachineArch } | Select-Object -First 1
            if (-not $chosen) { $chosen = $appPackages | Select-Object -First 1 }

            $depNames = @()
            if ($chosen) {
                $innerFileName = $chosen.Attributes["FileName"].Value
                $innerEntry = $zip.Entries | Where-Object { $_.FullName -eq $innerFileName }
                if ($innerEntry) {
                    $memStream = New-Object System.IO.MemoryStream
                    $innerEntry.Open().CopyTo($memStream)
                    $memStream.Position = 0
                    $innerZip = New-Object System.IO.Compression.ZipArchive($memStream)
                    try {
                        $manifestEntry = $innerZip.Entries | Where-Object { $_.FullName -eq "AppxManifest.xml" }
                        if ($manifestEntry) {
                            $r2 = New-Object System.IO.StreamReader($manifestEntry.Open())
                            $manifestText = $r2.ReadToEnd()
                            $r2.Close()
                            $depNames = @([regex]::Matches($manifestText, '<PackageDependency[^>]*\sName="([^"]+)"') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
                        }
                    }
                    finally { $innerZip.Dispose() }
                }
            }

            return [PSCustomObject]@{ Extension = ".msixbundle"; DependencyNames = $depNames }
        }
        else {
            $manifestEntry = $zip.Entries | Where-Object { $_.FullName -eq "AppxManifest.xml" }
            $depNames = @()
            if ($manifestEntry) {
                $reader = New-Object System.IO.StreamReader($manifestEntry.Open())
                $manifestText = $reader.ReadToEnd()
                $reader.Close()
                $depNames = @([regex]::Matches($manifestText, '<PackageDependency[^>]*\sName="([^"]+)"') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
            }
            return [PSCustomObject]@{ Extension = ".msix"; DependencyNames = $depNames }
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Get-ExistingFileForMoniker {
    param([string]$Folder, [string]$Moniker)

    if (-not (Test-Path -Path $Folder)) { return $null }
    return Get-ChildItem -Path $Folder -File -ErrorAction SilentlyContinue | Where-Object { $_.BaseName -eq $Moniker } | Select-Object -First 1
}

function Test-RemoteSizeMatchesLocal {
    param([string]$Url, [long]$LocalSize)

    try {
        $head = Invoke-WebRequest -Uri $Url -Method Head -UseBasicParsing -Headers $script:HttpHeaders
        $remoteSize = [long]$head.Headers["Content-Length"]
        return $remoteSize -eq $LocalSize
    }
    catch {
        return $false
    }
}

function Export-PackageViaStoreApi {
    param([string]$Entry, [string]$ProductId, [string]$ExportRoot, [switch]$SkipExisting)

    $DestFolder = $null
    try {
        $catalogInfo = Get-StoreCatalogInfo -ProductId $ProductId
        Write-Host "    Store-Katalog: $($catalogInfo.Title) ($($catalogInfo.MainName))"

        # Ordner nach dem Paketnamen benennen (z. B. "Microsoft.WindowsCalculator"),
        # nicht nach der rohen Store-Produkt-ID.
        $folderSafeName = ($catalogInfo.MainName -replace '[<>:"/\\|?*]', "_")
        if (-not $folderSafeName) { $folderSafeName = $ProductId }
        $DestFolder = Join-Path -Path $ExportRoot -ChildPath $folderSafeName
        if (-not (Test-Path -Path $DestFolder)) {
            New-Item -Path $DestFolder -ItemType Directory -Force | Out-Null
        }

        $candidates = @(Get-WuSyncCandidates -WuCategoryId $catalogInfo.WuCategoryId)
        if ($candidates.Count -eq 0) {
            throw "Keine herunterladbaren Paketdateien fuer WuCategoryId $($catalogInfo.WuCategoryId) gefunden."
        }

        $main = Select-BestCandidate -Candidates $candidates -Name $catalogInfo.MainName
        if (-not $main) {
            throw "Kein Hauptpaket mit Namen '$($catalogInfo.MainName)' in den Sync-Ergebnissen gefunden."
        }

        $mainUrls = @(Get-WuFileUrls -UpdateId $main.UpdateId -RevisionId $main.RevisionId)
        if ($mainUrls.Count -eq 0) {
            throw "Keine Download-URL fuer Hauptpaket '$($main.Moniker)' erhalten."
        }

        $existingMain = Get-ExistingFileForMoniker -Folder $DestFolder -Moniker $main.Moniker
        if ($SkipExisting -and $existingMain -and (Test-RemoteSizeMatchesLocal -Url $mainUrls[0] -LocalSize $existingMain.Length)) {
            $mainFilePath = $existingMain.FullName
            $mainFileName = $existingMain.Name
            Write-Host "[SKIP] Hauptpaket bereits vorhanden (gleiche Version/Groesse): $mainFileName"
            $packageInfo = Get-PackageInfoFromZip -ZipPath $mainFilePath
        }
        else {
            $tempMainPath = Join-Path -Path $DestFolder -ChildPath "_main.tmp"
            Invoke-WebRequest -Uri $mainUrls[0] -OutFile $tempMainPath -UseBasicParsing -Headers $script:HttpHeaders

            $packageInfo = Get-PackageInfoFromZip -ZipPath $tempMainPath
            $mainFileName = "$($main.Moniker)$($packageInfo.Extension)"
            $mainFilePath = Join-Path -Path $DestFolder -ChildPath $mainFileName
            Move-Item -Path $tempMainPath -Destination $mainFilePath -Force

            Write-Host "[OK] Hauptpaket heruntergeladen: $mainFileName"
        }

        # Abhaengigkeiten liegen gemeinsam auf ExportRoot-Ebene (nicht pro Paket redundant).
        $depFolder = Join-Path -Path $ExportRoot -ChildPath "Dependencies"
        if ($packageInfo.DependencyNames.Count -gt 0 -and -not (Test-Path -Path $depFolder)) {
            New-Item -Path $depFolder -ItemType Directory -Force | Out-Null
        }
        $collectedDeps = @()
        foreach ($depName in $packageInfo.DependencyNames) {
            $depCandidate = Select-BestCandidate -Candidates $candidates -Name $depName
            if (-not $depCandidate) {
                Write-Warning "    [WARN] Abhaengigkeit '$depName' wurde in den Sync-Ergebnissen nicht gefunden, wird uebersprungen."
                continue
            }
            $existingDep = Get-ExistingFileForMoniker -Folder $depFolder -Moniker $depCandidate.Moniker

            # Bereits in diesem Lauf sichergestellte Abhaengigkeit nicht erneut laden
            # (verhindert redundante Downloads, wenn mehrere Pakete dieselbe
            # Abhaengigkeit benoetigen).
            if ($existingDep -and $script:ProcessedDepMonikers.ContainsKey($depCandidate.Moniker)) {
                Write-Host "    [SKIP] Abhaengigkeit bereits im gemeinsamen Ordner vorhanden: $($existingDep.Name)"
                $collectedDeps += $existingDep.Name
                continue
            }

            $depUrls = @(Get-WuFileUrls -UpdateId $depCandidate.UpdateId -RevisionId $depCandidate.RevisionId)
            if ($SkipExisting -and $existingDep -and $depUrls.Count -gt 0 -and (Test-RemoteSizeMatchesLocal -Url $depUrls[0] -LocalSize $existingDep.Length)) {
                Write-Host "    [SKIP] Abhaengigkeit bereits vorhanden (gleiche Version/Groesse): $($existingDep.Name)"
                $script:ProcessedDepMonikers[$depCandidate.Moniker] = $true
                $collectedDeps += $existingDep.Name
                continue
            }
            foreach ($depUrl in $depUrls) {
                $tempDepPath = Join-Path -Path $depFolder -ChildPath "_dep.tmp"
                Invoke-WebRequest -Uri $depUrl -OutFile $tempDepPath -UseBasicParsing -Headers $script:HttpHeaders
                $depInfo = Get-PackageInfoFromZip -ZipPath $tempDepPath
                $depFileName = "$($depCandidate.Moniker)$($depInfo.Extension)"
                Move-Item -Path $tempDepPath -Destination (Join-Path $depFolder $depFileName) -Force
                Write-Host "    Abhaengigkeit heruntergeladen: $depFileName"
                $collectedDeps += $depFileName
            }
            $script:ProcessedDepMonikers[$depCandidate.Moniker] = $true
        }
        $collectedDeps = @($collectedDeps | Select-Object -Unique)

        $hash = Get-FileHash -Path $mainFilePath -Algorithm SHA256
        return New-ManifestEntry -PackageName $Entry -Status "OK" -Detail "" -Version $main.Version -ExportPath $DestFolder -MainFile $mainFileName -SHA256 $hash.Hash -Dependencies $collectedDeps
    }
    catch {
        # Liegengebliebene Teil-Downloads entfernen
        if ($DestFolder) {
            Remove-Item -Path (Join-Path -Path $DestFolder -ChildPath "_main.tmp") -Force -ErrorAction SilentlyContinue
        }
        Remove-Item -Path (Join-Path -Path (Join-Path -Path $ExportRoot -ChildPath "Dependencies") -ChildPath "_dep.tmp") -Force -ErrorAction SilentlyContinue

        Write-Warning "[FAIL] Store-API-Download fehlgeschlagen fuer '$Entry': $($_.Exception.Message)"
        return New-ManifestEntry -PackageName $Entry -Status "FAIL" -Detail "Store-API Fehler: $($_.Exception.Message)" -Version $null -ExportPath $DestFolder -MainFile $null -SHA256 $null
    }
}

#endregion

# SOAP-Request-Vorlage fuer SyncUpdates. Die lange Liste bekannter Update-Kategorie-IDs
# ist erforderlich, damit der Windows-Update-Dienst die Anfrage als "vollstaendig
# gepatchter" Client akzeptiert; sie enthaelt keine geraete- oder kontospezifischen
# Daten (rein generische Protokoll-Boilerplate).
$script:SyncUpdatesTemplate = @'
<s:Envelope
	xmlns:a="http://www.w3.org/2005/08/addressing"
	xmlns:s="http://www.w3.org/2003/05/soap-envelope">
	<s:Header>
		<a:Action s:mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/SyncUpdates</a:Action>
		<a:MessageID>urn:uuid:175df68c-4b91-41ee-b70b-f2208c65438e</a:MessageID>
		<a:To s:mustUnderstand="1">https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx</a:To>
		<o:Security s:mustUnderstand="1"
			xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
			<Timestamp
				xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
				<Created>2017-08-05T02:03:05.038Z</Created>
				<Expires>2017-08-05T02:08:05.038Z</Expires>
			</Timestamp>
			<wuws:WindowsUpdateTicketsToken wsu:id="ClientMSA"
				xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"
				xmlns:wuws="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization">
				<TicketType Name="MSA" Version="1.0" Policy="MBI_SSL">
					<User />
				</TicketType>
			</wuws:WindowsUpdateTicketsToken>
		</o:Security>
	</s:Header>
	<s:Body>
		<SyncUpdates
			xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService">
			<cookie>
				<Expiration>2045-03-11T02:02:48Z</Expiration>
				<EncryptedData>__COOKIE__</EncryptedData>
			</cookie>
			<parameters>
				<ExpressQuery>false</ExpressQuery>
				<InstalledNonLeafUpdateIDs>
					<int>1</int>
					<int>2</int>
					<int>3</int>
					<int>11</int>
					<int>19</int>
					<int>544</int>
					<int>549</int>
					<int>2359974</int>
					<int>2359977</int>
					<int>5169044</int>
					<int>8788830</int>
					<int>23110993</int>
					<int>23110994</int>
					<int>54341900</int>
					<int>54343656</int>
					<int>59830006</int>
					<int>59830007</int>
					<int>59830008</int>
					<int>60484010</int>
					<int>62450018</int>
					<int>62450019</int>
					<int>62450020</int>
					<int>66027979</int>
					<int>66053150</int>
					<int>97657898</int>
					<int>98822896</int>
					<int>98959022</int>
					<int>98959023</int>
					<int>98959024</int>
					<int>98959025</int>
					<int>98959026</int>
					<int>104433538</int>
					<int>104900364</int>
					<int>105489019</int>
					<int>117765322</int>
					<int>129905029</int>
					<int>130040031</int>
					<int>132387090</int>
					<int>132393049</int>
					<int>133399034</int>
					<int>138537048</int>
					<int>140377312</int>
					<int>143747671</int>
					<int>158941041</int>
					<int>158941042</int>
					<int>158941043</int>
					<int>158941044</int>
					<int>159123858</int>
					<int>159130928</int>
					<int>164836897</int>
					<int>164847386</int>
					<int>164848327</int>
					<int>164852241</int>
					<int>164852246</int>
					<int>164852252</int>
					<int>164852253</int>
				</InstalledNonLeafUpdateIDs>
				<OtherCachedUpdateIDs>
					<int>10</int>
					<int>17</int>
					<int>2359977</int>
					<int>5143990</int>
					<int>5169043</int>
					<int>5169047</int>
					<int>8806526</int>
					<int>9125350</int>
					<int>9154769</int>
					<int>10809856</int>
					<int>23110995</int>
					<int>23110996</int>
					<int>23110999</int>
					<int>23111000</int>
					<int>23111001</int>
					<int>23111002</int>
					<int>23111003</int>
					<int>23111004</int>
					<int>24513870</int>
					<int>28880263</int>
					<int>30077688</int>
					<int>30486944</int>
					<int>30526991</int>
					<int>30528442</int>
					<int>30530496</int>
					<int>30530501</int>
					<int>30530504</int>
					<int>30530962</int>
					<int>30535326</int>
					<int>30536242</int>
					<int>30539913</int>
					<int>30545142</int>
					<int>30545145</int>
					<int>30545488</int>
					<int>30546212</int>
					<int>30547779</int>
					<int>30548797</int>
					<int>30548860</int>
					<int>30549262</int>
					<int>30551160</int>
					<int>30551161</int>
					<int>30551164</int>
					<int>30553016</int>
					<int>30553744</int>
					<int>30554014</int>
					<int>30559008</int>
					<int>30559011</int>
					<int>30560006</int>
					<int>30560011</int>
					<int>30561006</int>
					<int>30563261</int>
					<int>30565215</int>
					<int>30578059</int>
					<int>30664998</int>
					<int>30677904</int>
					<int>30681618</int>
					<int>30682195</int>
					<int>30685055</int>
					<int>30702579</int>
					<int>30708772</int>
					<int>30709591</int>
					<int>30711304</int>
					<int>30715418</int>
					<int>30720106</int>
					<int>30720273</int>
					<int>30732075</int>
					<int>30866952</int>
					<int>30866964</int>
					<int>30870749</int>
					<int>30877852</int>
					<int>30878437</int>
					<int>30890151</int>
					<int>30892149</int>
					<int>30990917</int>
					<int>31049444</int>
					<int>31190936</int>
					<int>31196961</int>
					<int>31197811</int>
					<int>31198836</int>
					<int>31202713</int>
					<int>31203522</int>
					<int>31205442</int>
					<int>31205557</int>
					<int>31207585</int>
					<int>31208440</int>
					<int>31208451</int>
					<int>31209591</int>
					<int>31210536</int>
					<int>31211625</int>
					<int>31212713</int>
					<int>31213588</int>
					<int>31218518</int>
					<int>31219420</int>
					<int>31220279</int>
					<int>31220302</int>
					<int>31222086</int>
					<int>31227080</int>
					<int>31229030</int>
					<int>31238236</int>
					<int>31254198</int>
					<int>31258008</int>
					<int>36436779</int>
					<int>36437850</int>
					<int>36464012</int>
					<int>41916569</int>
					<int>47249982</int>
					<int>47283134</int>
					<int>58577027</int>
					<int>58578040</int>
					<int>58578041</int>
					<int>58628920</int>
					<int>59107045</int>
					<int>59125697</int>
					<int>59142249</int>
					<int>60466586</int>
					<int>60478936</int>
					<int>66450441</int>
					<int>66467021</int>
					<int>66479051</int>
					<int>75202978</int>
					<int>77436021</int>
					<int>77449129</int>
					<int>85159569</int>
					<int>90199702</int>
					<int>90212090</int>
					<int>96911147</int>
					<int>97110308</int>
					<int>98528428</int>
					<int>98665206</int>
					<int>98837995</int>
					<int>98842922</int>
					<int>98842977</int>
					<int>98846632</int>
					<int>98866485</int>
					<int>98874250</int>
					<int>98879075</int>
					<int>98904649</int>
					<int>98918872</int>
					<int>98945691</int>
					<int>98959458</int>
					<int>98984707</int>
					<int>100220125</int>
					<int>100238731</int>
					<int>100662329</int>
					<int>100795834</int>
					<int>100862457</int>
					<int>103124811</int>
					<int>103348671</int>
					<int>104369981</int>
					<int>104372472</int>
					<int>104385324</int>
					<int>104465831</int>
					<int>104465834</int>
					<int>104467697</int>
					<int>104473368</int>
					<int>104482267</int>
					<int>104505005</int>
					<int>104523840</int>
					<int>104550085</int>
					<int>104558084</int>
					<int>104659441</int>
					<int>104659675</int>
					<int>104664678</int>
					<int>104668274</int>
					<int>104671092</int>
					<int>104673242</int>
					<int>104674239</int>
					<int>104679268</int>
					<int>104686047</int>
					<int>104698649</int>
					<int>104751469</int>
					<int>104752478</int>
					<int>104755145</int>
					<int>104761158</int>
					<int>104762266</int>
					<int>104786484</int>
					<int>104853747</int>
					<int>104873258</int>
					<int>104983051</int>
					<int>105063056</int>
					<int>105116588</int>
					<int>105178523</int>
					<int>105318602</int>
					<int>105362613</int>
					<int>105364552</int>
					<int>105368563</int>
					<int>105369591</int>
					<int>105370746</int>
					<int>105373503</int>
					<int>105373615</int>
					<int>105376634</int>
					<int>105377546</int>
					<int>105378752</int>
					<int>105379574</int>
					<int>105381626</int>
					<int>105382587</int>
					<int>105425313</int>
					<int>105495146</int>
					<int>105862607</int>
					<int>105939029</int>
					<int>105995585</int>
					<int>106017178</int>
					<int>106129726</int>
					<int>106768485</int>
					<int>107825194</int>
					<int>111906429</int>
					<int>115121473</int>
					<int>115578654</int>
					<int>116630363</int>
					<int>117835105</int>
					<int>117850671</int>
					<int>118638500</int>
					<int>118662027</int>
					<int>118872681</int>
					<int>118873829</int>
					<int>118879289</int>
					<int>118889092</int>
					<int>119501720</int>
					<int>119551648</int>
					<int>119569538</int>
					<int>119640702</int>
					<int>119667998</int>
					<int>119674103</int>
					<int>119697201</int>
					<int>119706266</int>
					<int>119744627</int>
					<int>119773746</int>
					<int>120072697</int>
					<int>120144309</int>
					<int>120214154</int>
					<int>120357027</int>
					<int>120392612</int>
					<int>120399120</int>
					<int>120553945</int>
					<int>120783545</int>
					<int>120797092</int>
					<int>120881676</int>
					<int>120889689</int>
					<int>120999554</int>
					<int>121168608</int>
					<int>121268830</int>
					<int>121341838</int>
					<int>121729951</int>
					<int>121803677</int>
					<int>122165810</int>
					<int>125408034</int>
					<int>127293130</int>
					<int>127566683</int>
					<int>127762067</int>
					<int>127861893</int>
					<int>128571722</int>
					<int>128647535</int>
					<int>128698922</int>
					<int>128701748</int>
					<int>128771507</int>
					<int>129037212</int>
					<int>129079800</int>
					<int>129175415</int>
					<int>129317272</int>
					<int>129319665</int>
					<int>129365668</int>
					<int>129378095</int>
					<int>129424803</int>
					<int>129590730</int>
					<int>129603714</int>
					<int>129625954</int>
					<int>129692391</int>
					<int>129714980</int>
					<int>129721097</int>
					<int>129886397</int>
					<int>129968371</int>
					<int>129972243</int>
					<int>130009862</int>
					<int>130033651</int>
					<int>130040030</int>
					<int>130040032</int>
					<int>130040033</int>
					<int>130091954</int>
					<int>130100640</int>
					<int>130131267</int>
					<int>130131921</int>
					<int>130144837</int>
					<int>130171030</int>
					<int>130172071</int>
					<int>130197218</int>
					<int>130212435</int>
					<int>130291076</int>
					<int>130402427</int>
					<int>130405166</int>
					<int>130676169</int>
					<int>130698471</int>
					<int>130713390</int>
					<int>130785217</int>
					<int>131396908</int>
					<int>131455115</int>
					<int>131682095</int>
					<int>131689473</int>
					<int>131701956</int>
					<int>132142800</int>
					<int>132525441</int>
					<int>132765492</int>
					<int>132801275</int>
					<int>133399034</int>
					<int>134522926</int>
					<int>134524022</int>
					<int>134528994</int>
					<int>134532942</int>
					<int>134536993</int>
					<int>134538001</int>
					<int>134547533</int>
					<int>134549216</int>
					<int>134549317</int>
					<int>134550159</int>
					<int>134550214</int>
					<int>134550232</int>
					<int>134551154</int>
					<int>134551207</int>
					<int>134551390</int>
					<int>134553171</int>
					<int>134553237</int>
					<int>134554199</int>
					<int>134554227</int>
					<int>134555229</int>
					<int>134555240</int>
					<int>134556118</int>
					<int>134557078</int>
					<int>134560099</int>
					<int>134560287</int>
					<int>134562084</int>
					<int>134562180</int>
					<int>134563287</int>
					<int>134565083</int>
					<int>134566130</int>
					<int>134568111</int>
					<int>134624737</int>
					<int>134666461</int>
					<int>134672998</int>
					<int>134684008</int>
					<int>134916523</int>
					<int>135100527</int>
					<int>135219410</int>
					<int>135222083</int>
					<int>135306997</int>
					<int>135463054</int>
					<int>135779456</int>
					<int>135812968</int>
					<int>136097030</int>
					<int>136131333</int>
					<int>136146907</int>
					<int>136157556</int>
					<int>136320962</int>
					<int>136450641</int>
					<int>136466000</int>
					<int>136745792</int>
					<int>136761546</int>
					<int>136840245</int>
					<int>138160034</int>
					<int>138181244</int>
					<int>138210071</int>
					<int>138210107</int>
					<int>138232200</int>
					<int>138237088</int>
					<int>138277547</int>
					<int>138287133</int>
					<int>138306991</int>
					<int>138324625</int>
					<int>138341916</int>
					<int>138372035</int>
					<int>138372036</int>
					<int>138375118</int>
					<int>138378071</int>
					<int>138380128</int>
					<int>138380194</int>
					<int>138534411</int>
					<int>138618294</int>
					<int>138931764</int>
					<int>139536037</int>
					<int>139536038</int>
					<int>139536039</int>
					<int>139536040</int>
					<int>140367832</int>
					<int>140406050</int>
					<int>140421668</int>
					<int>140422973</int>
					<int>140423713</int>
					<int>140436348</int>
					<int>140483470</int>
					<int>140615715</int>
					<int>140802803</int>
					<int>140896470</int>
					<int>141189437</int>
					<int>141192744</int>
					<int>141382548</int>
					<int>141461680</int>
					<int>141624996</int>
					<int>141627135</int>
					<int>141659139</int>
					<int>141872038</int>
					<int>141993721</int>
					<int>142006413</int>
					<int>142045136</int>
					<int>142095667</int>
					<int>142227273</int>
					<int>142250480</int>
					<int>142518788</int>
					<int>142544931</int>
					<int>142546314</int>
					<int>142555433</int>
					<int>142653044</int>
					<int>143191852</int>
					<int>143258496</int>
					<int>143299722</int>
					<int>143331253</int>
					<int>143432462</int>
					<int>143632431</int>
					<int>143695326</int>
					<int>144219522</int>
					<int>144590916</int>
					<int>145410436</int>
					<int>146720405</int>
					<int>150810438</int>
					<int>151258773</int>
					<int>151315554</int>
					<int>151400090</int>
					<int>151429441</int>
					<int>151439617</int>
					<int>151453617</int>
					<int>151466296</int>
					<int>151511132</int>
					<int>151636561</int>
					<int>151823192</int>
					<int>151827116</int>
					<int>151850642</int>
					<int>152016572</int>
					<int>153111675</int>
					<int>153114652</int>
					<int>153123147</int>
					<int>153267108</int>
					<int>153389799</int>
					<int>153395366</int>
					<int>153718608</int>
					<int>154171028</int>
					<int>154315227</int>
					<int>154559688</int>
					<int>154978771</int>
					<int>154979742</int>
					<int>154985773</int>
					<int>154989370</int>
					<int>155044852</int>
					<int>155065458</int>
					<int>155578573</int>
					<int>156403304</int>
					<int>159085959</int>
					<int>159776047</int>
					<int>159816630</int>
					<int>160733048</int>
					<int>160733049</int>
					<int>160733050</int>
					<int>160733051</int>
					<int>160733056</int>
					<int>164824922</int>
					<int>164824924</int>
					<int>164824926</int>
					<int>164824930</int>
					<int>164831646</int>
					<int>164831647</int>
					<int>164831648</int>
					<int>164831650</int>
					<int>164835050</int>
					<int>164835051</int>
					<int>164835052</int>
					<int>164835056</int>
					<int>164835057</int>
					<int>164835059</int>
					<int>164836898</int>
					<int>164836899</int>
					<int>164836900</int>
					<int>164845333</int>
					<int>164845334</int>
					<int>164845336</int>
					<int>164845337</int>
					<int>164845341</int>
					<int>164845342</int>
					<int>164845345</int>
					<int>164845346</int>
					<int>164845349</int>
					<int>164845350</int>
					<int>164845353</int>
					<int>164845355</int>
					<int>164845358</int>
					<int>164845361</int>
					<int>164845364</int>
					<int>164847387</int>
					<int>164847388</int>
					<int>164847389</int>
					<int>164847390</int>
					<int>164848328</int>
					<int>164848329</int>
					<int>164848330</int>
					<int>164849448</int>
					<int>164849449</int>
					<int>164849451</int>
					<int>164849452</int>
					<int>164849454</int>
					<int>164849455</int>
					<int>164849457</int>
					<int>164849461</int>
					<int>164850219</int>
					<int>164850220</int>
					<int>164850222</int>
					<int>164850223</int>
					<int>164850224</int>
					<int>164850226</int>
					<int>164850227</int>
					<int>164850228</int>
					<int>164850229</int>
					<int>164850231</int>
					<int>164850236</int>
					<int>164850237</int>
					<int>164850240</int>
					<int>164850242</int>
					<int>164850243</int>
					<int>164852242</int>
					<int>164852243</int>
					<int>164852244</int>
					<int>164852247</int>
					<int>164852248</int>
					<int>164852249</int>
					<int>164852250</int>
					<int>164852251</int>
					<int>164852254</int>
					<int>164852256</int>
					<int>164852257</int>
					<int>164852258</int>
					<int>164852259</int>
					<int>164852260</int>
					<int>164852261</int>
					<int>164852262</int>
					<int>164853061</int>
					<int>164853063</int>
					<int>164853071</int>
					<int>164853072</int>
					<int>164853075</int>
					<int>168118980</int>
					<int>168118981</int>
					<int>168118983</int>
					<int>168118984</int>
					<int>168180375</int>
					<int>168180376</int>
					<int>168180378</int>
					<int>168180379</int>
					<int>168270830</int>
					<int>168270831</int>
					<int>168270833</int>
					<int>168270834</int>
					<int>168270835</int>
				</OtherCachedUpdateIDs>
				<SkipSoftwareSync>false</SkipSoftwareSync>
				<NeedTwoGroupOutOfScopeUpdates>true</NeedTwoGroupOutOfScopeUpdates>
				<FilterAppCategoryIds>
					<CategoryIdentifier>
						<Id>__CATEGORYID__</Id>
					</CategoryIdentifier>
				</FilterAppCategoryIds>
				<TreatAppCategoryIdsAsInstalled>true</TreatAppCategoryIdsAsInstalled>
				<AlsoPerformRegularSync>false</AlsoPerformRegularSync>
				<ComputerSpec/>
				<ExtendedUpdateInfoParameters>
					<XmlUpdateFragmentTypes>
						<XmlUpdateFragmentType>Extended</XmlUpdateFragmentType>
					</XmlUpdateFragmentTypes>
					<Locales>
						<string>en-US</string>
						<string>en</string>
					</Locales>
				</ExtendedUpdateInfoParameters>
				<ClientPreferredLanguages>
					<string>en-US</string>
				</ClientPreferredLanguages>
				<ProductsParameters>
					<SyncCurrentVersionOnly>false</SyncCurrentVersionOnly>
					<DeviceAttributes>BranchReadinessLevel=CB;CurrentBranch=rs_prerelease;OEMModel=Virtual Machine;FlightRing=WIS;AttrDataVer=21;SystemManufacturer=Microsoft Corporation;InstallLanguage=en-US;OSUILocale=en-US;InstallationType=Client;FlightingBranchName=external;FirmwareVersion=Hyper-V UEFI Release v2.5;SystemProductName=Virtual Machine;OSSkuId=48;FlightContent=Branch;App=WU;OEMName_Uncleaned=Microsoft Corporation;AppVer=10.0.16184.1001;OSArchitecture=AMD64;SystemSKU=None;UpdateManagementGroup=2;IsFlightingEnabled=1;IsDeviceRetailDemo=0;TelemetryLevel=3;OSVersion=10.0.16184.1001;DeviceFamily=Windows.Desktop;</DeviceAttributes>
					<CallerAttributes>Interactive=1;IsSeeker=0;</CallerAttributes>
					<Products/>
				</ProductsParameters>
			</parameters>
		</SyncUpdates>
	</s:Body>
</s:Envelope>
'@

$script:FileUrlTemplate = @'
<s:Envelope
	xmlns:a="http://www.w3.org/2005/08/addressing"
	xmlns:s="http://www.w3.org/2003/05/soap-envelope">
    <s:Header>
        <a:Action s:mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/GetExtendedUpdateInfo2</a:Action>
        <a:MessageID>urn:uuid:2cc99c2e-3b3e-4fb1-9e31-0cd30e6f43a0</a:MessageID>
        <a:To s:mustUnderstand="1">https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx/secured</a:To>
        <o:Security s:mustUnderstand="1"
			xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
            <Timestamp
				xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
                <Created>2017-08-01T00:29:01.868Z</Created>
                <Expires>2017-08-01T00:34:01.868Z</Expires>
            </Timestamp>
            <wuws:WindowsUpdateTicketsToken wsu:id="ClientMSA"
				xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"
				xmlns:wuws="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization">
                <TicketType Name="MSA" Version="1.0" Policy="MBI_SSL">
                    <User />
                </TicketType>
            </wuws:WindowsUpdateTicketsToken>
        </o:Security>
    </s:Header>
    <s:Body>
        <GetExtendedUpdateInfo2
			xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService">
            <updateIDs>
                <UpdateIdentity>
                    <UpdateID>__UPDATEID__</UpdateID>
                    <RevisionNumber>__REVISIONID__</RevisionNumber>
                </UpdateIdentity>
            </updateIDs>
            <infoTypes>
                <XmlUpdateFragmentType>FileUrl</XmlUpdateFragmentType>
                <XmlUpdateFragmentType>FileDecryption</XmlUpdateFragmentType>
            </infoTypes>
            <deviceAttributes>BranchReadinessLevel=CB;CurrentBranch=rs_prerelease;OEMModel=Virtual Machine;FlightRing=WIS;AttrDataVer=21;SystemManufacturer=Microsoft Corporation;InstallLanguage=en-US;OSUILocale=en-US;InstallationType=Client;FlightingBranchName=external;FirmwareVersion=Hyper-V UEFI Release v2.5;SystemProductName=Virtual Machine;OSSkuId=48;FlightContent=Branch;App=WU;OEMName_Uncleaned=Microsoft Corporation;AppVer=10.0.16184.1001;OSArchitecture=AMD64;SystemSKU=None;UpdateManagementGroup=2;IsFlightingEnabled=1;IsDeviceRetailDemo=0;TelemetryLevel=3;OSVersion=10.0.16184.1001;DeviceFamily=Windows.Desktop;</deviceAttributes>
        </GetExtendedUpdateInfo2>
    </s:Body>
</s:Envelope>
'@

$manifestEntries = @()

if (-not (Test-Path -Path $ExportRoot)) {
    New-Item -Path $ExportRoot -ItemType Directory -Force | Out-Null
}

# Vorhandenes Manifest einlesen: wird beim Schreiben zusammengefuehrt (damit
# Teillaeufe die Eintraege anderer Pakete nicht verwerfen) und liefert bei
# -SkipExisting die Dependencies-Spalte uebersprungener winget-Pakete.
$manifestPath = Join-Path -Path $ExportRoot -ChildPath "export-manifest.csv"
$script:PreviousManifest = @()
if (Test-Path -Path $manifestPath) {
    $script:PreviousManifest = @(Import-Csv -Path $manifestPath)
}

foreach ($entry in $PackageNames) {

    $id = Resolve-PackageIdentifier -Entry $entry

    if ($id -match "\.") {
        Write-Host "---> Verarbeite Paket: $entry (winget-Community-ID: $id)"
        $manifestEntries += Export-PackageViaWinget -Entry $entry -Id $id -ExportRoot $ExportRoot -SkipExisting:$SkipExisting
    }
    else {
        Write-Host "---> Verarbeite Paket: $entry (Store-Produkt-ID: $id)"
        $manifestEntries += Export-PackageViaStoreApi -Entry $entry -ProductId $id -ExportRoot $ExportRoot -SkipExisting:$SkipExisting
    }
}

# Zusammenfassungs-Zahlen beziehen sich nur auf die in DIESEM Lauf
# verarbeiteten Pakete (vor dem Merge mit alten Manifest-Zeilen ermitteln).
$runTotal = $manifestEntries.Count
$failCount = @($manifestEntries | Where-Object Status -ne "OK").Count

# Manifest mergen statt ueberschreiben: Zeilen frueher exportierter Pakete,
# die in diesem Lauf nicht verarbeitet wurden, bleiben erhalten - sonst
# verliert Import-AppxPackagesOffline.ps1 deren Abhaengigkeits-Infos und
# installiert diese Pakete still ohne Abhaengigkeiten.
$allEntries = $manifestEntries
if ($script:PreviousManifest.Count -gt 0) {
    $processedNames = @($manifestEntries | ForEach-Object { $_.PackageName })
    $manifestColumns = "PackageName", "Status", "Detail", "Version", "ExportPath", "MainFile", "SHA256", "Dependencies", "ExportedAtUtc"
    $kept = @($script:PreviousManifest |
        Where-Object { $processedNames -notcontains $_.PackageName } |
        Select-Object -Property $manifestColumns)
    $allEntries = @($kept) + @($manifestEntries)
}
$allEntries | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "---> Manifest geschrieben: $manifestPath ($($allEntries.Count) Eintraege, davon $runTotal aus diesem Lauf)"

if ($failCount -gt 0) {
    Write-Warning "$failCount von $runTotal Paket(en) konnten nicht vollstaendig heruntergeladen werden. Details siehe Manifest."
}
else {
    Write-Host "Alle $runTotal Paket(e) erfolgreich heruntergeladen."
}
