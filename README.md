# Appx Offline Updater

Zwei PowerShell-Skripte, um APPX/MSIX-Pakete (z. B. Standard-Windows-Apps wie
Rechner, Sticky Notes, Paint, Notepad, ...) auf einem Online-Geraet
herunterzuladen und anschliessend auf einem Air-Gapped (netzwerkgetrennten)
Zielsystem offline zu installieren.

- **Export-AppxPackages.ps1** – laedt Pakete inkl. aller Abhaengigkeiten herunter
- **Import-AppxPackagesOffline.ps1** – installiert die heruntergeladenen Pakete systemweit
- **packages.txt** – Liste der zu exportierenden Pakete

## Ablauf

```
1. Export-AppxPackages.ps1  (auf einem PC mit Internetzugang ausfuehren)
        │
        ▼  Ordner mit .msixbundle/.msix-Dateien + export-manifest.csv
        │
   USB-Stick / Netzwerkfreigabe uebertragen
        │
        ▼
2. Import-AppxPackagesOffline.ps1  (auf dem Air-Gapped Zielsystem ausfuehren, als Administrator)
```

---

## 1. Export-AppxPackages.ps1

Laedt Pakete herunter, ohne dass sie auf diesem Geraet installiert sein
muessen. Erkennt automatisch zwei Arten von Paket-Bezeichnern:

| Typ | Beispiel | Wie es geladen wird |
|---|---|---|
| Microsoft-Store-Produkt-ID oder -Link | `9WZDNCRFHVN5` oder `https://apps.microsoft.com/detail/9WZDNCRFHVN5` | Direkt ueber die oeffentliche Microsoft-Store-Katalog-/Windows-Update-Schnittstelle, **ohne Anmeldung** |
| Winget-Community-ID (enthaelt einen Punkt) | `Microsoft.WindowsTerminal` | Ueber `winget download` |

> Hintergrund zum Store-Weg: `Save-AppxPackage` gibt es in aktuellen
> Windows-Versionen (24H2+) nicht mehr, und `winget download` verlangt fuer
> Microsoft-Store-Pakete zwingend ein organisatorisches Microsoft-Entra-ID-Konto
> (private Konten werden abgelehnt). Der Store-Weg funktioniert daher komplett
> ohne Anmeldung, ist aber inoffiziell und koennte sich durch Microsoft
> jederzeit aendern. Nur fuer **kostenlose** Pakete geeignet.

### Verwendung

```powershell
# Nutzt automatisch packages.txt im Skriptverzeichnis
.\Export-AppxPackages.ps1 -ExportRoot "D:\AppxExport"

# Einzelne Pakete direkt angeben (ignoriert packages.txt)
.\Export-AppxPackages.ps1 -PackageNames "9WZDNCRFHVN5","Microsoft.WindowsTerminal" -ExportRoot "D:\AppxExport"

# Andere Paketliste verwenden
.\Export-AppxPackages.ps1 -PackageListFile "D:\meine-pakete.txt" -ExportRoot "D:\AppxExport"

# Wiederholter Lauf: bereits vorhandene, unveraenderte Pakete werden uebersprungen
.\Export-AppxPackages.ps1 -ExportRoot "D:\AppxExport" -SkipExisting
```

### Parameter

| Parameter | Beschreibung |
|---|---|
| `-PackageNames` | Ein oder mehrere Paket-Bezeichner (Store-ID, Store-Link oder winget-ID). Hat Vorrang vor `packages.txt`. |
| `-PackageListFile` | Pfad zu einer eigenen Paketliste. Ohne Angabe wird `packages.txt` im Skriptverzeichnis gesucht. |
| `-ExportRoot` | Zielordner fuer den Export (Default: `.\AppxExport`). Wird angelegt, falls nicht vorhanden. |
| `-SkipExisting` | Ueberspringt Dateien, die bereits mit gleicher Version (und bei Store-Paketen zusaetzlich gleicher Groesse) im Zielordner liegen. Ohne diesen Schalter wird immer neu heruntergeladen. |

### Ergebnis

Pro Paket entsteht unter `-ExportRoot` ein Unterordner, benannt nach dem
Appx-Paketnamen (z. B. `Microsoft.WindowsCalculator`, nicht nach der
Store-ID):

```
AppxExport/
├── export-manifest.csv
├── Microsoft.WindowsCalculator/
│   ├── Microsoft.WindowsCalculator_2021.2605.9.0_neutral_~_8wekyb3d8bbwe.msixbundle
│   └── Dependencies/
│       ├── Microsoft.UI.Xaml.2.8_..._x64....msix
│       ├── Microsoft.VCLibs.140.00_..._x64....msix
│       └── ...
└── Microsoft.WindowsTerminal/
    ├── Windows Terminal_..._x64....msix
    └── Dependencies/
        └── Microsoft.UI.Xaml_..._x64....msix
```

`export-manifest.csv` enthaelt pro Paket Status, Version, Hauptdatei und
SHA256-Hash – wird von `Import-AppxPackagesOffline.ps1 -VerifyHash` genutzt.

---

## 2. packages.txt

Einfache Textdatei, ein Paket-Bezeichner pro Zeile. Erlaubt sind Store-Links,
reine Store-Produkt-IDs oder winget-Community-IDs (gemischt in derselben
Datei moeglich).

```
# Kommentarzeilen beginnen mit "#" und werden ignoriert
# Sticky Notes
https://apps.microsoft.com/detail/9NBLGGH4QGHW?hl=de&gl=DE&ocid=pdpshare

# Paint - vorerst deaktiviert (auskommentiert, wird nicht heruntergeladen)
#https://apps.microsoft.com/detail/9PCFS5B6T72H?hl=de-de&gl=DE&ocid=pdpshare

# Windows Terminal ueber winget-Community-Repo
Microsoft.WindowsTerminal
```

- Leere Zeilen werden ignoriert.
- Eine Zeile mit `#` am Anfang (auch direkt vor einem Link, ohne Leerzeichen)
  wird komplett uebersprungen – so lassen sich einzelne Pakete vorbereiten,
  ohne sie gleich mit herunterzuladen.
- **Store-Produkt-ID finden:** Auf [apps.microsoft.com](https://apps.microsoft.com)
  zur gewuenschten App navigieren und den kompletten Link kopieren (die ID
  wird automatisch aus dem Link herausgelesen) – oder per
  `winget search "<Suchbegriff>" --source msstore` die ID in der Ergebnisliste
  ablesen.

---

## 3. Import-AppxPackagesOffline.ps1

Installiert die von `Export-AppxPackages.ps1` heruntergeladenen Pakete
systemweit (fuer alle bestehenden Benutzer) auf dem Air-Gapped Zielsystem.

### Voraussetzungen

- Muss **als Administrator** ausgefuehrt werden (`-AllUsers` erfordert Elevation).
- Muss in **Windows PowerShell 5.1** laufen (`powershell.exe`), nicht in
  PowerShell 7 (`pwsh.exe`) – in PowerShell 7 gibt es einen bekannten Fehler,
  bei dem `-DependencyPath` fehlerhaft verarbeitet wird. Das Skript bricht
  automatisch ab, wenn es unter PowerShell 7 gestartet wird (siehe `-Force`).

### Verwendung

```powershell
# Empfohlen: mit Hash-Pruefung gegen export-manifest.csv
.\Import-AppxPackagesOffline.ps1 -SourceRoot "D:\AppxExport" -VerifyHash

# Zusaetzlich fuer zukuenftig neu angelegte Benutzer bereitstellen
.\Import-AppxPackagesOffline.ps1 -SourceRoot "D:\AppxExport" -VerifyHash -ProvisionForFutureUsers
```

### Parameter

| Parameter | Beschreibung |
|---|---|
| `-SourceRoot` | Ordner mit den exportierten Paket-Unterordnern (z. B. der per USB uebertragene `AppxExport`-Ordner). |
| `-VerifyHash` | Prueft vor der Installation die SHA256-Pruefsumme jeder Hauptpaketdatei gegen `export-manifest.csv`. Empfohlen. |
| `-ProvisionForFutureUsers` | Registriert das Paket zusaetzlich per `Add-AppxProvisionedPackage`, damit es auch fuer erst spaeter angelegte Benutzerkonten verfuegbar ist. |
| `-Force` | Ueberspringt die PowerShell-5.1-Pruefung (nur verwenden, wenn unter PowerShell 7 getestet). |

Das Skript durchsucht `-SourceRoot` nach Unterordnern, findet darin jeweils
die Hauptpaketdatei (`.msixbundle`/`.appxbundle` bevorzugt, sonst
`.msix`/`.appx`) sowie die Abhaengigkeiten im `Dependencies`-Unterordner und
installiert alles per `Add-AppxPackage -AllUsers`.

---

## Hinweise

- **Microsoft.VCLibs / Microsoft.NET.Native / Microsoft.UI.Xaml** als
  Abhaengigkeiten sind eigenstaendige APPX-Framework-Pakete und unabhaengig
  von einer klassisch installierten Visual-C++-Redistributable
  (`vc_redist.x64.exe` o. ae.) – beide Formen existieren parallel
  nebeneinander, ohne sich zu stoeren.
- Der Store-Download-Weg ist nur fuer **kostenlose** Apps geeignet
  (keine Lizenz-/DRM-Verarbeitung).
