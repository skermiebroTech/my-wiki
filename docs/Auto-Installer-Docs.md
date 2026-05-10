# Auto Driver Installer

A fully automated PowerShell driver installer for **Dell**, **HP**, and **Lenovo** machines. Designed to run in Windows audit mode with a single command
code can be found here: https://github.com/skermiebroTech/my-wiki/raw/refs/heads/main/Install-Drivers-auto.ps1

---

## Quick Start

Run this from the **Win+R** box on any supported machine:

```powershell
powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1 | iex"
```

!!! tip ""
    This command works without any prior software installation. `irm` (Invoke-RestMethod) fetches the script directly from GitHub and `iex` runs it in memory — no file is saved to disk.

---

## How It Works

The script follows the same pipeline regardless of manufacturer:

```
Detect → Catalog → Download → Extract → Install
```

### 1. Detect

On launch the script reads the machine's manufacturer and model from WMI:

- **Manufacturer** — from `Win32_ComputerSystem`
- **Model / Machine Type** — from `Win32_ComputerSystemProduct` (Lenovo) or `Win32_ComputerSystem` (Dell/HP)
- The model name is automatically **copied to clipboard** for convenience
- A full **device information dump** is written to the log at startup, covering CPU, GPU, RAM, disks, NICs, battery, PnP problem devices, and driver store contents
- Log is stored in the current users downloads folder

If the script is not running as Administrator it re-launches itself elevated automatically.

---

### 2. Catalog Lookup

Each OEM uses a different catalog source:

#### Lenovo

Downloads `catalogv2.xml` directly from Lenovo's CDN:

```
https://download.lenovo.com/cdrt/td/catalogv2.xml
```

The first 4 characters of the machine type (e.g. `20W0` for a ThinkPad T14 Gen 2i) are used to match the correct model entry. The script prefers a **Win11** SCCM pack and falls back to **Win10** automatically if one isn't available.

!!! note ""
    Lenovo's catalog XML contains a UTF-8 BOM that breaks PowerShell's XML parser. The script strips it at the byte level before parsing to prevent this.

#### Dell

Downloads `DriverPackCatalog.cab` from Dell's CDN and extracts it:

```
https://downloads.dell.com/catalog/DriverPackCatalog.cab
```

The catalog is searched for a `DriverPackage` entry whose `<Model name="...">` attribute exactly matches the WMI model string (case-insensitive). Substring matching is intentionally avoided — for example, "Latitude 7330" will not match a "Latitude 7330 Rugged Extreme" pack.

The Service Tag is read from `Win32_BIOS.SerialNumber` for logging. If no matching pack is found in the catalog, the script opens the Dell support page for that Service Tag automatically.

#### HP

Downloads the HP Driver Pack Matrix HTML page from HP's FTP server:

```
https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html
```

The page is parsed for a table row matching the detected model name. The script extracts a direct SoftPaq `.exe` download URL from the matched row, preferring a **Win11** pack and falls back to **Win10** automatically if one isn't available.

---

### 3. Download

The driver pack is downloaded using `curl.exe` (available on Windows 10 1803+).

The download bar shows live progress including speed:

```
[Time] 500 MB received 376.2 Mbps
```

---

### 4. Extract

The pack is extracted to `C:\DRIVERS\{OEM}_Extracted` using the appropriate method per format:

| Format | Method | Flags |
|---|---|---|
| `.exe` (Lenovo / Inno Setup) | Pack itself | `/VERYSILENT /DIR="..." /EXTRACT=YES` |
| `.exe` (Dell) | Pack itself | `/s /e="..."` |
| `.exe` (HP SoftPaq) | Pack itself | `/s /e /f "..."` (async) |
| `.zip` | `System.IO.Compression.ZipFile` | Background job |
| `.cab` | `expand.exe` | `-F:* "..."` |

The vendor-specific extraction format is always tried first. If it produces zero files, the remaining formats are tried as fallbacks in order. A final legacy fallback of `-s -f"<dest>"` is attempted for older Lenovo packs if all other formats fail.

---

### 5. Install INFs

All `.inf` files found recursively under the extraction path are installed using:

```powershell
pnputil /add-driver "path\to\driver.inf" /install
```

---

## Sound Events

| Event | Sound file |
|---|---|
| Script started | `Windows Notify.wav` |
| Download complete | `Windows Print complete.wav` |
| Extraction complete | `Windows Print complete.wav` |
| Each INF installed | `Windows Navigation Start.wav` |
| Installation complete | `Windows Logon.wav` |
| Failure / cancelled | `Windows Critical Stop.wav` |

Sounds are played asynchronously using `System.Media.SoundPlayer`. All files are sourced from `C:\Windows\Media\` — no external audio files are required. Each event has a preference list of fallback filenames in case a specific file is absent (e.g. on minimal or freshly imaged systems).

---

## Cancellation

Clicking **Cancel** during any stage will:

1. Set a cancellation flag checked between all major operations
2. Re enable the Install Drivers button

!!! warning
    Cancellation during extraction may leave a partial `C:\DRIVERS\{OEM}_Extracted` folder. This is safe to delete manually and will be overwritten on the next run.

---

## Log File

Every run writes a timestamped log to:

```
C:\Users\{user}\Downloads\DriverInstaller_YYYYMMDD_HHmmss.log
```

The path is shown at the bottom of the window. The log includes the full device information dump, all download progress entries, extraction file counts, and pnputil output for every INF.
The log file is deleted automatically when the device is syspreped 

---

## Supported Manufacturers

| Manufacturer | Catalog Source | Pack Format |
|---|---|---|
| Lenovo | `catalogv2.xml` (CDN) | Inno Setup `.exe` |
| Dell | `DriverPackCatalog.cab` (CDN) | Self-extracting `.exe` |
| HP | `HP_Driverpack_Matrix_x64.html` (FTP) | SoftPaq `.exe` |

!!! warning "Unsupported machines"
    If the manufacturer is not Dell, HP, or Lenovo the script will display an error and exit gracefully without making any changes to the system.

---

## Requirements

- Windows 10 1803 or later (for inbox `curl.exe`)
- Administrator privileges (auto-elevated on launch)
- Internet access to OEM CDN/FTP servers
- ~3–5 GB free disk space in `C:\DRIVERS\` for download + extraction

!!! warning "Use outside of audit mode is currently untested"

---

Author: Joel Skerman | Date: 07 May 2026 | Modified: 11 May 2026
