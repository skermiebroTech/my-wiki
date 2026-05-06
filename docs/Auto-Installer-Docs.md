# Auto Driver Installer

A fully automated PowerShell driver installer for **Dell**, **HP**, and **Lenovo** machines. Designed to run in Windows audit mode with a single command — no manual downloads, no clicking through wizards.

---

## Quick Start

Run this from the **Win+R** box on any supported machine:

```powershell
powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1 | iex"
```

!!! tip "Audit Mode"
    This command works without any prior software installation. `irm` (Invoke-RestMethod) fetches the script directly from GitHub and `iex` runs it in memory — no file is saved to disk first.

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

!!! note "BOM handling"
    Lenovo's catalog XML contains a UTF-8 BOM that breaks PowerShell's XML parser. The script strips it at the byte level before parsing to prevent this.

#### Dell

Downloads `DriverPackCatalog.cab` from Dell's CDN and extracts it:

```
https://downloads.dell.com/catalog/DriverPackCatalog.cab
```

!!! info "Driver Pack Catalog vs CatalogPC"
    The script uses `DriverPackCatalog.cab` — not `CatalogPC.cab`. The driver pack catalog contains only full SCCM-style driver packs, which is what this tool requires. `CatalogPC.cab` lists individual drivers and firmware and is not used.

The catalog is searched for a `DriverPackage` entry whose `<Model name="...">` attribute exactly matches the WMI model string (case-insensitive). Substring matching is intentionally avoided — for example, "Latitude 7330" will not match a "Latitude 7330 Rugged Extreme" pack.

The Service Tag is read from `Win32_BIOS.SerialNumber` for logging. If no matching pack is found in the catalog, the script opens the Dell support page for that Service Tag automatically.

#### HP

Downloads the HP Driver Pack Matrix HTML page from HP's FTP server:

```
https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html
```

The page is parsed for a table row matching the detected model name. The script extracts a direct SoftPaq `.exe` download URL from the matched row, preferring a **Win10** pack when running on Windows 10. If no match is found the matrix page is opened in the browser for manual selection.

---

### 3. Download

The driver pack is downloaded using `curl.exe` (available on Windows 10 1803+).

The download bar shows live progress including speed:

```
1234.5 MB / 2100.0 MB  (58%)  376.2 Mbps
```

!!! info "Large packs"
    SCCM packs are typically 1.5–3 GB. Downloads retry automatically up to 10 times on failure, resume interrupted transfers with `--continue-at -`, and abort if the connection stalls for more than 3.5 minutes.

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

The extract bar shows a live file count as files land in the destination folder.

!!! note "HP async extraction"
    HP SoftPaqs extract asynchronously — the installer process exits before extraction is complete. The script polls the destination folder for up to 30 seconds after the process exits before falling back to the next format.

---

### 5. Install INFs

All `.inf` files found recursively under the extraction path are installed using:

```powershell
pnputil /add-driver "path\to\driver.inf" /install
```

The install bar tracks progress through each INF. Each is processed whether it results in a new install, an update, or is already current — pnputil handles all cases gracefully.

---

## User Interface

The tool has a simple GUI built with Windows Forms:

| Section | Description |
|---|---|
| **Status box** | Dark terminal-style log with timestamped entries (also saved to `Downloads\`) |
| **Download bar** | Marquee while connecting, shows MB received + speed once underway |
| **Extract bar** | Marquee with live file count during extraction; switches to proportional bar during INF install |
| **Overall bar** | End-to-end progress across all stages |
| **Spinners** | Braille dot spinners on each group box title animate during active work, show ✓ or ✗ on completion |
| **Sound FX** | Sound effects tied to key events (on by default, toggle via checkbox) |

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
2. Kill the active `curl.exe` download process immediately
3. Kill any running extractor process
4. Reset all bars and spinners to a failed state
5. Re-enable the Install button

!!! warning
    Cancellation during extraction may leave a partial `C:\DRIVERS\{OEM}_Extracted` folder. This is safe to delete manually and will be overwritten on the next run.

---

## Log File

Every run writes a timestamped log to:

```
C:\Users\{user}\Downloads\DriverInstaller_YYYYMMDD_HHmmss.log
```

The path is shown at the bottom of the window. The log includes the full device information dump, all download progress entries, extraction file counts, and pnputil output for every INF.

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

Author: Joel Skerman | Date: 07 May 2026
