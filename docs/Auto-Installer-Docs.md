# Auto Driver Installer

A fully automated PowerShell driver installer for **Dell**, **HP**, and **Lenovo** machines. Designed to run in Windows audit mode with a single command no manual downloads, no clicking through wizards.

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

The script follows the same four-stage pipeline regardless of manufacturer:

```
Detect → Catalog → Download → Extract → Install
```

### 1. Detect

On launch the script reads the machine's manufacturer and model from WMI:

- **Manufacturer** — from `Win32_ComputerSystem`
- **Model / Machine Type** — from `Win32_ComputerSystemProduct` (Lenovo) or `Win32_ComputerSystem` (Dell/HP)
- The model name is automatically **copied to clipboard** for convenience

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
Downloads `CatalogPC.cab` from Dell's CDN and extracts it:

```
https://downloads.dell.com/catalog/CatalogPC.cab
```

!!! info "Service Tags"
    The Service Tag is read from `Win32_BIOS.SerialNumber` and the System SKU from `Win32_ComputerSystem.SystemSKUNumber`. The catalog is searched for a `SoftwareComponent` with "Driver Pack" in its display name matching the detected SKU.

#### HP
Downloads a platform-specific catalog cab using the BaseBoard product ID:

```
https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/{platformId}/{platformId}.cab
```

!!! info "note"
    The catalog is parsed for a SoftPaq with a "Driver Pack" category, and the direct download URL is extracted.

---

### 3. Download

The driver pack is downloaded using `curl.exe` (available on Windows 10 1803+).

The download bar shows live progress:

```
1234.5 MB Received 376 Mbps
```

!!! info "Large packs"
     SCCM packs are typically ~2 GB. At 100 Mbps this takes around 2–3 minutes. Downloads have no time limit and will automatically resume if the connection drops, retrying up to 10 times before giving up.

---

### 4. Extract

The pack is extracted to `C:\DRIVERS\{OEM}_Extracted` using the appropriate flags per format:

|Format|Extractor|Flags|
|---|---|---|
|`.exe` (Inno Setup)|Pack itself|`/VERYSILENT /DIR="..." /EXTRACT=YES`|
|`.exe` (Dell/HP)|Pack itself|`/s /e="..."` or `/s /e /f "..."`|
|`.zip`|`System.IO.Compression.ZipFile`|Background job|
|`.cab`|`expand.exe`|`-F:* "..."`|

The extract bar shows a live file count as files land in the destination folder. For ZIP files the total is known exactly (pre-scanned via the ZipFile API). For EXE/CAB a running count is displayed.

---

### 5. Install INFs

All `.inf` files found recursively under the extraction path are installed using:

```powershell
pnputil /add-driver "path\to\driver.inf" /install
```

The install bar tracks progress. Each INF is processed whether it results in a new install, an update, or is already up to date — pnputil handles all cases gracefully.

---

## User Interface

The tool has a simple GUI built with Windows Forms:

|Section|Description|
|---|---|
|**Status box**|Dark terminal-style log with timestamped entries (also saved to `Downloads\`)|
|**Download bar**|Marquee while connecting, fills with % + Mbps once underway|
|**Extract bar**|Marquee with live file count during extraction; switches to proportional bar during INF install|
|**Overall bar**|End-to-end progress across all stages|
|**Spinners**|Braille dot spinners on each group box title animate during active work, show ✓ or ✗ on completion|
|**Sound FX**|Space-age sound effects tied to key events (on by default, toggle via checkbox)|

---

## Cancellation

Clicking **Cancel** during any stage will:

1. Set a cancellation flag
2. Kill the active curl download process immediately
3. Kill any running extractor process
4. Reset all bars and spinners
5. Re-enable the Install button

!!! warning "warning"
    Cancellation during extraction may leave a partial `C:\DRIVERS\{OEM}_Extracted` folder. This is safe to delete manually and will be overwritten on the next run.

---

## Log File

Every run appends a timestamped log to:

```
C:\Users\{user}\Downloads\DriverInstaller_YYYYMMDD_HHmmss.log
```

The path is shown at the bottom of the window and is useful for debugging failures.

---

## Sound Events

| Event                 | Sound                     |
| --------------------- | ------------------------- |
| Each INF installed    | Tight keyclick            |
| Installation complete | Triumphant 4-note fanfare |
| Failure / cancelled   | Descending dissonant buzz |

---

## Supported Manufacturers

|Manufacturer|Catalog Source|Pack Format|
|---|---|---|
|Lenovo|`catalogv2.xml` (CDN)|Inno Setup `.exe`|
|Dell|`CatalogPC.cab` (CDN)|Self-extracting `.exe`|
|HP|Platform catalog `.cab` (FTP)|SoftPaq `.exe`|

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
Author: Joel Skerman | Date: 06 May 2026