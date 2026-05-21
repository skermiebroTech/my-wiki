# =============================================================
# Install-Drivers-auto.ps1
# Version: 1.13.3
# Author:  skermiebroTech
# Repo:    https://github.com/skermiebroTech/my-wiki
#
# Run from Win+R in audit mode (GUI):
#   powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1 | iex"
#
# Run headlessly with arguments:
#   powershell -ExecutionPolicy Bypass -File Install-Drivers-auto.ps1 -Manufacturer Dell
#   powershell -ExecutionPolicy Bypass -File Install-Drivers-auto.ps1 -Manufacturer HP -Model "EliteBook x360 1030 G8 Notebook PC"
#   powershell -ExecutionPolicy Bypass -File Install-Drivers-auto.ps1 -Manufacturer Lenovo -MachineType 20XX -SkipInstall -SkipCleanup
#   powershell -ExecutionPolicy Bypass -File Install-Drivers-auto.ps1 -TestMode -Diagnostic
#   powershell -ExecutionPolicy Bypass -File Install-Drivers-auto.ps1 -Silent -NoAnalytics
#   powershell -ExecutionPolicy Bypass -File Install-Drivers-auto.ps1 -MaxParallelDownloads 5
#   powershell -ExecutionPolicy Bypass -File Install-Drivers-auto.ps1 -PromptWindowsUpdate:$false
#
# Parameters:
#   -Manufacturer  Override WMI manufacturer detection (Dell, HP, Lenovo, Microsoft)
#   -Model         Override WMI model detection
#   -Headless      Skip GUI, write to console only (auto-set when any param is passed)
#   -Silent        Like -Headless but ALSO suppresses console output (log file only).
#                  Forces -Headless on. Useful when launched from another script that
#                  shouldn't have its stdout polluted, or for unattended scheduled tasks.
#   -TestMode      Dry-run: detect manufacturer, snapshot missing drivers, fetch catalog
#                  metadata, but DO NOT download, extract, or install anything that
#                  mutates the system. Analytics still sent with result="testmode"
#                  (unless -NoAnalytics). Safe to run on production machines.
#   -Diagnostic    Verbose logging: extra environment dump, expanded network timings,
#                  per-step duration logging, full pnputil enum on entry. Doubles log
#                  size but is invaluable for triaging vendor-catalog regressions.
#   -NoAnalytics   Disable the Google Sheets webhook entirely. Local analytics JSON
#                  export is still written to Downloads. Use for offline / air-gapped
#                  machines or when testing without polluting production telemetry.
#   -MaxParallelDownloads <int>  v1.11.0 - concurrency cap for parallel downloads
#                  in the HP per-machine catalog and Dell CatalogPC individual-driver
#                  paths. Default 3. Clamped to 1..6 (above 6 just thrashes the CDN
#                  and triggers rate limiting). Has no effect on single-file full
#                  driver pack downloads.
#   -SkipInstall   Download and extract only, skip pnputil driver installation
#   -SkipCleanup   Keep C:\DRIVERS after run for inspection
#   -PromptWindowsUpdate  v1.12.1 - ENABLED BY DEFAULT. When drivers remain missing
#                  at end of install, offer to open Windows Update to search for
#                  additional drivers. In GUI mode, shows a Yes/No/Cancel dialog.
#                  In headless mode, automatically opens Windows Update if drivers
#                  remain unresolved. Pass -PromptWindowsUpdate:$false to disable.
#
# Supports: Dell, HP, Lenovo, Microsoft (Surface)
#
# Output files written to %USERPROFILE%\Downloads\ (timestamped, one set per run):
#   DriverInstaller_<ts>.log         - human-readable text log (always)
#   DriverInstaller_<ts>.events.json - NDJSON structured event log (always)
#   DriverInstaller_<ts>.analytics.json - final analytics payload (always)
#   DriverInstaller_<ts>.report.html - install summary report (on completion)
#
# v1.13.3 - Dell: form-factor suffix fallback in driver-pack search. The Dell
#           DriverPackCatalog occasionally lists a model under its base name
#           only (e.g. "Latitude 5310") with no separate entry for the
#           convertible variant ("Latitude 5310 2-in-1"), even though the
#           convertible exists as a real SKU and ships with its own service
#           tag. Previously Start-DellDriverInstall did a case-insensitive
#           equality match on the model name with no normalisation beyond
#           stripping a trailing " Notebook..." or a leading "Dell " - so a
#           "2-in-1" suffix would cause zero catalog candidates, the function
#           would return false, and the operator would land on the Dell
#           support page with nothing installed (observed on Latitude 5310
#           2-in-1 / SKU 099E with 14 missing devices).
#           Fix: factored the candidate-collection loop into a scriptblock so
#           it can be re-run. The primary search (full model string + existing
#           cleanups) runs first. If it returns zero candidates AND the model
#           name carries a "2-in-1" suffix (tolerant of "2 in 1", "2in1",
#           "2-in-1" spacing), the search is retried with the suffix stripped.
#           The primary search runs first so that when BOTH variants exist in
#           the catalog the convertible-specific pack still wins (different
#           chassis -> different touchscreen/digitizer/sensor drivers). Only
#           the Dell full-pack path is affected: the individual-driver
#           CatalogPC path (1-3 missing devices) keys off SystemSKUNumber, not
#           model name, and is unchanged.
# v1.13.2 - Spreadsheet payload: missing_before_list replaced with driver_urls.
#           The Google Sheets webhook field "missing_before_list" (the list of
#           device descriptions present BEFORE the install run) is dropped from
#           the JSON payload and replaced with "driver_urls" - a comma-separated
#           list of every kind='driver' URL recorded by v1.13.1's
#           AnalyticsDownloadUrls list (vendor catalog/descriptor/matrix URLs
#           are excluded). The list still goes through the existing
#           cleanForJsonList helper (control-char strip, JSON-escape, 8000-char
#           cap) so the cell stays inside Google Sheets' per-cell limits.
#           Rationale: actionable URLs that an operator can re-fetch are more
#           valuable in the spreadsheet than device descriptions that are
#           already implicit in inf_count / missing_before / missing_after.
#           IMPACT: the Apps Script doPost handler + Google Sheet column header
#           will need updating - rename the column from "missing_before_list"
#           (or whatever local name) to "driver_urls" and add the new key to
#           the row builder. The local artefacts are unchanged:
#           $script:AnalyticsMissingBeforeList is still populated and still
#           drives the "Missing drivers before run" section of the HTML report,
#           and the local analytics JSON is regenerated from the new payload
#           shape (so old vs new JSON files will differ - expected).
# v1.13.1 - Re-downloadable driver URLs in the HTML report. Every file
#           pulled by Invoke-CurlDownload / Invoke-CurlDownloadParallel is
#           now recorded (URL + filename + size) and classified as either
#           a DRIVER payload (.exe/.msi/.cab driver pack, softpaq, MSI) or
#           CATALOG metadata (vendor catalog .cab, descriptor/catalog .xml,
#           HP matrix .html). The HTML report gains a "Driver download URLs"
#           section listing each driver file as a clickable link plus the
#           raw URL text (so it survives copy/paste from a printed/saved
#           report), letting an operator re-fetch the exact same driver
#           pack later WITHOUT re-running the script or re-resolving the
#           vendor catalog. A smaller muted "Catalog / metadata sources"
#           list is included below it for traceability. New script-scope
#           list $script:AnalyticsDownloadUrls + helper Add-DownloadRecord
#           (dedupes by URL). Reset per run alongside the other Analytics
#           lists. No analytics-payload/webhook shape change - this is a
#           local-report-only addition; the Google Sheet schema is
#           untouched so existing dashboards keep working. No vendor
#           resolution logic touched - purely an observability/record add.
# v1.13.0 - Surface: CPU-aware driver pack selection for dual-CPU models.
#           Surface Laptop 3 and Surface Laptop 4 each ship as two distinct
#           hardware variants (Intel vs AMD) with SEPARATE Microsoft Download
#           Center packs, but only one model string ("Surface Laptop 4") is
#           reported by WMI for both. Pre-v1.13.0 the table mapped that single
#           string to the Intel pack unconditionally, so AMD units (e.g. the
#           "AMD Ryzen ... Microsoft Surface Edition") downloaded the wrong
#           Intel MSI - none of the INFs matched the AMD platform devices and
#           the run produced 0 installed drivers.
#           Fix: new Get-CpuVendor helper reads Win32_Processor.Manufacturer
#           (AuthenticAMD / GenuineIntel) with a Name-string + env fallback.
#           New $SurfaceCpuVariantIds map holds the Intel/AMD Download Center
#           IDs for the dual-CPU models. After the model -> ID match in
#           Start-MicrosoftSurfaceDriverInstall, if the matched model has CPU
#           variants the page ID is overridden to the pack matching the
#           detected CPU vendor. Confirmed IDs: SL4 Intel=102924 AMD=102923,
#           SL3 Intel=100429 AMD=100428. $SurfaceDownloadIds keeps the Intel
#           ID as the default so behaviour on Intel / undetectable-CPU units
#           is unchanged; only AMD units are redirected. The picker dropdown
#           and headless .Keys matchers are unaffected (keys stay model
#           strings). No other vendor logic touched.
# v1.12.1 - Windows Update feature enabled by default + centralized version variable.
#           (1) -PromptWindowsUpdate now defaults to $true (enabled) instead of requiring
#               the flag. Users who want to disable it can pass -PromptWindowsUpdate:$false.
#           (2) Single $SCRIPT_VERSION variable at top of script (line ~368) is now the
#               only place version needs updating. All references throughout the script
#               ($form text, analytics, logs, HTML report) automatically use this value.
#               Eliminates multi-place version bumps and sync errors.
# v1.12.0 - Windows Update integration. New -PromptWindowsUpdate switch enables
#           end-of-run Windows Update prompting when drivers remain unresolved.
#           On success: If any drivers are still missing, GUI shows a 3-button
#           dialog (Yes/No/Cancel) to open Windows Update, skip, or reboot.
#           Headless mode automatically opens Windows Update (unless -Silent).
#           On failure: Similar prompt offered if any drivers are still missing.
#           Helpful for cases where vendor driver packs don't cover all hardware
#           but Windows Update has additional matches in its catalog.
#           New helper function Open-WindowsUpdate launches ms-settings:windowsupdate.
# v1.11.1 - Analytics error legibility. When the Google Apps Script doPost
#           throws server-side, Apps Script returns HTTP 200 with an HTML
#           error page instead of plain "OK". Pre-v1.11.1 the script dumped
#           the entire multi-line HTML into the .log, which buried the actual
#           error message. v1.11.1 extracts the message (e.g. "ReferenceError:
#           Cannot access 'data' before initialization (line 4, file Code)")
#           and logs it on its own line at error level, with a pointer to the
#           Apps Script editor. The full HTML body is still captured in the
#           .events.json file for deep debugging. No client-side payload
#           changes - this is purely better error surfacing for server-side
#           Apps Script failures.
# v1.11.0 - Parallel downloads + Lenovo Recipe Card priority restored.
#           (1) Parallel curl downloads via new Invoke-CurlDownloadParallel helper.
#               Takes an array of {Url, OutFile, Label} items and runs up to
#               $MaxParallelDownloads curl.exe processes concurrently, polling
#               them together with the same per-process stall detection
#               (~3.5 min) and cancel handling as the serial download. On cancel:
#               every active curl is killed and every partial file deleted in one
#               pass before returning - same "no broken state" guarantee as the
#               v1.9.2 serial cancel fix.
#               Wired into TWO call sites where it actually helps:
#               - HP per-machine catalog softpaq loop (was the worst serial
#                 offender - N small softpaqs each from a fresh CDN connection,
#                 paying TCP slow-start N times)
#               - Dell CatalogPC individual-driver loop (same shape: 1-3 small
#                 INF cabinets each from a separate CDN edge)
#               NOT wired into: full driver pack downloads (single file, nothing
#               to parallelise), descriptor XMLs, BIOS update CABs (those are
#               one-shot and need to finish before anything else can plan).
#               Analytics: parallel batches SUM their bytes into AnalyticsDownloadMB
#               (more accurate for multi-file paths). Serial Invoke-CurlDownload
#               keeps its long-standing MAX semantics so existing single-pack
#               analytics rows don't change shape - documented quirk, not unified
#               to avoid changing downstream sheet/dashboard semantics.
#           (2) Lenovo: PATH PRIORITY REVERSED.
#               Pre-v1.11.0:  consumer catalog (PATH 1) -> catalogv2.xml (PATH 2 fallback)
#               v1.11.0+:    catalogv2.xml / Recipe Card (PATH 1) -> consumer catalog (PATH 2 fallback)
#               The consumer-catalog stack (v1.8.0 onwards: DetectInstall evaluator,
#               NVIDIA short-circuit, BIOS deferral, force-bind unbound, dock skip)
#               is ALL preserved; it just only runs now when the machine type
#               has no Recipe Card / catalogv2.xml entry.
#               Motivation: ThinkPad / ThinkCentre / commercial models that DO
#               have Recipe Card packs are far more reliable on the single-pack
#               path (one Inno Setup extract -> pnputil INFs) than the consumer
#               flow's per-package descriptor dance. The consumer path stays the
#               only option for ThinkBook / IdeaPad / consumer models that
#               aren't in catalogv2.xml (the use case v1.8.0 originally added).
#               Last-resort fallback unchanged: if neither catalog has the
#               machine type, opens RecipeCardWeb.html in a browser for manual
#               selection (skipped in headless mode, URL logged instead).
#               Refactor: factored the v1.7.x inline catalogv2.xml fetch+parse
#               into Get-LenovoFullPackUrl, and the download+extract+install into
#               Install-LenovoFullPack. Start-LenovoDriverInstall is now ~30
#               lines of clear dispatch logic instead of two interleaved paths.
# v1.10.0 - New switches and observability layer (additive; no vendor logic changed):
#           (1) -Silent: log-only mode. Disables both GUI and console output; the log
#               file is still written. Implies -Headless. Internally, the Log() helper
#               now respects $script:Silent and short-circuits the Write-Host branch.
#           (2) -TestMode: dry-run. Start-Install short-circuits before each vendor
#               install function and reports what WOULD be done. Analytics records
#               result="testmode" so dry-runs are distinguishable in the Sheet.
#           (3) -Diagnostic: verbose mode. Existing Write-DeviceInfo is unchanged;
#               new Write-VerboseDiagnostics adds per-run timings, expanded curl/
#               network info, and a final perf summary. Controlled by $script:Diagnostic.
#           (4) -NoAnalytics: skips the Sheets webhook POST. Local
#               .analytics.json is still written so the data isn't lost if you
#               later decide to backfill.
#           (5) Structured NDJSON event log (DriverInstaller_<ts>.events.json):
#               every Log() call also writes a {ts, level, event, msg, ...ctx} JSON
#               line. Makes downstream tooling (jq, log shipping, anomaly detection)
#               trivial without disturbing the existing human log format.
#           (6) Analytics JSON saved to Downloads (previously written to %TEMP%
#               and deleted after the curl POST). Now persists alongside the log.
#           (7) HTML install report (DriverInstaller_<ts>.report.html) generated
#               at end-of-run. Self-contained single file - useful for tech-stack
#               sign-off, customer hand-over, or as an artifact attached to a
#               support ticket.
#           No changes to: Dell/HP/Lenovo/Surface vendor logic, curl download loop,
#           extraction watchers, pnputil bind-marker parser, or DetectInstall
#           evaluator. The new features compose with all existing behaviour.
# v1.9.2 - Cancel: honour the cancel flag mid-download. In v1.9.1 and prior,
#           Invoke-CurlDownload's poll loop only watched for $proc.HasExited
#           and the 3.5-min stall timer - $script:CancelRequested was checked
#           by callers but never inside the loop itself. So a user clicking
#           Cancel during a 2 GB driver pack download would set the flag and
#           log "Cancel requested by user", but curl kept streaming until it
#           finished naturally (real case: ~60 MB and 3 seconds of additional
#           download after the click in the 1.9.0 EliteBook 830 G8 log).
#           Fix: poll loop now checks $script:CancelRequested every 700ms,
#           kills the curl process via Process.Kill() with a 2-second join,
#           deletes the partial output file (so a re-run doesn't accidentally
#           resume a half-baked download via curl's --continue-at), and
#           returns $false the same as a stall-abort would.
# v1.9.1 - HP catalog: fix CAB extraction silently failing. On v1.9.0 a real
#           EliteBook 830 G8 (SysID 880d) reached the catalog path, got a
#           valid HEAD 200 + a 100 KB download, but expand.exe extracted
#           nothing - log showed "Adding <cab> to Extraction Queue" then
#           "Expanding Files Complete" with no XML produced. Two changes:
#           (1) Pre-flight validation: check downloaded file >=1 KB and that
#               its first 4 bytes are the CAB magic 'MSCF'. This catches CDN
#               edges returning small text/HTML error bodies with a 200 OK.
#           (2) Three-tier extraction with fallback chain:
#               - Attempt 1: bare `expand <src> <dst-dir>` (canonical MS form)
#               - Attempt 2: `expand -F:* <src> <dst-dir>` (multi-file CABs)
#               - Attempt 3: Shell.Application COM CopyHere() (works on any
#                 Windows with Explorer regardless of expand quirks)
#           Each attempt clears the dest dir first so we can tell which one
#           produced the XML. If all three fail, dest contents are dumped to
#           the log for diagnosis before falling back to the full-pack scraper.
#           Also: now invoking expand.exe via Start-Process with split
#           -ArgumentList instead of PowerShell's quote-stuffed `& expand ...`,
#           which avoids the quoting confusion that contributed to v1.9.0's
#           silent failure.
# v1.9.0 - HP: per-machine reference catalog (HPIA backend) for selective driver
#           downloads. Mirrors the Dell CatalogPC and Lenovo consumer-catalog
#           patterns. Catalog lives at
#               https://hpia.hpcloud.hp.com/ref/<sysid>/<sysid>_64_<10|11>.0.<ver>.cab
#           where <sysid> is the 4-char hex Win32_BaseBoard.Product (e.g. 8B41
#           covers all G10 EliteBook/ZBook variants). Inside is a single XML
#           with two key sections: <Devices> (every supported PnP DeviceId on
#           that platform, each referencing one or more Softpaqs) and
#           <Solutions> (every Softpaq with direct URL, SHA256, file size, and
#           SilentInstall command).
#           Flow: SysID -> fetch CAB w/ OS version fallback chain (25h2 ->
#           24h2 -> 23h2 -> ...) -> expand.exe -> build O(1) DeviceId and SP
#           indexes -> enumerate missing devices (ConfigManagerErrorCode != 0)
#           -> walk each device's HardwareID/CompatibleID array (most-specific
#           first), also stripping &REV_XX, against the DeviceId index -> filter
#           matched SPs to "Driver - *" categories only -> total-size guardrail
#           (HpCatalogBudgetMB, default 800) -> download/SHA256-verify/extract/
#           SilentInstall each matched SP -> record into AnalyticsInstalledDrivers.
#           Returns $null on any of: no SystemID, no reference CAB at any tested
#           OS version, no missing devices, no catalog matches, total exceeds
#           budget, or zero softpaqs installed. $null falls through to the
#           existing HP_Driverpack_Matrix scraper (full-pack path) unchanged.
#           Bonus: SHA256 verification on every download (matrix-path full pack
#           still uses curl integrity only - could be unified in a future rev).
# v1.8.9 - Analytics: fix massive over-count of "Installed Drivers" in the
#           Dell/HP/Surface pnputil path. Lenovo path is already accurate
#           because it uses per-package vendor exit codes, not pnputil.
#           Root cause: pnputil /add-driver /install returns exit 0 for BOTH
#           "added to driver store" AND "actually bound to a matching device".
#           OEM packs include drivers for every SKU/CPU variant, so on any
#           given machine the majority of INFs are staged-but-unbound. v1.8.8
#           logged all of them. Example: HP EliteBook 830 G10 reported 331
#           "installed" with only 1 missing-driver resolved.
#           Fix: parse pnputil stdout for an explicit device-bind marker
#           before appending the INF name. Looks for one of three known
#           variations: "Installed driver package on matching", "Installed on
#           N device(s)" (N > 0), or "Successfully installed driver". INFs
#           added to the store but never bound are no longer recorded. The
#           total now tracks much closer to the missing-driver delta.
#           Caveat: pnputil's exact output strings have varied across Windows
#           versions; the parser uses a lenient OR of known patterns. If a
#           future Windows version changes the wording, we may under-count
#           rather than over-count (the failure mode is now safe).
# v1.8.8 - Analytics: capture which drivers were actually installed, not just the
#           count. New $script:AnalyticsInstalledDrivers list is populated from
#           all three install paths:
#           (1) Install-DriversFromPath (Dell/HP/Surface): appends the INF filename
#               when pnputil exits 0, 1641, or 3010 (success, reboot-initiated,
#               reboot-required). Code 259 ("newer driver already installed") and
#               other failure codes are not recorded.
#           (2) Lenovo consumer catalog main loop: appends the package Title when
#               the install passes the per-package RcOk check.
#           (3) Lenovo consumer catalog deferred BIOS phase: same Title append on
#               successful exit.
#           Sent to the webhook as a comma-separated "installed_drivers" string;
#           Google Sheet gets a matching "Installed Drivers" column at the end so
#           existing column positions don't shift.
# v1.8.7 - Lenovo consumer catalog: NVIDIA-on-non-NVIDIA-hardware short-circuit.
#           Title-based skip (mirrors the dock-skip pattern from v1.8.2): if a
#           package title matches *NVIDIA* but no Win32_VideoController on the
#           machine reports an NVIDIA / GeForce / VEN_10DE adapter, skip before
#           download. Saves 637MB+ on the integrated-only ThinkBook 14 G2 ITL
#           variant. Coarse but correct - the proper per-descriptor fix waits
#           on seeing one of Lenovo's "indeterminate" descriptors to know what
#           XML element to support in <DetectInstall>.
# v1.8.6 - End-of-run diagnostic: after the "Missing drivers AFTER install" count,
#           print a detailed list of every device still in problem state with
#           ErrorCode, Caption, DeviceID, HardwareID(s), and CompatibleID(s).
#           Works for all manufacturers (Dell/HP/Lenovo/Surface), not just the
#           Lenovo consumer-catalog path. Makes "still missing N drivers" runs
#           debuggable without a follow-up Get-PnpDevice query.
# v1.8.5 - Lenovo consumer catalog: two fixes for the unbound-ACPI-device case.
#           (1) Test-LenovoDPInstStaged bound-check bug: 0xFFFFFFFF parses as
#               Int32 -1 in PowerShell, so the upper-bound test "rejected" every
#               positive exit code and the DPInst fallback never fired. Replaced
#               with a sign-aware check ($ExitCode -le 0 short-circuits, no
#               upper bound needed since param is [int64]).
#           (2) Invoke-LenovoForceBindUnbound final pass: after PnP rescan, if
#               any Win32_PnPEntity still has ConfigManagerErrorCode != 0, run
#               pnputil /add-driver /install on every INF under the consumer
#               package root - including INFs from packages whose own installer
#               reported success - then re-enumerate and log anything still
#               unbound with its full DeviceID so we can see exactly which
#               hardware IDs Lenovo's published drivers don't cover.
# v1.8.4 - Lenovo consumer catalog: extended <DetectInstall> evaluator to skip
#           downloads (not just installs) for two more common element types.
#           (1) _DriverFileVersion: looks up the currently-bound driver for a
#               matching PnP hardware ID and compares its version against the
#               target. If hardware isn't enumerated on this machine OR an
#               equal/newer driver is already bound, the package is treated
#               as already-installed (no download, no install).
#           (2) _PnPID: checks for any PnP device whose hardware ID matches the
#               pattern. Used by Lenovo to gate packages on per-SKU components.
#           These two elements cover most of the v1.8.1 "indeterminate" cases.
#           Saves bandwidth on dual-SKU MTMs (e.g. 20VD covers integrated and
#           NVIDIA-discrete ThinkBook variants - the discrete-only drivers no
#           longer get downloaded on integrated-only hardware).
# v1.8.3 - Lenovo consumer catalog: three post-install reliability fixes.
#           (1) DPInst "staged but not bound" handling: when an install exits with
#               a DPInst-style packed code where bytes mean (0 bound, >0 staged,
#               0 failed), retry by running pnputil /add-driver /install on every
#               INF in the package dir. Fixes ThinkBook Fn-and-Function-Keys
#               driver leaving ACPI\VPC2004 + ACPI\IDEA2004 unbound.
#           (2) Hardware-not-present whitelist: nvsetup -436207360 (0xE6000100)
#               and similar "no compatible NVIDIA hardware" codes are now logged
#               as [N/A] (not present on this machine), not [FAIL]. Stops the
#               catalog's discrete-GPU drivers being counted as failures on
#               integrated-graphics-only variants of dual-SKU MTMs.
#           (3) Final pnputil /scan-devices after everything (main loop +
#               deferred BIOS phase) to bind any drivers that were staged but
#               not yet matched to their devices.
# v1.8.2 - Lenovo consumer catalog: skip dock-related packages entirely (no download,
#           no install). Title-based match on '*Dock*' covers ThinkPad USB-C/Thunderbolt
#           docks, hybrid docks, etc. - dock firmware doesn't belong in a one-shot
#           driver install run.
# v1.8.1 - Lenovo consumer catalog: two behavioural improvements.
#           (1) Evaluates each package's <DetectInstall> block - if the package
#               is already installed, skip download + install entirely. Supports
#               _Bios (BIOS level glob), _File (path + optional version), _FileVersion,
#               _Registry, _RegistryKey, and And/Or/Not combinators. Anything we can't
#               evaluate is fail-open (install anyway) for safety.
#           (2) BIOS packages (PackageType=3) are now deferred: downloaded + verified
#               in the main loop with everything else, but their install runs only
#               after all driver/firmware/app installs finish. Stops BIOS GUI from
#               interrupting the rest of the install queue.
# v1.8.0 - Lenovo: added consumer catalog support (per-machine-type XML used by Lenovo
#           System Update). Fixes coverage for ThinkBook/IdeaPad/consumer models missing
#           from catalogv2.xml (e.g. ThinkBook 14 G2 ITL / 20VD).
#           Flow: try https://download.lenovo.com/catalog/<MTM>_<Win10|Win11>.xml first,
#           parse each package descriptor, SHA256-verify, run vendor install commands
#           (with %PACKAGEPATH% / %WINDOWS% substitution). Falls back to legacy SCCM
#           driver pack via catalogv2.xml when consumer catalog has no entries.
#           Installs everything offered (drivers, firmware, BIOS); -SkipInstall still
#           honoured for download-only runs.
# v1.7.3 - Surface: removed OSDCatalog JSON path entirely (unreliable); Download Center is now
#           the sole method; MSI URL extracted from window.__DLCDetails__ JSON blob first
#           (reliable primary file), with href regex as fallback; driver version secondary sort;
#           audited and corrected all SurfaceDownloadIds against live Microsoft download pages
# v1.6.9 - Unknown manufacturer (e.g. OEMBY) shows Surface model picker dialog; headless auto-detects from -Model
# v1.6.8 - Surface: OSDCatalog JSON primary (SystemSKU match, MD5 verify, msiexec /a extract + pnputil)
# v1.6.7 - Pre-screen missing devices for parseable VEN/DEV before downloading CatalogPC
# v1.6.6 - Fixed INTELAUDIO bus HW ID parsing (CTLR_DEV_xxxx) for HDA audio devices
# v1.6.5 - Fixed param() position in Stop-DlSpinner/Stop-ExSpinner/Stop-OverallSpinner causing PS error
# v1.6.4 - Fixed CatalogPC parsing: UTF-16 encoding, PCIInfo VEN+DEV matching, osCode Win11 detection, systemID model filter
# v1.6.3 - Dell: individual driver lookup via CatalogPC when 1-3 devices missing; falls back to full pack
# v1.6.2 - Fixed: MessageBox DialogResult compared to enum not string (prompt was always falling through)
# v1.6.1 - Fixed: prompt blocks correctly before UI locks; auto-run restored when drivers missing
# v1.6.0 - GUI no longer auto-runs on launch; waits for user to click Install Drivers
# v1.5.9 - Fixed: missing driver prompt now shows before 7-Zip install and download
# v1.5.8 - Prompt to skip install if no missing drivers detected (GUI); continues in headless
# v1.5.7 - Added -DriverRoot param for parallel testing
# v1.5.6 - Added -MachineType param for Lenovo machine type override
# v1.5.5 - Added headless/parameter mode for testing and automation
# v1.5.4 - 7-Zip integration for Dell and HP extraction
#   Dell: 7-Zip pass-1 replaces /s /e= (verified identical output, 1.1x faster)
#   HP:   7-Zip pass-1 replaces /s /e /f (verified identical output, 4.7x faster)
#   Lenovo: unchanged - Inno Setup proprietary format, 7-Zip cannot extract
#   7-Zip is installed silently at start and removed before cleanup
# =============================================================

param(
    [string]$Manufacturer = "",
    [string]$Model        = "",
    [string]$MachineType  = "",   # Lenovo only: override 4-char machine type prefix (e.g. 20XX)
    [string]$DriverRoot   = "C:\DRIVERS",  # Override default driver root (useful for parallel testing)
    [switch]$Headless,
    [switch]$Silent,        # v1.10.0 - log-only mode (no GUI, no console). Implies -Headless.
    [switch]$TestMode,      # v1.10.0 - dry-run: detect + report but do not mutate system.
    [switch]$Diagnostic,    # v1.10.0 - verbose logging (per-step timings, env extras).
    [switch]$NoAnalytics,   # v1.10.0 - skip the Sheets webhook (local JSON still written).
    [int]$MaxParallelDownloads = 3,  # v1.11.0 - cap for parallel downloads in HP catalog / Dell individual paths.
    [switch]$SkipInstall,
    [switch]$SkipCleanup,
    [bool]$PromptWindowsUpdate = $true  # v1.12.1 - enabled by default: offer Windows Update if drivers still missing
)

# Auto-enable headless when any override param is passed.
# Silent always implies Headless (no GUI is even more "headless" than -Headless alone).
if ($Manufacturer -or $Model -or $MachineType -or $DriverRoot -ne "C:\DRIVERS" `
    -or $SkipInstall -or $SkipCleanup -or $Silent -or $TestMode -or $Diagnostic -or $NoAnalytics -or $PromptWindowsUpdate -eq $false) {
    $Headless = $true
}
if ($Silent) { $Headless = $true }

# =============================================================
# VERSION DEFINITION - Single source of truth for all version refs
# Update this number when making changes to the script
# =============================================================
$SCRIPT_VERSION = "1.13.3"

# =============================================================
$SpinnerFrames   = @('⠋','⠙','⠹','⠸','⠼','⠴','⠦','⠧','⠇','⠏')
$SpinnerIndex    = 0
$CancelRequested = $false

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

powercfg /change standby-timeout-ac 0
powercfg /change monitor-timeout-ac 0

# Set timezone to Brisbane (UTC+10, no DST) and sync clock
# Runs before log file creation so the filename timestamp is correct
tzutil /s "E. Australia Standard Time"
Start-Service w32tm -ErrorAction SilentlyContinue
w32tm /resync /force | Out-Null

# =========================
# LOG FILE SETUP
# v1.10.0 - three sibling files share a common timestamp/base:
#   <base>.log              - the canonical human-readable log (unchanged)
#   <base>.events.json      - NDJSON structured events (one JSON object per line)
#   <base>.analytics.json   - the final analytics payload (written at end-of-run)
#   <base>.report.html      - end-of-run HTML report (written on completion)
# =========================
$LogStamp        = Get-Date -Format 'yyyyMMdd_HHmmss'
$DownloadsDir    = Join-Path ([Environment]::GetFolderPath("UserProfile")) "Downloads"
$LogBase         = Join-Path $DownloadsDir ("DriverInstaller_" + $LogStamp)
$LogFile         = "$LogBase.log"
$EventsLogFile   = "$LogBase.events.json"
$AnalyticsFile   = "$LogBase.analytics.json"
$ReportFile      = "$LogBase.report.html"
New-Item -ItemType File -Path $LogFile       -Force | Out-Null
New-Item -ItemType File -Path $EventsLogFile -Force | Out-Null

# Mode flags lifted to script scope so helpers below can see them without
# parameter plumbing. Set in Start-Install as well (covers irm|iex re-entry).
$script:Silent     = [bool]$Silent
$script:TestMode   = [bool]$TestMode
$script:Diagnostic = [bool]$Diagnostic
$script:NoAnalytics = [bool]$NoAnalytics
$script:MaxParallelDownloads = $MaxParallelDownloads  # v1.11.0
$script:AutoScroll = $true   # v1.11.0 - GUI console auto-scrolls to latest line; toggle via checkbox or Set-AutoScroll

# =========================
# FONT CONSTANTS
# =========================
$FontMono       = New-Object System.Drawing.Font("Cascadia Mono", 9,    [System.Drawing.FontStyle]::Regular)
if ($FontMono.Name -ne "Cascadia Mono") {
    # Cascadia ships with Windows Terminal but isn't on every fresh install.
    # Fall back to Consolas, then Courier New, both ubiquitous.
    $FontMono   = New-Object System.Drawing.Font("Consolas",      9,    [System.Drawing.FontStyle]::Regular)
    if ($FontMono.Name -ne "Consolas") {
        $FontMono = New-Object System.Drawing.Font("Courier New", 9,    [System.Drawing.FontStyle]::Regular)
    }
}
$FontMonoSm     = New-Object System.Drawing.Font($FontMono.FontFamily, 8.25, [System.Drawing.FontStyle]::Regular)
$FontUI         = New-Object System.Drawing.Font("Segoe UI",    9,    [System.Drawing.FontStyle]::Regular)
$FontUIBold     = New-Object System.Drawing.Font("Segoe UI",    9,    [System.Drawing.FontStyle]::Bold)
$FontUIBoldSm   = New-Object System.Drawing.Font("Segoe UI",    8.25, [System.Drawing.FontStyle]::Bold)
$FontUISmall    = New-Object System.Drawing.Font("Segoe UI",    8,    [System.Drawing.FontStyle]::Regular)
$FontTitleBold  = New-Object System.Drawing.Font("Segoe UI Semibold", 17, [System.Drawing.FontStyle]::Bold)
$FontSubtitle   = New-Object System.Drawing.Font("Segoe UI",    9,    [System.Drawing.FontStyle]::Regular)
$FontSectionHdr = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Bold)
$FontStatusDot  = New-Object System.Drawing.Font("Segoe UI",    13,   [System.Drawing.FontStyle]::Regular)
$FontButton     = New-Object System.Drawing.Font("Segoe UI Semibold", 9.5, [System.Drawing.FontStyle]::Bold)

# =========================
# COLOR PALETTE (v1.11.0 UI refresh)
# Cool, modern, content-forward. Inspired by Tailwind gray/blue/emerald families
# so the script doesn't look like an abandoned 2014 utility on a 2026 monitor.
# Roles are stable so the rest of the script can keep referencing the same names.
# =========================
$ColorBg          = [System.Drawing.Color]::FromArgb(248, 249, 251)   # surface (gray-50)
$ColorCardBorder  = [System.Drawing.Color]::FromArgb(229, 231, 235)   # card border (gray-200)
$ColorTextHi      = [System.Drawing.Color]::FromArgb(17,  24,  39)    # primary text (gray-900)
$ColorTextMid     = [System.Drawing.Color]::FromArgb(75,  85,  99)    # secondary text (gray-600)
$ColorTextLo      = [System.Drawing.Color]::FromArgb(107, 114, 128)   # tertiary text (gray-500)
$ColorPrimary     = [System.Drawing.Color]::FromArgb(37,  99,  235)   # blue-600
$ColorPrimaryDk   = [System.Drawing.Color]::FromArgb(29,  78,  216)   # blue-700 (hover)
$ColorPrimaryPd   = [System.Drawing.Color]::FromArgb(30,  64,  175)   # blue-800 (pressed)
$ColorSuccess     = [System.Drawing.Color]::FromArgb(16,  185, 129)   # emerald-500
$ColorAccent      = [System.Drawing.Color]::FromArgb(99,  102, 241)   # indigo-500
$ColorDanger      = [System.Drawing.Color]::FromArgb(220, 38,  38)    # red-600
$ColorDangerDk    = [System.Drawing.Color]::FromArgb(185, 28,  28)    # red-700 (hover)
$ColorMuted       = [System.Drawing.Color]::FromArgb(156, 163, 175)   # gray-400 (idle dot)
$ColorMutedBg     = [System.Drawing.Color]::FromArgb(107, 114, 128)   # gray-500 (disabled btn bg - dark enough for WHITE text @ 4.5:1)
$ColorDisabledFg  = [System.Drawing.Color]::FromArgb(243, 244, 246)   # gray-100 (near-white text on disabled btn)
$ColorConsoleBg   = [System.Drawing.Color]::FromArgb(17,  24,  39)    # console bg (gray-900)
$ColorConsoleFg   = [System.Drawing.Color]::FromArgb(209, 250, 229)   # console fg (green-100 - bright mint, ~16:1 contrast on bg)

# =========================
# FORM
# =========================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Driver Installer  v$SCRIPT_VERSION"
$form.Size            = New-Object System.Drawing.Size(612, 628)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
$form.BackColor       = $ColorBg
$form.Font            = $FontUI   # default for any child control that doesn't override

# Top accent stripe (4px) - cheap branding cue, modern apps all do this.
$accentBar           = New-Object System.Windows.Forms.Panel
$accentBar.Size      = New-Object System.Drawing.Size(612, 4)
$accentBar.Location  = New-Object System.Drawing.Point(0, 0)
$accentBar.BackColor = $ColorPrimary
$form.Controls.Add($accentBar)

# =========================
# HEADER ROW (title + dynamic subtitle + version pill)
# =========================
$title           = New-Object System.Windows.Forms.Label
$title.AutoSize  = $true
$title.Font      = $FontTitleBold
$title.ForeColor = $ColorTextHi
$title.Location  = New-Object System.Drawing.Point(16, 22)
$title.Text      = "Driver Installer"
$title.UseCompatibleTextRendering = $false
$form.Controls.Add($title)

# Subtitle below the title - shows model after detection. Helpers can set
# $subtitle.Text to refresh; it's tiny and easy to miss but really pulls the
# whole layout together.
$subtitle           = New-Object System.Windows.Forms.Label
$subtitle.AutoSize  = $true
$subtitle.Font      = $FontSubtitle
$subtitle.ForeColor = $ColorTextMid
$subtitle.Location  = New-Object System.Drawing.Point(16, 56)
$subtitle.Text      = "Detecting hardware…"
$subtitle.UseCompatibleTextRendering = $false
$form.Controls.Add($subtitle)

# Version "pill" in the corner. Right-aligned via fixed coords (cheap), small.
$versionLabel           = New-Object System.Windows.Forms.Label
$versionLabel.AutoSize  = $true
$versionLabel.Font      = $FontUISmall
$versionLabel.ForeColor = $ColorTextMid           # v1.11.0 bumped from TextLo (gray-500) for readability
$versionLabel.Text      = "v$SCRIPT_VERSION"
$versionLabel.Location  = New-Object System.Drawing.Point(540, 30)
$versionLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($versionLabel)

# =========================
# CONSOLE / STATUS CARD
# Dark console box wrapped in a faint border so it reads as a "card".
# The status box is a CHILD of the border panel, not a sibling on the form —
# v1.11.0 fix: siblings on $form fight over Z-order (first-added wins in
# WinForms), which made the border panel render OVER the textbox and the
# whole console appeared as a grey rectangle.
# =========================
$statusCardBorder            = New-Object System.Windows.Forms.Panel
$statusCardBorder.Size       = New-Object System.Drawing.Size(556, 224)
$statusCardBorder.Location   = New-Object System.Drawing.Point(28, 88)
$statusCardBorder.BackColor  = $ColorCardBorder
$form.Controls.Add($statusCardBorder)

$statusBox             = New-Object System.Windows.Forms.RichTextBox
$statusBox.Multiline   = $true
$statusBox.ScrollBars  = "Vertical"
$statusBox.Size        = New-Object System.Drawing.Size(554, 222)
$statusBox.Location    = New-Object System.Drawing.Point(1, 1)    # relative to parent panel = 1px frame
$statusBox.ReadOnly    = $true
$statusBox.BackColor   = $ColorConsoleBg
$statusBox.ForeColor   = $ColorConsoleFg
$statusBox.Font        = $FontMono
$statusBox.BorderStyle = "None"
$statusCardBorder.Controls.Add($statusBox)        # CHILD of the panel - no Z-order conflict

# =========================
# PROGRESS SECTIONS
#
# Three sections (Download / Extract / Overall) share a layout: a colored
# status dot, a small uppercase header label, an inline braille spinner, a
# right-aligned mono caption with stats, and a progress bar below.
#
# Variable names are preserved from the pre-v1.11.0 layout so the helpers
# (SetDownload, SetExtract, Step-*Spinner, Stop-*Spinner, Test-Cancelled)
# work without modification. New variables introduced for the redesign:
#   $dlStatusDot / $exStatusDot / $overallStatusDot   - small "●" labels
#   $dlHeaderLabel / $exHeaderLabel / $overallHeaderLabel
#     (replaces $dlGroupBox.Text / $exGroupBox.Text / $overallGroupBox.Text;
#      five external call sites that wrote to $exGroupBox.Text have been
#      redirected to $exHeaderLabel.Text)
# =========================

# ---- DOWNLOAD section (y=324) ----
$dlStatusDot           = New-Object System.Windows.Forms.Label
$dlStatusDot.AutoSize  = $true
$dlStatusDot.Font      = $FontStatusDot
$dlStatusDot.ForeColor = $ColorMuted
$dlStatusDot.Text      = "●"
$dlStatusDot.Location  = New-Object System.Drawing.Point(28, 322)
$dlStatusDot.UseCompatibleTextRendering = $false
$form.Controls.Add($dlStatusDot)

$dlHeaderLabel           = New-Object System.Windows.Forms.Label
$dlHeaderLabel.AutoSize  = $true
$dlHeaderLabel.Font      = $FontSectionHdr
$dlHeaderLabel.ForeColor = $ColorTextHi
$dlHeaderLabel.Text      = "Download"
$dlHeaderLabel.Location  = New-Object System.Drawing.Point(46, 326)
$dlHeaderLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($dlHeaderLabel)

$dlSpinnerLabel           = New-Object System.Windows.Forms.Label
$dlSpinnerLabel.AutoSize  = $true
$dlSpinnerLabel.Font      = $FontUIBold
$dlSpinnerLabel.ForeColor = $ColorPrimary
$dlSpinnerLabel.Text      = ""
$dlSpinnerLabel.Location  = New-Object System.Drawing.Point(116, 326)
$dlSpinnerLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($dlSpinnerLabel)

$dlLabel           = New-Object System.Windows.Forms.Label
$dlLabel.AutoSize  = $false
$dlLabel.Size      = New-Object System.Drawing.Size(396, 18)
$dlLabel.Location  = New-Object System.Drawing.Point(188, 326)
$dlLabel.Font      = $FontMonoSm
$dlLabel.ForeColor = $ColorTextHi
$dlLabel.Text      = "Waiting…"
$dlLabel.TextAlign = "MiddleRight"
$dlLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($dlLabel)

$dlBar                       = New-Object System.Windows.Forms.ProgressBar
$dlBar.Size                  = New-Object System.Drawing.Size(556, 14)
$dlBar.Location              = New-Object System.Drawing.Point(28, 350)
$dlBar.Style                 = "Marquee"
$dlBar.MarqueeAnimationSpeed = 30
$dlBar.Minimum               = 0
$dlBar.Maximum               = 100
$form.Controls.Add($dlBar)

# ---- EXTRACT section (y=378) ----
$exStatusDot           = New-Object System.Windows.Forms.Label
$exStatusDot.AutoSize  = $true
$exStatusDot.Font      = $FontStatusDot
$exStatusDot.ForeColor = $ColorMuted
$exStatusDot.Text      = "●"
$exStatusDot.Location  = New-Object System.Drawing.Point(28, 376)
$exStatusDot.UseCompatibleTextRendering = $false
$form.Controls.Add($exStatusDot)

$exHeaderLabel           = New-Object System.Windows.Forms.Label
$exHeaderLabel.AutoSize  = $true
$exHeaderLabel.Font      = $FontSectionHdr
$exHeaderLabel.ForeColor = $ColorTextHi
$exHeaderLabel.Text      = "Extract"
$exHeaderLabel.Location  = New-Object System.Drawing.Point(46, 380)
$exHeaderLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($exHeaderLabel)

$exSpinnerLabel           = New-Object System.Windows.Forms.Label
$exSpinnerLabel.AutoSize  = $true
$exSpinnerLabel.Font      = $FontUIBold
$exSpinnerLabel.ForeColor = $ColorSuccess
$exSpinnerLabel.Text      = ""
$exSpinnerLabel.Location  = New-Object System.Drawing.Point(116, 380)
$exSpinnerLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($exSpinnerLabel)

$exLabel           = New-Object System.Windows.Forms.Label
$exLabel.AutoSize  = $false
$exLabel.Size      = New-Object System.Drawing.Size(396, 18)
$exLabel.Location  = New-Object System.Drawing.Point(188, 380)
$exLabel.Font      = $FontMonoSm
$exLabel.ForeColor = $ColorTextHi
$exLabel.Text      = "Waiting…"
$exLabel.TextAlign = "MiddleRight"
$exLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($exLabel)

$exBar                       = New-Object System.Windows.Forms.ProgressBar
$exBar.Size                  = New-Object System.Drawing.Size(556, 14)
$exBar.Location              = New-Object System.Drawing.Point(28, 404)
$exBar.Style                 = "Marquee"
$exBar.MarqueeAnimationSpeed = 30
$exBar.Minimum               = 0
$exBar.Maximum               = 100
$form.Controls.Add($exBar)

# ---- OVERALL section (y=432) ----
$overallStatusDot           = New-Object System.Windows.Forms.Label
$overallStatusDot.AutoSize  = $true
$overallStatusDot.Font      = $FontStatusDot
$overallStatusDot.ForeColor = $ColorMuted
$overallStatusDot.Text      = "●"
$overallStatusDot.Location  = New-Object System.Drawing.Point(28, 430)
$overallStatusDot.UseCompatibleTextRendering = $false
$form.Controls.Add($overallStatusDot)

$overallHeaderLabel           = New-Object System.Windows.Forms.Label
$overallHeaderLabel.AutoSize  = $true
$overallHeaderLabel.Font      = $FontSectionHdr
$overallHeaderLabel.ForeColor = $ColorTextHi
$overallHeaderLabel.Text      = "Overall"
$overallHeaderLabel.Location  = New-Object System.Drawing.Point(46, 434)
$overallHeaderLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($overallHeaderLabel)

$overallSpinnerLabel           = New-Object System.Windows.Forms.Label
$overallSpinnerLabel.AutoSize  = $true
$overallSpinnerLabel.Font      = $FontUIBold
$overallSpinnerLabel.ForeColor = $ColorAccent
$overallSpinnerLabel.Text      = ""
$overallSpinnerLabel.Location  = New-Object System.Drawing.Point(116, 434)
$overallSpinnerLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($overallSpinnerLabel)

$progress          = New-Object System.Windows.Forms.ProgressBar
$progress.Size     = New-Object System.Drawing.Size(556, 14)
$progress.Location = New-Object System.Drawing.Point(28, 458)
$progress.Style    = "Continuous"
$progress.Minimum  = 0
$progress.Maximum  = 100
$form.Controls.Add($progress)

# =========================
# FOOTER: log path + sound toggle + buttons
# =========================
$logLabel           = New-Object System.Windows.Forms.Label
$logLabel.AutoSize  = $false
$logLabel.Size      = New-Object System.Drawing.Size(556, 16)
$logLabel.Location  = New-Object System.Drawing.Point(28, 490)
$logLabel.ForeColor = $ColorTextMid    # v1.11.0 bumped from TextLo for readability
$logLabel.Font      = $FontUISmall
$logLabel.Text      = "Log: $LogFile"
$logLabel.UseCompatibleTextRendering = $false
$form.Controls.Add($logLabel)

# =========================
# SOUND TOGGLE CHECKBOX + ACTION BUTTONS
# =========================
$soundCheckbox                   = New-Object System.Windows.Forms.CheckBox
$soundCheckbox.Text              = "Sound FX"
$soundCheckbox.Checked           = $true
$soundCheckbox.Font              = $FontUI
$soundCheckbox.ForeColor         = $ColorTextHi     # v1.11.0 - was TextMid; darker for clearer label
$soundCheckbox.AutoSize          = $true
$soundCheckbox.Location          = New-Object System.Drawing.Point(28, 538)
$soundCheckbox.UseCompatibleTextRendering = $false
$form.Controls.Add($soundCheckbox)

# v1.11.0 - Auto-scroll toggle. Default on (matches pre-v1.11.0 behaviour).
# Unchecking lets the user scroll back through the log while the run continues -
# new lines still arrive but the textbox doesn't yank them down to the bottom.
$autoScrollCheckbox                   = New-Object System.Windows.Forms.CheckBox
$autoScrollCheckbox.Text              = "Auto-scroll"
$autoScrollCheckbox.Checked           = $true
$autoScrollCheckbox.Font              = $FontUI
$autoScrollCheckbox.ForeColor         = $ColorTextHi
$autoScrollCheckbox.AutoSize          = $true
$autoScrollCheckbox.Location          = New-Object System.Drawing.Point(126, 538)
$autoScrollCheckbox.UseCompatibleTextRendering = $false
$autoScrollCheckbox.Add_CheckedChanged({
    # Sync the script-scope flag with the checkbox state. Set-AutoScroll
    # also re-snaps the textbox to bottom on enable, which is the natural
    # expectation when the user re-checks the box mid-run.
    if ($script:AutoScroll -ne $autoScrollCheckbox.Checked) {
        Set-AutoScroll -Enabled $autoScrollCheckbox.Checked
    }
})
$form.Controls.Add($autoScrollCheckbox)

$button            = New-Object System.Windows.Forms.Button
$button.Text       = "Install Drivers"
$button.Size       = New-Object System.Drawing.Size(146, 36)
$button.Location   = New-Object System.Drawing.Point(290, 530)
$button.Font       = $FontButton
$button.BackColor  = $ColorPrimary
$button.ForeColor  = [System.Drawing.Color]::White
$button.FlatStyle  = "Flat"
$button.FlatAppearance.BorderSize         = 0
# Built-in hover/press states - cheaper than wiring MouseEnter/MouseLeave.
$button.FlatAppearance.MouseOverBackColor = $ColorPrimaryDk
$button.FlatAppearance.MouseDownBackColor = $ColorPrimaryPd
$button.Cursor     = "Hand"
$form.Controls.Add($button)

$cancelButton            = New-Object System.Windows.Forms.Button
$cancelButton.Text       = "Cancel"
$cancelButton.Size       = New-Object System.Drawing.Size(120, 36)
$cancelButton.Location   = New-Object System.Drawing.Point(446, 530)
$cancelButton.Font       = $FontButton
$cancelButton.BackColor  = $ColorMutedBg                        # idle: dark muted gray (v1.11.0)
$cancelButton.ForeColor  = $ColorDisabledFg                     # white text - readable when disabled
$cancelButton.FlatStyle  = "Flat"
$cancelButton.FlatAppearance.BorderSize         = 0
$cancelButton.FlatAppearance.MouseOverBackColor = $ColorDangerDk
$cancelButton.FlatAppearance.MouseDownBackColor = $ColorDangerDk
$cancelButton.Enabled    = $false
$form.Controls.Add($cancelButton)

# =========================
# SOUND HELPER
# =========================
function Play-Sound {
    param(
        [ValidateSet("Start","DownloadComplete","ExtractComplete","DriverAdded","Success","Failure","Cancel")]
        [string]$Event
    )
    if (-not $soundCheckbox.Checked) { return }
    $mediaDir = "$env:SystemRoot\Media"
    $wavCandidates = switch ($Event) {
        "Start"            { @("Windows Notify.wav", "Windows Notify System Generic.wav", "chimes.wav") }
        "DownloadComplete" { @("Windows Print complete.wav", "Windows Notify.wav", "chimes.wav") }
        "ExtractComplete"  { @("Windows Print complete.wav", "Windows Notify.wav", "chimes.wav") }
        "DriverAdded"      { @("Windows Navigation Start.wav", "Windows Notify Calendar.wav", "Windows Notify.wav") }
        "Success"          { @("Windows Logon.wav", "Windows Notify.wav", "tada.wav") }
        "Failure"          { @("Windows Critical Stop.wav", "Windows Foreground.wav", "chord.wav") }
        "Cancel"           { @("Windows Critical Stop.wav", "Windows Foreground.wav", "chord.wav") }
    }
    $wavFile = $null
    foreach ($candidate in $wavCandidates) {
        $path = Join-Path $mediaDir $candidate
        if (Test-Path $path) { $wavFile = $path; break }
    }
    if (-not $wavFile) { return }
    try { $player = New-Object System.Media.SoundPlayer $wavFile; $player.Play() } catch {}
}

# =========================
# WINDOWS UPDATE HELPER
# =========================
function Open-WindowsUpdate {
    <#
    .SYNOPSIS
    Opens Windows Update Settings to allow user to check for and install additional drivers
    #>
    Log "Opening Windows Update..."
    try {
        # Windows 10/11: launch Settings app to Windows Update
        # ms-settings:windowsupdate opens directly to Windows Update
        Start-Process "ms-settings:windowsupdate" -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
    } catch {
        Log "WARNING: Could not open Windows Update via ms-settings - $($_.Exception.Message)"
        try {
            # Fallback: try opening via control.exe
            Start-Process "control" -ArgumentList "sysdm.cpl" -ErrorAction SilentlyContinue
        } catch {
            Log "WARNING: Could not open system settings - $($_.Exception.Message)"
        }
    }
}


function Set-ButtonRunning {
    if ($script:Headless) { return }
    $button.Enabled         = $false
    $button.BackColor       = $ColorMutedBg
    $button.ForeColor       = $ColorDisabledFg     # v1.11.0 - white on gray-500 (4.5:1, readable)
    $cancelButton.Enabled   = $true
    $cancelButton.BackColor = $ColorDanger
    $cancelButton.ForeColor = [System.Drawing.Color]::White
    # Reset section dots back to muted-then-they-light-up-as-spinners-fire
    $dlStatusDot.ForeColor      = $ColorMuted
    $exStatusDot.ForeColor      = $ColorMuted
    $overallStatusDot.ForeColor = $ColorMuted
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-ButtonIdle {
    if ($script:Headless) { return }
    $button.Enabled         = $true
    $button.BackColor       = $ColorPrimary
    $button.ForeColor       = [System.Drawing.Color]::White
    $cancelButton.Enabled   = $false
    $cancelButton.BackColor = $ColorMutedBg
    $cancelButton.ForeColor = $ColorDisabledFg     # v1.11.0 - white on gray-500
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-AutoScroll {
    # v1.11.0 - enable or disable auto-scroll of the GUI console box.
    # Usage:
    #   Set-AutoScroll $true     # snap to bottom on every Log line (default)
    #   Set-AutoScroll $false    # leave the scroll position alone - user can
    #                            # browse history while the run continues
    #
    # The checkbox in the footer toggles this too; calling this function and
    # toggling the checkbox stay in sync because the checkbox's CheckedChanged
    # handler calls back into this function.
    param([bool]$Enabled)
    $script:AutoScroll = $Enabled
    if (-not $script:Headless -and $autoScrollCheckbox -and $autoScrollCheckbox.Checked -ne $Enabled) {
        # Don't trigger a CheckedChanged event loop: only assign if the checkbox
        # state doesn't already match. (The CheckedChanged handler also guards
        # against this via state comparison, but belt-and-braces is cheap.)
        $autoScrollCheckbox.Checked = $Enabled
    }
    if (-not $script:Headless -and $Enabled -and $statusBox) {
        # When re-enabling, jump straight to the bottom so the next log line
        # doesn't appear orphaned in the middle of the scrollback.
        $statusBox.SelectionStart = $statusBox.TextLength
        $statusBox.ScrollToCaret()
    }
    Log "Auto-scroll: $(if ($Enabled) { 'on' } else { 'off (browse history freely)' })" `
        -Level "info" -Event "auto_scroll_toggle" -Context @{ enabled = $Enabled }
}

# =========================
# HELPERS
# =========================
function Log {
    # v1.10.0 - now writes to three sinks:
    #   1. console / GUI (suppressed by -Silent)
    #   2. canonical .log text file (always)
    #   3. .events.json NDJSON stream (always)
    # The 'level' is heuristically inferred from the message text - existing
    # call sites don't need to change. Pass -Level / -Event / -Context to a
    # future structured caller and they'll override the inferred values.
    param(
        [Parameter(Mandatory=$true, Position=0)] [string]$msg,
        [string]$Level   = $null,
        [string]$Event   = $null,
        [hashtable]$Context = $null
    )
    $now  = Get-Date
    $ts   = $now.ToString('HH:mm:ss')
    $line = "[$ts] $msg"

    # Human sinks
    if (-not $script:Silent) {
        if ($script:Headless) {
            Write-Host $line
        } else {
            $statusBox.AppendText("$line`r`n")
            # v1.11.0 - only snap to bottom when auto-scroll is on. When the user
            # has scrolled back to read history, $script:AutoScroll = $false stops
            # the textbox yanking them down on every new log line.
            if ($script:AutoScroll) { $statusBox.ScrollToCaret() }
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    Add-Content -Path $LogFile -Value $line -Encoding UTF8

    # Structured NDJSON sink. Heuristic level inference keeps existing
    # one-arg Log calls working; explicit -Level wins when supplied.
    if (-not $Level) {
        if ($msg -match '(?i)\bERROR\b|FATAL|FAILED|cannot') { $Level = 'error' }
        elseif ($msg -match '(?i)\bWARNING\b|WARN\b')        { $Level = 'warn'  }
        elseif ($msg -match '(?i)\bcancel(led)?\b')          { $Level = 'cancel' }
        else                                                  { $Level = 'info' }
    }
    try {
        $evt = [ordered]@{
            ts             = $now.ToString('o')
            level          = $Level
            script_version = $SCRIPT_VERSION
            msg            = $msg
        }
        if ($Event)   { $evt['event']   = $Event }
        if ($Context) { foreach ($k in $Context.Keys) { $evt[$k] = $Context[$k] } }
        $json = $evt | ConvertTo-Json -Compress -Depth 4
        Add-Content -Path $EventsLogFile -Value $json -Encoding UTF8
    } catch {
        # Never let event-log failure interrupt the script. Worst case: the
        # NDJSON sidecar is incomplete, but the canonical .log still has the line.
    }
}

function Log-Diag {
    # Diagnostic-only log: silently dropped unless -Diagnostic is set.
    # Use this for high-volume / low-signal trace info (per-iteration timings,
    # raw HTTP headers, etc.) that would otherwise bloat the standard log.
    param([string]$msg, [hashtable]$Context = $null)
    if (-not $script:Diagnostic) { return }
    Log -msg "[diag] $msg" -Level "debug" -Event "diagnostic" -Context $Context
}

function SetProgress($val) {
    if ($script:Headless) { return }
    $progress.Value = [math]::Min([math]::Max([int]$val, 0), 100)
    [System.Windows.Forms.Application]::DoEvents()
}

function SetDownload {
    param([int]$Pct, [string]$Label)
    if ($script:Headless) { return }
    if ($Pct -ge 100) {
        $dlBar.Style = "Continuous"
        $dlBar.Value = 100
    } elseif ($dlBar.Style -ne "Marquee") {
        $dlBar.Style                 = "Marquee"
        $dlBar.MarqueeAnimationSpeed = 25
    }
    $dlLabel.Text = $Label
    [System.Windows.Forms.Application]::DoEvents()
}

function SetExtract {
    param([int]$Pct, [string]$Label)
    if ($script:Headless) { return }
    if ($Pct -ge 100) {
        $exBar.Style = "Continuous"
        $exBar.Value = 100
    } elseif ($Pct -lt 0) {
        if ($exBar.Style -ne "Marquee") {
            $exBar.Style                 = "Marquee"
            $exBar.MarqueeAnimationSpeed = 30
        }
    } else {
        if ($exBar.Style -ne "Continuous") { $exBar.Style = "Continuous" }
        $exBar.Value = [math]::Min($Pct, 99)
    }
    $exLabel.Text = $Label
    [System.Windows.Forms.Application]::DoEvents()
}

function Step-DlSpinner {
    if ($script:Headless) { return }
    $script:SpinnerIndex = ($script:SpinnerIndex + 1) % $SpinnerFrames.Count
    $dlSpinnerLabel.Text = " " + $SpinnerFrames[$script:SpinnerIndex]
    # v1.11.0 - dot tracks state: working = primary blue
    $dlStatusDot.ForeColor = $ColorPrimary
    [System.Windows.Forms.Application]::DoEvents()
}
function Stop-DlSpinner {
    param([bool]$Success = $true)
    if ($script:Headless) { return }
    $dlSpinnerLabel.Text      = if ($Success) { " OK" } else { " XX" }
    $dlSpinnerLabel.ForeColor = if ($Success) { $ColorPrimary } else { $ColorDanger }
    # v1.11.0 - dot tracks state: done = success green, failed = danger red
    $dlStatusDot.ForeColor    = if ($Success) { $ColorSuccess } else { $ColorDanger }
    [System.Windows.Forms.Application]::DoEvents()
}

function Step-ExSpinner {
    if ($script:Headless) { return }
    $script:SpinnerIndex  = ($script:SpinnerIndex + 1) % $SpinnerFrames.Count
    $exSpinnerLabel.Text  = " " + $SpinnerFrames[$script:SpinnerIndex]
    $exStatusDot.ForeColor = $ColorPrimary
    [System.Windows.Forms.Application]::DoEvents()
}
function Stop-ExSpinner {
    param([bool]$Success = $true)
    if ($script:Headless) { return }
    $exSpinnerLabel.Text      = if ($Success) { " OK" } else { " XX" }
    $exSpinnerLabel.ForeColor = if ($Success) { $ColorSuccess } else { $ColorDanger }
    $exStatusDot.ForeColor    = if ($Success) { $ColorSuccess } else { $ColorDanger }
    [System.Windows.Forms.Application]::DoEvents()
}

function Step-OverallSpinner {
    if ($script:Headless) { return }
    $script:SpinnerIndex      = ($script:SpinnerIndex + 1) % $SpinnerFrames.Count
    $overallSpinnerLabel.Text = " " + $SpinnerFrames[$script:SpinnerIndex]
    $overallStatusDot.ForeColor = $ColorAccent
    [System.Windows.Forms.Application]::DoEvents()
}
function Stop-OverallSpinner {
    param([bool]$Success = $true)
    if ($script:Headless) { return }
    $overallSpinnerLabel.Text      = if ($Success) { " OK" } else { " XX" }
    $overallSpinnerLabel.ForeColor = if ($Success) { $ColorSuccess } else { $ColorDanger }
    $overallStatusDot.ForeColor    = if ($Success) { $ColorSuccess } else { $ColorDanger }
    [System.Windows.Forms.Application]::DoEvents()
}

function Step-AllSpinners {
    if ($script:Headless) { return }
    $script:SpinnerIndex      = ($script:SpinnerIndex + 1) % $SpinnerFrames.Count
    $f                        = $SpinnerFrames[$script:SpinnerIndex]
    $dlSpinnerLabel.Text      = " " + $f
    $exSpinnerLabel.Text      = " " + $f
    $overallSpinnerLabel.Text = " " + $f
    # While anything is spinning, every dot reflects "working" colour
    $dlStatusDot.ForeColor      = $ColorPrimary
    $exStatusDot.ForeColor      = $ColorPrimary
    $overallStatusDot.ForeColor = $ColorAccent
    [System.Windows.Forms.Application]::DoEvents()
}

function Test-Cancelled {
    if ($script:CancelRequested) {
        Log "Operation cancelled."
        SetDownload -Pct 0 -Label "Cancelled"
        SetExtract  -Pct 0 -Label "Cancelled"
        Stop-DlSpinner      -Success $false
        Stop-ExSpinner      -Success $false
        Stop-OverallSpinner -Success $false
        Play-Sound -Event "Cancel"
        Set-ButtonIdle
        return $true
    }
    return $false
}

function Assert-Curl {
    if (-not (Get-Command curl.exe -ErrorAction SilentlyContinue)) {
        Log "ERROR: curl.exe not found. Windows 10 1803+ required."
        return $false
    }
    return $true
}

function Get-MissingDriverCount {
    try {
        $count = @(Get-CimInstance Win32_PnPEntity |
            Where-Object { $_.ConfigManagerErrorCode -ne 0 }).Count
        return $count
    } catch {
        Log "  WARNING: Could not query PnP device status: $($_.Exception.Message)"
        return -1
    }
}

function Get-MissingDriverNames {
    # v1.11.0 - returns an array of device names for every PnP device whose
    # ConfigManagerErrorCode is non-zero (i.e. "missing/broken driver"). Used
    # to populate $script:AnalyticsMissingBeforeList (HTML report only as of
    # v1.13.2 - the spreadsheet's missing_before_list field was replaced with
    # driver_urls) and $script:AnalyticsMissingAfterList (HTML report + the
    # spreadsheet's missing_after_list field). Returns an empty array on any
    # failure so callers can blindly forward to ConvertTo-Json without
    # null-handling.
    try {
        $missing = @(Get-CimInstance Win32_PnPEntity -EA Stop |
                     Where-Object { $_.ConfigManagerErrorCode -ne 0 })
        $names = New-Object 'System.Collections.Generic.List[string]'
        foreach ($m in $missing) {
            $n = if ($m.Name)        { [string]$m.Name }
                 elseif ($m.Caption) { [string]$m.Caption }
                 else                { '(unnamed device)' }
            # Deduplicate trivially - many machines have multiple "Unknown device"
            # rows that all collapse to the same string; one row in the sheet is
            # plenty.
            if ($names -notcontains $n) { $names.Add($n) | Out-Null }
        }
        return ,@($names)
    } catch {
        Log "  WARNING: Could not enumerate missing device names: $($_.Exception.Message)"
        return ,@()
    }
}

function Write-MissingDriverDetails {
    # Detailed end-of-run dump of every device still in problem state.
    # Pulls HardwareID + CompatibleID arrays from Get-PnpDevice (Win32_PnPEntity
    # alone doesn't expose those) so the log shows exactly which IDs an INF
    # would need to claim to bind to each device.
    try {
        $missing = @(Get-CimInstance Win32_PnPEntity -EA Stop |
                     Where-Object { $_.ConfigManagerErrorCode -ne 0 })
    } catch {
        Log "-- Missing drivers AFTER install (enumeration error: $($_.Exception.Message)) --"
        return
    }
    if ($missing.Count -eq 0) {
        Log "-- Missing drivers AFTER install (0 - all present devices are bound) --"
        return
    }
    Log "-- Missing drivers AFTER install ($($missing.Count)) --"
    foreach ($m in $missing) {
        $name = if ($m.Name)         { $m.Name }
                elseif ($m.Caption)  { $m.Caption }
                else                 { '(unnamed device)' }
        Log "  [ERR $($m.ConfigManagerErrorCode)] $name"
        Log "    DeviceID:    $($m.DeviceID)"
        try {
            $pnp = Get-PnpDevice -InstanceId $m.DeviceID -EA Stop
            if ($pnp.HardwareID) {
                $ids = @($pnp.HardwareID)
                Log "    HardwareID:  $($ids[0])"
                for ($i = 1; $i -lt $ids.Count; $i++) { Log "                 $($ids[$i])" }
            }
            if ($pnp.CompatibleID) {
                $ids = @($pnp.CompatibleID)
                Log "    CompatID:    $($ids[0])"
                for ($i = 1; $i -lt $ids.Count; $i++) { Log "                 $($ids[$i])" }
            }
        } catch {
            # Get-PnpDevice failed for this InstanceId - DeviceID line above is
            # still enough to identify the device; skip the extra detail.
        }
    }
}

# =========================
# v1.10.0 - VERBOSE DIAGNOSTICS
# Only emits when -Diagnostic was passed. Complements (does not replace)
# Write-DeviceInfo, which always runs.
# =========================
function Write-VerboseDiagnostics {
    if (-not $script:Diagnostic) { return }
    Log "============================================" -Level "info" -Event "diag_header"
    Log "  VERBOSE DIAGNOSTICS (-Diagnostic)"
    Log "============================================"
    try {
        $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $proxyUri = $proxy.GetProxy("https://www.google.com").AbsoluteUri
        Log "  System proxy        : $proxyUri"
    } catch { Log "  System proxy        : (error: $($_.Exception.Message))" }
    try {
        $tls = [Net.ServicePointManager]::SecurityProtocol
        Log "  TLS protocols       : $tls"
    } catch {}
    try {
        $curlVer = (& curl.exe --version 2>$null) -join " | "
        Log "  curl.exe version    : $curlVer"
    } catch {}
    try {
        $route = (route print 0.0.0.0 2>$null | Select-String "0.0.0.0\s+0.0.0.0" | Select-Object -First 1).ToString().Trim()
        Log "  Default route       : $route"
    } catch {}
    try {
        $nic = Get-NetAdapter -Physical -EA Stop | Where-Object Status -eq 'Up' | Select-Object -First 1
        if ($nic) {
            Log "  Active NIC          : $($nic.Name) ($($nic.InterfaceDescription)) - LinkSpeed $($nic.LinkSpeed)"
        }
    } catch {}
    Log "============================================"
}

# =========================
# v1.10.0 - HTML INSTALL REPORT
# Self-contained single-file HTML with all the analytics state. Saved to
# Downloads alongside the .log so it can be attached to support tickets or
# shown to customers at handover. No external CSS/JS - styled inline.
# =========================
function Write-HtmlReport {
    param(
        [Parameter(Mandatory=$true)][ValidateSet("success","failure","cancelled","testmode")]
        [string]$Result
    )
    try {
        $durationSec = if ($script:AnalyticsStartTime) {
            [int]((Get-Date) - $script:AnalyticsStartTime).TotalSeconds
        } else { 0 }
        $durationDisplay = if ($durationSec -ge 60) {
            "{0}m {1}s" -f ([math]::Floor($durationSec/60)), ($durationSec % 60)
        } else { "${durationSec}s" }

        $missingDelta = if ($script:AnalyticsMissingBefore -ge 0 -and $script:AnalyticsMissingAfter -ge 0) {
            $script:AnalyticsMissingBefore - $script:AnalyticsMissingAfter
        } else { 0 }

        # Status colour palette - reused for the badge and the row accents.
        $statusColour = switch ($Result) {
            "success"   { "#1b873f" }
            "failure"   { "#c92a2a" }
            "cancelled" { "#a37b00" }
            "testmode"  { "#1864ab" }
        }
        $statusLabel = $Result.ToUpper()

        # HTML-escape user-controlled strings (model names sometimes contain &)
        function _He($s) {
            if ($null -eq $s) { return "" }
            $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
        }

        $installedRows = ""
        if ($script:AnalyticsInstalledDrivers -and $script:AnalyticsInstalledDrivers.Count -gt 0) {
            foreach ($d in $script:AnalyticsInstalledDrivers) {
                if ([string]::IsNullOrWhiteSpace($d)) { continue }
                $installedRows += "<li>$(_He $d)</li>`n"
            }
        } else {
            $installedRows = "<li class='muted'>(none recorded)</li>"
        }

        # v1.11.0 - missing-driver lists in the HTML report. Same <ul><li> shape
        # as installed_drivers; muted placeholder if a list happens to be empty.
        $missingBeforeRows = ""
        if ($script:AnalyticsMissingBeforeList -and $script:AnalyticsMissingBeforeList.Count -gt 0) {
            foreach ($d in $script:AnalyticsMissingBeforeList) {
                if ([string]::IsNullOrWhiteSpace($d)) { continue }
                $missingBeforeRows += "<li>$(_He $d)</li>`n"
            }
        } else {
            $missingBeforeRows = "<li class='muted'>(no missing devices detected before run)</li>"
        }
        $missingAfterRows = ""
        if ($script:AnalyticsMissingAfterList -and $script:AnalyticsMissingAfterList.Count -gt 0) {
            foreach ($d in $script:AnalyticsMissingAfterList) {
                if ([string]::IsNullOrWhiteSpace($d)) { continue }
                $missingAfterRows += "<li>$(_He $d)</li>`n"
            }
        } else {
            $missingAfterRows = "<li class='muted'>(all devices resolved)</li>"
        }

        # v1.13.1 - build the re-downloadable URL lists. Driver payloads get a
        # clickable link AND the raw URL printed beneath it, so the URL is
        # still recoverable from a printed or PDF'd copy of the report where
        # the hyperlink target would otherwise be lost. Catalog/metadata
        # sources are listed separately and muted - kept for traceability,
        # not for standalone re-download.
        $driverUrlRows  = ""
        $catalogUrlRows = ""
        if ($script:AnalyticsDownloadUrls -and $script:AnalyticsDownloadUrls.Count -gt 0) {
            foreach ($e in $script:AnalyticsDownloadUrls) {
                if (-not $e -or [string]::IsNullOrWhiteSpace($e.Url)) { continue }
                $u  = _He $e.Url
                $fn = _He $e.FileName
                $sz = if ($e.MB -gt 0) { " &middot; $($e.MB) MB" } else { "" }
                if ($e.Kind -eq 'catalog') {
                    $catalogUrlRows += "<li><span class='fn'>$fn</span><br><a href=`"$u`">$u</a></li>`n"
                } else {
                    $driverUrlRows += "<li><span class='fn'>$fn$sz</span><br><a href=`"$u`">$u</a><div class='rawurl'>$u</div></li>`n"
                }
            }
        }
        if ([string]::IsNullOrWhiteSpace($driverUrlRows)) {
            $driverUrlRows = "<li class='muted'>(no driver files were downloaded this run)</li>"
        }
        if ([string]::IsNullOrWhiteSpace($catalogUrlRows)) {
            $catalogUrlRows = "<li class='muted'>(none)</li>"
        }

        $html = @"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Driver Installer Report - $(_He $script:AnalyticsModel)</title>
<style>
  body { font: 14px/1.5 -apple-system, "Segoe UI", system-ui, sans-serif; color: #222; background: #f6f7f9; margin: 0; padding: 24px; }
  .wrap { max-width: 900px; margin: 0 auto; background: #fff; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,.08); overflow: hidden; }
  header { padding: 20px 28px; border-bottom: 1px solid #e5e7eb; display: flex; justify-content: space-between; align-items: center; }
  header h1 { margin: 0; font-size: 20px; font-weight: 600; }
  header .v { color: #6b7280; font-size: 12px; }
  .badge { display: inline-block; padding: 6px 14px; border-radius: 999px; color: #fff; font-weight: 600; font-size: 12px; letter-spacing: .05em; background: $statusColour; }
  section { padding: 18px 28px; border-bottom: 1px solid #f0f1f3; }
  section:last-child { border-bottom: 0; }
  h2 { font-size: 13px; text-transform: uppercase; letter-spacing: .06em; color: #6b7280; margin: 0 0 12px 0; font-weight: 600; }
  table { width: 100%; border-collapse: collapse; }
  td { padding: 6px 0; vertical-align: top; }
  td.k { width: 200px; color: #6b7280; }
  td.v { font-family: ui-monospace, "Cascadia Mono", "Consolas", monospace; font-size: 13px; word-break: break-word; }
  ul { margin: 0; padding-left: 20px; font-family: ui-monospace, "Cascadia Mono", "Consolas", monospace; font-size: 13px; }
  ul li { margin: 2px 0; }
  .muted { color: #9ca3af; font-style: italic; list-style: none; margin-left: -20px; }
  .delta-good { color: #1b873f; font-weight: 600; }
  .delta-bad  { color: #c92a2a; font-weight: 600; }
  /* v1.13.1 - download URL list */
  ul.urls li { margin: 8px 0; word-break: break-all; }
  ul.urls .fn { font-weight: 600; color: #222; }
  ul.urls a { color: #1864ab; }
  ul.urls .rawurl { display: none; color: #6b7280; font-size: 12px; margin-top: 2px; }
  @media print { ul.urls a { color: #222; text-decoration: none; }
                  ul.urls .rawurl { display: block; } }
  footer { padding: 12px 28px; color: #9ca3af; font-size: 11px; background: #fafbfc; }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div>
      <h1>Driver Installer Report</h1>
      <div class="v">Script v$SCRIPT_VERSION &middot; generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')</div>
    </div>
    <span class="badge">$statusLabel</span>
  </header>

  <section>
    <h2>Device</h2>
    <table>
      <tr><td class="k">Manufacturer</td><td class="v">$(_He $script:AnalyticsManufacturer)</td></tr>
      <tr><td class="k">Model</td><td class="v">$(_He $script:AnalyticsModel)</td></tr>
      <tr><td class="k">Serial / Service Tag</td><td class="v">$(_He $script:AnalyticsSerial)</td></tr>
      <tr><td class="k">OS Version</td><td class="v">$(_He $script:AnalyticsOsVersion) (build $($script:AnalyticsOsBuild))</td></tr>
    </table>
  </section>

  <section>
    <h2>Result</h2>
    <table>
      <tr>
        <td class="k">Missing drivers before</td>
        <td class="v">$($script:AnalyticsMissingBefore)</td>
      </tr>
      <tr>
        <td class="k">Missing drivers after</td>
        <td class="v">$($script:AnalyticsMissingAfter)</td>
      </tr>
      <tr>
        <td class="k">Resolved</td>
        <td class="v"><span class="$(if ($missingDelta -gt 0) {'delta-good'} elseif ($missingDelta -lt 0) {'delta-bad'} else {''})">$missingDelta</span></td>
      </tr>
      <tr><td class="k">INFs installed (bound)</td><td class="v">$($script:AnalyticsInfCount)</td></tr>
      <tr><td class="k">Total download size</td><td class="v">$($script:AnalyticsDownloadMB) MB</td></tr>
      <tr><td class="k">Duration</td><td class="v">$durationDisplay</td></tr>
    </table>
  </section>

  <section>
    <h2>Missing drivers before run</h2>
    <ul>$missingBeforeRows</ul>
  </section>

  <section>
    <h2>Missing drivers still unresolved</h2>
    <ul>$missingAfterRows</ul>
  </section>

  <section>
    <h2>Installed drivers / packages</h2>
    <ul>$installedRows</ul>
  </section>

  <section>
    <h2>Driver download URLs</h2>
    <p style="margin:0 0 10px 0;color:#6b7280;font-size:12px;">
      Direct links to the exact driver file(s) pulled this run. Use these to
      re-download the same pack later without re-running the script or
      re-resolving the vendor catalog. Vendor CDN links can expire or be
      superseded over time.
    </p>
    <ul class="urls">$driverUrlRows</ul>
  </section>

  <section>
    <h2>Catalog / metadata sources</h2>
    <p style="margin:0 0 10px 0;color:#9ca3af;font-size:12px;">
      Intermediate catalog/descriptor files used to locate the driver above.
      Listed for traceability only - not standalone driver downloads.
    </p>
    <ul class="urls">$catalogUrlRows</ul>
  </section>

  <section>
    <h2>Artifacts</h2>
    <table>
      <tr><td class="k">Text log</td><td class="v">$(_He $LogFile)</td></tr>
      <tr><td class="k">Structured events</td><td class="v">$(_He $EventsLogFile)</td></tr>
      <tr><td class="k">Analytics JSON</td><td class="v">$(_He $AnalyticsFile)</td></tr>
    </table>
  </section>

  <footer>Install-Drivers-auto.ps1 v$SCRIPT_VERSION &middot; skermiebroTech &middot; github.com/skermiebroTech/my-wiki</footer>
</div>
</body>
</html>
"@
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($ReportFile, $html, $utf8NoBom)
        Log "HTML report saved to: $ReportFile" -Level "info" -Event "report_written"
    } catch {
        Log "  WARNING: HTML report generation failed - $($_.Exception.Message)" -Level "warn"
    }
}

# =========================
# 7-ZIP HELPERS
# Installed at start of Start-Install for Dell/HP extraction.
# Removed before final cleanup so nothing persists to the customer.
# Lenovo uses Inno Setup - 7-Zip cannot extract its proprietary format.
# =========================
$script:7zExe         = "C:\Program Files\7-Zip\7z.exe"
$script:7zInstaller   = "$env:TEMP\7z-installer.exe"
$script:7zInstalled   = $false

function Install-7Zip {
    if (Test-Path $script:7zExe) {
        Log "7-Zip already present - skipping install."
        $script:7zInstalled = $true
        return $true
    }
    Log "Installing 7-Zip (temporary - will be removed after extraction)..."
    SetDownload -Pct 0 -Label "Downloading 7-Zip..."
    try {
        $psi                        = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = "curl.exe"
        $psi.Arguments              = "--silent --location --max-time 60 --connect-timeout 15 " +
                                      "--output `"$($script:7zInstaller)`" " +
                                      "`"https://www.7-zip.org/a/7z2409-x64.exe`""
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true
        $proc                       = New-Object System.Diagnostics.Process
        $proc.StartInfo             = $psi
        $proc.Start() | Out-Null
        while (-not $proc.HasExited) { Start-Sleep -Milliseconds 400; Step-AllSpinners }
        if ($proc.ExitCode -ne 0 -or -not (Test-Path $script:7zInstaller)) {
            Log "  7-Zip download failed (curl exit $($proc.ExitCode))."
            return $false
        }
        Start-Process $script:7zInstaller -ArgumentList "/S" -Wait
        if (-not (Test-Path $script:7zExe)) {
            Log "  7-Zip installer ran but 7z.exe not found."
            return $false
        }
        $script:7zInstalled = $true
        Log "  7-Zip installed OK."
        return $true
    } catch {
        Log "  7-Zip install error: $($_.Exception.Message)"
        return $false
    }
}

function Remove-7Zip {
    Log "Removing 7-Zip..."
    try {
        $uninstaller = "C:\Program Files\7-Zip\Uninstall.exe"
        if (Test-Path $uninstaller) {
            Start-Process $uninstaller -ArgumentList "/S" -Wait
            Start-Sleep -Seconds 2
        }
        # Belt-and-braces: remove folder if uninstaller left anything
        $folder = "C:\Program Files\7-Zip"
        if (Test-Path $folder) { Remove-Item $folder -Recurse -Force -EA SilentlyContinue }
    } catch {
        Log "  WARNING: 7-Zip uninstall error: $($_.Exception.Message)"
    }
    # Always clean up the installer temp file
    Remove-Item $script:7zInstaller -Force -EA SilentlyContinue
    $script:7zInstalled = $false

    # Verify
    if (Test-Path $script:7zExe) {
        Log "  WARNING: 7z.exe still present after uninstall - check manually before sysprep."
    } else {
        Log "  7-Zip removed OK."
    }
}

# =========================
# GOOGLE SHEETS ANALYTICS
# =========================
$SHEETS_WEBHOOK = "https://script.google.com/macros/s/AKfycbygEF0i6j_6rSstmfQ2sQPLn0KjkqxZwUwIRjyCsd911IP9kALucv2cImMFumGoUUs/exec"

$script:AnalyticsManufacturer    = ""
$script:AnalyticsModel           = ""
$script:AnalyticsSerial          = ""
$script:AnalyticsOsVersion       = ""
$script:AnalyticsOsBuild         = 0
$script:AnalyticsInfCount        = 0
$script:AnalyticsDownloadMB      = 0.0
$script:AnalyticsStartTime       = $null
$script:AnalyticsMissingBefore   = -1
$script:AnalyticsMissingAfter    = -1
$script:AnalyticsMissingBeforeList = New-Object System.Collections.Generic.List[string]   # v1.11.0
$script:AnalyticsMissingAfterList  = New-Object System.Collections.Generic.List[string]   # v1.11.0
$script:AnalyticsInstalledDrivers = New-Object System.Collections.Generic.List[string]
# v1.13.1 - every file we successfully pull is recorded here so the HTML
# report can list re-downloadable driver URLs. Each entry is a hashtable:
#   @{ Url=...; FileName=...; MB=<double>; Kind='driver'|'catalog' }
# Kind='catalog' = vendor catalog/descriptor/matrix metadata (not a driver
# payload you'd re-download standalone); Kind='driver' = the actual pack /
# softpaq / MSI / installer. Deduped by exact URL.
$script:AnalyticsDownloadUrls = New-Object System.Collections.Generic.List[hashtable]

# v1.13.1 - classify + record a completed download. Called from both the
# serial Invoke-CurlDownload and the parallel Invoke-CurlDownloadParallel
# success paths. Never throws into the caller - a bad record must not break
# a download that already succeeded.
function Add-DownloadRecord {
    param(
        [string]$Url,
        [string]$OutFile,
        [double]$MB = 0.0
    )
    try {
        if ([string]::IsNullOrWhiteSpace($Url)) { return }
        foreach ($e in $script:AnalyticsDownloadUrls) {
            if ($e.Url -eq $Url) { return }   # dedupe by exact URL
        }
        $fileName = ""
        try   { $fileName = [System.IO.Path]::GetFileName(([System.Uri]$Url).LocalPath) }
        catch { $fileName = [System.IO.Path]::GetFileName($OutFile) }
        if ([string]::IsNullOrWhiteSpace($fileName)) {
            $fileName = [System.IO.Path]::GetFileName($OutFile)
        }
        # Catalog / metadata heuristic: vendor catalog cabs, descriptor /
        # catalog XML, the HP driver-pack matrix HTML. Anything else
        # (.exe / .msi / non-catalog .cab) is a re-downloadable driver.
        $lf   = $fileName.ToLower()
        $kind = if ($lf -match 'catalog' -or $lf -match '\.xml$' -or $lf -match '\.html?$') {
            'catalog'
        } else {
            'driver'
        }
        $script:AnalyticsDownloadUrls.Add(@{
            Url      = $Url
            FileName = $fileName
            MB       = $MB
            Kind     = $kind
        }) | Out-Null
    } catch {
        Log-Diag "Add-DownloadRecord skipped a URL: $($_.Exception.Message)"
    }
}

function Send-AnalyticsEvent {
    param(
        [ValidateSet("success","failure","cancelled","testmode")]
        [string]$Result
    )
    $durationSec = 0
    if ($script:AnalyticsStartTime) {
        $durationSec = [int]((Get-Date) - $script:AnalyticsStartTime).TotalSeconds
    }
    # v1.11.0 - DRY helper. Was duplicated for installed_drivers; also used
    # for missing_after_list and (since v1.13.2) the driver_urls list. Strips
    # control chars, escapes JSON specials, joins, caps to a sane length so a
    # single spreadsheet cell can hold the result.
    $cleanForJsonList = {
        param($items)
        if (-not $items -or $items.Count -eq 0) { return "" }
        $cleaned = foreach ($d in $items) {
            if ([string]::IsNullOrWhiteSpace($d)) { continue }
            $t = ($d -replace '[\r\n\t]', ' ').Trim()
            $t = $t -replace '\\', '\\'
            $t = $t -replace '"',  '\"'
            $t
        }
        $joined = ($cleaned -join ', ')
        if ($joined.Length -gt 8000) {
            $joined = $joined.Substring(0, 8000) + '...[truncated]'
        }
        return $joined
    }

    $installedDriversStr   = & $cleanForJsonList $script:AnalyticsInstalledDrivers
    # v1.13.2 - missing_before_list dropped from the spreadsheet payload, replaced
    # with driver_urls. Extract just the URL strings from the v1.13.1
    # AnalyticsDownloadUrls list, filtered to kind='driver' (vendor catalog /
    # descriptor / matrix URLs are excluded - they aren't standalone drivers).
    # The local HTML report still uses $script:AnalyticsMissingBeforeList for
    # its "Missing drivers before run" section, so that variable stays populated.
    $driverUrlList = New-Object System.Collections.Generic.List[string]
    if ($script:AnalyticsDownloadUrls) {
        foreach ($e in $script:AnalyticsDownloadUrls) {
            if (-not $e -or $e.Kind -ne 'driver') { continue }
            if ([string]::IsNullOrWhiteSpace($e.Url)) { continue }
            $driverUrlList.Add($e.Url) | Out-Null
        }
    }
    $driverUrlsStr         = & $cleanForJsonList $driverUrlList
    $missingAfterListStr   = & $cleanForJsonList $script:AnalyticsMissingAfterList

    $payload = @"
{
  "result":              "$Result",
  "manufacturer":        "$($script:AnalyticsManufacturer -replace '"','\"')",
  "model":               "$($script:AnalyticsModel -replace '"','\"')",
  "serial":              "$($script:AnalyticsSerial -replace '"','\"')",
  "os_version":          "$($script:AnalyticsOsVersion -replace '"','\"')",
  "os_build":            $($script:AnalyticsOsBuild),
  "inf_count":           $($script:AnalyticsInfCount),
  "download_mb":         $($script:AnalyticsDownloadMB),
  "missing_before":      $($script:AnalyticsMissingBefore),
  "missing_after":       $($script:AnalyticsMissingAfter),
  "duration_sec":        $durationSec,
  "script_version":      "$SCRIPT_VERSION",
  "installed_drivers":   "$installedDriversStr",
  "driver_urls":         "$driverUrlsStr",
  "missing_after_list":  "$missingAfterListStr"
}
"@

    # v1.10.0 - persist payload to Downloads alongside the .log so analytics is
    # never lost (previously only existed in %TEMP% for the duration of the curl
    # POST, then deleted). This also gives -NoAnalytics runs a local record.
    try {
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($AnalyticsFile, $payload, $utf8NoBom)
        Log-Diag "Analytics payload persisted to: $AnalyticsFile"
    } catch {
        Log "  WARNING: could not write analytics JSON to Downloads - $($_.Exception.Message)"
    }

    Log "Sending analytics (result=$Result, model=$($script:AnalyticsModel), infs=$($script:AnalyticsInfCount), dl=$($script:AnalyticsDownloadMB)MB, missing=$($script:AnalyticsMissingBefore)->$($script:AnalyticsMissingAfter), installed=$($script:AnalyticsInstalledDrivers.Count), duration=${durationSec}s)..." `
        -Level "info" -Event "analytics_send" -Context @{
            result        = $Result
            inf_count     = $script:AnalyticsInfCount
            download_mb   = $script:AnalyticsDownloadMB
            duration_sec  = $durationSec
        }

    if ($script:NoAnalytics) {
        Log "  -NoAnalytics set: skipping Sheets webhook. Local copy at $AnalyticsFile"
        return
    }

    try {
        # Use the persisted file as curl's data source - no second temp file needed.
        $curlArgs = "--silent --max-time 15 --connect-timeout 10 " +
                    "--location " +
                    "-H `"Content-Type: application/json`" " +
                    "--data `@`"$AnalyticsFile`" " +
                    "`"$SHEETS_WEBHOOK`""
        $psi                        = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = "curl.exe"
        $psi.Arguments              = $curlArgs
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $proc                       = New-Object System.Diagnostics.Process
        $proc.StartInfo             = $psi
        $proc.Start() | Out-Null
        $stdout = $proc.StandardOutput.ReadToEnd().Trim()
        $stderr = $proc.StandardError.ReadToEnd().Trim()
        $proc.WaitForExit()
        if ($proc.ExitCode -ne 0) {
            Log "  Analytics warning: curl exit $($proc.ExitCode) - $stderr" `
                -Level "warn" -Event "analytics_fail"
        } elseif ($stdout -eq "OK") {
            Log "  Analytics sent OK - row written to Google Sheet." `
                -Level "info" -Event "analytics_ok"
        } else {
            # v1.11.1 - When the Apps Script throws server-side, it returns a
            # *200 OK* with an HTML error page instead of plain "OK". The full
            # HTML is useless in the human log (multi-line, inline CSS) but the
            # error message itself is sitting in the centered <div>. Extract it
            # for the .log; keep the full HTML in the structured event so it's
            # still recoverable for deep debugging.
            $brief = $stdout
            $isHtmlError = $stdout -match '(?i)<html|<!DOCTYPE'
            if ($isHtmlError) {
                # Apps Script's error page puts the message in a monospace div.
                # Match its content first; fall back to any <title>Error</title>
                # neighbour, and last to any "*Error:..." substring.
                if ($stdout -match '(?is)<div[^>]*monospace[^>]*>(.+?)</div>') {
                    $brief = $matches[1]
                } elseif ($stdout -match '(?i)(ReferenceError|TypeError|SyntaxError|RangeError|Exception)[^<]+') {
                    $brief = $matches[0]
                } else {
                    $brief = '(Apps Script returned an HTML error page - see events.json for full body)'
                }
                # Decode the most common HTML entities so the log is readable.
                $brief = $brief -replace '&#39;', "'" -replace '&quot;', '"' `
                                 -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>'
                $brief = ($brief -replace '\s+', ' ').Trim()
                Log "  Analytics FAILED (server-side Apps Script error):" `
                    -Level "error" -Event "analytics_fail"
                Log "    $brief" -Level "error"
                Log "    To fix: open your Apps Script (Extensions > Apps Script in the Sheet)" `
                    -Level "error"
                Log "    and check the doPost function. Full HTML body is in the .events.json file." `
                    -Level "error"
            } else {
                # Non-HTML, non-OK response - some other unexpected payload.
                Log "  Analytics unexpected response: $brief" `
                    -Level "warn" -Event "analytics_fail"
            }
            # Always include the full body in the structured event for forensics.
            Log "[diag] analytics raw response body follows" `
                -Level "debug" -Event "analytics_raw_body" -Context @{ body = $stdout }
        }
    } catch {
        # Analytics failures must NEVER stop script execution - just log + carry on.
        Log "  Analytics error: $($_.Exception.Message)" `
            -Level "error" -Event "analytics_error"
    }
}

# =========================
# CURL DOWNLOAD
# =========================
function Invoke-CurlDownload {
    param(
        [string]$Url,
        [string]$OutFile
    )
    $fileName = [System.IO.Path]::GetFileName($OutFile)
    Log "Downloading: $fileName"
    Log "  URL: $Url"
    SetDownload -Pct 0 -Label "Connecting..."

    $totalBytes = 0
    try {
        $headResult = & curl.exe --silent --head --max-time 15 --connect-timeout 10 `
            --user-agent "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" `
            --write-out "%{content_length_download}" --output NUL "$Url" 2>$null
        if ($headResult -match '^\d+$' -and [long]$headResult -gt 0) {
            $totalBytes = [long]$headResult
        }
    } catch {}

    $totalMB = if ($totalBytes -gt 0) { [math]::Round($totalBytes / 1MB, 1) } else { 0 }
    if ($totalMB -gt 0) { Log "  Expected size: $totalMB MB" }

    $psi                 = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName        = "curl.exe"
    $psi.Arguments       = "--location --fail --connect-timeout 30 " +
                           "--retry 10 --retry-delay 5 --retry-all-errors " +
                           "--continue-at - " +
                           "--user-agent `"Mozilla/5.0 (Windows NT 10.0; Win64; x64)`" " +
                           "--output `"$OutFile`" `"$Url`""
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    $proc                = New-Object System.Diagnostics.Process
    $proc.StartInfo      = $psi
    $proc.Start() | Out-Null

    $lastSize = 0; $stall = 0; $prevSize = 0
    while (-not $proc.HasExited) {
        Start-Sleep -Milliseconds 700
        # Honour cancel mid-download: kill curl immediately so the next ~10MB
        # don't keep streaming after the user clicked Cancel. Without this,
        # curl runs to completion (potentially hundreds of MB) while the
        # caller's Test-Cancelled check just sits waiting for HasExited.
        if ($script:CancelRequested) {
            Log "  Cancel detected - killing curl."
            try { $proc.Kill() } catch {}
            try { $proc.WaitForExit(2000) | Out-Null } catch {}
            # Remove the partial file so a retry doesn't try to resume from a
            # half-baked download (curl --continue-at would otherwise pick it
            # up and possibly succeed on a Range request from byte N).
            Remove-Item $OutFile -EA SilentlyContinue
            SetDownload -Pct 0 -Label "Cancelled."
            Stop-DlSpinner -Success $false
            return $false
        }
        $sz = if (Test-Path $OutFile) { (Get-Item $OutFile -EA SilentlyContinue).Length } else { 0 }
        if ($sz -gt $lastSize) { $stall = 0; $lastSize = $sz } else { $stall++ }
        $mbDone    = [math]::Round($sz / 1MB, 1)
        $speedMbps = [math]::Round(($sz - $prevSize) * 8 / 1MB / 0.7, 1)
        $prevSize  = $sz
        $speedStr  = if ($speedMbps -gt 0) { "  $speedMbps Mbps" } else { "" }
        if ($totalMB -gt 0) {
            $pct = [math]::Min([int](($sz / $totalBytes) * 100), 99)
            SetDownload -Pct $pct -Label "$mbDone MB / $totalMB MB  ($pct%)$speedStr"
            Log "  $mbDone MB / $totalMB MB ($pct%)$speedStr"
        } else {
            SetDownload -Pct 0 -Label "$mbDone MB received...$speedStr"
            Log "  $mbDone MB received$speedStr"
        }
        Step-AllSpinners
        if ($stall -gt 300) {
            Log "  WARNING: download stalled 3.5 min - aborting."
            $proc.Kill()
            SetDownload -Pct 0 -Label "Stalled - aborted."
            Stop-DlSpinner -Success $false
            return $false
        }
    }

    if ($proc.ExitCode -ne 0) {
        Log "  curl failed (exit $($proc.ExitCode))"
        SetDownload -Pct 0 -Label "Failed (curl exit $($proc.ExitCode))"
        Stop-DlSpinner -Success $false
        return $false
    }
    if (-not (Test-Path $OutFile) -or (Get-Item $OutFile).Length -eq 0) {
        Log "  File missing or empty after download."
        SetDownload -Pct 0 -Label "Failed - file empty."
        Stop-DlSpinner -Success $false
        return $false
    }

    $finalMB = [math]::Round((Get-Item $OutFile).Length / 1MB, 1)
    if ($finalMB -gt $script:AnalyticsDownloadMB) { $script:AnalyticsDownloadMB = $finalMB }
    Add-DownloadRecord -Url $Url -OutFile $OutFile -MB $finalMB   # v1.13.1
    Log "  Download complete: $finalMB MB"
    SetDownload -Pct 100 -Label "Complete - $finalMB MB"
    Stop-DlSpinner -Success $true
    Play-Sound -Event "DownloadComplete"
    return $true
}

# =========================
# v1.11.0 - PARALLEL CURL DOWNLOAD
#
# Launches up to $MaxConcurrency curl.exe processes at once and polls them
# together. Designed for the HP catalog softpaq loop and the Dell CatalogPC
# individual-driver loop, where N small files each come from a fresh CDN
# connection; running them serially serialises TCP slow-start overhead.
#
# Inputs:
#   $Items: array of @{ Url=...; OutFile=...; Label='Display name' }
#   $MaxConcurrency: cap (defaults to script-scope $MaxParallelDownloads, min 1)
#
# Returns: array of @{ Item=...; Success=$bool; Bytes=long; ErrorMsg=string }
# in the same order as $Items. Callers can then iterate the returned array
# and run any post-download serial work (hash verify, extract, install).
#
# Analytics: unlike Invoke-CurlDownload which uses MAX(downloads) for the
# AnalyticsDownloadMB field, this helper SUMS the bytes of the whole batch
# and adds the sum to AnalyticsDownloadMB. That's strictly more accurate
# for the multi-file catalog paths. Serial Invoke-CurlDownload keeps its
# MAX semantics so existing single-pack analytics rows don't change shape.
#
# Cancellation: $script:CancelRequested checked every 700ms; when set, every
# active curl process is killed in parallel and partial files deleted.
# =========================
function Invoke-CurlDownloadParallel {
    param(
        [Parameter(Mandatory=$true)] [hashtable[]]$Items,
        [int]$MaxConcurrency = 0
    )
    if (-not $Items -or $Items.Count -eq 0) { return @() }
    if ($MaxConcurrency -le 0) {
        $MaxConcurrency = if ($script:MaxParallelDownloads -gt 0) { $script:MaxParallelDownloads } else { 3 }
    }
    # Above 6 we're just thrashing the CDN and getting throttled. Below 1 makes no sense.
    $MaxConcurrency = [math]::Min([math]::Max($MaxConcurrency, 1), 6)
    $total = $Items.Count
    Log "Parallel download: $total file(s), concurrency=$MaxConcurrency" `
        -Level "info" -Event "parallel_dl_start" -Context @{ total=$total; concurrency=$MaxConcurrency }
    SetDownload -Pct 0 -Label "Parallel: 0/$total complete"

    # Result objects (one per input item, indexed identically)
    $results = New-Object 'System.Collections.Generic.List[hashtable]'
    foreach ($_ in $Items) {
        $results.Add(@{ Item=$_; Success=$false; Bytes=0L; ErrorMsg=$null; Proc=$null; Started=$null; LastSize=0L; Stall=0 }) | Out-Null
    }

    $nextIdx       = 0
    $activeIdxs    = New-Object 'System.Collections.Generic.List[int]'
    $completedCnt  = 0
    $failedCnt     = 0
    $batchBytes    = 0L
    $StallLimitSec = 210     # ~3.5 min, matches serial Invoke-CurlDownload

    while ($completedCnt -lt $total) {
        # Launch more processes up to the concurrency cap
        while ($activeIdxs.Count -lt $MaxConcurrency -and $nextIdx -lt $total) {
            $r    = $results[$nextIdx]
            $item = $r.Item
            $url  = $item.Url
            $out  = $item.OutFile
            $lbl  = if ($item.Label) { $item.Label } else { [System.IO.Path]::GetFileName($out) }

            $psi                 = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName        = "curl.exe"
            $psi.Arguments       = "--location --fail --connect-timeout 30 " +
                                   "--retry 10 --retry-delay 5 --retry-all-errors " +
                                   "--continue-at - " +
                                   "--user-agent `"Mozilla/5.0 (Windows NT 10.0; Win64; x64)`" " +
                                   "--output `"$out`" `"$url`""
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow  = $true
            try {
                $proc            = New-Object System.Diagnostics.Process
                $proc.StartInfo  = $psi
                $null = $proc.Start()
                $r.Proc    = $proc
                $r.Started = Get-Date
                $activeIdxs.Add($nextIdx) | Out-Null
                Log-Diag "  parallel[$nextIdx]: started '$lbl' -> $out"
            } catch {
                $r.ErrorMsg = "curl launch failed: $($_.Exception.Message)"
                $r.Success  = $false
                $completedCnt++
                $failedCnt++
                Log "  parallel[$nextIdx]: launch failed - $($r.ErrorMsg)" -Level "error"
            }
            $nextIdx++
        }

        Start-Sleep -Milliseconds 700

        # Cancel check - bail entire batch
        if ($script:CancelRequested) {
            Log "  Parallel cancel detected - killing all active curls." -Level "cancel"
            foreach ($i in $activeIdxs) {
                $r = $results[$i]
                if ($r.Proc -and -not $r.Proc.HasExited) {
                    try { $r.Proc.Kill() } catch {}
                    try { $r.Proc.WaitForExit(2000) | Out-Null } catch {}
                }
                $of = $r.Item.OutFile
                if ($of -and (Test-Path $of)) { Remove-Item $of -EA SilentlyContinue }
                $r.Success  = $false
                $r.ErrorMsg = "cancelled"
            }
            SetDownload -Pct 0 -Label "Cancelled."
            Stop-DlSpinner -Success $false
            return $results
        }

        # Poll each active process: handle exit, stall, progress aggregation
        $stillActive = New-Object 'System.Collections.Generic.List[int]'
        foreach ($i in $activeIdxs) {
            $r = $results[$i]
            $proc = $r.Proc
            $of   = $r.Item.OutFile
            $sz   = if ($of -and (Test-Path $of)) { (Get-Item $of -EA SilentlyContinue).Length } else { 0 }
            if ($sz -gt $r.LastSize) { $r.Stall = 0; $r.LastSize = $sz } else { $r.Stall++ }

            if ($proc.HasExited) {
                if ($proc.ExitCode -eq 0 -and (Test-Path $of) -and (Get-Item $of).Length -gt 0) {
                    $r.Bytes   = (Get-Item $of).Length
                    $r.Success = $true
                    $batchBytes += $r.Bytes
                    $mb = [math]::Round($r.Bytes / 1MB, 1)
                    $lbl = if ($r.Item.Label) { $r.Item.Label } else { [System.IO.Path]::GetFileName($of) }
                    Add-DownloadRecord -Url $r.Item.Url -OutFile $of -MB $mb   # v1.13.1
                    Log "  parallel[$i] OK: $lbl ($mb MB)" -Level "info" -Event "parallel_dl_ok"
                } else {
                    $r.Success  = $false
                    $r.ErrorMsg = "curl exit $($proc.ExitCode)"
                    Log "  parallel[$i] FAIL: $($r.ErrorMsg)" -Level "warn" -Event "parallel_dl_fail"
                    if (Test-Path $of) { Remove-Item $of -EA SilentlyContinue }
                    $failedCnt++
                }
                $completedCnt++
            } elseif ($r.Stall -gt ($StallLimitSec * 1000 / 700)) {
                # Per-process stall: kill just this one, don't affect siblings
                Log "  parallel[$i] STALLED >$StallLimitSec`s - killing." -Level "warn"
                try { $proc.Kill() } catch {}
                try { $proc.WaitForExit(2000) | Out-Null } catch {}
                if (Test-Path $of) { Remove-Item $of -EA SilentlyContinue }
                $r.Success  = $false
                $r.ErrorMsg = "stalled"
                $completedCnt++
                $failedCnt++
            } else {
                $stillActive.Add($i) | Out-Null
            }
        }
        $activeIdxs = $stillActive

        # UI update
        $okSoFar = $completedCnt - $failedCnt
        $pct     = if ($total -gt 0) { [int](($completedCnt / $total) * 100) } else { 0 }
        if ($pct -ge 100) { $pct = 99 }  # leave 100 for the post-loop final
        $activeMB = 0
        foreach ($i in $activeIdxs) { $activeMB += ($results[$i].LastSize / 1MB) }
        $batchMB  = [math]::Round(($batchBytes / 1MB) + $activeMB, 1)
        SetDownload -Pct $pct -Label "Parallel: $okSoFar/$total complete ($batchMB MB, $($activeIdxs.Count) active)"
        Step-AllSpinners
    }

    # Final accounting
    $batchMB = [math]::Round($batchBytes / 1MB, 1)
    $script:AnalyticsDownloadMB = [math]::Round($script:AnalyticsDownloadMB + $batchMB, 1)
    $ok = ($results | Where-Object { $_.Success }).Count
    Log "Parallel download finished: $ok/$total OK, total $batchMB MB" `
        -Level "info" -Event "parallel_dl_done" -Context @{ ok=$ok; total=$total; mb=$batchMB }
    SetDownload -Pct 100 -Label "Parallel: $ok/$total complete ($batchMB MB)"
    Stop-DlSpinner -Success ($ok -gt 0)
    if ($ok -gt 0) { Play-Sound -Event "DownloadComplete" }
    return $results
}

# =========================
# EXTRACTION WATCHER
# =========================
function Watch-Extraction {
    param(
        [System.Diagnostics.Process]$ExtractProc,
        [string]$DestPath,
        [int]$TotalFiles    = 0,
        [int]$StallLimitSec = 300
    )
    $stall     = 0
    $lastCount = 0
    $script:SpinnerIndex      = 0
    $exSpinnerLabel.Text      = " " + $SpinnerFrames[0]
    $overallSpinnerLabel.Text = " " + $SpinnerFrames[0]
    SetExtract -Pct -1 -Label "Extracting..."

    while (-not $ExtractProc.HasExited) {
        Start-Sleep -Milliseconds 700
        $count = if (Test-Path $DestPath) {
            (Get-ChildItem $DestPath -Recurse -ErrorAction SilentlyContinue).Count
        } else { 0 }
        if ($count -gt $lastCount) { $stall = 0; $lastCount = $count } else { $stall++ }
        if ($TotalFiles -gt 0) {
            $remaining = [math]::Max($TotalFiles - $count, 0)
            SetExtract -Pct -1 -Label "$count / $TotalFiles files  ($remaining remaining)"
        } else {
            SetExtract -Pct -1 -Label "$count files extracted..."
        }
        Step-ExSpinner
        [System.Windows.Forms.Application]::DoEvents()
        if ($script:CancelRequested) {
            Log "  Extraction cancelled by user."
            try { $ExtractProc.Kill() } catch {}
            break
        }
        if ($stall -gt [int]($StallLimitSec * 1.25)) {
            Log "  WARNING: extraction stalled - killing process."
            try { $ExtractProc.Kill() } catch {}
            break
        }
    }
    Start-Sleep -Seconds 2
    $finalCount = if (Test-Path $DestPath) {
        (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
    } else { 0 }
    SetExtract -Pct 100 -Label "Done - $finalCount files extracted"
    Stop-ExSpinner      -Success $true
    Stop-OverallSpinner -Success $true
    Play-Sound -Event "ExtractComplete"
    Log "  Extraction finished: $finalCount files in $DestPath"
}

# =========================
# INF INSTALLER
# =========================
function Install-DriversFromPath {
    param([string]$BasePath)
    $infs = Get-ChildItem $BasePath -Recurse -Filter *.inf -ErrorAction SilentlyContinue
    if (-not $infs -or $infs.Count -eq 0) {
        Log "No INF files found under: $BasePath"
        return $false
    }
    $total = $infs.Count; $i = 0
    Log "Found $total INF file(s)."
    if ($SkipInstall) {
        Log "  SkipInstall flag set - skipping pnputil. Extraction verified OK."
        $script:AnalyticsInfCount = 0
        SetProgress 100
        SetExtract -Pct 100 -Label "Extract complete ($total INFs found, install skipped)"
        Stop-ExSpinner      -Success $true
        Stop-OverallSpinner -Success $true
        return $true
    }
    Log "Installing via pnputil..."
    $exHeaderLabel.Text  = "Install INFs"
    $exSpinnerLabel.Text = " " + $SpinnerFrames[0]
    $script:SpinnerIndex = 0
    $exBar.Style         = "Continuous"
    $exBar.Value         = 0

    foreach ($inf in $infs) {
        $i++
        SetProgress (60 + [int](($i / $total) * 38))
        $infPct    = [int](($i / $total) * 100)
        $remaining = $total - $i
        SetExtract -Pct $infPct -Label "$i / $total INFs  ($remaining remaining) - $($inf.Name)"
        Log "[$i/$total] $($inf.Name)"
        $out = pnputil /add-driver "`"$($inf.FullName)`"" /install 2>&1
        $infExit = $LASTEXITCODE
        foreach ($l in $out) { Log "  $l" }
        # pnputil returns exit 0 for both "added to driver store" AND "installed
        # on a matching device". Most INFs in an OEM pack get added to the store
        # but never bind to a device on this machine - we don't want those in
        # the analytics. Parse stdout for an explicit device-bind marker; the
        # exact wording varies by Windows version so we accept any of three.
        if ($infExit -in @(0, 1641, 3010)) {
            $outText = ($out | Out-String)
            if ($outText -match 'Installed driver package on matching|Installed on\s+[1-9]\d*\s+device|Successfully installed driver') {
                $null = $script:AnalyticsInstalledDrivers.Add($inf.Name)
            }
        }
        Play-Sound -Event "DriverAdded"
        Step-ExSpinner
        Step-OverallSpinner
        [System.Windows.Forms.Application]::DoEvents()
        if ($script:CancelRequested) {
            Log "INF installation cancelled at $i / $total"
            $script:AnalyticsInfCount = $i
            break
        }
    }
    $script:AnalyticsInfCount = $i
    SetProgress 100
    SetExtract -Pct 100 -Label "All $total INFs installed."
    Stop-ExSpinner      -Success $true
    Stop-OverallSpinner -Success $true
    $exHeaderLabel.Text = "Extract & install"
    Log "All INFs processed."
    return $true
}

# =========================
# SHARED EXTRACT RUNNER
# =========================
function Start-PackExtraction {
    param(
        [string]$PackFile,
        [string]$DestPath,
        [int]$StallLimitSec = 300,
        [ValidateSet("Dell","HP","Lenovo","")]
        [string]$Vendor = ""
    )
    if (-not (Test-Path $DestPath)) {
        New-Item -Path $DestPath -ItemType Directory -Force | Out-Null
    }
    $ext = [System.IO.Path]::GetExtension($PackFile).ToLower()

    switch ($ext) {
        ".zip" {
            Log "Extracting ZIP..."
            SetExtract -Pct -1 -Label "Starting ZIP extraction..."
            $zipJob = Start-Job {
                param($src, $dst)
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                [System.IO.Compression.ZipFile]::ExtractToDirectory($src, $dst)
            } -ArgumentList $PackFile, $DestPath

            $stall = 0; $lastCount = 0
            $script:SpinnerIndex      = 0
            $exSpinnerLabel.Text      = " " + $SpinnerFrames[0]
            $overallSpinnerLabel.Text = " " + $SpinnerFrames[0]
            while ($zipJob.State -eq "Running") {
                Start-Sleep -Milliseconds 700
                $count = if (Test-Path $DestPath) {
                    (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
                } else { 0 }
                if ($count -gt $lastCount) { $stall = 0; $lastCount = $count } else { $stall++ }
                SetExtract -Pct -1 -Label "$count files extracted..."
                Step-ExSpinner
                Step-OverallSpinner
                [System.Windows.Forms.Application]::DoEvents()
                if ($stall -gt 375) { Log "  ZIP stalled - stopping."; Stop-Job $zipJob; break }
            }
            Receive-Job $zipJob -EA SilentlyContinue | Out-Null
            Remove-Job  $zipJob
            $finalCount = if (Test-Path $DestPath) {
                (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
            } else { 0 }
            SetExtract -Pct 100 -Label "Done - $finalCount files extracted"
            Stop-ExSpinner      -Success $true
            Stop-OverallSpinner -Success $true
            Play-Sound -Event "ExtractComplete"
            Log "  ZIP extraction complete. $finalCount files."
        }

        ".cab" {
            Log "Extracting CAB..."
            $psi                 = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName        = "expand.exe"
            $psi.Arguments       = "`"$PackFile`" -F:* `"$DestPath`""
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow  = $true
            $exProc              = New-Object System.Diagnostics.Process
            $exProc.StartInfo    = $psi
            $exProc.Start() | Out-Null
            Watch-Extraction -ExtractProc $exProc -DestPath $DestPath -StallLimitSec $StallLimitSec
        }

        default {
            # ------------------------------------------------------------------
            # Dell and HP: use 7-Zip pass-1 when available.
            #   Verified against real packs - byte-for-byte identical output,
            #   faster than vendor extractors (Dell 1.1x, HP 4.7x).
            #   Falls back to vendor extractor if 7-Zip unavailable or yields 0 files.
            #
            # Lenovo: always uses Inno Setup /VERYSILENT /EXTRACT=YES.
            #   7-Zip cannot extract Lenovo's proprietary Inno payload format.
            # ------------------------------------------------------------------

            $CountFiles = {
                if (Test-Path $DestPath) {
                    (Get-ChildItem $DestPath -Recurse -EA SilentlyContinue).Count
                } else { 0 }
            }

            # Try 7-Zip pass-1 for Dell and HP
            if (($Vendor -eq "Dell" -or $Vendor -eq "HP") -and (Test-Path $script:7zExe)) {
                Log "  Extracting with 7-Zip (pass 1)..."
                SetExtract -Pct -1 -Label "Extracting with 7-Zip..."
                $script:SpinnerIndex      = 0
                $exSpinnerLabel.Text      = " " + $SpinnerFrames[0]
                $overallSpinnerLabel.Text = " " + $SpinnerFrames[0]

                $psi                 = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName        = $script:7zExe
                $psi.Arguments       = "x `"$PackFile`" -o`"$DestPath`" -y"
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow  = $true
                $sevenProc           = New-Object System.Diagnostics.Process
                $sevenProc.StartInfo = $psi
                $sevenProc.Start() | Out-Null

                while (-not $sevenProc.HasExited) {
                    Start-Sleep -Milliseconds 700
                    $n = & $CountFiles
                    SetExtract -Pct -1 -Label "$n files extracted..."
                    Step-ExSpinner
                    Step-OverallSpinner
                    [System.Windows.Forms.Application]::DoEvents()
                    if ($script:CancelRequested) { try { $sevenProc.Kill() } catch {}; break }
                }
                Start-Sleep -Seconds 1
                $n7z = & $CountFiles

                if ($n7z -gt 0) {
                    Log "  7-Zip extraction complete: $n7z files."
                    SetExtract -Pct 100 -Label "Done - $n7z files extracted"
                    Stop-ExSpinner      -Success $true
                    Stop-OverallSpinner -Success $true
                    Play-Sound -Event "ExtractComplete"
                    return  # success - skip fallback chain
                }
                Log "  7-Zip yielded 0 files - falling back to vendor extractor."
            }

            # Vendor extractor fallback (always used for Lenovo, fallback for Dell/HP)
            Log "Extracting EXE pack (Vendor=$Vendor)..."

            function TrySync {
                param([string]$ExeArgs)
                $p                 = New-Object System.Diagnostics.ProcessStartInfo
                $p.FileName        = $PackFile
                $p.Arguments       = $ExeArgs
                $p.UseShellExecute = $true
                $p.WindowStyle     = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $proc = New-Object System.Diagnostics.Process
                $proc.StartInfo = $p
                $proc.Start() | Out-Null
                while (-not $proc.HasExited) {
                    Start-Sleep -Milliseconds 700
                    $n = & $CountFiles
                    SetExtract -Pct -1 -Label "$n files extracted..."
                    Step-ExSpinner
                    Step-OverallSpinner
                    [System.Windows.Forms.Application]::DoEvents()
                    if ($script:CancelRequested) { try { $proc.Kill() } catch {}; break }
                }
                Start-Sleep -Seconds 2
                return (& $CountFiles)
            }

            function TryAsyncHP {
                $p                 = New-Object System.Diagnostics.ProcessStartInfo
                $p.FileName        = $PackFile
                $p.Arguments       = "/s /e /f `"$DestPath`""
                $p.UseShellExecute = $true
                $p.WindowStyle     = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $proc = New-Object System.Diagnostics.Process
                $proc.StartInfo = $p
                $proc.Start() | Out-Null
                $start = Get-Date
                while (-not $proc.HasExited) {
                    Start-Sleep -Milliseconds 700
                    $n = & $CountFiles
                    SetExtract -Pct -1 -Label "$n files extracted..."
                    Step-ExSpinner
                    [System.Windows.Forms.Application]::DoEvents()
                    if (((Get-Date) - $start).TotalSeconds -gt 30 -and $n -eq 0) {
                        Log "  HP format timed out with no output - killing."
                        try { $proc.Kill() } catch {}
                        break
                    }
                }
                Start-Sleep -Seconds 2
                return (& $CountFiles)
            }

            $attempts = [System.Collections.Generic.List[object]]::new()
            switch ($Vendor) {
                "Dell"   {
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";           Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "Lenovo (Inno /VERYSILENT)";   Action = { TrySync "/VERYSILENT /DIR=`"$DestPath`" /EXTRACT=YES" } })
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)";          Action = { TryAsyncHP } })
                }
                "Lenovo" {
                    $attempts.Add(@{ Label = "Lenovo (Inno /VERYSILENT)";   Action = { TrySync "/VERYSILENT /DIR=`"$DestPath`" /EXTRACT=YES" } })
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";           Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)";          Action = { TryAsyncHP } })
                }
                "HP"     {
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)";          Action = { TryAsyncHP } })
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";           Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "Lenovo (Inno /VERYSILENT)";   Action = { TrySync "/VERYSILENT /DIR=`"$DestPath`" /EXTRACT=YES" } })
                }
                default  {
                    $attempts.Add(@{ Label = "Dell (/s /e=dest)";           Action = { TrySync "/s /e=`"$DestPath`"" } })
                    $attempts.Add(@{ Label = "Lenovo (Inno /VERYSILENT)";   Action = { TrySync "/VERYSILENT /DIR=`"$DestPath`" /EXTRACT=YES" } })
                    $attempts.Add(@{ Label = "HP (/s /e /f dest)";          Action = { TryAsyncHP } })
                }
            }

            $extracted = $false
            foreach ($attempt in $attempts) {
                Log "  Trying: $($attempt.Label)..."
                $n = & $attempt.Action
                Log "  Result: $n files in $DestPath"
                if ($n -gt 0) {
                    SetExtract -Pct 100 -Label "Done - $n files extracted"
                    Stop-ExSpinner -Success $true; Stop-OverallSpinner -Success $true
                    Play-Sound -Event "ExtractComplete"
                    Log "  Extraction finished: $n files in $DestPath"
                    $extracted = $true
                    break
                }
            }

            if (-not $extracted) {
                Log "  All primary formats failed - trying Lenovo legacy (-s -fdest)..."
                $p                 = New-Object System.Diagnostics.ProcessStartInfo
                $p.FileName        = $PackFile
                $p.Arguments       = "-s -f`"$DestPath`""
                $p.UseShellExecute = $true
                $p.WindowStyle     = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $proc = New-Object System.Diagnostics.Process
                $proc.StartInfo = $p
                $proc.Start() | Out-Null
                Watch-Extraction -ExtractProc $proc -DestPath $DestPath -StallLimitSec $StallLimitSec
            }
        }
    }
}

# =========================
# DELL - INDIVIDUAL DRIVER INSTALL (1-3 missing devices)
#
# Uses CatalogPC.cab which contains individual driver entries,
# each with <SupportedDevices> hardware IDs. Downloads only the
# specific drivers needed for the missing devices rather than the
# full driver pack.
#
# Falls back to full pack install if any device has no catalog match.
# =========================
function Start-DellIndividualDriverInstall {
    param([string]$DriverRoot, [string]$ServiceTag, [bool]$IsWin11)

    Log "=== DELL: Individual driver mode (<=3 missing devices) ==="

    # Get missing devices and their hardware IDs
    Log "Enumerating missing devices..."
    $missingDevices = @()
    try {
        $missingDevices = @(Get-CimInstance Win32_PnPEntity |
            Where-Object { $_.ConfigManagerErrorCode -ne 0 } |
            Select-Object Name, DeviceID,
                @{ N='HardwareIDs'; E={
                    try { (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Enum\$($_.DeviceID)" -EA Stop).HardwareID }
                    catch { @() }
                }}
        )
    } catch {
        Log "  Failed to enumerate missing devices: $($_.Exception.Message)"
        return $null
    }

    if ($missingDevices.Count -eq 0) {
        Log "  No missing devices found on re-check."
        return $true
    }

    foreach ($dev in $missingDevices) {
        $ids = @($dev.HardwareIDs) | Where-Object { $_ }
        Log "  Missing: $($dev.Name)"
        foreach ($id in $ids) { Log "    HW ID: $id" }
    }

    # Get the SystemSKUNumber - used to filter drivers to this specific model
    $systemSKU = ""
    try {
        $systemSKU = (Get-CimInstance Win32_ComputerSystem).SystemSKUNumber.Trim().ToUpper()
        Log "  SystemSKUNumber: $systemSKU"
    } catch { Log "  WARNING: Could not read SystemSKUNumber - model filtering disabled." }

    # Download CatalogPC.cab
    Log "Downloading Dell CatalogPC.cab..."
    $catalogCab = Join-Path $env:TEMP "DellCatalogPC.cab"
    $catalogXml = Join-Path $env:TEMP "CatalogPC.xml"
    Remove-Item $catalogCab -EA SilentlyContinue
    Remove-Item $catalogXml -EA SilentlyContinue

    SetDownload -Pct 0 -Label "Downloading Dell CatalogPC..."
    if (-not (Invoke-CurlDownload -Url "https://downloads.dell.com/catalog/CatalogPC.cab" -OutFile $catalogCab)) {
        Log "  Failed to download CatalogPC.cab - falling back to full pack."
        return $null
    }
    if (Test-Cancelled) { return $false }

    Log "Extracting CatalogPC.cab..."
    SetExtract -Pct 10 -Label "Extracting catalog..."
    $expandOut = & expand.exe "`"$catalogCab`"" "`"$catalogXml`"" 2>&1
    Log "  expand.exe: $expandOut"
    if (-not (Test-Path $catalogXml) -or (Get-Item $catalogXml).Length -eq 0) {
        Log "  CatalogPC CAB extraction failed - falling back to full pack."
        return $null
    }
    SetExtract -Pct 30 -Label "Catalog extracted OK"

    # CatalogPC.xml is UTF-16 encoded
    Log "Parsing CatalogPC.xml (UTF-16)..."
    try {
        $rawXml   = [System.IO.File]::ReadAllText($catalogXml, [System.Text.Encoding]::Unicode)
        [xml]$cat = $rawXml
    } catch {
        Log "  Failed to parse CatalogPC.xml: $($_.Exception.Message) - falling back to full pack."
        return $null
    }

    # Win11 osCodes in CatalogPC.xml use W21xx prefix (not Display text)
    $win11Codes = @('W21H4','W21P4','W21S4','W21S5','W11AH','W11AP','W11S5','W11TM','IOTL5')
    $win10Codes = @('W10H4','W10P4','W10H2','W10P2','IOT01','IOTL3','IOTL4','WTCLD')

    # For each missing device, find matching driver entries in the catalog
    $toDownload = [System.Collections.Generic.List[object]]::new()
    $allMatched = $true

    foreach ($dev in $missingDevices) {
        # Parse VEN and DEV from each hardware ID string.
        # Handles both standard PCI IDs (PCI\VEN_xxx&DEV_xxx) and
        # INTELAUDIO bus IDs (INTELAUDIO\CTLR_DEV_xxx&...VEN_xxx&DEV_xxx)
        # where CTLR_DEV is the HDA controller and DEV is the codec endpoint.
        $devVenDev = @()
        $seen = @{}
        foreach ($hwId in @($dev.HardwareIDs) | Where-Object { $_ }) {
            $ven = $null; $dev2 = $null
            if ($hwId -match 'VEN_([0-9A-Fa-f]+)') { $ven = $Matches[1].ToUpper() }
            # For INTELAUDIO, use CTLR_DEV as the primary device ID (matches catalog PCIInfo)
            if ($hwId -match 'CTLR_DEV_([0-9A-Fa-f]+)') {
                $dev2 = $Matches[1].ToUpper()
            } elseif ($hwId -match '(?:^|&)DEV_([0-9A-Fa-f]+)') {
                $dev2 = $Matches[1].ToUpper()
            }
            if ($ven -and $dev2) {
                $key = "$ven|$dev2"
                if (-not $seen[$key]) {
                    $seen[$key] = $true
                    $devVenDev += [PSCustomObject]@{ VEN = $ven; DEV = $dev2 }
                    Log "    Parsed VEN=$ven DEV=$dev2 from: $hwId"
                }
            }
        }
        if ($devVenDev.Count -eq 0) {
            Log "  No PCI VEN/DEV found for '$($dev.Name)' - falling back to full pack."
            $allMatched = $false
            break
        }

        $matched = $false
        foreach ($component in $cat.SelectNodes("//*[local-name()='SoftwareComponent']")) {

            # Must be a driver (DRVR), not firmware
            $typeNode = $component.SelectSingleNode("*[local-name()='ComponentType']")
            if ($typeNode -and $typeNode.GetAttribute("value") -ne "DRVR") { continue }

            # Check OS compatibility via osCode attribute
            $osMatch = $false
            $targetCodes = if ($IsWin11) { $win11Codes } else { $win10Codes }
            foreach ($osNode in $component.SelectNodes(".//*[local-name()='OperatingSystem']")) {
                if ($targetCodes -contains $osNode.GetAttribute("osCode")) { $osMatch = $true; break }
            }
            if (-not $osMatch) { continue }

            # Check model compatibility via systemID (matches SystemSKUNumber)
            if ($systemSKU) {
                $modelMatch = $false
                foreach ($modelNode in $component.SelectNodes(".//*[local-name()='Model']")) {
                    if ($modelNode.GetAttribute("systemID").ToUpper() -eq $systemSKU) { $modelMatch = $true; break }
                }
                if (-not $modelMatch) { continue }
            }

            # Match hardware ID via PCIInfo deviceID + vendorID attributes
            foreach ($pciNode in $component.SelectNodes(".//*[local-name()='PCIInfo']")) {
                $catVen = $pciNode.GetAttribute("vendorID").ToUpper()
                $catDev = $pciNode.GetAttribute("deviceID").ToUpper()
                foreach ($vd in $devVenDev) {
                    if ($vd.VEN -eq $catVen -and $vd.DEV -eq $catDev) {
                        $driverPath = $component.GetAttribute("path")
                        $driverName = ""
                        try { $driverName = $component.SelectSingleNode("*[local-name()='Name']/*[local-name()='Display']").InnerText } catch {}
                        Log "  Matched '$($dev.Name)'"
                        Log "    Driver : $driverName"
                        Log "    Path   : $driverPath"
                        if ($driverPath -and -not ($toDownload | Where-Object { $_.Path -eq $driverPath })) {
                            $toDownload.Add([PSCustomObject]@{
                                DeviceName = $dev.Name
                                DriverName = $driverName
                                Path       = $driverPath
                            })
                        }
                        $matched = $true
                        break
                    }
                }
                if ($matched) { break }
            }
            if ($matched) { break }
        }

        if (-not $matched) {
            Log "  No CatalogPC match for '$($dev.Name)' - will fall back to full pack."
            $allMatched = $false
            break
        }
    }

    if (-not $allMatched) { return $null }

    if ($toDownload.Count -eq 0) {
        Log "  No drivers to download - all devices may already be covered."
        return $true
    }

    Log "Found $($toDownload.Count) driver(s) to download individually."
    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }

    # v1.11.0 - PARALLEL DOWNLOAD PHASE for Dell CatalogPC individual drivers.
    # Same shape as HP: build manifest, batch-download, then iterate survivors
    # serially for extraction (extraction needs disk I/O exclusivity per file).
    $dlItems = New-Object 'System.Collections.Generic.List[hashtable]'
    foreach ($drv in $toDownload) {
        $driverUrl  = "https://downloads.dell.com/$($drv.Path)"
        $driverFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName($drv.Path))
        $dlItems.Add(@{
            Url     = $driverUrl
            OutFile = $driverFile
            Label   = $drv.DriverName
        }) | Out-Null
    }
    SetProgress 22
    $dlResults = Invoke-CurlDownloadParallel -Items $dlItems
    if (Test-Cancelled) { return $false }

    $allDownloaded = $true
    foreach ($r in $dlResults) {
        if (-not $r.Success) {
            Log "  Download failed for '$($r.Item.Label)': $($r.ErrorMsg) - falling back to full pack."
            $allDownloaded = $false
        }
    }
    if (-not $allDownloaded) { return $null }

    # SERIAL EXTRACT PHASE (one extract dir per file - reuses indexing convention)
    $i = 0
    foreach ($drv in $toDownload) {
        $i++
        if (Test-Cancelled) { return $false }
        $driverFile  = Join-Path $DriverRoot ([System.IO.Path]::GetFileName($drv.Path))
        $extractPath = Join-Path $DriverRoot "Dell_Individual_$i"
        Log "  Extracting [$i/$($toDownload.Count)]: $($drv.DriverName)..."
        SetProgress (40 + [int](($i / $toDownload.Count) * 20))
        Start-PackExtraction -PackFile $driverFile -DestPath $extractPath -StallLimitSec 120 -Vendor "Dell"
    }

    # Install all extracted INFs (unchanged)
    $anyInstalled = $false
    foreach ($extractDir in (Get-Item (Join-Path $DriverRoot "Dell_Individual_*") -EA SilentlyContinue)) {
        if (Install-DriversFromPath -BasePath $extractDir.FullName) { $anyInstalled = $true }
    }
    return $anyInstalled
}

# =========================
# DELL
# =========================
function Start-DellDriverInstall {
    param([string]$DriverRoot, [string]$ModelName)

    Log "=== DELL: Starting automated driver install ==="
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."

    $serviceTag = $null
    try {
        $serviceTag = (Get-CimInstance Win32_BIOS).SerialNumber.Trim()
        Log "Service Tag: $serviceTag"
        $script:AnalyticsSerial = $serviceTag
    } catch {
        Log "Could not read Service Tag: $($_.Exception.Message)"
        return $false
    }
    if (-not $serviceTag -or $serviceTag.Length -lt 4) { Log "Invalid Service Tag."; return $false }

    $isWin11 = $false
    try {
        $isWin11 = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber -ge 22000
        Log "OS: $(if ($isWin11) {'Windows 11'} else {'Windows 10'})"
    } catch {}

    # Threshold check: if 1-3 missing devices, try individual driver lookup first.
    # Pre-screen: skip individual lookup if any missing device has no parseable VEN/DEV
    # (e.g. proprietary bus devices like Qualcomm QMUX that will never be in CatalogPC).
    $missingCount = $script:AnalyticsMissingBefore
    if ($missingCount -ge 1 -and $missingCount -le 3) {
        $allHaveVenDev = $true
        $missingEntities = @(Get-CimInstance Win32_PnPEntity | Where-Object { $_.ConfigManagerErrorCode -ne 0 })
        foreach ($ent in $missingEntities) {
            try {
                $hwIds = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Enum\$($ent.DeviceID)" -EA Stop).HardwareID
                $hasVenDev = $false
                foreach ($id in @($hwIds) | Where-Object { $_ }) {
                    $ven = $null; $dev2 = $null
                    if ($id -match 'VEN_([0-9A-Fa-f]+)') { $ven = $Matches[1] }
                    if ($id -match 'CTLR_DEV_([0-9A-Fa-f]+)') { $dev2 = $Matches[1] }
                    elseif ($id -match '(?:^|&)DEV_([0-9A-Fa-f]+)') { $dev2 = $Matches[1] }
                    if ($ven -and $dev2) { $hasVenDev = $true; break }
                }
                if (-not $hasVenDev) {
                    Log "Device '$($ent.Name)' has no parseable VEN/DEV (proprietary bus) - skipping individual lookup, using full pack."
                    $allHaveVenDev = $false
                    break
                }
            } catch {
                Log "Could not read HW IDs for '$($ent.Name)' - skipping individual lookup."
                $allHaveVenDev = $false
                break
            }
        }

        if ($allHaveVenDev) {
            Log "Missing devices ($missingCount) within threshold - attempting individual driver lookup..."
            $individualResult = Start-DellIndividualDriverInstall -DriverRoot $DriverRoot -ServiceTag $serviceTag -IsWin11 $isWin11
            if ($individualResult -eq $true) {
                Log "Individual driver install complete."
                return $true
            } elseif ($individualResult -eq $false) {
                # Cancelled or hard failure
                return $false
            }
            # $null = no catalog match found, fall through to full pack
            Log "Falling back to full driver pack install..."
        }
    }

    Log "Downloading Dell DriverPackCatalog.cab..."
    $catalogCab = Join-Path $env:TEMP "DellDriverPackCatalog.cab"
    $catalogXml = Join-Path $env:TEMP "DriverPackCatalog.xml"
    Remove-Item $catalogCab -EA SilentlyContinue
    Remove-Item $catalogXml -EA SilentlyContinue

    if (-not (Invoke-CurlDownload -Url "https://downloads.dell.com/catalog/DriverPackCatalog.cab" -OutFile $catalogCab)) {
        Log "Failed to download Dell DriverPackCatalog."
        return $false
    }
    if (Test-Cancelled) { return $false }
    SetProgress 15

    Log "Extracting Dell DriverPackCatalog CAB..."
    SetExtract -Pct 10 -Label "Extracting catalog..."
    $expandOut = & expand.exe "`"$catalogCab`"" "`"$catalogXml`"" 2>&1
    Log "  expand.exe: $expandOut"
    if (-not (Test-Path $catalogXml) -or (Get-Item $catalogXml).Length -eq 0) {
        Log "DriverPackCatalog CAB extraction failed."
        return $false
    }
    SetExtract -Pct 30 -Label "Catalog extracted OK"
    SetProgress 20

    Log "Parsing DriverPackCatalog..."
    try {
        $rawXml   = [System.IO.File]::ReadAllText($catalogXml).TrimStart([char]0xFEFF)
        [xml]$cat = $rawXml
    } catch {
        Log "Failed to parse DriverPackCatalog XML: $($_.Exception.Message)"
        return $false
    }

    # Build primary search token list (exact model + light cleanups).
    $searchNames = @()
    if ($ModelName) {
        $searchNames += $ModelName
        $searchNames += ($ModelName -replace '\s+Notebook.*$','')
        $searchNames += ($ModelName -replace '^Dell\s+','')
        $searchNames = @($searchNames | Select-Object -Unique | Where-Object { $_.Length -gt 3 })
    }
    Log "Searching catalog - model tokens: $($searchNames -join ' | ')"

    # Reusable matcher: enumerates DriverPackage nodes and returns candidates
    # whose <Model name="..."> matches any token (case-insensitive equality).
    # Defined as a scriptblock so we can re-run it with a broader token list
    # if the first pass returns nothing.
    $findCandidates = {
        param([string[]]$tokens)
        $out = @()
        foreach ($pkg in $cat.SelectNodes("//*[local-name()='DriverPackage']")) {
            $modelMatched = $false
            foreach ($modelNode in $pkg.SelectNodes(".//*[local-name()='Model']")) {
                $nameAttr = $modelNode.GetAttribute("name")
                foreach ($tok in $tokens) {
                    if ($nameAttr -ieq $tok) { $modelMatched = $true; break }
                }
                if ($modelMatched) { break }
            }
            if (-not $modelMatched) { continue }

            $supportsWin11 = $false; $supportsWin10 = $false
            foreach ($osNode in $pkg.SelectNodes(".//*[local-name()='OperatingSystem']")) {
                $osDisp = ""
                try { $osDisp = $osNode.SelectSingleNode("*[local-name()='Display']").InnerText } catch {}
                if ($osDisp -match "(?i)windows 11") { $supportsWin11 = $true }
                if ($osDisp -match "(?i)windows 10") { $supportsWin10 = $true }
            }
            $pkgName = ""
            try { $pkgName = $pkg.SelectSingleNode("*[local-name()='Name']/*[local-name()='Display']").InnerText } catch {}
            $pkgPath = $pkg.GetAttribute("path")
            $out += [PSCustomObject]@{ Path = $pkgPath; DisplayName = $pkgName; Win11 = $supportsWin11; Win10 = $supportsWin10 }
            Log "  Candidate: $pkgName  [W11=$supportsWin11 W10=$supportsWin10]"
        }
        ,$out
    }

    $candidates = & $findCandidates $searchNames

    # v1.13.3 - form-factor suffix fallback. Dell catalogs occasionally list a
    # model under its base name only (e.g. "Latitude 5310") with no separate
    # entry for the convertible variant ("Latitude 5310 2-in-1"), even though
    # the convertible exists as a real SKU. If the primary search yielded no
    # candidates AND the model name carries a form-factor suffix, retry with
    # the suffix stripped. The primary search runs first so that when BOTH
    # variants exist in the catalog the convertible-specific pack still wins
    # (different chassis -> different touchscreen/digitizer/sensor drivers).
    if ($candidates.Count -eq 0 -and $ModelName -match '(?i)\s+2[\s-]?in[\s-]?1\b') {
        $stripped = ($ModelName -replace '(?i)\s+2[\s-]?in[\s-]?1\b','').Trim()
        if ($stripped -and $stripped.Length -gt 3 -and $stripped -ne $ModelName) {
            Log "No match for '$ModelName' - retrying without form-factor suffix: '$stripped'"
            $fallbackNames = @($stripped)
            $fallbackNames += ($stripped -replace '\s+Notebook.*$','')
            $fallbackNames += ($stripped -replace '^Dell\s+','')
            $fallbackNames = @($fallbackNames | Select-Object -Unique | Where-Object { $_.Length -gt 3 })
            Log "Fallback search tokens: $($fallbackNames -join ' | ')"
            $candidates = & $findCandidates $fallbackNames
        }
    }

    if ($candidates.Count -eq 0) {
        Log "No driver pack found in DriverPackCatalog for '$ModelName'."
        Log "Opening Dell support page..."
        Start-Process "https://www.dell.com/support/home/en-us/product-support/servicetag/$serviceTag/drivers"
        return $false
    }

    $chosen = $null
    if ($isWin11) {
        $chosen = $candidates | Where-Object { $_.Win11 } | Select-Object -First 1
        if (-not $chosen) { $chosen = $candidates[0] }
    } else {
        $chosen = $candidates | Where-Object { $_.Win10 } | Select-Object -First 1
        if (-not $chosen) { $chosen = $candidates[0] }
    }

    Log "Selected: $($chosen.DisplayName)"
    $packPath = $chosen.Path
    if (-not $packPath) { Log "Driver pack entry has no path - unexpected catalog format."; return $false }

    $packUrl  = "https://downloads.dell.com/$packPath"
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName($packPath))
    Log "Pack file: $([System.IO.Path]::GetFileName($packPath))"
    Log "Pack URL:  $packUrl"
    SetProgress 25

    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    SetProgress 30
    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile)) { Log "Dell driver pack download failed."; return $false }
    if (Test-Cancelled) { return $false }
    SetProgress 55

    $extractPath = Join-Path $DriverRoot "Dell_Extracted"
    Log "Extracting Dell pack..."
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 300 -Vendor "Dell"
    SetProgress 60
    return (Install-DriversFromPath -BasePath $extractPath)
}

# =========================
# HP REFERENCE CATALOG (HPIA backend)
# =========================
# v1.9.0+: per-machine reference XML from HP's HPIA backend, used to download
# only the Softpaqs that match this machine's missing devices.
#
# Pipeline:
#   1. SystemID via Win32_BaseBoard.Product (4-char hex, e.g. "8B41").
#   2. Fetch https://hpia.hpcloud.hp.com/ref/<sysid>/<sysid>_64_<10|11>.0.<ver>.cab
#      with a fallback chain on the OS code (25h2 -> 24h2 -> 23h2 -> ...).
#   3. expand.exe the CAB to a single XML.
#   4. Build O(1) lookup tables: catalog DeviceId -> SP IDs, SP ID -> metadata.
#   5. Enumerate missing devices (ConfigManagerErrorCode != 0), match each
#      device's HardwareID/CompatibleID array against the DeviceId index.
#   6. Filter matched SPs to "Driver - *" categories (skip dock firmware,
#      security software, BIOS, diagnostics - these aren't what a missing-driver
#      run is for, and they're huge).
#   7. Total download budget guardrail: if matched SPs exceed
#      HP_CATALOG_BUDGET_MB, fall back to the full driver pack (it'll usually
#      be smaller). Defaults to 800 MB.
#   8. Download each SP, SHA256-verify, extract via 7-Zip (HP softpaqs are
#      self-extracting), run the SilentInstall command from the catalog.
#   9. Record each successfully-installed SP into AnalyticsInstalledDrivers.
#
# Returns: $true on success, $null to signal "fall back to full driver pack".
# Caller (Start-HpDriverInstall) treats $null as fall-through.

$script:HpCatalogBudgetMB = 800   # tuning knob - over this, prefer full pack
$script:HpHpUpSuccessCodes = @(0, 1, 1641, 3010)  # HPUP/HpFirmwareUpdRec exit-OK set

function Get-HpSystemId {
    # The 4-char hex platform ID is in Win32_BaseBoard.Product per HPIA docs.
    try {
        $sysid = (Get-CimInstance Win32_BaseBoard -EA Stop).Product.Trim().ToLower()
        if ($sysid -match '^[0-9a-f]{4}$') { return $sysid }
        Log "  HP catalog: BaseBoard.Product '$sysid' isn't a 4-char hex SystemID - skipping catalog path."
        return $null
    } catch {
        Log "  HP catalog: cannot read BaseBoard.Product - $($_.Exception.Message)"
        return $null
    }
}

function Get-HpCatalogUrlCandidates {
    param([string]$SysId)
    # Win32_OperatingSystem.Version reports 10.0.x for both Win10 AND Win11.
    # HPIA's URL prefix is split by build (>=22000 = Win11 = "11.0.").
    $build = 0
    $display = $null
    try {
        $build = [int](Get-CimInstance Win32_OperatingSystem -EA Stop).BuildNumber
    } catch {}
    try {
        $display = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -EA Stop).DisplayVersion
    } catch {}
    $prefix = if ($build -ge 22000) { '11.0' } else { '10.0' }

    # Build fallback chain: machine's DisplayVersion first, then progressively
    # older Win11 / Win10 codes. Lowercase to match observed naming.
    $candidates = @()
    if ($display) {
        $dv = $display.ToLower()
        $candidates += "https://hpia.hpcloud.hp.com/ref/$SysId/${SysId}_64_${prefix}.${dv}.cab"
    }
    $win11Versions = @('25h2','24h2','23h2','22h2','21h2')
    $win10Versions = @('22h2','21h2','21h1','20h2','2004','1909','1903','1809','1803')
    $list = if ($build -ge 22000) { $win11Versions } else { $win10Versions }
    foreach ($v in $list) {
        $u = "https://hpia.hpcloud.hp.com/ref/$SysId/${SysId}_64_${prefix}.${v}.cab"
        if ($candidates -notcontains $u) { $candidates += $u }
    }
    return $candidates
}

function Get-HpReferenceCatalogXml {
    param(
        [string]$SysId,
        [string]$DriverRoot
    )
    # Returns parsed [xml] or $null. Probes candidate URLs with cheap HEAD
    # requests (Invoke-CurlDownload retries 10 times, way too slow for
    # probing), then full-downloads only the first one that returns 2xx.
    $candidates = Get-HpCatalogUrlCandidates -SysId $SysId
    if (-not $candidates -or $candidates.Count -eq 0) {
        Log "  HP catalog: no candidate URLs built - skipping."
        return $null
    }

    $cabPath = Join-Path $env:TEMP "hp_ref_${SysId}.cab"
    $xmlDir  = Join-Path $env:TEMP "hp_ref_${SysId}_x"
    Remove-Item $cabPath -EA SilentlyContinue
    if (Test-Path $xmlDir) { Remove-Item $xmlDir -Recurse -Force -EA SilentlyContinue }

    $winnerUrl = $null
    foreach ($u in $candidates) {
        if (Test-Cancelled) { return $null }
        # HEAD probe: --head + --fail makes a 404 exit non-zero immediately;
        # --max-time bounds the worst case per URL to 10 seconds total.
        $code = & curl.exe --silent --head --fail --location `
                           --max-time 10 --connect-timeout 5 `
                           --user-agent "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" `
                           --output NUL `
                           --write-out "%{http_code}" "$u" 2>$null
        $exitOk = ($LASTEXITCODE -eq 0)
        if ($exitOk -and $code -match '^2\d\d$') {
            Log "  HEAD 200: $u"
            $winnerUrl = $u
            break
        } else {
            Log "  HEAD ${code}: $u"
        }
    }
    if (-not $winnerUrl) {
        Log "  HP catalog: no reference CAB found for SystemID $SysId at any tested OS version."
        return $null
    }

    Log "  Downloading reference CAB..."
    if (-not (Invoke-CurlDownload -Url $winnerUrl -OutFile $cabPath)) {
        Log "  HP catalog: HEAD succeeded but full download failed - skipping."
        return $null
    }
    if (Test-Cancelled) { return $null }

    # Validate the download before trying to expand it. Some HP CDN edges return
    # a small HTML/text error body with a 200 status code (e.g. geo-blocking
    # surrogates). Quick sanity: a real reference CAB starts with "MSCF" magic.
    $cabSize = (Get-Item $cabPath).Length
    Log "    Downloaded $([math]::Round($cabSize/1KB,1)) KB"
    if ($cabSize -lt 1024) {
        $preview = ''
        try { $preview = (Get-Content $cabPath -TotalCount 1 -EA SilentlyContinue) -as [string] } catch {}
        Log "  HP catalog: download too small to be a real CAB (got $cabSize bytes). Preview: '$preview'"
        Log "  Skipping catalog path."
        return $null
    }
    $magic = $null
    try {
        $fs = [System.IO.File]::OpenRead($cabPath)
        $buf = New-Object byte[] 4
        [void]$fs.Read($buf, 0, 4)
        $fs.Close()
        $magic = -join ($buf | ForEach-Object { [char]$_ })
    } catch {}
    if ($magic -ne 'MSCF') {
        Log "  HP catalog: downloaded file isn't a CAB (magic='$magic', expected 'MSCF'). Skipping catalog path."
        return $null
    }

    New-Item -ItemType Directory -Path $xmlDir -Force | Out-Null
    # expand.exe is notoriously fussy about argument passing. Different
    # combinations work depending on PowerShell quoting behavior. Try the
    # cleanest form first (Start-Process with split args, no -F filter),
    # check for output, then fall back to alternative forms if needed.
    function _TryExpand {
        param([string[]]$Args)
        $expandLog = Join-Path $env:TEMP "hp_expand_${SysId}.log"
        Remove-Item $expandLog -EA SilentlyContinue
        try {
            $proc = Start-Process -FilePath "expand.exe" -ArgumentList $Args `
                -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $expandLog -EA Stop
            if (Test-Path $expandLog) {
                foreach ($l in (Get-Content $expandLog)) { Log "    expand: $l" }
                Remove-Item $expandLog -EA SilentlyContinue
            }
            return $proc.ExitCode
        } catch {
            Log "    expand error: $($_.Exception.Message)"
            return -1
        }
    }

    # Attempt 1: bare "expand <src> <dst-dir>" - the canonical form per MS docs.
    Log "    Extracting via: expand <src> <dst-dir>"
    $null = _TryExpand @($cabPath, $xmlDir)
    $xmlFile = Get-ChildItem $xmlDir -Filter "*.xml" -Recurse -EA SilentlyContinue | Select-Object -First 1

    # Attempt 2: with -F:* filter (some Windows builds need it for multi-file CABs).
    if (-not $xmlFile) {
        Log "    Retrying with: expand -F:* <src> <dst-dir>"
        Get-ChildItem $xmlDir -EA SilentlyContinue | Remove-Item -Recurse -Force -EA SilentlyContinue
        $null = _TryExpand @("-F:*", $cabPath, $xmlDir)
        $xmlFile = Get-ChildItem $xmlDir -Filter "*.xml" -Recurse -EA SilentlyContinue | Select-Object -First 1
    }

    # Attempt 3: Shell.Application COM (works on any Windows with Explorer).
    # CABs are valid Shell namespaces; CopyHere extracts contents to a folder.
    if (-not $xmlFile) {
        Log "    Retrying via Shell.Application COM..."
        Get-ChildItem $xmlDir -EA SilentlyContinue | Remove-Item -Recurse -Force -EA SilentlyContinue
        try {
            $shell = New-Object -ComObject Shell.Application
            $srcFolder = $shell.NameSpace($cabPath)
            $dstFolder = $shell.NameSpace($xmlDir)
            if ($srcFolder -and $dstFolder) {
                # 0x14 = no progress UI + auto-overwrite
                $dstFolder.CopyHere($srcFolder.Items(), 0x14)
                # CopyHere is async even when sync'd; wait briefly for files.
                $wait = 0
                while (-not (Get-ChildItem $xmlDir -EA SilentlyContinue) -and $wait -lt 30) {
                    Start-Sleep -Milliseconds 250
                    $wait++
                }
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
        } catch {
            Log "    Shell.Application error: $($_.Exception.Message)"
        }
        $xmlFile = Get-ChildItem $xmlDir -Filter "*.xml" -Recurse -EA SilentlyContinue | Select-Object -First 1
    }

    if (-not $xmlFile) {
        # All three methods failed - log directory contents for diagnosis.
        $contents = @(Get-ChildItem $xmlDir -Recurse -EA SilentlyContinue)
        Log "  HP catalog: extraction failed by all methods. Dest dir contains $($contents.Count) item(s):"
        foreach ($f in $contents | Select-Object -First 10) {
            Log "    - $($f.FullName)  ($([math]::Round($f.Length/1KB,1)) KB)"
        }
        Log "  Skipping catalog path."
        return $null
    }
    Log "  Parsing reference XML: $($xmlFile.Name) ($([math]::Round($xmlFile.Length/1MB,1)) MB)"
    try {
        [xml]$catalog = Get-Content $xmlFile.FullName -Raw -Encoding UTF8
        return $catalog
    } catch {
        Log "  HP catalog: XML parse failed - $($_.Exception.Message)"
        return $null
    }
}

function Get-HpMissingDevicesWithHwIds {
    # Same enum approach as Start-DellIndividualDriverInstall, but we keep the
    # full HardwareID + CompatibleID arrays since HP matches by full DeviceId
    # string rather than just VEN/DEV.
    $list = New-Object System.Collections.Generic.List[object]
    try {
        $entities = @(Get-CimInstance Win32_PnPEntity -EA Stop |
            Where-Object { $_.ConfigManagerErrorCode -ne 0 })
    } catch {
        Log "  HP catalog: cannot enumerate Win32_PnPEntity - $($_.Exception.Message)"
        return $list
    }
    foreach ($e in $entities) {
        $hwIds   = @()
        $compIds = @()
        try {
            $pnp = Get-PnpDevice -InstanceId $e.DeviceID -EA Stop
            if ($pnp.HardwareID)   { $hwIds   = @($pnp.HardwareID) }
            if ($pnp.CompatibleID) { $compIds = @($pnp.CompatibleID) }
        } catch {
            # Get-PnpDevice can fail for some phantom entries; fall back to
            # registry like the Dell path does for max compatibility.
            try {
                $reg = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Enum\$($e.DeviceID)" -EA Stop
                if ($reg.HardwareID)   { $hwIds   = @($reg.HardwareID) }
                if ($reg.CompatibleIDs) { $compIds = @($reg.CompatibleIDs) }
            } catch {}
        }
        $list.Add([pscustomobject]@{
            Name         = if ($e.Name)    { $e.Name }    elseif ($e.Caption) { $e.Caption } else { '(unnamed)' }
            DeviceID     = $e.DeviceID
            HardwareIDs  = $hwIds
            CompatibleIDs = $compIds
        })
    }
    return $list
}

function Build-HpDeviceIndex {
    param([xml]$Catalog)
    # Map: upper(DeviceId) -> list of SP IDs.
    # [xml] auto-decodes &amp; to & so we get the raw form Windows uses.
    $index = @{}
    foreach ($dev in $Catalog.SelectNodes('/ImagePal/Devices/Device')) {
        $did = $dev.SelectSingleNode('DeviceId')
        if (-not $did -or [string]::IsNullOrWhiteSpace($did.InnerText)) { continue }
        $key = $did.InnerText.Trim().ToUpper()
        $sps = New-Object System.Collections.Generic.List[string]
        foreach ($u in $dev.SelectNodes('Solutions/UpdateInfo')) {
            $idref = $u.GetAttribute('IdRef')
            if ($idref) { $sps.Add($idref) }
        }
        if ($sps.Count -gt 0) {
            if ($index.ContainsKey($key)) {
                foreach ($s in $sps) { if ($index[$key] -notcontains $s) { $index[$key] += $s } }
            } else {
                $index[$key] = @($sps)
            }
        }
    }
    return $index
}

function Build-HpSolutionIndex {
    param([xml]$Catalog)
    # Map: SP ID -> properties hashtable. Defensive on null fields - if the
    # catalog schema drifts or an entry is malformed, skip that one entry
    # rather than crashing the whole index build.
    $index = @{}
    $skipped = 0
    foreach ($sol in $Catalog.SelectNodes('/ImagePal/Solutions/UpdateInfo')) {
        try {
            $id = $sol.SelectSingleNode('Id')
            if (-not $id -or [string]::IsNullOrWhiteSpace($id.InnerText)) { $skipped++; continue }
            $sp = $id.InnerText.Trim()

            $urlNode  = $sol.SelectSingleNode('Url')
            $shaNode  = $sol.SelectSingleNode('SHA256')
            $sizeNode = $sol.SelectSingleNode('Size')
            $instNode = $sol.SelectSingleNode('SilentInstall')
            $nameNode = $sol.SelectSingleNode('Name')
            $catNode  = $sol.SelectSingleNode('Category')
            $verNode  = $sol.SelectSingleNode('Version')

            $url = if ($urlNode) { $urlNode.InnerText } else { '' }
            # Catalog URLs are schemeless: prepend https:// (HP serves both HTTP and HTTPS).
            if ($url -and -not ($url -match '^https?://')) { $url = "https://$url" }

            $size = 0
            if ($sizeNode -and $sizeNode.InnerText -match '^\d+$') {
                $size = [int64]$sizeNode.InnerText
            }

            $index[$sp] = @{
                Name          = if ($nameNode) { $nameNode.InnerText } else { '(unnamed)' }
                Category      = if ($catNode)  { $catNode.InnerText  } else { '' }
                Version       = if ($verNode)  { $verNode.InnerText  } else { '' }
                Url           = $url
                SHA256        = if ($shaNode)  { $shaNode.InnerText  } else { '' }
                Size          = $size
                SilentInstall = if ($instNode) { $instNode.InnerText } else { '' }
            }
        } catch {
            $skipped++
        }
    }
    if ($skipped -gt 0) {
        Log "  HP catalog: skipped $skipped malformed Solutions entries during indexing."
    }
    return $index
}

function Find-HpApplicableSoftpaqs {
    param(
        [object[]]$MissingDevices,
        [hashtable]$DeviceIndex,
        [hashtable]$SolutionIndex
    )
    # For each missing device, walk its HardwareID + CompatibleID arrays in order
    # (most-specific first - that's how Windows presents them) and look up each
    # in the catalog's DeviceId index. First match per device wins.
    #
    # Returns deduplicated list of SP entries, only Driver categories.
    $matched = [ordered]@{}   # SP ID -> SP entry
    $unmatchedDevices = New-Object System.Collections.Generic.List[string]

    foreach ($dev in $MissingDevices) {
        $deviceHit = $false
        $allIds = @()
        if ($dev.HardwareIDs)   { $allIds += @($dev.HardwareIDs) }
        if ($dev.CompatibleIDs) { $allIds += @($dev.CompatibleIDs) }
        # Also try the bare DeviceID (e.g. "ACPI\HPIC000C" type entries match this directly).
        if ($dev.DeviceID) { $allIds += $dev.DeviceID }

        foreach ($rawId in $allIds) {
            if ([string]::IsNullOrWhiteSpace($rawId)) { continue }
            $candidates = @($rawId.Trim().ToUpper())
            # Also try stripping the &REV_XX suffix - catalog often omits it.
            $stripped = $candidates[0] -replace '&REV_[0-9A-F]+$', ''
            if ($stripped -ne $candidates[0]) { $candidates += $stripped }

            foreach ($key in $candidates) {
                if ($DeviceIndex.ContainsKey($key)) {
                    foreach ($spId in $DeviceIndex[$key]) {
                        if (-not $matched.Contains($spId) -and $SolutionIndex.ContainsKey($spId)) {
                            $entry = $SolutionIndex[$spId]
                            # Filter: only "Driver - *" categories. Skip docks, BIOS,
                            # security software, diagnostics, etc. They're either
                            # huge or not what a missing-driver run is for.
                            if ($entry.Category -like 'Driver - *') {
                                $matched[$spId] = $entry + @{ Id = $spId; MatchedDevice = $dev.Name }
                            }
                        }
                    }
                    Log "    Matched device '$($dev.Name)' via '$key'"
                    $deviceHit = $true
                    break
                }
            }
            if ($deviceHit) { break }
        }
        if (-not $deviceHit) {
            $unmatchedDevices.Add("$($dev.Name) [$($dev.DeviceID)]")
        }
    }
    return @{
        Softpaqs   = @($matched.Values)
        Unmatched  = $unmatchedDevices
    }
}

function Get-HpSoftpaqOutFile {
    # Single source of truth for the local SP filename so the parallel download
    # planner and the post-download installer agree on the path.
    param([hashtable]$Sp, [string]$DriverRoot)
    return (Join-Path $DriverRoot "$($Sp.Id).exe")
}

function Install-HpSoftpaqPostDownload {
    # v1.11.0 - the post-download (serial) half of the old Install-HpSoftpaq.
    # Assumes $spFile already exists on disk from a prior (serial or parallel)
    # download. Does: SHA256 verify -> extract -> SilentInstall.
    # Returns $true on success, $false on any failure.
    param(
        [hashtable]$Sp,
        [string]$DriverRoot,
        [int]$Index,
        [int]$Total
    )
    Log ""
    Log "----- HP CATALOG [$Index/$Total] $($Sp.Name) (v$($Sp.Version)) -----"
    Log "  SP ID:   $($Sp.Id)"
    Log "  Category:$($Sp.Category)"
    Log "  Size:    $([math]::Round($Sp.Size/1MB,1)) MB"
    Log "  Install: $($Sp.SilentInstall)"

    if (-not $Sp.SilentInstall -or $Sp.SilentInstall -eq 'NA') {
        Log "  No SilentInstall command - skipping (would need user interaction)."
        return $false
    }
    $spFile = Get-HpSoftpaqOutFile -Sp $Sp -DriverRoot $DriverRoot
    if (-not (Test-Path $spFile)) {
        Log "  Expected file missing on disk ($spFile) - skipping."
        return $false
    }

    Log "  Verifying SHA256..."
    $actual = (Get-FileHash -Path $spFile -Algorithm SHA256).Hash.ToUpper()
    $expect = ($Sp.SHA256).ToUpper()
    if ($actual -ne $expect) {
        Log "  SHA256 mismatch: got $actual  expected $expect - skipping."
        return $false
    }
    Log "  SHA256 OK."

    # Extract via 7-Zip pass-1 (works on HP self-extracting softpaqs).
    $extractDir = Join-Path $DriverRoot "$($Sp.Id)_x"
    if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force -EA SilentlyContinue }
    Log "  Extracting to $extractDir..."
    Start-PackExtraction -PackFile $spFile -DestPath $extractDir -StallLimitSec 180 -Vendor "HP"
    if (-not (Test-Path $extractDir) -or @(Get-ChildItem $extractDir -EA SilentlyContinue).Count -eq 0) {
        Log "  Extraction produced no files - skipping."
        return $false
    }

    # Resolve SilentInstall command. Most HP softpaqs use HPUP.exe at the
    # extraction root; some put the installer in a subdir. If the exact command
    # binary isn't at the root, search recursively for it.
    $cmdBinary = $Sp.SilentInstall.Trim()
    if ($cmdBinary -match '^"([^"]+)"') { $cmdBinary = $matches[1] }
    elseif ($cmdBinary -match '^(\S+)')  { $cmdBinary = $matches[1] }
    $cmdBinName = [System.IO.Path]::GetFileName($cmdBinary)
    $cmdRoot    = $extractDir
    if (-not (Test-Path (Join-Path $cmdRoot $cmdBinName))) {
        $found = Get-ChildItem $extractDir -Recurse -Filter $cmdBinName -EA SilentlyContinue | Select-Object -First 1
        if ($found) {
            $cmdRoot = $found.Directory.FullName
            Log "  Note: $cmdBinName not at root, running from subdir: $cmdRoot"
        } else {
            Log "  WARNING: $cmdBinName not found in extraction - install will likely fail."
        }
    }

    Log "  Running: $($Sp.SilentInstall)"
    $exit = Invoke-LenovoPackageCommand -Command $Sp.SilentInstall -WorkingDir $cmdRoot -TimeoutSec 900
    Log "  Exit: $exit"
    if ($script:HpHpUpSuccessCodes -contains $exit) {
        Log "  Install OK."
        $null = $script:AnalyticsInstalledDrivers.Add("$($Sp.Name) (v$($Sp.Version))")
        if ($exit -eq 3010 -or $exit -eq 1641) { Log "  (reboot required to finalize)" }
        return $true
    }
    Log "  Install reported failure (exit $exit)."
    return $false
}

function Install-HpSoftpaq {
    # v1.11.0 - kept as a back-compat wrapper for any callers that still want
    # the old "do everything for one SP, serially" behaviour. Internally just
    # runs a 1-item parallel download then the post-download phase.
    param(
        [hashtable]$Sp,
        [string]$DriverRoot,
        [int]$Index,
        [int]$Total
    )
    if (-not $Sp.Url) { Log "  ERROR: no URL in catalog - skipping."; return $false }
    $spFile = Get-HpSoftpaqOutFile -Sp $Sp -DriverRoot $DriverRoot
    if (-not (Invoke-CurlDownload -Url $Sp.Url -OutFile $spFile)) {
        Log "  Download failed."
        return $false
    }
    if (Test-Cancelled) { return $false }
    return (Install-HpSoftpaqPostDownload -Sp $Sp -DriverRoot $DriverRoot -Index $Index -Total $Total)
}

function Start-HpReferenceCatalogInstall {
    param([string]$DriverRoot)
    # Returns:
    #   $true  - catalog path handled the install (full or partial)
    #   $null  - couldn't use catalog path; caller should fall back to full pack
    # Never returns $false - any in-catalog failure still counts as "handled" if
    # at least one SP installed. If zero installed, we return $null to give the
    # full-pack fallback a chance.
    #
    # Whole flow is wrapped in try/catch: anything unexpected (schema drift,
    # network hiccup mid-parse, etc.) returns $null so the caller can fall
    # through to the legacy scraper rather than dying.
    Log "--- HP reference catalog (HPIA backend) ---"
    try {
        $sysid = Get-HpSystemId
        if (-not $sysid) { return $null }
        Log "SystemID: $sysid"

        $catalog = Get-HpReferenceCatalogXml -SysId $sysid -DriverRoot $DriverRoot
        if (-not $catalog) { return $null }

        Log "Indexing catalog..."
        $devIdx = Build-HpDeviceIndex   -Catalog $catalog
        $solIdx = Build-HpSolutionIndex -Catalog $catalog
        Log "  Devices in catalog:   $($devIdx.Count)"
        Log "  Solutions in catalog: $($solIdx.Count)"

        Log "Enumerating missing devices on this machine..."
        $missing = Get-HpMissingDevicesWithHwIds
        Log "  Missing devices: $($missing.Count)"
        if ($missing.Count -eq 0) {
            Log "  No missing devices - nothing for the catalog path to do."
            return $null    # let full-pack path handle "no missing drivers" prompt
        }
        foreach ($m in $missing) {
            Log "  - $($m.Name) [$($m.DeviceID)]"
        }

        Log "Matching missing devices against catalog..."
        $result   = Find-HpApplicableSoftpaqs -MissingDevices $missing -DeviceIndex $devIdx -SolutionIndex $solIdx
        $softpaqs = @($result.Softpaqs)
        if ($softpaqs.Count -eq 0) {
            Log "  No catalog matches for any missing device - falling back to full pack."
            return $null
        }

        $totalMB = [math]::Round((($softpaqs | Measure-Object -Property Size -Sum).Sum) / 1MB, 1)
        Log "Matched $($softpaqs.Count) driver Softpaq(s), total $totalMB MB:"
        foreach ($sp in $softpaqs) {
            Log "  $($sp.Id)  $([math]::Round($sp.Size/1MB,1).ToString().PadLeft(7)) MB  $($sp.Category) - $($sp.Name)"
        }
        if ($result.Unmatched.Count -gt 0) {
            Log "Unmatched missing devices ($($result.Unmatched.Count)): catalog has no driver for these"
            foreach ($u in $result.Unmatched) { Log "  - $u" }
        }
        if ($totalMB -gt $script:HpCatalogBudgetMB) {
            Log "Catalog total ($totalMB MB) exceeds budget ($($script:HpCatalogBudgetMB) MB) - falling back to full pack."
            return $null
        }

        if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }

        # v1.11.0 - PARALLEL DOWNLOAD PHASE
        # Build a download manifest (skipping anything that's missing a URL or
        # SilentInstall - those would be no-ops in the install phase anyway),
        # fetch them all concurrently, then iterate the survivors serially for
        # SHA verify + extract + install. The serial install phase is unchanged.
        $dlItems = New-Object 'System.Collections.Generic.List[hashtable]'
        $dlMap   = @{}   # spId -> sp object, for re-lookup after the download batch
        foreach ($sp in $softpaqs) {
            if (-not $sp.Url) {
                Log "  SKIP: $($sp.Name) has no URL in catalog."
                continue
            }
            if (-not $sp.SilentInstall -or $sp.SilentInstall -eq 'NA') {
                # Matches the old serial Install-HpSoftpaq's pre-download check -
                # don't waste bandwidth on a SP we'd just reject in the install phase.
                Log "  SKIP: $($sp.Name) has no SilentInstall command (would need user interaction)."
                continue
            }
            $outFile = Get-HpSoftpaqOutFile -Sp $sp -DriverRoot $DriverRoot
            $dlItems.Add(@{
                Url     = $sp.Url
                OutFile = $outFile
                Label   = "$($sp.Id) - $($sp.Name)"
            }) | Out-Null
            $dlMap[$sp.Id] = $sp
        }
        if ($dlItems.Count -eq 0) {
            Log "All matched softpaqs were filtered out before download - falling back to full pack."
            return $null
        }
        SetProgress 32
        $dlResults = Invoke-CurlDownloadParallel -Items $dlItems
        if (Test-Cancelled) { return $false }

        # Reduce results to the SPs whose download actually succeeded, preserving
        # the original catalog order so per-INF logs match the manifest.
        $okSet = @{}
        foreach ($r in $dlResults) {
            if ($r.Success) {
                $name = [System.IO.Path]::GetFileNameWithoutExtension($r.Item.OutFile)
                $okSet[$name] = $true
            }
        }
        $readySps = @($softpaqs | Where-Object { $okSet[$_.Id] -eq $true })
        if ($readySps.Count -eq 0) {
            Log "Parallel download phase produced no usable files - falling back to full pack."
            return $null
        }
        Log "Parallel download phase: $($readySps.Count)/$($softpaqs.Count) softpaq(s) ready for install."
        SetProgress 50

        # SERIAL INSTALL PHASE (unchanged semantics - pnputil/HPUP need exclusive access)
        $okCount = 0
        $i = 0
        foreach ($sp in $readySps) {
            $i++
            if (Test-Cancelled) { return $false }
            SetProgress (50 + [int](($i / $readySps.Count) * 45))
            SetExtract -Pct ([int](($i / $readySps.Count) * 100)) -Label "Catalog SP [$i/$($readySps.Count)]: $($sp.Name)"
            if (Install-HpSoftpaqPostDownload -Sp $sp -DriverRoot $DriverRoot -Index $i -Total $readySps.Count) {
                $okCount++
            }
        }

        Log ""
        Log "=== HP CATALOG SUMMARY ==="
        Log "Downloaded: $($readySps.Count) / $($softpaqs.Count)  (parallel, concurrency=$($script:MaxParallelDownloads))"
        Log "Installed : $okCount / $($readySps.Count)"

        if ($okCount -eq 0) {
            Log "No softpaqs installed via catalog - falling back to full pack."
            return $null
        }
        return $true
    } catch {
        Log "  HP catalog: unexpected error - $($_.Exception.Message)"
        Log "  Falling back to full pack."
        return $null
    }
}

# =========================
# HP
# =========================
function Start-HpDriverInstall {
    param([string]$DriverRoot, [string]$ModelName)

    Log "=== HP: Starting automated driver install ==="
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."

    try { $script:AnalyticsSerial = (Get-CimInstance Win32_BIOS).SerialNumber.Trim() } catch {}

    # v1.9.0: try the HPIA reference catalog first.
    # Returns $true if it handled the install, $null to fall through to the
    # full-pack matrix scraper below. $false reserved for explicit cancellation.
    $catalogResult = Start-HpReferenceCatalogInstall -DriverRoot $DriverRoot
    if ($catalogResult -eq $true)  { return $true }
    if ($catalogResult -eq $false) { return $false }
    Log "Falling back to full-pack driver matrix..."

    $osBuild = $null; $isWin11 = $false
    try {
        $osBuild = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber
        $isWin11 = $osBuild -ge 22000
        Log "OS Build: $osBuild  ($(if ($isWin11) {'Win11'} else {'Win10'}))"
    } catch { Log "Could not read OS build: $($_.Exception.Message)" }

    $matrixUrl  = "https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html"
    $matrixFile = Join-Path $env:TEMP "HP_DPMatrix.html"
    Remove-Item $matrixFile -EA SilentlyContinue

    Log "Downloading HP Driver Pack Matrix..."
    SetExtract -Pct 5 -Label "Downloading matrix page..."
    if (-not (Invoke-CurlDownload -Url $matrixUrl -OutFile $matrixFile)) {
        Log "Failed to download HP Driver Pack Matrix."
        Start-Process $matrixUrl
        return $false
    }
    if (Test-Cancelled) { return $false }

    $matrixHtml = [System.IO.File]::ReadAllText($matrixFile)
    Log "  Matrix HTML: $([math]::Round($matrixHtml.Length/1KB)) KB"
    SetExtract -Pct 20 -Label "Parsing matrix..."
    SetProgress 20

    $searchTokens = @()
    if ($ModelName) {
        $stripped = $ModelName -replace '(?i)^HP\s+', ''
        $searchTokens += $stripped
        $searchTokens += ($stripped -replace '\s+Notebook.*$', '')
        $searchTokens += ($stripped -replace '\s+PC.*$',       '')
        $searchTokens = $searchTokens | Select-Object -Unique | Where-Object { $_.Length -gt 4 }
    }
    Log "Search tokens: $($searchTokens -join ' | ')"

    $packUrl = $null; $packSpNum = $null
    $flat    = $matrixHtml -replace "`r`n|`r|`n", " " -replace "\s{2,}", " "
    $rows    = [regex]::Matches($flat, '(?i)<tr[^>]*>(.*?)</tr>')

    foreach ($row in $rows) {
        $rowHtml = $row.Groups[1].Value
        $cells   = [regex]::Matches($rowHtml, '(?i)<t[dh][^>]*>(.*?)</t[dh]>')
        if ($cells.Count -lt 2) { continue }
        $modelCell = [regex]::Replace($cells[0].Groups[1].Value, '<[^>]+>', ' ')
        $modelCell = [System.Net.WebUtility]::HtmlDecode($modelCell) -replace '\s+', ' '
        $matched = $false
        foreach ($tok in $searchTokens) {
            if ($modelCell -match [regex]::Escape($tok)) { $matched = $true; break }
        }
        if (-not $matched) { continue }
        Log "  Matched matrix row: $($modelCell.Trim() -replace '\s+',' ')"
        $allLinks = [regex]::Matches($rowHtml, '(?i)href="([^"]*sp\d+\.exe)"')
        if ($allLinks.Count -eq 0) { Log "  Row matched but contains no .exe links - skipping."; continue }
        $bestUrl = $allLinks[0].Groups[1].Value
        if (-not $bestUrl.StartsWith("http")) { $bestUrl = "https://ftp.hp.com$bestUrl" }
        if (-not $isWin11) {
            foreach ($lm in $allLinks) {
                $href = $lm.Groups[1].Value
                if (-not $href.StartsWith("http")) { $href = "https://ftp.hp.com$href" }
                $aTag  = [regex]::Match($rowHtml, "(?i)<a[^>]+href=""[^""]*$([regex]::Escape([System.IO.Path]::GetFileName($href)))[^""]*""[^>]*>")
                $title = if ($aTag.Success) { $aTag.Value } else { "" }
                if ($title -eq "" -or $title -match "(?i)windows 10") { $bestUrl = $href; break }
            }
        }
        $packUrl   = $bestUrl
        $packSpNum = [regex]::Match($packUrl, '(?i)(sp\d+)\.exe').Groups[1].Value
        Log "  Selected SoftPaq: $packSpNum"
        Log "  URL: $packUrl"
        break
    }

    if (Test-Cancelled) { return $false }
    if (-not $packUrl) {
        Log "Model '$ModelName' not found in HP Driver Pack Matrix."
        Log "Opening HP Driver Pack Matrix for manual selection..."
        Start-Process $matrixUrl
        return $false
    }

    SetExtract  -Pct 40 -Label "Matrix OK - $packSpNum"
    SetProgress 25
    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    $packFile = Join-Path $DriverRoot "$packSpNum.exe"
    SetProgress 30
    if (-not (Invoke-CurlDownload -Url $packUrl -OutFile $packFile)) { Log "HP driver pack download failed."; return $false }
    if (Test-Cancelled) { return $false }
    SetProgress 55

    $extractPath = Join-Path $DriverRoot "HP_Extracted"
    Log "Extracting HP SoftPaq..."
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 300 -Vendor "HP"
    SetProgress 60
    return (Install-DriversFromPath -BasePath $extractPath)
}

# =========================
# LENOVO
# =========================
# Two-path Lenovo flow (priority reversed in v1.11.0):
#
#   Path 1 (PRIMARY, v1.11.0+): Recipe Card / catalogv2.xml SCCM driver pack
#     Lenovo's "Recipe Card" web tool at
#     https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html lists the same
#     packs as catalogv2.xml. One monolithic .exe -> Inno Setup extract ->
#     pnputil INFs. Simpler and more reliable than per-package descriptors;
#     covers most commercial ThinkPad / ThinkCentre models.
#
#   Path 2 (FALLBACK): Consumer catalog (download.lenovo.com/catalog/<MTM>_<OS>.xml)
#     Same catalog Lenovo System Update uses. Used only when path 1 has no
#     entry - i.e. consumer / ThinkBook / IdeaPad models that don't appear
#     in catalogv2.xml. The v1.8.x consumer-catalog stack (DetectInstall
#     evaluator, NVIDIA short-circuit, BIOS deferral, force-bind unbound) all
#     stay in place and still run when this fallback fires.
#
# Pre-v1.11.0 the order was reversed: v1.8.0 made consumer catalog primary
# because catalogv2.xml lacks ThinkBook/IdeaPad coverage. The trade-off bit
# back on ThinkPads where the consumer-catalog path had many more failure
# modes than a single full-pack install. v1.11.0 restores the
# full-pack-when-available behaviour while keeping consumer catalog for the
# models that genuinely need it.
#
# Helper Invoke-LenovoPackageCommand runs each consumer-catalog package's
# <ExtractCommand> and <Install><Cmdline> via cmd.exe with the placeholder
# substitutions Lenovo's installer authors expect (%PACKAGEPATH% / %WINDOWS%).
# =========================

# =========================
# LENOVO DETECTION HELPERS (used by consumer catalog flow)
#
# Test-LenovoDetectInstall walks a <DetectInstall> element and returns:
#   $true  -> package is already installed -> caller should SKIP
#   $false -> package is not installed     -> caller should INSTALL
#   $null  -> indeterminate (unknown element / error) -> caller should INSTALL
#             (fail-open: a redundant install is much cheaper than a missing one)
#
# Supported elements: _Bios, _File, _FileVersion, _Registry, _RegistryKey,
# and the And / Or / Not combinators. Multiple direct children of <DetectInstall>
# are AND'd together (Lenovo convention).
# =========================

function Test-LenovoBiosLevel {
    param([System.Xml.XmlElement]$Node)
    try {
        $currentBios = [string](Get-CimInstance Win32_BIOS -EA Stop).SMBIOSBIOSVersion
        if (-not $currentBios) { return $null }
        foreach ($lvl in @($Node.Level)) {
            $pat = ([string]$lvl).Trim()
            if (-not $pat) { continue }
            if ($currentBios -like $pat) { return $true }
        }
        return $false
    } catch { return $null }
}

function Expand-LenovoPath {
    param([string]$Path)
    $Path = $Path.Replace('%WINDOWS%',          $env:WINDIR)
    $Path = $Path.Replace('%PROGRAMFILES%',     ${env:ProgramFiles})
    $Path = $Path.Replace('%PROGRAMFILES(X86)%', ${env:ProgramFiles(x86)})
    $Path = $Path.Replace('%SYSTEMROOT%',       $env:SystemRoot)
    $Path = $Path.Replace('%SYSTEMDRIVE%',      $env:SystemDrive)
    return [System.Environment]::ExpandEnvironmentVariables($Path)
}

function Compare-LenovoVersionString {
    # Returns: $true if currentVer satisfies expectedVer (with optional ^ suffix = ">=")
    param([string]$Current, [string]$Expected)
    if (-not $Current -or -not $Expected) { return $null }
    $needsGE = $Expected.EndsWith('^')
    $exp     = $Expected.TrimEnd('^').Trim()
    try {
        $cv = [version]$Current
        $ev = [version]$exp
        if ($needsGE) { return ($cv -ge $ev) } else { return ($cv -eq $ev) }
    } catch {
        # Fall back to string compare for non-dotted values
        if ($needsGE) { return ($Current -ge $exp) } else { return ($Current -eq $exp) }
    }
}

function Test-LenovoFileDetect {
    param([System.Xml.XmlElement]$Node)
    try {
        # _File may carry the path in <Name>, <File>, or directly as inner text
        $name = ""
        if ($Node.Name)      { $name = [string]$Node.Name }
        elseif ($Node.File)  { $name = [string]$Node.File }
        elseif ($Node.InnerText) { $name = [string]$Node.InnerText.Trim() }
        if (-not $name) { return $null }

        $path = Expand-LenovoPath -Path $name
        if (-not (Test-Path -LiteralPath $path)) { return $false }

        $verNode = $Node.SelectSingleNode('Version')
        if (-not $verNode) { return $true }   # file exists, no version constraint = match

        $expected = [string]$verNode.InnerText.Trim()
        $actual = ""
        try { $actual = (Get-Item -LiteralPath $path).VersionInfo.FileVersion } catch {}
        if (-not $actual) { return $null }
        return (Compare-LenovoVersionString -Current $actual -Expected $expected)
    } catch { return $null }
}

function Test-LenovoFileVersionDetect {
    param([System.Xml.XmlElement]$Node)
    try {
        # Two shapes seen in the wild:
        #   <_FileVersion><FileVersion><Name>X</Name><Version>Y</Version></FileVersion></_FileVersion>
        #   <_FileVersion><Name>X</Name><Version>Y</Version></_FileVersion>
        $fvNode = $Node.SelectSingleNode('FileVersion')
        $src    = if ($fvNode) { $fvNode } else { $Node }
        $name = [string]$src.Name
        $ver  = [string]$src.Version
        if (-not $name -or -not $ver) { return $null }

        $path = Expand-LenovoPath -Path $name
        if (-not (Test-Path -LiteralPath $path)) { return $false }

        $actual = ""
        try { $actual = (Get-Item -LiteralPath $path).VersionInfo.FileVersion } catch {}
        if (-not $actual) { return $null }
        return (Compare-LenovoVersionString -Current $actual -Expected $ver)
    } catch { return $null }
}

function ConvertTo-LenovoRegPath {
    param([string]$Raw)
    if (-not $Raw) { return $null }
    $r = $Raw.Trim()
    if ($r -match '^HKEY_LOCAL_MACHINE\\') { return ($r -replace '^HKEY_LOCAL_MACHINE\\', 'HKLM:\') }
    if ($r -match '^HKEY_CURRENT_USER\\')  { return ($r -replace '^HKEY_CURRENT_USER\\',  'HKCU:\') }
    if ($r -match '^HKEY_CLASSES_ROOT\\')  { return ($r -replace '^HKEY_CLASSES_ROOT\\',  'HKCR:\') }
    if ($r -match '^HKLM\\')               { return ($r -replace '^HKLM\\', 'HKLM:\') }
    if ($r -match '^HKCU\\')               { return ($r -replace '^HKCU\\', 'HKCU:\') }
    return "HKLM:\$r"   # Lenovo descriptors commonly omit the hive
}

function Test-LenovoRegistryDetect {
    param([System.Xml.XmlElement]$Node)
    try {
        $keyPath = ConvertTo-LenovoRegPath -Raw ([string]$Node.Key)
        $valName = [string]$Node.KeyName
        $ver     = [string]$Node.Version
        if (-not $keyPath) { return $null }
        if (-not (Test-Path -LiteralPath $keyPath)) { return $false }
        if (-not $valName) { return $true }

        $actual = $null
        try { $actual = (Get-ItemProperty -LiteralPath $keyPath -Name $valName -EA Stop).$valName } catch { return $false }
        if ($null -eq $actual) { return $false }
        if (-not $ver) { return $true }
        return (Compare-LenovoVersionString -Current ([string]$actual) -Expected $ver)
    } catch { return $null }
}

function Test-LenovoRegistryKeyDetect {
    param([System.Xml.XmlElement]$Node)
    try {
        $keyPath = ConvertTo-LenovoRegPath -Raw ([string]$Node.InnerText)
        if (-not $keyPath) { return $null }
        return (Test-Path -LiteralPath $keyPath)
    } catch { return $null }
}

function Get-LenovoHasNvidiaGpu {
    # True if any video controller looks like an NVIDIA GPU. Fail-open: if WMI
    # query errors, return $true so we don't accidentally skip a needed driver.
    try {
        $gpus = Get-CimInstance Win32_VideoController -EA Stop
        foreach ($g in $gpus) {
            if ($g.PNPDeviceID          -like '*VEN_10DE*') { return $true }
            if ($g.AdapterCompatibility -like '*NVIDIA*')   { return $true }
            if ($g.Name                 -like '*NVIDIA*')   { return $true }
            if ($g.Name                 -like '*GeForce*')  { return $true }
            if ($g.Name                 -like '*Quadro*')   { return $true }
        }
        return $false
    } catch { return $true }
}

function Get-LenovoPnpDevicesByHwId {
    # Returns @() of PnP devices whose HardwareID (or CompatibleID) contains the
    # given pattern as a substring (case-insensitive). Cached per-call: callers
    # in a hot loop should grab the device list once and filter manually if perf
    # matters - but each <DetectInstall> only runs a couple of these.
    param([string]$Pattern)
    if (-not $Pattern) { return @() }
    try {
        $all = Get-PnpDevice -EA Stop
    } catch { return @() }
    return @($all | Where-Object {
        $hit = $false
        if ($_.HardwareID) {
            foreach ($id in $_.HardwareID) { if ($id -like "*$Pattern*") { $hit = $true; break } }
        }
        if (-not $hit -and $_.CompatibleID) {
            foreach ($id in $_.CompatibleID) { if ($id -like "*$Pattern*") { $hit = $true; break } }
        }
        $hit
    })
}

function Test-LenovoDriverFileVersionDetect {
    # <_DriverFileVersion>
    #   <HardwareID>VEN_8086&DEV_xxxx</HardwareID>   (or <Hardware> / <PnPID>)
    #   <Version>30.0.101.1960^</Version>
    # </_DriverFileVersion>
    # Semantics: returns TRUE (already installed) if either
    #   (a) no PnP device matches the hardware ID on this machine (driver
    #       not applicable - common on dual-SKU MTMs), OR
    #   (b) at least one matching device already has a driver bound whose
    #       version satisfies the constraint.
    param([System.Xml.XmlElement]$Node)
    try {
        $hwId = $null
        if ($Node.HardwareID)   { $hwId = [string]$Node.HardwareID }
        elseif ($Node.Hardware) { $hwId = [string]$Node.Hardware }
        elseif ($Node.PnPID)    { $hwId = [string]$Node.PnPID }
        $hwId = if ($hwId) { $hwId.Trim() } else { "" }
        $ver  = if ($Node.Version) { ([string]$Node.Version).Trim() } else { "" }
        if (-not $hwId -or -not $ver) { return $null }

        $matched = Get-LenovoPnpDevicesByHwId -Pattern $hwId
        if ($matched.Count -eq 0) {
            # Driver isn't applicable to this machine - treat as "already handled"
            # so we skip download + install entirely.
            return $true
        }

        # Find the highest installed driver version across all matched devices.
        $highest = $null
        foreach ($d in $matched) {
            try {
                $p = Get-PnpDeviceProperty -InstanceId $d.InstanceId -KeyName 'DEVPKEY_Device_DriverVersion' -EA SilentlyContinue
                if (-not $p -or -not $p.Data) { continue }
                $dv = [string]$p.Data
                try {
                    $vObj = [version]$dv
                    if (-not $highest -or $vObj -gt $highest) { $highest = $vObj }
                } catch {}
            } catch {}
        }
        if (-not $highest) { return $null }
        return (Compare-LenovoVersionString -Current $highest.ToString() -Expected $ver)
    } catch { return $null }
}

function Test-LenovoPnPIDDetect {
    # <_PnPID>VEN_xxxx&DEV_yyyy</_PnPID>
    # In <DetectInstall>, Lenovo uses this to confirm the package's target
    # PnP entry exists - which we read as "the package has already installed
    # whatever PnP device it was supposed to install".
    param([System.Xml.XmlElement]$Node)
    try {
        $pat = [string]$Node.InnerText
        if ($pat) { $pat = $pat.Trim() }
        if (-not $pat) { return $null }
        $matched = Get-LenovoPnpDevicesByHwId -Pattern $pat
        return ($matched.Count -gt 0)
    } catch { return $null }
}

function Test-LenovoDetectNode {
    param([System.Xml.XmlElement]$Node)
    if (-not $Node) { return $null }
    $kids = @($Node.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })

    switch ($Node.LocalName) {
        'And' {
            $sawNull = $false
            foreach ($k in $kids) {
                $r = Test-LenovoDetectNode -Node $k
                if ($r -eq $false) { return $false }
                if ($null -eq $r)  { $sawNull = $true }
            }
            if ($sawNull) { return $null }
            return $true
        }
        'Or' {
            $allNull = $true
            foreach ($k in $kids) {
                $r = Test-LenovoDetectNode -Node $k
                if ($r -eq $true)  { return $true }
                if ($null -ne $r)  { $allNull = $false }
            }
            if ($allNull) { return $null }
            return $false
        }
        'Not' {
            if ($kids.Count -ne 1) { return $null }
            $r = Test-LenovoDetectNode -Node $kids[0]
            if ($null -eq $r) { return $null }
            return (-not $r)
        }
        '_Bios'              { return (Test-LenovoBiosLevel              -Node $Node) }
        '_File'              { return (Test-LenovoFileDetect             -Node $Node) }
        '_FileVersion'       { return (Test-LenovoFileVersionDetect      -Node $Node) }
        '_Registry'          { return (Test-LenovoRegistryDetect         -Node $Node) }
        '_RegistryKey'       { return (Test-LenovoRegistryKeyDetect      -Node $Node) }
        '_DriverFileVersion' { return (Test-LenovoDriverFileVersionDetect -Node $Node) }
        '_PnPID'             { return (Test-LenovoPnPIDDetect            -Node $Node) }
        default              { return $null }   # unknown element -> fail-open
    }
}

function Test-LenovoDetectInstall {
    param([System.Xml.XmlElement]$Node)
    if (-not $Node) { return $null }
    $kids = @($Node.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
    # Empty <DetectInstall/> means "no detection info" - tell caller to install.
    if ($kids.Count -eq 0) { return $false }
    # Multiple top-level children are AND'd (Lenovo convention).
    if ($kids.Count -eq 1) { return (Test-LenovoDetectNode -Node $kids[0]) }
    $sawNull = $false
    foreach ($k in $kids) {
        $r = Test-LenovoDetectNode -Node $k
        if ($r -eq $false) { return $false }
        if ($null -eq $r)  { $sawNull = $true }
    }
    if ($sawNull) { return $null }
    return $true
}

# =========================
# LENOVO POST-INSTALL HELPERS (used by consumer catalog flow)
#
# Test-LenovoDPInstStaged   - True if exit code looks like "DPInst copied driver
#                              to the store but didn't bind it to any device".
# Test-LenovoHwMismatchCode - True if exit code looks like "the target hardware
#                              isn't present" (NVIDIA installer on iGPU-only,
#                              Realtek installer on non-Realtek, etc).
# Invoke-LenovoPnputilInstall - Walks a package dir, runs `pnputil /add-driver
#                              /install` on every INF. Used as a fallback when
#                              DPInst stages but doesn't bind.
# Invoke-LenovoPnpRescan    - Triggers `pnputil /scan-devices` to catch any
#                              staged drivers still waiting for a device match.
# =========================

function Test-LenovoDPInstStaged {
    param([int64]$ExitCode)
    # DPInst packs: bits 0-7 = bound, bits 8-15 = staged, bits 16-23 = failed.
    # "Staged but not bound" = nothing bound, something staged, nothing failed.
    # Short-circuit on <=0: success codes are 0 and negative codes are typically
    # vendor-specific errors (e.g. NVIDIA's -436207360) that don't follow
    # DPInst's packing convention. The HW-mismatch whitelist catches those.
    # NOTE: do NOT bound-check against 0xFFFFFFFF - in PowerShell that literal
    # parses as Int32 -1 (signed), which made the previous version of this
    # function return $false for every positive exit code.
    if ($ExitCode -le 0) { return $false }
    $bound  = $ExitCode -band 0xFF
    $staged = ($ExitCode -shr 8) -band 0xFF
    $failed = ($ExitCode -shr 16) -band 0xFF
    return ($bound -eq 0 -and $staged -gt 0 -and $failed -eq 0)
}

# Codes that mean "target hardware not present on this machine". Lenovo's
# per-MTM catalogs list every variant's drivers; on a single-GPU machine the
# discrete-GPU installer correctly fails because there's nothing to install to.
$script:LenovoHwMismatchCodes = @(
    -436207360    # nvsetup.exe: 0xE6000100 - no compatible NVIDIA hardware
)

function Test-LenovoHwMismatchCode {
    param([int64]$ExitCode, [string]$Title = "")
    return ($script:LenovoHwMismatchCodes -contains [int]$ExitCode)
}

function Invoke-LenovoPnputilInstall {
    param([string]$PackagePath)
    # Returns $true if at least one INF was successfully added; $false otherwise.
    $infs = @(Get-ChildItem -Path $PackagePath -Recurse -Filter '*.inf' -EA SilentlyContinue)
    if ($infs.Count -eq 0) {
        Log "    pnputil fallback: no .inf files under $PackagePath - nothing to do."
        return $false
    }
    Log "    pnputil fallback: $($infs.Count) INF file(s) to add."
    $anyOk = $false
    foreach ($inf in $infs) {
        Log "    pnputil /add-driver `"$($inf.Name)`" /install"
        $out = pnputil /add-driver "`"$($inf.FullName)`"" /install 2>&1
        foreach ($l in $out) { Log "      $l" }
        # 0 = success, 259 = no more items (often benign), 3010 = success+reboot
        if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 259 -or $LASTEXITCODE -eq 3010) { $anyOk = $true }
    }
    return $anyOk
}

function Invoke-LenovoPnpRescan {
    Log ""
    Log "Triggering PnP rescan (pnputil /scan-devices) to bind any staged drivers..."
    try {
        $out = pnputil /scan-devices 2>&1
        foreach ($l in $out) { Log "  $l" }
        Log "  pnputil /scan-devices exit: $LASTEXITCODE"
    } catch {
        Log "  pnputil /scan-devices failed: $($_.Exception.Message)"
    }
}

function Get-LenovoUnboundDevices {
    # Match the script's existing Get-MissingDriverCount approach exactly so
    # counts agree between the pre/post snapshot and this helper.
    try {
        return @(Get-CimInstance Win32_PnPEntity -EA Stop |
                 Where-Object { $_.ConfigManagerErrorCode -ne 0 })
    } catch { return @() }
}

function Invoke-LenovoForceBindUnbound {
    # Final brute-force pass. If anything is still unbound after the rescan,
    # walk every INF in the consumer-catalog package tree and run
    # `pnputil /add-driver /install`. Catches the case where DPInst staged a
    # driver but the binding never happened (or the install command was a
    # vendor wrapper that bypassed proper INF registration).
    param([string]$PkgRoot)

    $before = Get-LenovoUnboundDevices
    if ($before.Count -eq 0) {
        Log "Force-bind check: no devices in problem state. Nothing to do."
        return
    }

    Log ""
    Log "Force-bind check: $($before.Count) device(s) still in problem state after rescan:"
    foreach ($d in $before) {
        $cap = if ($d.Caption) { $d.Caption } elseif ($d.Name) { $d.Name } else { '(unnamed)' }
        Log "  - $cap"
        if ($d.DeviceID)              { Log "      DeviceID: $($d.DeviceID)" }
        if ($d.ConfigManagerErrorCode){ Log "      ErrorCode: $($d.ConfigManagerErrorCode)" }
    }

    if (-not (Test-Path $PkgRoot)) {
        Log "  Consumer pkg root not found ($PkgRoot) - cannot retry."
        return
    }

    $infs = @(Get-ChildItem -Path $PkgRoot -Recurse -Filter '*.inf' -EA SilentlyContinue)
    if ($infs.Count -eq 0) {
        Log "  No INF files under $PkgRoot - nothing to retry."
        return
    }

    Log ""
    Log "Force-bind pass: pnputil /add-driver /install on $($infs.Count) INF(s)..."
    foreach ($inf in $infs) {
        if (Test-Cancelled) { Log "  Cancelled."; break }
        $out = pnputil /add-driver "`"$($inf.FullName)`"" /install 2>&1
        # Filter pnputil's banner + blanks for log readability
        foreach ($l in $out) {
            $line = ([string]$l).Trim()
            if (-not $line)                            { continue }
            if ($line -like 'Microsoft PnP Utility*')  { continue }
            Log "    [$($inf.Name)] $line"
        }
    }

    # One more scan after the brute-force pass
    Log ""
    Log "Final scan-devices after force-bind..."
    try {
        $out = pnputil /scan-devices 2>&1
        foreach ($l in $out) { Log "  $l" }
    } catch {}

    $after = Get-LenovoUnboundDevices
    $delta = $before.Count - $after.Count
    if ($delta -gt 0) {
        Log "Force-bind resolved $delta device(s)."
    }
    if ($after.Count -gt 0) {
        Log ""
        Log "$($after.Count) device(s) STILL unbound:"
        foreach ($d in $after) {
            $cap = if ($d.Caption) { $d.Caption } elseif ($d.Name) { $d.Name } else { '(unnamed)' }
            Log "  - $cap"
            if ($d.DeviceID) { Log "      DeviceID: $($d.DeviceID)" }
        }
        Log ""
        Log "These devices have no INF in the consumer catalog whose HardwareID or"
        Log "CompatibleID list matches. Either Lenovo's per-MTM catalog does not"
        Log "publish a driver for this hardware variant, or a system reboot is"
        Log "required before binding can complete."
    }
}

function Invoke-LenovoPackageCommand {
    param(
        [string]$Command,
        [string]$WorkingDir,
        [int]$TimeoutSec = 1200
    )
    # cmd.exe /c handles quoting + path-with-spaces cleanly. The downside is
    # cmd swallows '^' as escape - Lenovo install lines do not use it, so safe.
    try {
        $psi                  = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName         = $env:ComSpec
        $psi.Arguments        = "/c $Command"
        $psi.WorkingDirectory = $WorkingDir
        $psi.UseShellExecute  = $false
        $psi.CreateNoWindow   = $true
        $proc                 = New-Object System.Diagnostics.Process
        $proc.StartInfo       = $psi
        $proc.Start() | Out-Null

        $deadline = (Get-Date).AddSeconds($TimeoutSec)
        while (-not $proc.HasExited) {
            Start-Sleep -Milliseconds 500
            Step-AllSpinners
            [System.Windows.Forms.Application]::DoEvents()
            if ((Get-Date) -gt $deadline) {
                Log "  WARNING: process exceeded $TimeoutSec s - killing."
                try { $proc.Kill() } catch {}
                return -9999
            }
            if ($script:CancelRequested) {
                try { $proc.Kill() } catch {}
                return -9998
            }
        }
        return $proc.ExitCode
    } catch {
        Log "  Command exec error: $($_.Exception.Message)"
        return -9997
    }
}

function Start-LenovoConsumerCatalogInstall {
    param(
        [string]$MachineType,   # 4-char (e.g. "20VD")
        [string]$OsCode,        # "Win10" or "Win11"
        [string]$DriverRoot
    )

    Log "--- Lenovo consumer catalog (System Update mechanism) ---"
    $catalogUrl  = "https://download.lenovo.com/catalog/${MachineType}_${OsCode}.xml"
    $catalogFile = Join-Path $env:TEMP "lenovo_${MachineType}_${OsCode}.xml"
    Log "Catalog URL: $catalogUrl"

    if (-not (Invoke-CurlDownload -Url $catalogUrl -OutFile $catalogFile)) {
        Log "Consumer catalog not available for ${MachineType}_${OsCode} (404 or fetch failed)."
        return $false
    }
    if (Test-Cancelled) { return $false }

    # Parse catalog (BOM-safe)
    try {
        $bytes    = [System.IO.File]::ReadAllBytes($catalogFile)
        $rawText  = [System.Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF)
        [xml]$cat = $rawText
    } catch {
        Log "Consumer catalog parse failed: $($_.Exception.Message)"
        return $false
    }

    $pkgNodes = @($cat.packages.package)
    if ($pkgNodes.Count -eq 0) {
        Log "Consumer catalog has 0 packages - cannot proceed."
        return $false
    }
    Log "Consumer catalog lists $($pkgNodes.Count) package(s)."
    SetProgress 10

    # Working dir for all consumer-catalog payloads
    $pkgRoot = Join-Path $DriverRoot "Lenovo_Consumer"
    if (-not (Test-Path $pkgRoot)) { New-Item -ItemType Directory -Path $pkgRoot -Force | Out-Null }

    $totalPkgs    = $pkgNodes.Count
    $pkgIdx       = 0
    $okCount      = 0
    $failCount    = 0
    $skipCount    = 0
    $alreadyInstalledCount = 0
    $dockSkipCount = 0
    $naCount      = 0
    $rebootNeeded = $false
    $results      = New-Object System.Collections.Generic.List[object]
    # BIOS (PackageType=3) installs are queued here and run AFTER everything else,
    # so the BIOS GUI doesn't interrupt the driver/firmware install sequence.
    $deferredBios = New-Object System.Collections.Generic.List[object]

    foreach ($pkg in $pkgNodes) {
        $pkgIdx++
        if (Test-Cancelled) { Log "Cancelled at package $pkgIdx/$totalPkgs."; break }

        $descUrl         = ([string]$pkg.location).Trim()
        $category        = if ($pkg.category) { ([string]$pkg.category).Trim() } else { "" }
        $expectedDescSha = ""
        if ($pkg.checksum) {
            $expectedDescSha = if ($pkg.checksum.'#text') { $pkg.checksum.'#text' } else { [string]$pkg.checksum }
            $expectedDescSha = $expectedDescSha.Trim().ToLower()
        }

        # Map overall progress 15-95 across packages
        $overall = 15 + [int](($pkgIdx / $totalPkgs) * 80)
        SetProgress $overall
        Log ""
        Log "----- [$pkgIdx/$totalPkgs] $category -----"

        # 1. Descriptor XML
        $descName = [System.IO.Path]::GetFileName(([System.Uri]$descUrl).LocalPath)
        $descFile = Join-Path $pkgRoot $descName
        if (-not (Invoke-CurlDownload -Url $descUrl -OutFile $descFile)) {
            Log "  Descriptor download failed - skipping."
            $failCount++; continue
        }
        if ($expectedDescSha) {
            $actualSha = (Get-FileHash -Path $descFile -Algorithm SHA256).Hash.ToLower()
            if ($actualSha -ne $expectedDescSha) {
                Log "  Descriptor SHA256 mismatch - skipping."
                Log "    expected: $expectedDescSha"
                Log "    actual:   $actualSha"
                $failCount++; continue
            }
        }

        # 2. Parse descriptor (BOM-safe)
        try {
            $descBytes = [System.IO.File]::ReadAllBytes($descFile)
            $descText  = [System.Text.Encoding]::UTF8.GetString($descBytes).TrimStart([char]0xFEFF)
            [xml]$desc = $descText
        } catch {
            Log "  Descriptor parse failed: $($_.Exception.Message)"
            $failCount++; continue
        }

        $pkgName     = [string]$desc.Package.name
        $pkgId       = [string]$desc.Package.id
        $pkgVer      = [string]$desc.Package.version
        $releaseDate = [string]$desc.Package.ReleaseDate
        $severity    = if ($desc.Package.Severity)    { [string]$desc.Package.Severity.type }    else { '?' }
        $packageType = if ($desc.Package.PackageType) { [string]$desc.Package.PackageType.type } else { '?' }
        $rebootType  = if ($desc.Package.Reboot)      { [string]$desc.Package.Reboot.type }      else { '0' }

        # Title (EN preferred)
        $title = ""
        try {
            $descNodes = @($desc.Package.Title.Desc)
            $enDesc = $descNodes | Where-Object { $_.id -eq 'EN' } | Select-Object -First 1
            if (-not $enDesc) { $enDesc = $descNodes | Select-Object -First 1 }
            if ($enDesc) { $title = [string]$enDesc.InnerText }
        } catch {}
        if (-not $title) { $title = $pkgName }

        $typeName = switch ($packageType) {
            '1' { 'Application' }
            '2' { 'Driver' }
            '3' { 'BIOS' }
            '4' { 'Firmware' }
            default { "Type$packageType" }
        }
        $sevName = switch ($severity) {
            '1' { 'Critical' }
            '2' { 'Recommended' }
            '3' { 'Optional' }
            default { "Sev$severity" }
        }

        Log "  $title"
        Log "  id=$pkgId  version=$pkgVer  released=$releaseDate"
        Log "  $typeName / $sevName  (reboot type $rebootType)"

        # 2a. Skip dock-related packages entirely. Dock firmware/drivers don't
        #     belong in a one-shot driver run - they're hardware-specific to whatever
        #     dock the user has plugged in (or doesn't have plugged in).
        if ($title -like '*Dock*') {
            Log "  Dock package - skipping per script policy (no download, no install)."
            $dockSkipCount++
            $results.Add([pscustomobject]@{
                Index=$pkgIdx; Title=$title; Category=$category; Version=$pkgVer; Exit=$null; Status='SKIPPED-DOCK'
            })
            continue
        }

        # 2a2. NVIDIA driver packages on a machine with no NVIDIA GPU. The 20VD
        #      MTM covers both integrated-only and integrated+MX450 variants;
        #      Lenovo lists both sets of drivers per MTM. Catch this at catalog
        #      parse time so we don't waste 600+ MB downloading a driver that
        #      nvsetup will reject at install time anyway.
        if ($title -like '*NVIDIA*' -and -not (Get-LenovoHasNvidiaGpu)) {
            Log "  Package targets NVIDIA but no NVIDIA GPU detected on this machine. Skipping (no download)."
            $naCount++
            $results.Add([pscustomobject]@{
                Index=$pkgIdx; Title=$title; Category=$category; Version=$pkgVer; Exit=$null; Status='N/A-NO-HW'
            })
            continue
        }

        # 2b. Evaluate <DetectInstall> - skip if package is already installed.
        #     fail-open: $null (indeterminate) -> install anyway, log a note.
        $detectNode = $null
        try { $detectNode = $desc.Package.DetectInstall } catch {}
        if ($detectNode -is [System.Xml.XmlElement]) {
            $detectResult = Test-LenovoDetectInstall -Node $detectNode
            if ($detectResult -eq $true) {
                Log "  DetectInstall -> ALREADY INSTALLED. Skipping (no download)."
                $alreadyInstalledCount++
                $results.Add([pscustomobject]@{
                    Index=$pkgIdx; Title=$title; Category=$category; Version=$pkgVer; Exit=$null; Status='ALREADY-INSTALLED'
                })
                continue
            } elseif ($null -eq $detectResult) {
                Log "  DetectInstall -> indeterminate (unsupported element). Will install."
            } else {
                Log "  DetectInstall -> not installed. Will install."
            }
        }

        # 3. Installer file metadata
        $installerNode = $null
        try { $installerNode = $desc.Package.Files.Installer.File } catch {}
        if (-not $installerNode -or -not $installerNode.Name) {
            Log "  No <Installer><File> entry - cannot install. Skipping."
            $skipCount++; continue
        }
        $installerName = [string]$installerNode.Name
        $expectedCrc   = if ($installerNode.CRC) { ([string]$installerNode.CRC).ToLower() } else { "" }

        # Installer URL = same directory as descriptor URL + installer filename
        $descDir      = $descUrl.Substring(0, $descUrl.LastIndexOf('/') + 1)
        $installerUrl = $descDir + $installerName

        # Per-package extract directory == %PACKAGEPATH%
        $packagePath   = Join-Path $pkgRoot $pkgId
        if (-not (Test-Path $packagePath)) { New-Item -ItemType Directory -Path $packagePath -Force | Out-Null }
        $installerFile = Join-Path $packagePath $installerName

        # 4. Download installer
        if (-not (Invoke-CurlDownload -Url $installerUrl -OutFile $installerFile)) {
            Log "  Installer download failed - skipping."
            $failCount++; continue
        }
        if ($expectedCrc) {
            $actualCrc = (Get-FileHash -Path $installerFile -Algorithm SHA256).Hash.ToLower()
            if ($actualCrc -ne $expectedCrc) {
                Log "  Installer SHA256 mismatch - skipping."
                Log "    expected: $expectedCrc"
                Log "    actual:   $actualCrc"
                $failCount++; continue
            }
            Log "  SHA256 OK."
        }

        # 5. Download External helpers (version checkers etc.) if present
        try {
            if ($desc.Package.Files.External) {
                foreach ($ext in @($desc.Package.Files.External.File)) {
                    if (-not $ext -or -not $ext.Name) { continue }
                    $extName = [string]$ext.Name
                    $extUrl  = $descDir + $extName
                    $extFile = Join-Path $packagePath $extName
                    Log "  External helper: $extName"
                    if (-not (Invoke-CurlDownload -Url $extUrl -OutFile $extFile)) {
                        Log "  WARNING: external helper download failed (continuing)."
                        continue
                    }
                    if ($ext.CRC) {
                        $extSha    = (Get-FileHash -Path $extFile -Algorithm SHA256).Hash.ToLower()
                        $extExpect = ([string]$ext.CRC).ToLower()
                        if ($extSha -ne $extExpect) { Log "  WARNING: external helper SHA256 mismatch (continuing)." }
                    }
                }
            }
        } catch { Log "  External helper handling error: $($_.Exception.Message)" }

        if ($SkipInstall) {
            Log "  -SkipInstall set: download + verify only."
            $okCount++
            $results.Add([pscustomobject]@{
                Index=$pkgIdx; Title=$title; Category=$category; Version=$pkgVer; Exit=$null; Status='DOWNLOAD-ONLY'
            })
            continue
        }

        # 6. Run extract command (if present) - unpacks payload into %PACKAGEPATH%
        $extractCmd = [string]$desc.Package.ExtractCommand
        if ($extractCmd) {
            $cmd = $extractCmd.Replace('%PACKAGEPATH%', $packagePath).Replace('%WINDOWS%', $env:WINDIR)
            Log "  Extract: $cmd"
            SetExtract -Pct $overall -Label "Extracting [$pkgIdx/$totalPkgs]: $title"
            $extExit = Invoke-LenovoPackageCommand -Command $cmd -WorkingDir $packagePath -TimeoutSec 600
            Log "  Extract exit code: $extExit"
            # Extract exit codes vary by vendor - non-zero isn't necessarily failure, continue regardless.
        }

        # 7. Run install command
        $installNode = $null
        try { $installNode = $desc.Package.Install } catch {}
        if (-not $installNode) {
            Log "  No <Install> block - skipping."
            $skipCount++; continue
        }

        # Cmdline: prefer id="EN", else first one
        $installCmdRaw = ""
        try {
            $cmdNodes = @($installNode.Cmdline)
            $enCmd = $cmdNodes | Where-Object { $_.id -eq 'EN' } | Select-Object -First 1
            if (-not $enCmd) { $enCmd = $cmdNodes | Select-Object -First 1 }
            if ($enCmd) {
                $installCmdRaw = if ($enCmd -is [System.Xml.XmlElement]) { $enCmd.InnerText } else { [string]$enCmd }
            }
        } catch {}
        if (-not $installCmdRaw) {
            Log "  No <Cmdline> text - skipping."
            $skipCount++; continue
        }

        # Acceptable exit codes: 0 plus Windows reboot codes plus whatever rc="" lists
        $rcOk = New-Object System.Collections.Generic.HashSet[int64]
        [void]$rcOk.Add(0); [void]$rcOk.Add(3010); [void]$rcOk.Add(1641)
        if ($installNode.rc) {
            foreach ($r in (([string]$installNode.rc) -split ',')) {
                $r = $r.Trim()
                if ($r -match '^-?\d+$') { [void]$rcOk.Add([int64]$r) }
            }
        }

        $installCmd = $installCmdRaw.Replace('%PACKAGEPATH%', $packagePath).Replace('%WINDOWS%', $env:WINDIR)

        # BIOS packages (PackageType=3): download is done, but defer the install
        # itself until after every other package has finished. BIOS updaters pop
        # their own GUI which would otherwise stall the rest of the queue.
        if ($packageType -eq '3') {
            Log "  BIOS package - install deferred until end of run."
            Log "  Deferred install: $installCmd"
            $deferredBios.Add([pscustomobject]@{
                Index       = $pkgIdx
                Title       = $title
                Category    = $category
                Version     = $pkgVer
                PackagePath = $packagePath
                InstallCmd  = $installCmd
                RcOk        = $rcOk
                RebootType  = $rebootType
                PackageType = $packageType
            })
            continue
        }

        Log "  Install: $installCmd"
        SetExtract -Pct $overall -Label "Installing [$pkgIdx/$totalPkgs]: $title"

        # Driver / Firmware / App install. Extract already ran in step 6 if needed.
        $timeout = if ($packageType -eq '4') { 1800 } else { 1200 }
        $exit = Invoke-LenovoPackageCommand -Command $installCmd -WorkingDir $packagePath -TimeoutSec $timeout
        $isOk = $rcOk.Contains([int64]$exit)
        $statusLabel = if ($isOk) { 'OK' } else { 'FAIL' }

        # Special case: DPInst exit indicates the driver was staged in the driver
        # store but never bound to a device. Retry via pnputil, which explicitly
        # tries to bind staged drivers to matching hardware.
        if (-not $isOk -and (Test-LenovoDPInstStaged -ExitCode $exit)) {
            Log "  Install exit $exit -> driver staged in store but not bound. Trying pnputil fallback..."
            $fbOk = Invoke-LenovoPnputilInstall -PackagePath $packagePath
            if ($fbOk) {
                Log "  pnputil fallback succeeded - driver re-staged for binding at next PnP scan."
                $isOk = $true
                $statusLabel = 'OK-FALLBACK'
            } else {
                Log "  pnputil fallback did not succeed."
            }
        }

        # Special case: known hardware-not-present codes (e.g. NVIDIA installer
        # on iGPU-only machines). Lenovo's per-MTM catalog lists every variant's
        # drivers; a mismatch here is expected, not a failure.
        if (-not $isOk -and (Test-LenovoHwMismatchCode -ExitCode $exit -Title $title)) {
            Log "  Install exit $exit -> target hardware not present on this machine."
            $statusLabel = 'N/A'
        }

        Log "  Install exit: $exit  [$statusLabel]"

        if ($isOk) {
            $okCount++
            if ($exit -eq 3010 -or $exit -eq 1641 -or $rebootType -in @('1','3','4','5')) {
                $rebootNeeded = $true
            }
            $null = $script:AnalyticsInstalledDrivers.Add($title)
        } elseif ($statusLabel -eq 'N/A') {
            $naCount++
        } else {
            $failCount++
        }
        $results.Add([pscustomobject]@{
            Index=$pkgIdx; Title=$title; Category=$category; Version=$pkgVer; Exit=$exit; Status=$statusLabel
        })

        Play-Sound -Event "DriverAdded"
    }

    # ------------------------------------------------------------------
    # Deferred phase: now run all BIOS installs that we queued above.
    # By this point every driver/firmware/app has finished, so the BIOS
    # GUI can take over the screen without interrupting anything.
    # ------------------------------------------------------------------
    if ($deferredBios.Count -gt 0 -and -not $SkipInstall -and -not (Test-Cancelled)) {
        Log ""
        Log "=== DEFERRED PHASE: $($deferredBios.Count) BIOS package(s) ==="
        SetProgress 96
        $biosIdx = 0
        foreach ($b in $deferredBios) {
            $biosIdx++
            if (Test-Cancelled) { Log "Cancelled before BIOS install $biosIdx/$($deferredBios.Count)."; break }
            Log ""
            Log "----- DEFERRED BIOS [$biosIdx/$($deferredBios.Count)] $($b.Title) (v$($b.Version)) -----"
            SetExtract -Pct 100 -Label "BIOS [$biosIdx/$($deferredBios.Count)]: $($b.Title)"

            # Extract already ran during the main loop's step 6, so go straight to install.
            Log "  Install: $($b.InstallCmd)"
            $bExit = Invoke-LenovoPackageCommand -Command $b.InstallCmd -WorkingDir $b.PackagePath -TimeoutSec 1800
            $bOk   = $b.RcOk.Contains([int64]$bExit)
            $bStatus = if ($bOk) { 'OK' } else { 'FAIL' }
            Log "  Install exit: $bExit  [$bStatus]"

            if ($bOk) {
                $okCount++
                if ($bExit -eq 3010 -or $bExit -eq 1641 -or $b.RebootType -in @('1','3','4','5')) {
                    $rebootNeeded = $true
                }
                $null = $script:AnalyticsInstalledDrivers.Add($b.Title)
            } else {
                $failCount++
            }
            $results.Add([pscustomobject]@{
                Index=$b.Index; Title=$b.Title; Category=$b.Category; Version=$b.Version; Exit=$bExit; Status=$bStatus
            })
            Play-Sound -Event "DriverAdded"
        }
    }

    # Final pass: kick a PnP rescan so any drivers staged in the store (DPInst
    # fallback path, plus anything else) get matched to their devices.
    if (-not $SkipInstall -and -not (Test-Cancelled)) {
        Invoke-LenovoPnpRescan
        # If anything still isn't bound, force-retry every INF in the consumer
        # pkg tree to give Windows one more chance to find a match.
        Invoke-LenovoForceBindUnbound -PkgRoot $pkgRoot
    }

    # Summary
    Log ""
    Log "=== LENOVO CONSUMER CATALOG SUMMARY ==="
    Log "Machine type: $MachineType  OS: $OsCode"
    Log "Total packages:     $totalPkgs"
    Log "Installed:          $okCount"
    Log "Already installed:  $alreadyInstalledCount  (detected via <DetectInstall>)"
    Log "Dock packages:      $dockSkipCount  (skipped by policy)"
    Log "Not applicable:     $naCount  (hardware not present)"
    Log "Failed:             $failCount"
    Log "Skipped (no data):  $skipCount"
    Log "BIOS deferred:      $($deferredBios.Count)"
    if ($rebootNeeded) { Log "*** REBOOT REQUIRED to finalize one or more updates ***" }
    foreach ($r in $results) {
        $mark = switch ($r.Status) {
            'OK'                { '[OK]  ' }
            'OK-FALLBACK'       { '[OK*] ' }
            'DOWNLOAD-ONLY'     { '[DL]  ' }
            'ALREADY-INSTALLED' { '[SKIP]' }
            'SKIPPED-DOCK'      { '[DOCK]' }
            'N/A'               { '[N/A] ' }
            'N/A-NO-HW'         { '[N/A] ' }
            default             { '[FAIL]' }
        }
        $exitTxt = if ($null -eq $r.Exit) { '-' } else { "exit $($r.Exit)" }
        Log "  $mark [$($r.Index)/$totalPkgs] $($r.Category) - $($r.Title) (v$($r.Version)) $exitTxt"
    }

    $script:AnalyticsInfCount = $okCount
    SetProgress 100
    SetExtract -Pct 100 -Label "Done - $okCount installed, $alreadyInstalledCount already current"
    Stop-ExSpinner      -Success ($failCount -eq 0)
    Stop-OverallSpinner -Success ($failCount -eq 0)

    # Treat the run as a success if we either installed something OR confirmed
    # everything was already up to date OR detected hardware-mismatch correctly.
    return (($okCount + $alreadyInstalledCount + $naCount) -gt 0)
}

function Get-LenovoFullPackUrl {
    # v1.11.0 - factored out of the old inline Path-2 code in
    # Start-LenovoDriverInstall. Returns the SCCM driver pack URL from
    # catalogv2.xml (Lenovo's "Recipe Card" listing) for the given machine
    # type, or $null if the machine type isn't listed. Tries the detected
    # OS first, then the other Windows generation as a fallback.
    param(
        [Parameter(Mandatory=$true)] [string]$MachineType,
        [Parameter(Mandatory=$true)] [ValidateSet('Win10','Win11')] [string]$OsCode
    )
    Log "Checking Lenovo Recipe Card / catalogv2.xml for full driver pack..."
    $catalogFile = Join-Path $env:TEMP "lenovo_catalogv2.xml"
    if (-not (Invoke-CurlDownload -Url "https://download.lenovo.com/cdrt/td/catalogv2.xml" -OutFile $catalogFile)) {
        Log "  catalogv2.xml download failed."
        return $null
    }
    if (Test-Cancelled) { return $null }

    try {
        $bytes    = [System.IO.File]::ReadAllBytes($catalogFile)
        $rawText  = [System.Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF)
        [xml]$cat = $rawText
    } catch {
        Log "  catalogv2.xml parse failed: $($_.Exception.Message)"
        return $null
    }

    $osAttr     = $OsCode.ToLower()
    $osFallback = if ($osAttr -eq 'win11') { 'win10' } else { 'win11' }

    foreach ($model in $cat.ModelList.Model) {
        $types = @($model.Types.Type)
        if (-not ($types | Where-Object { $_ -like "$MachineType*" })) { continue }
        Log "  Matched Recipe Card model: $($model.name)"
        foreach ($os in @($osAttr, $osFallback)) {
            $nodes = @($model.SCCM | Where-Object { $_.os -eq $os })
            if ($nodes.Count -gt 0) {
                $url = ($nodes | Select-Object -Last 1)."#text"
                if ($url -match "^https?://") {
                    Log "  Recipe Card pack URL [$os]: $url"
                    return $url
                }
            }
        }
        # Matched the model but no SCCM URL for either OS - no point checking more models
        break
    }
    Log "  No Recipe Card / SCCM pack listed for machine type '$MachineType'."
    return $null
}

function Install-LenovoFullPack {
    # v1.11.0 - factored out of Start-LenovoDriverInstall's inline path-2 code.
    # Download -> Inno-Setup extract -> pnputil. Same flow that's been used
    # since pre-v1.8.0; just isolated for clarity.
    param(
        [Parameter(Mandatory=$true)] [string]$PackUrl,
        [Parameter(Mandatory=$true)] [string]$DriverRoot
    )
    SetProgress 28
    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }
    $packFile = Join-Path $DriverRoot ([System.IO.Path]::GetFileName(([System.Uri]$PackUrl).LocalPath))
    SetProgress 30
    if (-not (Invoke-CurlDownload -Url $PackUrl -OutFile $packFile)) {
        Log "Lenovo driver pack download failed."
        return $false
    }
    if (Test-Cancelled) { return $false }
    SetProgress 55

    $extractPath = Join-Path $DriverRoot "Lenovo_Extracted"
    Log "Extracting Lenovo pack..."
    Start-PackExtraction -PackFile $packFile -DestPath $extractPath -StallLimitSec 300 -Vendor "Lenovo"
    SetProgress 60
    return (Install-DriversFromPath -BasePath $extractPath)
}

function Start-LenovoDriverInstall {
    param([string]$DriverRoot)

    Log "=== LENOVO: Starting automated driver install ==="
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."

    $machineType = $null
    if ($MachineType) {
        # Use override from -MachineType param - take first 4 chars uppercased
        $machineType = $MachineType.Substring(0, [math]::Min(4, $MachineType.Length)).ToUpper()
        Log "Machine type: $MachineType  ->  prefix: $machineType  [overridden via param]"
    } else {
        try {
            $sku = (Get-CimInstance Win32_ComputerSystemProduct).Name.Trim()
            if ($sku.Length -ge 4) {
                $machineType = $sku.Substring(0, 4).ToUpper()
                Log "Machine type: $sku  ->  prefix: $machineType"
            }
        } catch { Log "Could not read machine type: $($_.Exception.Message)" }
    }
    if (-not $machineType) { Log "Cannot determine Lenovo machine type."; return $false }

    try { $script:AnalyticsSerial = (Get-CimInstance Win32_BIOS).SerialNumber.Trim() } catch {}

    # BuildNumber is the unambiguous Win10 vs Win11 discriminator (>=22000 = Win11).
    $buildNum = 0
    try { $buildNum = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber } catch {}
    $osCode = if ($buildNum -ge 22000) { 'Win11' } else { 'Win10' }
    Log "Detected OS: $osCode (build $buildNum)"

    # ------------------------------------------------------------------
    # v1.11.0 PATH PRIORITY REVERSED FROM v1.8.0:
    #
    #   PATH 1 (primary): Recipe Card / catalogv2.xml SCCM full driver pack.
    #     Lenovo's own "Recipe Card" web tool at
    #     https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html lists the
    #     same packs. Single .exe -> Inno Setup extract -> pnputil INFs.
    #     Simpler and more reliable than the per-descriptor consumer flow;
    #     covers most ThinkPad/ThinkCentre commercial models.
    #
    #   PATH 2 (fallback): Consumer catalog (download.lenovo.com/catalog/<MTM>_<OS>.xml).
    #     Only used when the machine type ISN'T listed in catalogv2.xml -
    #     i.e. consumer / ThinkBook / IdeaPad models that don't get Recipe
    #     Card packs (the use case that v1.8.0 originally added support for).
    #     Per-package descriptors, individual installers, the whole v1.8.x
    #     consumer-catalog stack stays in place untouched but only runs
    #     when path 1 has nothing.
    #
    #   PATH 3: open Lenovo's Recipe Card web tool for manual selection.
    # ------------------------------------------------------------------
    SetProgress 5
    $packUrl = Get-LenovoFullPackUrl -MachineType $machineType -OsCode $osCode
    if ($packUrl) {
        Log "Recipe Card / catalogv2.xml has full pack for '$machineType' - using it."
        return (Install-LenovoFullPack -PackUrl $packUrl -DriverRoot $DriverRoot)
    }
    Log "No Recipe Card pack for '$machineType' - trying consumer catalog (v1.8.0 path)..."
    SetProgress 15

    $consumerOK = Start-LenovoConsumerCatalogInstall -MachineType $machineType -OsCode $osCode -DriverRoot $DriverRoot
    if ($consumerOK) {
        Log "Consumer-catalog flow succeeded."
        return $true
    }

    Log "Neither Recipe Card nor consumer catalog had drivers for '$machineType'."
    if (-not $script:Headless) {
        Log "Opening Lenovo Recipe Card web tool for manual selection..."
        try { Start-Process "https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html" } catch {}
    } else {
        Log "  (headless: skipping web-tool launch)"
        Log "  Manual URL: https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html"
    }
    return $false
}

# =========================
# MICROSOFT (SURFACE)
#
# Uses Microsoft Download Center IDs to find the correct driver MSI for the model.
# Fetches the Download Center details page, extracts the MSI URL from the embedded
# window.__DLCDetails__ JSON blob (primary), falls back to href regex scraping.
# Selects best MSI by OS build (highest build <= device OS build) with secondary
# sort on driver version. Downloads MSI, extracts with msiexec /a, installs via pnputil.
#
# Surface Pro X (ARM/SQ processor): not supported.
#   Microsoft requires Windows Update for Pro X driver delivery.
#   No MSI is published; script opens support page and exits gracefully.
#
# To add new Surface models: add an entry to $SurfaceDownloadIds below
#   (value = numeric ID from details.aspx?id=XXXXXX).
#
# Dual-CPU models: a few Surface models (Surface Laptop 3, Surface Laptop 4)
#   report ONE WMI model string but ship as separate Intel and AMD hardware
#   with separate Download Center packs. For those, keep the Intel ID as the
#   default in $SurfaceDownloadIds and add an Intel/AMD pair to
#   $SurfaceCpuVariantIds (below the table). Get-CpuVendor decides which one
#   to use at runtime.
# =========================

$SurfaceDownloadIds = [ordered]@{
    # Surface Pro
    "Surface Pro 12"                          = "108199"   # verified
    "Surface Pro for Business (11th Edition)" = "108013"   # verified (Intel)
    "Surface Pro (11th Edition)"              = "106119"   # verified (Snapdragon)
    "Surface Pro 10 with 5G"                  = "106292"   # verified
    "Surface Pro 10"                          = "105947"   # verified
    "Surface Pro 9 with 5G"                   = "105941"   # verified
    "Surface Pro 9"                           = "104680"   # verified
    "Surface Pro 8"                           = "103503"   # verified
    "Surface Pro 7+"                          = "102633"   # verified
    "Surface Pro 7"                           = "100419"   # verified
    "Surface Pro 6"                           = "57514"    # verified
    "Surface Pro with LTE"                    = "56278"    # verified
    "Surface Pro (5th Gen)"                   = "55484"    # verified
    "Surface Pro 5"                           = "55484"    # verified
    "Surface Pro 4"                           = "49498"    # verified
    "Surface Pro 3"                           = "38826"    # verified
    "Surface Pro 2"                           = "49042"    # verified
    # Surface Laptop
    "Surface Laptop 7 with Intel"             = "108014"   # verified
    "Surface Laptop 7"                        = "106120"   # verified (Snapdragon)
    "Surface Laptop 6"                        = "105946"   # verified
    "Surface Laptop 5"                        = "104679"   # verified
    "Surface Laptop 4"                        = "102924"   # Intel default - AMD override via $SurfaceCpuVariantIds
    "Surface Laptop 3"                        = "100429"   # Intel default - AMD override via $SurfaceCpuVariantIds
    "Surface Laptop 2"                        = "57515"    # verified
    "Surface Laptop Studio 2"                 = "105610"   # verified
    "Surface Laptop Studio"                   = "103505"   # verified
    "Surface Laptop Go 3"                     = "105608"   # verified
    "Surface Laptop Go 2"                     = "104251"   # verified
    "Surface Laptop Go"                       = "102261"   # verified
    # Surface Book
    "Surface Book 3"                          = "101315"   # verified
    "Surface Book 2"                          = "56261"    # verified
    "Surface Book"                            = "49497"    # verified
    # Surface Go
    "Surface Go 4"                            = "105609"   # verified
    "Surface Go 3"                            = "103504"   # verified
    "Surface Go 2"                            = "101304"   # verified
    "Surface Go"                              = "57439"    # verified (Wi-Fi)
    # Surface Studio
    "Surface Studio 2+"                       = "104681"   # verified
    "Surface Studio 2"                        = "57593"    # verified
    "Surface Studio"                          = "54311"    # verified
}

# ------------------------------------------------------------------
# Dual-CPU Surface models: same WMI model string, two different
# hardware platforms (Intel vs AMD), each with its OWN Microsoft
# Download Center pack. WMI Win32_ComputerSystem.Model reports just
# "Surface Laptop 4" (or 3) regardless of CPU, so the base
# $SurfaceDownloadIds table can only hold one ID per model - it keeps
# the Intel ID as the default. This map supplies the per-CPU IDs;
# Get-CpuVendor picks the right one at runtime in
# Start-MicrosoftSurfaceDriverInstall. Keys MUST exactly match the
# corresponding $SurfaceDownloadIds key (regex/-ilike model match).
# IDs verified against live Microsoft Download Center pages (May 2026).
# ------------------------------------------------------------------
$SurfaceCpuVariantIds = @{
    "Surface Laptop 4" = @{ Intel = "102924"; AMD = "102923" }
    "Surface Laptop 3" = @{ Intel = "100429"; AMD = "100428" }
}

# Helper: determine the CPU vendor ("Intel", "AMD", or "Unknown").
# Used to pick the correct Microsoft driver pack for dual-CPU Surface
# models. Primary signal is Win32_Processor.Manufacturer, which on
# real hardware is the CPUID vendor string "GenuineIntel" or
# "AuthenticAMD" (this is what the v1.12.x device dump already logs).
# Falls back to the processor brand string (e.g. the Surface AMD parts
# report "AMD Ryzen 5 Microsoft Surface (R) Edition") and finally to
# $env:PROCESSOR_IDENTIFIER. Returns "Unknown" rather than guessing if
# nothing matches, so callers can keep the safe Intel default.
function Get-CpuVendor {
    $vendorRaw = ""
    $nameRaw   = ""
    try {
        $cpu = Get-CimInstance Win32_Processor -EA Stop | Select-Object -First 1
        if ($cpu) {
            $vendorRaw = "$($cpu.Manufacturer)".Trim()
            $nameRaw   = "$($cpu.Name)".Trim()
        }
    } catch {
        Log "  Get-CpuVendor: Win32_Processor query failed: $($_.Exception.Message)" -Level "warn" -Event "cpu_vendor_wmi_error"
    }

    $probe = "$vendorRaw $nameRaw $env:PROCESSOR_IDENTIFIER"

    if ($probe -match '(?i)AuthenticAMD|\bAMD\b|Ryzen')        { return "AMD" }
    if ($probe -match '(?i)GenuineIntel|\bIntel\b|Core\(TM\)') { return "Intel" }

    Log "  Get-CpuVendor: could not classify CPU (vendor='$vendorRaw' name='$nameRaw' env='$env:PROCESSOR_IDENTIFIER')" -Level "warn" -Event "cpu_vendor_unknown"
    return "Unknown"
}

# Helper: extract MSI contents using msiexec /a (admin install), then install INFs via pnputil.
# /a unpacks the MSI payload into a flat directory tree with INF + driver files,
# consistent with how Dell/HP/Lenovo packs are handled.
function Install-SurfaceMsi {
    param(
        [string]$MsiFile,
        [string]$DriverRoot,
        [string]$FileName
    )

    $extractPath = Join-Path $DriverRoot "Surface_Extracted"
    if (-not (Test-Path $extractPath)) {
        New-Item -Path $extractPath -ItemType Directory -Force | Out-Null
    }

    Log "Extracting Surface MSI: $FileName"
    Log "  msiexec /a `"$MsiFile`" /qn TARGETDIR=`"$extractPath`""
    $exHeaderLabel.Text          = "Extract MSI"
    $exSpinnerLabel.Text         = " " + $SpinnerFrames[0]
    $script:SpinnerIndex         = 0
    $exBar.Style                 = "Marquee"
    $exBar.MarqueeAnimationSpeed = 30
    SetExtract -Pct -1 -Label "Extracting MSI contents..."

    $psi                 = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName        = "msiexec.exe"
    $psi.Arguments       = "/a `"$MsiFile`" /qn TARGETDIR=`"$extractPath`""
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow  = $true
    $msiProc             = New-Object System.Diagnostics.Process
    $msiProc.StartInfo   = $psi
    $msiProc.Start() | Out-Null

    $elapsed = 0; $lastCount = 0; $stall = 0
    while (-not $msiProc.HasExited) {
        Start-Sleep -Milliseconds 700
        $elapsed += 0.7
        $count = if (Test-Path $extractPath) {
            (Get-ChildItem $extractPath -Recurse -EA SilentlyContinue).Count
        } else { 0 }
        if ($count -gt $lastCount) { $stall = 0; $lastCount = $count } else { $stall++ }
        $mins = [int]($elapsed / 60); $secs = [int]($elapsed % 60)
        SetExtract -Pct -1 -Label "Extracting... $count files  ($mins`m $secs`s)"
        Step-ExSpinner
        Step-OverallSpinner
        [System.Windows.Forms.Application]::DoEvents()
        if ($script:CancelRequested) {
            Log "  MSI extraction cancelled by user."
            try { $msiProc.Kill() } catch {}
            return $false
        }
        if ($elapsed -gt 600) {
            Log "  WARNING: MSI extraction exceeded 10 minutes - aborting."
            try { $msiProc.Kill() } catch {}
            return $false
        }
    }

    $exitCode = $msiProc.ExitCode
    Log "  msiexec /a exit code: $exitCode"

    if ($exitCode -ne 0) {
        Log "  MSI extraction failed (exit $exitCode)."
        SetExtract -Pct 0 -Label "MSI extraction failed (exit $exitCode)"
        Stop-ExSpinner -Success $false
        return $false
    }

    $finalCount = if (Test-Path $extractPath) {
        (Get-ChildItem $extractPath -Recurse -EA SilentlyContinue).Count
    } else { 0 }
    Log "  MSI extracted: $finalCount files in $extractPath"
    SetExtract -Pct 60 -Label "Extracted $finalCount files - installing INFs..."
    Play-Sound -Event "ExtractComplete"

    $exHeaderLabel.Text = "Install INFs"
    SetProgress 60
    return (Install-DriversFromPath -BasePath $extractPath)
}

# Helper: show a modal dialog with a sorted ComboBox of all known Surface models.
# Returns the selected model name string, or $null if the user cancels.
# Called when WMI reports an unrecognised/generic manufacturer (e.g. "OEMBY").
function Show-SurfaceModelPicker {
    param([string]$DetectedManufacturer = "", [string]$DetectedModel = "")

    $pickerForm                  = New-Object System.Windows.Forms.Form
    $pickerForm.Text             = "Select Surface Model"
    $pickerForm.Size             = New-Object System.Drawing.Size(420, 200)
    $pickerForm.StartPosition    = "CenterParent"
    $pickerForm.FormBorderStyle  = "FixedDialog"
    $pickerForm.MaximizeBox      = $false
    $pickerForm.MinimizeBox      = $false
    $pickerForm.BackColor        = [System.Drawing.Color]::FromArgb(245, 245, 245)

    $lbl           = New-Object System.Windows.Forms.Label
    $lbl.AutoSize  = $false
    $lbl.Size      = New-Object System.Drawing.Size(380, 36)
    $lbl.Location  = New-Object System.Drawing.Point(16, 12)
    $lbl.Font      = $FontUI
    $lbl.Text      = "Manufacturer '$DetectedManufacturer' was not recognised.`nSelect the Surface model to install drivers for:"
    $lbl.UseCompatibleTextRendering = $false
    $pickerForm.Controls.Add($lbl)

    $combo               = New-Object System.Windows.Forms.ComboBox
    $combo.Size          = New-Object System.Drawing.Size(380, 24)
    $combo.Location      = New-Object System.Drawing.Point(16, 56)
    $combo.Font          = $FontUI
    $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    # Populate sorted model list from the global $SurfaceDownloadIds hashtable
    $sortedModels = $SurfaceDownloadIds.Keys | Sort-Object
    foreach ($m in $sortedModels) { $combo.Items.Add($m) | Out-Null }
    $combo.SelectedIndex = 0
    $pickerForm.Controls.Add($combo)

    $okBtn            = New-Object System.Windows.Forms.Button
    $okBtn.Text       = "Install Drivers"
    $okBtn.Size       = New-Object System.Drawing.Size(140, 32)
    $okBtn.Location   = New-Object System.Drawing.Point(130, 96)
    $okBtn.Font       = $FontUIBold
    $okBtn.BackColor  = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $okBtn.ForeColor  = [System.Drawing.Color]::White
    $okBtn.FlatStyle  = "Flat"
    $okBtn.FlatAppearance.BorderSize = 0
    $okBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $pickerForm.Controls.Add($okBtn)

    $cancelBtn            = New-Object System.Windows.Forms.Button
    $cancelBtn.Text       = "Cancel"
    $cancelBtn.Size       = New-Object System.Drawing.Size(80, 32)
    $cancelBtn.Location   = New-Object System.Drawing.Point(286, 96)
    $cancelBtn.Font       = $FontUIBold
    $cancelBtn.BackColor  = [System.Drawing.Color]::FromArgb(160, 160, 160)
    $cancelBtn.ForeColor  = [System.Drawing.Color]::White
    $cancelBtn.FlatStyle  = "Flat"
    $cancelBtn.FlatAppearance.BorderSize = 0
    $cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $pickerForm.Controls.Add($cancelBtn)

    $pickerForm.AcceptButton = $okBtn
    $pickerForm.CancelButton = $cancelBtn

    $result = $pickerForm.ShowDialog($form)
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $combo.SelectedItem
    }
    return $null
}

function Start-MicrosoftSurfaceDriverInstall {
    param([string]$DriverRoot, [string]$ModelName)

    Log "=== MICROSOFT SURFACE: Starting automated driver install ==="
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."

    try { $script:AnalyticsSerial = (Get-CimInstance Win32_BIOS).SerialNumber.Trim() } catch {}

    $osBuild = 0
    try {
        $osBuild = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber
        Log "OS Build: $osBuild"
    } catch { Log "Could not read OS build: $($_.Exception.Message)" }

    # Read SystemSKU for logging purposes
    $systemSKU = ""
    try {
        $systemSKU = (Get-CimInstance Win32_ComputerSystem).SystemSKUNumber.Trim()
        Log "SystemSKU: $systemSKU"
    } catch { Log "  WARNING: Could not read SystemSKU." }

    if (-not (Test-Path $DriverRoot)) { New-Item -Path $DriverRoot -ItemType Directory -Force | Out-Null }

    # --------------------------------------------------
    # Download Center lookup: match model -> page ID -> fetch page ->
    # extract MSI URL from __DLCDetails__ JSON blob -> download -> msiexec /a -> pnputil
    # --------------------------------------------------
    Log "Looking up Surface driver pack via Download Center..."
    SetExtract -Pct 5 -Label "Looking up Download Center ID..."
    SetProgress 10

    $pageId    = $null
    $matchedKey = $null
    foreach ($entry in $SurfaceDownloadIds.GetEnumerator()) {
        if ($ModelName -match [regex]::Escape($entry.Key)) {
            $pageId     = $entry.Value
            $matchedKey = $entry.Key
            Log "Matched model '$($entry.Key)' -> Download Center ID: $pageId"
            break
        }
    }
    if (-not $pageId) {
        foreach ($entry in $SurfaceDownloadIds.GetEnumerator()) {
            if ($ModelName -ilike "*$($entry.Key)*") {
                $pageId     = $entry.Value
                $matchedKey = $entry.Key
                Log "Fuzzy-matched '$($entry.Key)' -> Download Center ID: $pageId"
                break
            }
        }
    }

    if (-not $pageId) {
        Log "No Download Center entry found for: '$ModelName'"
        Log "Note: Surface Pro X uses Windows Update only (ARM - no MSI available)."
        Log "Opening Surface driver downloads page for manual selection..."
        Start-Process "https://support.microsoft.com/en-us/surface/drivers-firmware/download-drivers-and-firmware-for-surface-pro"
        return $false
    }

    # --------------------------------------------------
    # CPU-aware override for dual-CPU models (Surface Laptop 3 / 4).
    # WMI reports a single model string for both the Intel and AMD
    # hardware, but Microsoft publishes a SEPARATE driver pack for
    # each. The base table defaulted to the Intel ID; if this model
    # has Intel/AMD variants, switch to the pack matching the CPU.
    # --------------------------------------------------
    if ($matchedKey -and $SurfaceCpuVariantIds.ContainsKey($matchedKey)) {
        $cpuVendor = Get-CpuVendor
        $variants  = $SurfaceCpuVariantIds[$matchedKey]
        Log "Model '$matchedKey' has CPU-specific driver packs. Detected CPU vendor: $cpuVendor"
        if ($cpuVendor -ne "Unknown" -and $variants.ContainsKey($cpuVendor)) {
            $variantId = $variants[$cpuVendor]
            if ($variantId -ne $pageId) {
                Log "  $cpuVendor CPU detected -> overriding Download Center ID $pageId -> $variantId ($matchedKey, $cpuVendor pack)" -Level "info" -Event "surface_cpu_variant_override" -Context @{ model = $matchedKey; cpu = $cpuVendor; from_id = $pageId; to_id = $variantId }
                $pageId = $variantId
            } else {
                Log "  $cpuVendor CPU detected -> already on the correct pack (ID $pageId)"
            }
        } else {
            Log "  WARNING: '$matchedKey' ships in Intel and AMD variants but the CPU vendor could not be determined. Falling back to the default pack (ID $pageId). If this machine is AMD and the install finds no matching drivers, re-run on the device (WMI CPU detection should work there) or download the AMD pack manually from microsoft.com/download." -Level "warn" -Event "surface_cpu_variant_unresolved" -Context @{ model = $matchedKey; default_id = $pageId }
        }
    }

    $detailsUrl  = "https://www.microsoft.com/en-us/download/details.aspx?id=$pageId"
    $detailsFile = Join-Path $env:TEMP "surface_dl_page_$pageId.html"
    Remove-Item $detailsFile -EA SilentlyContinue

    Log "Fetching download details page (ID=$pageId)..."
    SetExtract -Pct 5 -Label "Fetching download page..."
    SetProgress 15

    $psi                        = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = "curl.exe"
    $psi.Arguments              = "--silent --location --max-time 30 --connect-timeout 15 " +
                                  "--user-agent `"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36`" " +
                                  "--output `"$detailsFile`" `"$detailsUrl`""
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $psi.RedirectStandardError  = $true
    $fetchProc2                 = New-Object System.Diagnostics.Process
    $fetchProc2.StartInfo       = $psi
    $fetchProc2.Start() | Out-Null

    $start2 = Get-Date
    while (-not $fetchProc2.HasExited) {
        Start-Sleep -Milliseconds 400
        Step-AllSpinners
        if (((Get-Date) - $start2).TotalSeconds -gt 35) {
            Log "  Timeout fetching download page."
            try { $fetchProc2.Kill() } catch {}
            break
        }
    }
    $fetchProc2.WaitForExit()

    if (-not (Test-Path $detailsFile) -or (Get-Item $detailsFile).Length -lt 1000) {
        Log "Failed to fetch download details page for ID=$pageId"
        Log "Opening page manually: $detailsUrl"
        Start-Process $detailsUrl
        return $false
    }
    if (Test-Cancelled) { return $false }

    Log "Parsing download page for MSI links..."
    SetExtract -Pct 20 -Label "Parsing download page..."
    $pageHtml = [System.IO.File]::ReadAllText($detailsFile)

    # Helper: parse driver version from MSI filename e.g. "26.040.371.0" -> [Version]
    function Parse-MsiDriverVersion {
        param([string]$FileName)
        $m = [regex]::Match($FileName, '_(\d+\.\d+\.\d+\.\d+)\.msi$')
        if ($m.Success) { try { return [Version]$m.Groups[1].Value } catch {} }
        return [Version]"0.0.0.0"
    }

    $msiCandidates = [System.Collections.Generic.List[object]]::new()
    $seen          = @{}

    function Add-MsiCandidate {
        param([string]$Url)
        if (-not $Url -or $seen.ContainsKey($Url)) { return }
        $seen[$Url] = $true
        $fileName   = [System.IO.Path]::GetFileName(([System.Uri]$Url).LocalPath)
        $buildMatch = [regex]::Match($fileName, '_Win\d+_(\d{5})_')
        $msiOsBuild = if ($buildMatch.Success) { [int]$buildMatch.Groups[1].Value } else { 0 }
        Log "  MSI: $fileName  (target build: $(if ($msiOsBuild -gt 0) { $msiOsBuild } else { 'unknown' }))"
        $msiCandidates.Add([PSCustomObject]@{ Url = $Url; FileName = $fileName; OsBuild = $msiOsBuild })
    }

    # PRIMARY: extract from window.__DLCDetails__ JSON blob (contains the primary download file)
    $dlcMatch = [regex]::Match($pageHtml, 'window\.__DLCDetails__\s*=\s*(\{.+?"detailsId":.+?\})\s*</script>')
    if ($dlcMatch.Success) {
        try {
            $dlcJson = $dlcMatch.Groups[1].Value | ConvertFrom-Json
            foreach ($f in $dlcJson.dlcDetailsView.downloadFile) {
                if ($f.url -match '\.msi$') {
                    Log "  DLCDetails primary: $($f.name)"
                    Add-MsiCandidate -Url $f.url
                }
            }
        } catch { Log "  WARNING: __DLCDetails__ parse error: $($_.Exception.Message)" }
    }

    # FALLBACK: regex scrape all download.microsoft.com MSI hrefs (catches archive links too)
    foreach ($m in [regex]::Matches($pageHtml, 'href="(https://download\.microsoft\.com/[^"]+\.msi)"')) {
        Add-MsiCandidate -Url $m.Groups[1].Value
    }
    if ($msiCandidates.Count -eq 0) {
        foreach ($m in [regex]::Matches($pageHtml, '(https://download\.microsoft\.com/[^\s"<>]+\.msi)')) {
            Add-MsiCandidate -Url $m.Groups[1].Value
        }
    }

    if ($msiCandidates.Count -eq 0) {
        Log "No MSI links found on page ID=$pageId - page may require JavaScript."
        Log "Opening page for manual download: $detailsUrl"
        Start-Process $detailsUrl
        return $false
    }
    Log "  Found $($msiCandidates.Count) MSI candidate(s)."

    # Select best: OSBuild desc (primary), driver version desc (secondary)
    $chosen = $null
    if ($osBuild -gt 0) {
        $eligible = @($msiCandidates | Where-Object { $_.OsBuild -gt 0 -and $_.OsBuild -le $osBuild })
        if ($eligible.Count -eq 0) {
            $eligible = @($msiCandidates | Where-Object { $_.OsBuild -gt 0 } | Sort-Object OsBuild)
        }
        if ($eligible.Count -eq 0) { $eligible = @($msiCandidates) }
        $chosen = ($eligible | Sort-Object `
            @{ E = { $_.OsBuild };                         Descending = $true },
            @{ E = { Parse-MsiDriverVersion $_.FileName }; Descending = $true }
        )[0]
        Log "Selected MSI: $($chosen.FileName)  [OSBuild=$($chosen.OsBuild)]"
    } else {
        $chosen = $msiCandidates[0]
        Log "OS build unknown - using first MSI found: $($chosen.FileName)"
    }

    Log "MSI URL: $($chosen.Url)"
    SetExtract -Pct 40 -Label "MSI selected: $($chosen.FileName)"
    SetProgress 25

    $msiFile = Join-Path $DriverRoot $chosen.FileName
    SetProgress 30

    if (-not (Invoke-CurlDownload -Url $chosen.Url -OutFile $msiFile)) {
        Log "Surface MSI download failed."
        return $false
    }
    if (Test-Cancelled) { return $false }
    SetProgress 55

    return (Install-SurfaceMsi -MsiFile $msiFile -DriverRoot $DriverRoot -FileName $chosen.FileName)
}

# =========================
# DEVICE INFO DUMP
# =========================
function Write-DeviceInfo {
    Log "============================================"
    Log "  DEVICE INFORMATION DUMP"
    Log "============================================"

    Log "-- System --"
    try {
        $cs = Get-CimInstance Win32_ComputerSystem
        Log "  Manufacturer       : $($cs.Manufacturer)"
        Log "  Model              : $($cs.Model)"
        Log "  SystemSKUNumber    : $($cs.SystemSKUNumber)"
        Log "  SystemFamily       : $($cs.SystemFamily)"
        Log "  PCSystemType       : $($cs.PCSystemType)"
        Log "  TotalPhysRAM (GB)  : $([math]::Round($cs.TotalPhysicalMemory/1GB,2))"
        Log "  Domain             : $($cs.Domain)"
        Log "  UserName           : $($cs.UserName)"
    } catch { Log "  [Win32_ComputerSystem ERROR] $($_.Exception.Message)" }

    try {
        $csp = Get-CimInstance Win32_ComputerSystemProduct
        Log "  CSProduct.Name     : $($csp.Name)"
        Log "  CSProduct.Version  : $($csp.Version)"
        Log "  CSProduct.UUID     : $($csp.UUID)"
        Log "  CSProduct.Vendor   : $($csp.Vendor)"
    } catch { Log "  [Win32_ComputerSystemProduct ERROR] $($_.Exception.Message)" }

    Log "-- BIOS --"
    try {
        $bios = Get-CimInstance Win32_BIOS
        Log "  SerialNumber       : $($bios.SerialNumber)"
        Log "  SMBIOSBIOSVersion  : $($bios.SMBIOSBIOSVersion)"
        Log "  ReleaseDate        : $($bios.ReleaseDate)"
        Log "  Manufacturer       : $($bios.Manufacturer)"
        Log "  Name               : $($bios.Name)"
        Log "  Version            : $($bios.Version)"
    } catch { Log "  [Win32_BIOS ERROR] $($_.Exception.Message)" }

    Log "-- Baseboard --"
    try {
        $bb = Get-CimInstance Win32_BaseBoard
        Log "  Product            : $($bb.Product)"
        Log "  Manufacturer       : $($bb.Manufacturer)"
        Log "  SerialNumber       : $($bb.SerialNumber)"
        Log "  Version            : $($bb.Version)"
    } catch { Log "  [Win32_BaseBoard ERROR] $($_.Exception.Message)" }

    Log "-- Enclosure --"
    try {
        $enc = Get-CimInstance Win32_SystemEnclosure
        Log "  ChassisTypes       : $($enc.ChassisTypes -join ',')"
        Log "  SMBIOSAssetTag     : $($enc.SMBIOSAssetTag)"
        Log "  SerialNumber       : $($enc.SerialNumber)"
        Log "  Manufacturer       : $($enc.Manufacturer)"
    } catch { Log "  [Win32_SystemEnclosure ERROR] $($_.Exception.Message)" }

    Log "-- Operating System --"
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        Log "  Caption            : $($os.Caption)"
        Log "  Version            : $($os.Version)"
        Log "  BuildNumber        : $($os.BuildNumber)"
        Log "  OSArchitecture     : $($os.OSArchitecture)"
        Log "  SystemDrive        : $($os.SystemDrive)"
        Log "  WindowsDirectory   : $($os.WindowsDirectory)"
        Log "  InstallDate        : $($os.InstallDate)"
        Log "  LastBootUpTime     : $($os.LastBootUpTime)"
    } catch { Log "  [Win32_OperatingSystem ERROR] $($_.Exception.Message)" }

    Log "-- Processor --"
    try {
        $cpus = Get-CimInstance Win32_Processor
        foreach ($cpu in $cpus) {
            Log "  Name               : $($cpu.Name)"
            Log "  DeviceID           : $($cpu.DeviceID)"
            Log "  Manufacturer       : $($cpu.Manufacturer)"
            Log "  MaxClockSpeed      : $($cpu.MaxClockSpeed) MHz"
            Log "  NumberOfCores      : $($cpu.NumberOfCores)"
            Log "  NumberOfLogical    : $($cpu.NumberOfLogicalProcessors)"
            Log "  ProcessorId        : $($cpu.ProcessorId)"
        }
    } catch { Log "  [Win32_Processor ERROR] $($_.Exception.Message)" }

    Log "-- Video Controller --"
    try {
        $gpus = Get-CimInstance Win32_VideoController
        foreach ($gpu in $gpus) {
            Log "  Name               : $($gpu.Name)"
            Log "  DeviceID           : $($gpu.DeviceID)"
            Log "  AdapterRAM         : $([math]::Round($gpu.AdapterRAM/1MB))MB"
            Log "  DriverVersion      : $($gpu.DriverVersion)"
            Log "  DriverDate         : $($gpu.DriverDate)"
            Log "  VideoModeDesc      : $($gpu.VideoModeDescription)"
            Log "  ---"
        }
    } catch { Log "  [Win32_VideoController ERROR] $($_.Exception.Message)" }

    Log "-- Network Adapters --"
    try {
        $nics = Get-CimInstance Win32_NetworkAdapter | Where-Object { $_.PhysicalAdapter -eq $true }
        foreach ($nic in $nics) {
            Log "  Name               : $($nic.Name)"
            Log "  MACAddress         : $($nic.MACAddress)"
            Log "  AdapterType        : $($nic.AdapterType)"
            Log "  ---"
        }
    } catch { Log "  [Win32_NetworkAdapter ERROR] $($_.Exception.Message)" }

    Log "-- Disk Drives --"
    try {
        $disks = Get-CimInstance Win32_DiskDrive
        foreach ($d in $disks) {
            $sizeGB = if ($d.Size) { [math]::Round($d.Size/1GB,1) } else { "?" }
            Log "  Model              : $($d.Model)"
            Log "  SerialNumber       : $($d.SerialNumber)"
            Log "  InterfaceType      : $($d.InterfaceType)"
            Log "  Size               : $sizeGB GB"
            Log "  MediaType          : $($d.MediaType)"
            Log "  ---"
        }
    } catch { Log "  [Win32_DiskDrive ERROR] $($_.Exception.Message)" }

    Log "-- Physical Memory --"
    try {
        $dimms = Get-CimInstance Win32_PhysicalMemory
        foreach ($d in $dimms) {
            $sz = if ($d.Capacity) { [math]::Round($d.Capacity/1GB,1) } else { "?" }
            Log "  BankLabel          : $($d.BankLabel)"
            Log "  DeviceLocator      : $($d.DeviceLocator)"
            Log "  Capacity           : $sz GB"
            Log "  Speed              : $($d.Speed) MHz"
            Log "  Manufacturer       : $($d.Manufacturer)"
            Log "  PartNumber         : $($d.PartNumber)"
            Log "  ---"
        }
    } catch { Log "  [Win32_PhysicalMemory ERROR] $($_.Exception.Message)" }

    Log "-- Battery --"
    try {
        $batts = Get-CimInstance Win32_Battery
        if ($batts) {
            foreach ($b in $batts) {
                Log "  Name               : $($b.Name)"
                Log "  EstimatedRuntime   : $($b.EstimatedRunTime) min"
                Log "  BatteryStatus      : $($b.BatteryStatus)"
                Log "  DesignCapacity     : $($b.DesignCapacity) mWh"
                Log "  FullChargeCapacity : $($b.FullChargeCapacity) mWh"
            }
        } else { Log "  (no battery detected - desktop?)" }
    } catch { Log "  [Win32_Battery ERROR] $($_.Exception.Message)" }

    Log "-- PnP Devices (problem state / no driver) --"
    try {
        $problem = Get-CimInstance Win32_PnPEntity |
            Where-Object { $_.ConfigManagerErrorCode -ne 0 } |
            Select-Object -Property Name, DeviceID, ConfigManagerErrorCode
        if ($problem) {
            foreach ($p in $problem) {
                Log "  [ERR $($p.ConfigManagerErrorCode)] $($p.Name)"
                Log "    DeviceID: $($p.DeviceID)"
            }
        } else { Log "  (none - all PnP devices have drivers)" }
    } catch { Log "  [Win32_PnPEntity ERROR] $($_.Exception.Message)" }

    Log "-- pnputil driver store (first 20 OEM INFs) --"
    try {
        $pnpOut = & pnputil /enum-drivers 2>&1 | Select-Object -First 60
        foreach ($line in $pnpOut) { Log "  $line" }
    } catch { Log "  [pnputil ERROR] $($_.Exception.Message)" }

    Log "-- Environment --"
    Log "  TEMP               : $env:TEMP"
    Log "  COMPUTERNAME       : $env:COMPUTERNAME"
    Log "  USERNAME           : $env:USERNAME"
    Log "  PROCESSOR_ARCH     : $env:PROCESSOR_ARCHITECTURE"
    Log "  PS Version         : $($PSVersionTable.PSVersion)"
    Log "  curl.exe path      : $($(Get-Command curl.exe -EA SilentlyContinue).Source)"

    Log "-- Disk Space --"
    try {
        $c = Get-PSDrive C -EA Stop
        Log "  C: Used (GB)       : $([math]::Round($c.Used/1GB,2))"
        Log "  C: Free (GB)       : $([math]::Round($c.Free/1GB,2))"
    } catch { Log "  [PSDrive C: ERROR] $($_.Exception.Message)" }

    Log "============================================"
    Log "  END DEVICE INFORMATION DUMP"
    Log "============================================"
}

# =========================
# MAIN
# =========================
function Start-Install {

    $script:CancelRequested          = $false
    $script:AnalyticsInfCount        = 0
    $script:AnalyticsDownloadMB      = 0.0
    $script:AnalyticsSerial          = ""
    $script:AnalyticsManufacturer    = ""
    $script:AnalyticsModel           = ""
    $script:AnalyticsOsVersion       = ""
    $script:AnalyticsOsBuild         = 0
    $script:AnalyticsMissingBefore   = -1
    $script:AnalyticsMissingAfter    = -1
    $script:AnalyticsMissingBeforeList = New-Object System.Collections.Generic.List[string]
    $script:AnalyticsMissingAfterList  = New-Object System.Collections.Generic.List[string]
    $script:AnalyticsInstalledDrivers = New-Object System.Collections.Generic.List[string]
    $script:AnalyticsDownloadUrls    = New-Object System.Collections.Generic.List[hashtable]   # v1.13.1
    $script:AnalyticsStartTime       = Get-Date
    $script:7zInstalled              = $false
    $script:Headless                 = [bool]$Headless
    # v1.10.0 - re-anchor mode flags on each Start-Install invocation. The
    # script-scope assignments at top-of-file cover the irm|iex path; this
    # covers the GUI-rerun-via-button-click path where users might re-launch
    # without restarting PowerShell.
    $script:Silent      = [bool]$Silent
    $script:TestMode    = [bool]$TestMode
    $script:Diagnostic  = [bool]$Diagnostic
    $script:NoAnalytics = [bool]$NoAnalytics
    $script:MaxParallelDownloads = $MaxParallelDownloads  # v1.11.0

    $id  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pri = New-Object Security.Principal.WindowsPrincipal($id)
    if (-not $pri.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        if ($script:Headless) {
            Write-Error "ERROR: Script must be run as Administrator."
            exit 1
        }
        Log "Not running as admin - re-launching elevated..."
        Start-Process powershell `
            "-WindowStyle Hidden -ExecutionPolicy Bypass -Command `"irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1 | iex`"" `
            -Verb RunAs
        $form.Close()
        exit
    }

    Log "Driver Installer v$SCRIPT_VERSION" -Level "info" -Event "run_start" -Context @{
        log_file       = $LogFile
        events_file    = $EventsLogFile
        analytics_file = $AnalyticsFile
        report_file    = $ReportFile
        headless       = $script:Headless
        silent         = $script:Silent
        test_mode      = $script:TestMode
        diagnostic     = $script:Diagnostic
        no_analytics   = $script:NoAnalytics
        skip_install   = [bool]$SkipInstall
        skip_cleanup   = [bool]$SkipCleanup
        max_parallel_downloads = $script:MaxParallelDownloads
    }
    Log "Log: $LogFile"
    # Surface the active modes prominently - matters most when someone is
    # reading the log later trying to work out why an install "didn't happen".
    $modeFlags = @()
    if ($script:Silent)      { $modeFlags += "SILENT" }
    if ($script:TestMode)    { $modeFlags += "TEST-MODE (dry-run)" }
    if ($script:Diagnostic)  { $modeFlags += "DIAGNOSTIC" }
    if ($script:NoAnalytics) { $modeFlags += "NO-ANALYTICS" }
    if ($SkipInstall)        { $modeFlags += "SKIP-INSTALL" }
    if ($SkipCleanup)        { $modeFlags += "SKIP-CLEANUP" }
    if ($modeFlags.Count -gt 0) { Log "Mode flags  : $($modeFlags -join ', ')" }
    Log "--------------------------------------------"

    Play-Sound -Event "Start"

    $cs = Get-CimInstance Win32_ComputerSystem

    # Use param overrides if provided, otherwise read from WMI
    $manufacturer = if ($Manufacturer) { $Manufacturer } else { $cs.Manufacturer.Trim() }
    $model        = if ($Model)        { $Model }        else { $cs.Model.Trim() }

    $script:AnalyticsManufacturer = $manufacturer
    $script:AnalyticsModel        = $model
    try {
        $osObj = Get-CimInstance Win32_OperatingSystem
        $script:AnalyticsOsVersion = $osObj.Version
        $script:AnalyticsOsBuild   = [int]$osObj.BuildNumber
    } catch {}

    if (-not $script:Headless) {
        try { [System.Windows.Forms.Clipboard]::SetText($model) } catch {}
        # v1.11.0 - title stays static; the detected model lives in the subtitle row
        # for a cleaner header hierarchy. The window title bar now matches the subtitle
        # format for consistency.
        $subtitle.Text  = "$manufacturer  ·  $model"
        $form.Text      = "Driver Installer  -  $manufacturer  ·  $model"
    }

    $overrideNote = if ($Manufacturer -or $Model) { "  [OVERRIDDEN via param]" } else { "  (from WMI)" }
    Log "Manufacturer : $manufacturer$overrideNote"
    Log "Model        : $model$overrideNote"

    Write-DeviceInfo
    Write-VerboseDiagnostics
    SetProgress 5

    # Snapshot missing drivers before install
    Log "Checking for devices with missing drivers..."
    $script:AnalyticsMissingBefore = Get-MissingDriverCount
    # v1.11.0 - also capture names for the analytics sheet
    $beforeNames = Get-MissingDriverNames
    foreach ($n in $beforeNames) { $script:AnalyticsMissingBeforeList.Add($n) | Out-Null }
    Log "Missing drivers BEFORE install: $($script:AnalyticsMissingBefore)"
    if ($beforeNames.Count -gt 0) {
        # Mirror the at-a-glance list into the log for the human reader too,
        # so the .log and the analytics sheet tell the same story.
        Log "Missing devices (before):"
        foreach ($n in $beforeNames) { Log "  - $n" }
    }

    # Offer to skip if no missing drivers detected
    if ($script:AnalyticsMissingBefore -eq 0) {
        if ($script:Headless) {
            Log "No missing drivers detected - continuing anyway (headless mode)."
        } else {
            $skipResult = [System.Windows.Forms.MessageBox]::Show(
                "No missing drivers were detected on this device.`n`nRun driver installation anyway?",
                "No Missing Drivers", "YesNo", "Question"
            )
            if ($skipResult -eq [System.Windows.Forms.DialogResult]::No) {
                Log "User chose to skip - no missing drivers detected."
                SetProgress 100
                SetDownload -Pct 100 -Label "Skipped - no missing drivers"
                SetExtract  -Pct 100 -Label "Skipped - no missing drivers"
                $script:AnalyticsMissingAfter = 0
                Send-AnalyticsEvent -Result "success"
                Play-Sound -Event "Success"
                Set-ButtonIdle
                return
            }
            Log "User chose to run anyway despite no missing drivers."
        }
    }

    # User confirmed (or drivers were missing) - lock UI and begin
    Set-ButtonRunning
    SetProgress 0
    SetDownload -Pct 0 -Label "Waiting..."
    SetExtract  -Pct 0 -Label "Waiting..."
    $exHeaderLabel.Text       = "Extract"
    $dlSpinnerLabel.Text      = ""
    $exSpinnerLabel.Text      = ""
    $overallSpinnerLabel.Text = ""
    $script:SpinnerIndex      = 0

    # v1.10.0 - TEST MODE SHORT-CIRCUIT
    # Bail before any system-mutating work happens. We've already done all the
    # safe stuff: detected manufacturer, snapshotted missing drivers, written
    # diagnostics. That's enough to tell the operator what a real run would
    # have done. NB we deliberately fire this AFTER the missing-driver snapshot
    # so the analytics row shows the real before/after gap (after == before in
    # test mode, since we touched nothing).
    if ($script:TestMode) {
        Log "============================================" -Level "info"
        Log "  TEST MODE - no downloads, no installs"
        Log "============================================"
        Log "  Would dispatch to vendor handler: $manufacturer / $model"
        Log "  Missing drivers (would attempt to resolve): $($script:AnalyticsMissingBefore)"
        Log "  Driver root would be: $($DriverRoot)"
        $script:AnalyticsMissingAfter = $script:AnalyticsMissingBefore
        SetProgress 100
        SetDownload -Pct 100 -Label "Test mode - skipped"
        SetExtract  -Pct 100 -Label "Test mode - skipped"
        Send-AnalyticsEvent -Result "testmode"
        Write-HtmlReport -Result "testmode"
        Play-Sound -Event "Success"
        if ($script:Headless) {
            Write-Host "TEST MODE: dry-run complete. See log: $LogFile"
            Write-Host "Report:    $ReportFile"
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Test mode complete.`n`nNo downloads or installs were performed.`n`nReport: $ReportFile",
                "Test Mode Complete", "OK", "Information"
            ) | Out-Null
            Set-ButtonIdle
        }
        return
    }

    # Install 7-Zip for Dell and HP extraction (not needed for Lenovo or Surface)
    if ($manufacturer -match "Dell|HP|Hewlett") {
        Log "Installing 7-Zip for fast extraction..."
        if (-not (Install-7Zip)) {
            Log "WARNING: 7-Zip unavailable - will fall back to vendor extractor."
        }
    }

    $driverRoot = $DriverRoot
    $success    = $false

    if ($manufacturer -match "Dell") {
        if (-not (Assert-Curl)) { Send-AnalyticsEvent -Result "failure"; Set-ButtonIdle; return }
        $success = Start-DellDriverInstall -DriverRoot $driverRoot -ModelName $model
    } elseif ($manufacturer -match "HP|Hewlett") {
        if (-not (Assert-Curl)) { Send-AnalyticsEvent -Result "failure"; Set-ButtonIdle; return }
        $success = Start-HpDriverInstall -DriverRoot $driverRoot -ModelName $model
    } elseif ($manufacturer -match "Lenovo") {
        if (-not (Assert-Curl)) { Send-AnalyticsEvent -Result "failure"; Set-ButtonIdle; return }
        $success = Start-LenovoDriverInstall -DriverRoot $driverRoot
    } elseif ($manufacturer -match "Microsoft") {
        if (-not (Assert-Curl)) { Send-AnalyticsEvent -Result "failure"; Set-ButtonIdle; return }
        $success = Start-MicrosoftSurfaceDriverInstall -DriverRoot $driverRoot -ModelName $model
    } else {
        # Unknown manufacturer (e.g. "OEMBY", blank, generic OEM string).
        # In headless mode: if -Model was explicitly passed and looks like a Surface, run it.
        # Otherwise error out with instructions to pass -Manufacturer/-Model explicitly.
        # In GUI mode: show a Surface model picker dialog and proceed if user selects one.
        Log "Unrecognised manufacturer: '$manufacturer'"
        Log "  Model reported as: '$model'"

        if ($script:Headless) {
            # Allow headless override: if -Model param contains a known Surface name, proceed.
            $headlessSurfaceMatch = $SurfaceDownloadIds.Keys | Where-Object { $model -ilike "*$_*" } | Select-Object -First 1
            if ($headlessSurfaceMatch) {
                Log "  Headless: model '$model' looks like a Surface - proceeding as Microsoft Surface."
                if (-not (Assert-Curl)) { Send-AnalyticsEvent -Result "failure"; return }
                $success = Start-MicrosoftSurfaceDriverInstall -DriverRoot $driverRoot -ModelName $model
            } else {
                Write-Host "ERROR: Manufacturer '$manufacturer' not recognised and model '$model' doesn't match a known Surface."
                Write-Host "Pass -Manufacturer and -Model explicitly, e.g.:"
                Write-Host "  -Manufacturer Microsoft -Model `"Surface Pro 9`""
                Send-AnalyticsEvent -Result "failure"
                return
            }
        } else {
            Log "  Showing Surface model picker..."
            Set-ButtonIdle
            $pickedModel = Show-SurfaceModelPicker -DetectedManufacturer $manufacturer -DetectedModel $model
            if (-not $pickedModel) {
                Log "  User cancelled Surface model selection."
                Send-AnalyticsEvent -Result "cancelled"
                return
            }
            Log "  User selected: $pickedModel"
            # Update analytics model to the picked Surface model
            $script:AnalyticsManufacturer = "Microsoft"
            $script:AnalyticsModel        = $pickedModel
            $subtitle.Text = "Microsoft  ·  $pickedModel"
            $form.Text     = "Driver Installer  -  Microsoft  ·  $pickedModel"
            Set-ButtonRunning
            if (-not (Assert-Curl)) { Send-AnalyticsEvent -Result "failure"; Set-ButtonIdle; return }
            $success = Start-MicrosoftSurfaceDriverInstall -DriverRoot $driverRoot -ModelName $pickedModel
        }
    }

    # Remove 7-Zip before cleanup - must happen before C:\DRIVERS is deleted
    if ($script:7zInstalled) { Remove-7Zip }

    # Snapshot missing drivers after install
    $script:AnalyticsMissingAfter = Get-MissingDriverCount
    # v1.11.0 - capture names of devices still unresolved so the analytics
    # sheet shows exactly what the run failed to fix. Write-MissingDriverDetails
    # below dumps the full HardwareID/CompatibleID detail to the .log.
    $afterNames = Get-MissingDriverNames
    foreach ($n in $afterNames) { $script:AnalyticsMissingAfterList.Add($n) | Out-Null }
    $missingDelta = if ($script:AnalyticsMissingBefore -ge 0 -and $script:AnalyticsMissingAfter -ge 0) {
        $script:AnalyticsMissingBefore - $script:AnalyticsMissingAfter
    } else { 0 }
    Log "Missing drivers AFTER  install: $($script:AnalyticsMissingAfter)"
    Log "Devices resolved by this install: $missingDelta"
    Write-MissingDriverDetails

    if ($script:CancelRequested) { Send-AnalyticsEvent -Result "cancelled" }

    Log "--------------------------------------------"
    if ($success) {
        SetProgress 100
        SetDownload -Pct 100 -Label "Complete"
        SetExtract  -Pct 100 -Label "Complete"
        Log "Driver installation complete!"
        Log "Log saved to: $LogFile"
        Send-AnalyticsEvent -Result "success"
        Write-HtmlReport    -Result "success"  # v1.10.0
        Play-Sound -Event "Success"

        if ($SkipCleanup) {
            Log "SkipCleanup flag set - keeping $driverRoot for inspection."
        } elseif (Test-Path $driverRoot) {
            Log "Cleaning up $driverRoot..."
            try {
                Remove-Item $driverRoot -Recurse -Force -ErrorAction Stop
                Log "  $driverRoot removed."
            } catch { Log "  WARNING: Could not remove $driverRoot - $($_.Exception.Message)" }
        }

        $missingLine = if ($script:AnalyticsMissingBefore -ge 0 -and $script:AnalyticsMissingAfter -ge 0) {
            "`n`nMissing drivers:  $($script:AnalyticsMissingBefore) -> $($script:AnalyticsMissingAfter)  ($missingDelta resolved)"
        } else { "" }

        # v1.12.0 - Check if Windows Update should be prompted
        $stillMissing = if ($script:AnalyticsMissingAfter -ge 0) { $script:AnalyticsMissingAfter } else { -1 }
        if ($PromptWindowsUpdate -and $stillMissing -gt 0) {
            Log "PromptWindowsUpdate enabled and $stillMissing drivers still missing - will offer Windows Update"
        }

        if ($script:Headless) {
            if (-not $script:Silent) {
                Write-Host "SUCCESS: Drivers installed for $model.$missingLine"
                Write-Host "Run complete. Reboot when ready."
                Write-Host "Report: $ReportFile"
            }
            # v1.12.0 - In headless mode with PromptWindowsUpdate, open Windows Update if drivers still missing
            if ($PromptWindowsUpdate -and $stillMissing -gt 0 -and -not $script:Silent) {
                Start-Sleep -Milliseconds 500
                Open-WindowsUpdate
            }
        } else {
            $msgText = "Drivers installed successfully for:`n$model$missingLine`n`nReport saved to:`n$ReportFile"
            # v1.12.0 - Add Windows Update prompt if drivers still missing
            if ($PromptWindowsUpdate -and $stillMissing -gt 0) {
                $msgText += "`n`n$stillMissing drivers still missing.`nWould you like to open Windows Update to search for additional drivers?`n`nSelect Yes to open Windows Update, No to skip, or Cancel to reboot."
                $result = [System.Windows.Forms.MessageBox]::Show(
                    $msgText,
                    "Installation Complete", "YesNoCancel", "Information"
                )
                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Open-WindowsUpdate
                    Set-ButtonIdle
                } elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
                    Restart-Computer -Force
                } else {
                    Set-ButtonIdle
                }
            } else {
                # Standard completion dialog if no missing drivers or PromptWindowsUpdate not enabled
                $msgText += "`n`nReboot now to complete installation?"
                $result = [System.Windows.Forms.MessageBox]::Show(
                    $msgText,
                    "Installation Complete", "YesNo", "Information"
                )
                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { Restart-Computer -Force }
                else { Set-ButtonIdle }
            }
        }
    } else {
        # On cancel, Send-AnalyticsEvent was already fired above with result="cancelled".
        # On non-cancel failure, fire it here. Either way, write the HTML report next
        # so the operator has a permanent record even on broken/cancelled runs.
        if (-not $script:CancelRequested) { Send-AnalyticsEvent -Result "failure" }
        $reportResult = if ($script:CancelRequested) { "cancelled" } else { "failure" }
        Write-HtmlReport -Result $reportResult  # v1.10.0
        SetDownload -Pct 0 -Label "Failed - see log"
        SetExtract  -Pct 0 -Label "Failed - see log"
        Stop-DlSpinner      -Success $false
        Stop-ExSpinner      -Success $false
        Stop-OverallSpinner -Success $false
        Play-Sound -Event "Failure"
        Log "Driver installation did not complete. Check log: $LogFile"
        
        # v1.12.0 - Check if Windows Update should be prompted on failure
        $stillMissing = if ($script:AnalyticsMissingAfter -ge 0) { $script:AnalyticsMissingAfter } else { -1 }
        
        if ($script:Headless) {
            if (-not $script:Silent) {
                Write-Host "FAILED: Driver installation did not complete. Check log: $LogFile"
                Write-Host "Report: $ReportFile"
            }
            # v1.12.0 - In headless mode with PromptWindowsUpdate, offer Windows Update on failure if drivers missing
            if ($PromptWindowsUpdate -and $stillMissing -gt 0 -and -not $script:Silent) {
                Write-Host "Note: $stillMissing drivers still missing. Consider using Windows Update to find additional drivers."
                Start-Sleep -Milliseconds 500
                Open-WindowsUpdate
            }
        } else {
            $msgText = "Driver installation failed or no pack was found.`nCheck the log:`n`n$LogFile`n`nReport: $ReportFile"
            # v1.12.0 - Add Windows Update prompt if drivers still missing
            if ($PromptWindowsUpdate -and $stillMissing -gt 0) {
                $msgText += "`n`n$stillMissing drivers still missing.`nWould you like to open Windows Update to search for additional drivers?"
                $result = [System.Windows.Forms.MessageBox]::Show(
                    $msgText,
                    "Installation Failed", "YesNo", "Warning"
                )
                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Open-WindowsUpdate
                    Set-ButtonIdle
                } else {
                    Set-ButtonIdle
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    $msgText,
                    "Installation Failed", "OK", "Error"
                )
                Set-ButtonIdle
            }
        }
    }
}

# =========================
# WIRE UP + LAUNCH
# =========================
if ($Headless) {
    # Headless mode - run directly, no GUI
    Start-Install
} else {
    # GUI mode - wire up form and show
    $button.Add_Click({ Start-Install })

    $cancelButton.Add_Click({
        if ($cancelButton.Enabled) {
            $script:CancelRequested = $true
            Log "--- Cancel requested by user ---"
            Play-Sound -Event "Cancel"
            $cancelButton.Enabled   = $false
            $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
            [System.Windows.Forms.Application]::DoEvents()
        }
    })

    $form.Add_Shown({
        $form.Activate()
        Start-Sleep -Milliseconds 300
        Log "Running startup checks..."
        Start-Install
    })

    [void]$form.ShowDialog()
}