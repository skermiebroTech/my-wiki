# Windows Fingerprint Reader Test

A single-file PowerShell diagnostic that checks whether the fingerprint reader on a Windows PC is detected, configured, and working. Runs from the **Win+R** Run dialog with zero installation — PowerShell is already on every supported version of Windows.

## Run it

Press **Win+R**, paste this

```
powershell -ep bypass -nop -c "irm https://raw.githubusercontent.com/SkermiebroTech/my-wiki/main/fp-test.ps1 | iex"
```

A console window opens, runs the diagnostic, and at the end pops up the Windows Hello prompt so you can physically touch the sensor and confirm authentication works.

## What it checks

1. **Hardware** — enumerates devices in the `Biometric` PnP class and reports their driver status.
2. **Service** — confirms the Windows Biometric Service (`WbioSrvc`) is running.
3. **Windows Hello** — queries `UserConsentVerifier.CheckAvailabilityAsync()` to report `Available`, `DeviceNotPresent`, `NotConfiguredForUser`, `DisabledByPolicy`, or `DeviceBusy`.
4. **Enrollment** — checks `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WinBio\AccountInfo\<SID>` for the current user.
5. **Live test** — calls `UserConsentVerifier.RequestVerificationAsync()` to trigger an actual Windows Hello prompt and reports the result.

## Notes

- `UserConsentVerifier` authenticates with whichever Windows Hello modality the user has enrolled (fingerprint, face, or PIN). To isolate the fingerprint sensor specifically, make sure only a fingerprint is enrolled, or watch which method the Hello prompt asks for.
- Uses Windows PowerShell 5.1 (the built-in `powershell.exe`). The WinRT interop pattern used here does not work reliably in PowerShell 7 (`pwsh`).
- No admin rights required.
- Network: the `irm | iex` invocation requires internet access to reach `raw.githubusercontent.com`. On air-gapped machines, save `fp-test.ps1` to disk and run `powershell -ep bypass -f fp-test.ps1`.