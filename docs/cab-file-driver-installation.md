# Installing Drivers from a `.cab` File Using PowerShell (Audit Mode)

This guide explains how to extract and install drivers from a `.cab` file located in the Downloads folder while in Audit Mode.
- ⚠️ these pre-made commands are specifically for the Panasonic FZ-G1

---
## One-Liner (Extract + Install)

```powershell
$cab="C:\Users\Administrator\Downloads\G1F_Mk2_Win10x64_1803_V1.00.cab"
$out="C:\Drivers\Extracted"
mkdir $out -Force | Out-Null
expand -F:* $cab $out
pnputil /add-driver "$out\*.inf" /subdirs /install
```

 - ⚠️ Replace `G1F_Mk2_Win10x64_1803_V1.00.cab` with the filename of the CAB file you downloaded 
	 - ⚠️ you can do this more easily by pasting into `notepad.exe` and editing there

---
### Explanation:
- `/add-driver` → Adds drivers to the Windows driver store
- `/subdirs` → Searches all subfolders for `.inf` files
- `/install` → Installs drivers if matching hardware is found

---

Author: Joel Skerman | Date: 27 Apr 2026 | Updated 28 Apr 2026