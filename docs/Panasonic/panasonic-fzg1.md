# Preinstalling the Panasonic FZ-G1 (mk2)

## Step 1

Install Windows 10 from a USB

---
## Step 2

Enter Audit Mode by pressing `ctrl+shift+f3`

---
## Step 3 

Disable `Microsoft Serial Ballpoint` in device manager
- ⚠️  Note: Pressing `win+x` brings up a menu where you can click on device manager 


---
## Step 4
- Download this file [Download](https://na.panasonic.com/computer/cab/G1F_Mk2_Win10x64_1803_V1.00.cab)
- **After Download Is complete** run the following in powershell as Administrator

```powershell
$cab="C:\Users\Administrator\Downloads\G1F_Mk2_Win10x64_1803_V1.00.cab"
$out="C:\Drivers\FZ-G1"
mkdir $out -Force | Out-Null
expand -F:* $cab $out
pnputil /add-driver "$out\*.inf" /subdirs /install
```

- Reboot after complete
## Step 5 

Run Windows update and reboot if prompted

---
Author: Joel Skerman | Date: 28 Apr 2026