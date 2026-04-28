# Preinstalling the Panasonic FZ-G1 (mk2)
⚠️ The MK2 has a 4th gen Intel CPU
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

- ⚠️ Reboot after complete

If you need to rerun use the following:

```powershell
pnputil /add-driver "C:\Drivers\FZ-G1\*.inf" /subdirs /install
```

---
## Step 5

Run Windows Update and reboot If prompted 
### Step 5.1
- You may ignore this step if all devices show up in device manager
1. Download the [USB Serial Driver](https://pc-dl.panasonic.co.jp/public/soft_update/d_driver/usb_serial/usbser_2.12.16_5086.exe) from Panasonic
2. open device manager and install driver from this directory `c:\util2\drivers\usbser`
---
## Step 6
verify all drivers are installed by checking in device manager

---
## Step 7
Run Sysprep!

---
## Links
[Panasonic Driver search](https://global-pc-support.connect.panasonic.com/search)

Author: Joel Skerman | Date: 28 Apr 2026