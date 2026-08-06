# ⚡ Quick Driver Installation
## Automatic one liner:
```
powershell -WindowStyle Hidden -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1 | iex"
```

## Automatic one liner + auto reboot:
Same as above but the **Auto-reboot** toggle starts ON — 15 seconds after a successful install the machine restarts itself (`shutdown /a` to abort). The toggle can still be unticked in the app before the run finishes. Failed or cancelled runs never reboot.
```
powershell -WindowStyle Hidden -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto.ps1))) -AutoReboot"
```

## DEV Version (could be unstable)
```
powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Install-Drivers-auto-dev.ps1 | iex"
```
## Links

- [Dell Driver Packs ](https://www.dell.com/support/kbdoc/en-au/000124139/dell-command-deploy-driver-packs-for-enterprise-client-os-deployment)
- [HP Driver Packs](https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html)
- [Lenovo Driver Packs](https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html)
- [Panasonic Driver Packs](https://global-pc-support.connect.panasonic.com/driver/deployment-support-tools/driver-pack#topics03)
- [Lenovo System Update](https://download.lenovo.com/pccbbs/thinkvantage_en/system_update_5.08.03.59.exe)
- [Camera Test](https://openthecamera.com/camtest)

## Dell

```bash
pnputil /add-driver "C:\Users\Administrator\<model>\*.inf" /subdirs /install
```

⚠️ Folder name changes per model. Replace ``<model>`` with folder name.

---

## Lenovo

```bash
pnputil /add-driver "C:\DRIVERS\*.inf" /subdirs /install
```

---

## HP

```bash
pnputil /add-driver "C:\SWSetup\*.inf" /subdirs /install
```

---
Author: Joel Skerman | Date: 23 Apr 2026 | updated: 06 Aug 2026