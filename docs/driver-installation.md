# Driver Installation Guide

## Step 1 – Open Driver Resource Page

Go to the appropriate driver download page:

- [Dell Driver Packs ](https://www.dell.com/support/kbdoc/en-au/000124139/dell-command-deploy-driver-packs-for-enterprise-client-os-deployment)
- [HP Driver Packs](https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html)
- [Lenovo Driver Packs](https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html)

---

##  Step 2 – Find Your Device & Download Drivers

- Search for your **exact model**
- Download the **driver pack (.exe)**

---

##  Step 3 – Extract the Driver Package

Run the `.exe` file to extract drivers.

- This does not install drivers yet

---

## Step 4 – Verify Extraction Location

| Brand  | Location                       |
| ------ | ------------------------------ |
| Dell   | `C:\Users\<username>\<model>\` |
| HP     | `C:\SWSetup\`                  |
| Lenovo | `C:\DRIVERS\`                  |

---

## Step 5 – Install Drivers

Open **Command Prompt as Administrator**.

### Dell

```bash
pnputil /add-driver "C:\Users\Administrator\<model>\*.inf" /subdirs /install
```

⚠️ Folder name changes per model. Replace ``<model>`` with folder name.
### Lenovo

```bash
pnputil /add-driver "C:\DRIVERS\*.inf" /subdirs /install
```
### HP

```bash
pnputil /add-driver "C:\SWSetup\*.inf" /subdirs /install
```

---
## Step 6 – Wait for Installation

Drivers will install automatically.

---
## Step 7 – Restart

Restart the laptop after installation.

---
## Step 8 – Run Windows Update

stop blanco from continuing and run regular windows update once

---
## Step 8 – Reboot

Reboot system and let blanco run!