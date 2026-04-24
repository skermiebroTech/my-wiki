# Driver Installation Guide

## Step 1 – Open Driver Resource Page

Go to the appropriate driver download page:

- [Dell Driver Packs ](https://www.dell.com/support/kbdoc/en-au/000124139/dell-command-deploy-driver-packs-for-enterprise-client-os-deployment)
- [HP Driver Packs](https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html)
- [Lenovo Driver Packs](https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html)

---

##  Step 2 – Find Your Device & Download Drivers!

- Search for your **exact model**
- Download the **driver pack (.exe)**
	notes:
	- ⚠️ **Dell Driver Packs** require you to select the "Line Of Business" before searching for the specific model eg: XPS or Latitude 
	- ⚠️ **Lenovo** puts their most recent driver packages at the bottom of the page 
	- ⚠️ It is best to use **Machine Type** to find **Lenovo** driver packs
	- ⚠️ Use `ctrl+f` to find laptop models for **Dell** and **HP** 

---

##  Step 3 – Extract the Driver Package

Run the `.exe` file you just downloaded to extract drivers.

- ⚠️ This does not install drivers yet

---

## Step 4 – Install Drivers

Open **Powershell as Administrator**.
- ⚠️ If you are in audit mode you can do this by pressing `win+r` and typing `powershell`

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

### ⚠️ Troubleshooting 
If these commands do not work verify the driver folder exists at one of these locations

| Brand  | Location                       |
| ------ | ------------------------------ |
| Dell   | `C:\Users\<username>\<model>\` |
| HP     | `C:\SWSetup\`                  |
| Lenovo | `C:\DRIVERS\`                  |

---
## Step 5 – Wait for Installation

Drivers will install automatically. Be patient this can take a while 

---
## Step 6 – Restart

Restart the computer after installation if prompted.

---
## Step 7 – Run Windows Update

stop blancco from continuing (if it is running) and run regular windows update once

---
## Step 8 – Reboot

Reboot system and let blancco run!
Now that you have read the documentation you can find a minified version of this page at `9dtr.com/w`

---
Author: Joel Skerman | Date: 23 Apr 2026