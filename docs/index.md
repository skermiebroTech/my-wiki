# ⚡ Quick Driver Installation

## Links

- [Dell Driver Packs ](https://www.dell.com/support/kbdoc/en-au/000124139/dell-command-deploy-driver-packs-for-enterprise-client-os-deployment)
- [HP Driver Packs](https://ftp.hp.com/pub/caps-softpaq/cmit/HP_Driverpack_Matrix_x64.html)
- [Lenovo Driver Packs](https://download.lenovo.com/cdrt/ddrc/RecipeCardWeb.html)
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
Author: Joel Skerman | Date: 23 Apr 2026