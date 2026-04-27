# Installing Drivers from a `.cab` File Using PowerShell (Audit Mode)

This guide explains how to extract and install drivers from a `.cab` file located in the Downloads folder while in Audit Mode.

---

## 📦 Driver File

```
C:\Users\Administrator\Downloads\PDP_FZ-G1mk1_Win10x64_V1.02L10M00.cab
```

---

## 🔧 Step 1 — Extract the `.cab` File

Use the built-in `expand` command:

```powershell
mkdir "C:\Drivers\FZ-G1" -Force
expand -F:* "C:\Users\Administrator\Downloads\PDP_FZ-G1mk1_Win10x64_V1.02L10M00.cab" "C:\Drivers\FZ-G1\"
```

### Explanation:
- `-F:*` → Extracts all files from the `.cab`
- First path → Location of the `.cab` file
- Second path → Destination folder for extracted files

---

## 🔧 Step 2 — Install the Drivers

After extracting, install the drivers using `pnputil`:

```powershell
pnputil /add-driver "C:\Drivers\FZ-G1\*.inf" /subdirs /install
```

### Explanation:
- `/add-driver` → Adds drivers to the Windows driver store
- `/subdirs` → Searches all subfolders for `.inf` files
- `/install` → Installs drivers if matching hardware is found

---

## ⚡ One-Liner (Extract + Install)

```powershell
$cab="C:\Users\Administrator\Downloads\PDP_FZ-G1mk1_Win10x64_V1.02L10M00.cab"
$out="C:\Drivers\FZ-G1"
mkdir $out -Force | Out-Null
expand -F:* $cab $out
pnputil /add-driver "$out\*.inf" /subdirs /install
```

---

## ⚠️ Important Notes

- Run PowerShell **as Administrator**
- If drivers do not install:
  - The hardware may not be connected
  - Drivers may already be installed

---

## 🔍 Optional — List Installed Drivers

```powershell
pnputil /enum-drivers
```

---

## 📁 Example Workflow

```powershell
expand -F:* "C:\Users\Administrator\Downloads\PDP_FZ-G1mk1_Win10x64_V1.02L10M00.cab" "C:\Drivers\FZ-G1"
pnputil /add-driver "C:\Drivers\FZ-G1\*.inf" /subdirs /install
```

---

## ✅ Summary

1. Extract the `.cab` file from Downloads  
2. Install drivers using `pnputil`  
3. Verify installation if needed  

---