# Windows 11 OOBE Bypass

One command that skips the whole Windows 11 out-of-box experience. It creates a
local administrator named `user` with a blank password, marks setup as complete,
and reboots. The machine comes back at the normal sign-in screen. No Microsoft
account, no network, no privacy pages, no region or keyboard pages.

!!! warning "The account has no password"
    Anyone at the keyboard gets administrator rights. Use this on bench machines
    only. Set a password before the machine leaves the workshop.

## Run it

Press **Shift+F10** at any Windows 11 OOBE screen. A command prompt opens with
SYSTEM rights. Paste this:

```
powershell -ep bypass -nop -c "irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Bypass-OOBE.ps1 | iex"
```

The machine reboots after 5 seconds. Sign in as `user` and leave the password
box empty.

### No network at OOBE

The `irm | iex` command needs internet access. OOBE has no Wi-Fi until you join
a network, so on a Wi-Fi-only machine use the offline version instead. Copy
[Bypass-OOBE.cmd](https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Bypass-OOBE.cmd)
to a USB stick, then run it from the Shift+F10 prompt:

```
X:\Bypass-OOBE.cmd
```

Replace `X:` with the drive letter of the USB stick. Use `wmic logicaldisk get
name` or `diskpart` to find it.

### Options

`iex` cannot pass parameters, so the one-liner uses a script block:

```
powershell -ep bypass -nop -c "& ([scriptblock]::Create((irm https://raw.githubusercontent.com/skermiebroTech/my-wiki/main/Bypass-OOBE.ps1))) -AutoLogon"
```

| Parameter | Effect |
|---|---|
| `-UserName <name>` | Account name. Default `user`. |
| `-AutoLogon` | Also boot straight to the desktop with no click. |
| `-NoReboot` | Apply the changes but stay in OOBE. |
| `-Force` | Allow the script to run on a machine that is already set up. |

The offline `.cmd` takes the same two options as arguments:

```
X:\Bypass-OOBE.cmd tech /autologon
```

## What it changes

**1. The account**

```
net user user /add /passwordreq:no
net localgroup Administrators user /add
```

The script finds the Administrators group by its SID (`S-1-5-32-544`), so a
non-English Windows also works. It then confirms the membership before it
continues.

**2. Setup state**

| Value | Data | Purpose |
|---|---|---|
| `HKLM\SYSTEM\Setup\CmdLine` | *(empty)* | Stops `oobe\windeploy.exe` on the next boot. |
| `HKLM\SYSTEM\Setup\SetupType` | `0` | No setup phase pending. |
| `HKLM\SYSTEM\Setup\OOBEInProgress` | `0` | OOBE is not running. |
| `HKLM\SYSTEM\Setup\SystemSetupInProgress` | `0` | Setup is not running. |
| `HKLM\SYSTEM\Setup\Status\ChildCompletion\setup.exe` | `3` | Marks `setup.exe` as complete. |
| `HKLM\SYSTEM\Setup\Status\SysprepStatus\GeneralizationState` | `7` | Stops a sysprepped image from resealing to OOBE. |
| `HKLM\SYSTEM\Setup\Status\SysprepStatus\CleanupState` | `2` | Sysprep cleanup done. |
| `...\CurrentVersion\Setup\State\ImageState` | `IMAGE_STATE_COMPLETE` | Image is deployed. |
| `...\CurrentVersion\Setup\OOBE\UnattendCreatedUser` | `1` | An account already exists. |
| `...\CurrentVersion\OOBE\BypassNRO` | `1` | No network requirement. |
| `...\Policies\Microsoft\Windows\OOBE\DisablePrivacyExperience` | `1` | No privacy pages. |
| `...\Policies\System\EnableFirstLogonAnimation` | `0` | No "Hi, we're setting things up" screen. |

**3. Clean-up**

The script deletes the `defaultuser0` placeholder account that OOBE leaves
behind.

## Notes

- **Windows 11 24H2 and later.** Microsoft removed the `oobe\bypassnro` command
  from the image. This script writes the `BypassNRO` value directly, so it still
  works.
- **Region and keyboard.** The script skips the region page, so Windows keeps
  the default of the install media. That is usually `en-US`. Correct it after
  sign-in under **Settings → Time & language**.
- **Computer name.** Windows keeps the random `DESKTOP-XXXXXXX` name. Rename the
  machine after sign-in.
- **Guard.** The PowerShell script refuses to run when Windows is already set
  up. Add `-Force` to override that. Windows 11 24H2 and LTSC 26100 clear
  `SystemSetupInProgress` and set `ImageState` to `IMAGE_STATE_COMPLETE` before
  the OOBE pages appear, so the script does not trust those two values alone. It
  also looks for the `defaultuser0` account, which exists only during OOBE. The
  script prints every signal it finds.
- **Undo the automatic sign-in.** Set
  `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoAdminLogon`
  to `0` and delete `DefaultPassword`.
- **Next step.** Run the [driver installer](Auto-Installer-Docs.md) after the
  first sign-in.
