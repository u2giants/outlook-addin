---
summary: How a version reaches PCs — build, copy into the Dropbox package folder, first install by double-click, silent self-update, rollback, per-PC checks.
kind: ops
status: current
read-when: publishing a version, installing on a PC, rolling back, or confirming what a PC is running
---

# Deployment

There is no CI build, image, hosting platform or server. Shipping is a file copy.

## Publish a version

1. Bump the version ([`development.md`](development.md#bump-the-version)).
2. `dotnet build src -c Release`.
3. Copy `PopOutlookAddin.dll` into the package folder `ai\outlook-addin` in Albert's Dropbox. Copy `Microsoft.Office.Interop.Outlook.dll`, `deploy/Install.ps1` and `deploy/Install.cmd` too whenever they changed (all four must be present in the folder).

Nothing else is needed. Each PC installs the new version by itself.

## How a PC updates

- At startup, and again at shutdown, the add-in finds the Dropbox root (from Dropbox's `info.json`, else `%USERPROFILE%\Dropbox`) and compares its own version with the file version of `ai\outlook-addin\PopOutlookAddin.dll`.
- If the Dropbox copy is newer, it logs `Update pending: <old> -> <new>`. When Outlook closes it copies the package to a temporary folder and runs `Install.ps1 -Silent`, which waits for Outlook to exit and then replaces the files.
- The new version is therefore live on the Outlook start after the next close.

## First install on a PC

Double-click `Install.cmd` in the Dropbox package folder. It is per-user and needs no admin rights. It:

- copies the two DLLs to `%LOCALAPPDATA%\PopOutlookAddin` and registers the add-in under HKCU;
- adds the add-in to Outlook's do-not-disable list and clears any earlier "disabled" mark;
- removes ribbon buttons that pointed at the old macros (the ribbon file is backed up as `*.pre-addin.bak` first);
- renames an existing `VbaProject.OTM` to `VbaProject.OTM.pre-addin.bak` so the old macros do not run twice.

Uninstall: run `Install.ps1 -Uninstall`. It removes the registration and the install folder; it does not restore the two backups.

## Roll back

The updater installs only a strictly higher version, so putting an older DLL back in Dropbox does nothing. Check out the older code, give it a version number higher than the bad one, build and publish.

To get one PC back to the old macros: uninstall, then rename `VbaProject.OTM.pre-addin.bak` back to `VbaProject.OTM` in `%APPDATA%\Microsoft\Outlook`.

## Confirm what a PC is running

- File version of `%LOCALAPPDATA%\PopOutlookAddin\PopOutlookAddin.dll`.
- `LoadBehavior` = `3` under `HKCU\Software\Microsoft\Office\Outlook\Addins\PopOutlookAddin.Connect`.
- No Outlook event Id 1000 in the Windows Application log.
- `%LOCALAPPDATA%\PopOutlookAddin\addin.log` shows no fresh errors.

## Environment

No secrets and no configuration file. Runtime needs classic Outlook for Windows, .NET Framework 4.8, and Dropbox signed in on the PC (without Dropbox the add-in still runs; it just never updates).
