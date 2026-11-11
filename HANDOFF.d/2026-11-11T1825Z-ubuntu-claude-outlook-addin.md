# Outlook add-in (POP Tools) — handoff

## 1. Goal
Replace Albert's per-PC Outlook VBA macros with a classic Outlook COM add-in that installs once per PC and then updates itself.

## 2. Current state (verified 2026-11-11 1:23 PM EST)
- Add-in v1.0.6 running on al8960ofc (SSH alias `4837`, user ahazan2). No Outlook crashes since the v1.0.2 fix.
- Self-update proven live: v1.0.5 -> v1.0.6 installed automatically after an Outlook restart (log line "Update pending: 1.0.5.0 -> 1.0.6.0").
- Retention live stamping proven: an Inbox item with a 365-day policy got Retention = "12 Mos".
- Ribbon: Home tab group "POP Tools" = Remove Pictures, Remove Attachments. Separate "POP Tools" tab = Retention Column (3 Fill In buttons) + One-time Setup (2 buttons). Layout requested by Albert; "Remove Both" removed at his request (code handler still exists, unused).
- Published package: Albert's Dropbox `ai\outlook-addin\` (two DLLs, Install.ps1, Install.cmd). Other PCs: double-click Install.cmd once.

## 3. Key files
- [Connect.cs](https://github.com/u2giants/outlook-addin/blob/main/src/Connect.cs) — COM entry, event wiring, ribbon callbacks
- [Retention.cs](https://github.com/u2giants/outlook-addin/blob/main/src/Retention.cs), [AttachmentKit.cs](https://github.com/u2giants/outlook-addin/blob/main/src/AttachmentKit.cs) — ports of the VBA in [legacy-vba/](https://github.com/u2giants/outlook-addin/blob/main/legacy-vba)
- [Updater.cs](https://github.com/u2giants/outlook-addin/blob/main/src/Updater.cs) — Dropbox version check at startup and shutdown
- [Ribbon.xml](https://github.com/u2giants/outlook-addin/blob/main/src/Ribbon.xml), [deploy/Install.ps1](https://github.com/u2giants/outlook-addin/blob/main/deploy/Install.ps1)

## 4. Unfinished
1. Remove Pictures / Remove Attachments never exercised on a real email. Test on a COPY of an email with an inline image and an attachment on al8960ofc (drive Outlook via a scheduled task with /IT in the console session; SSH runs in a non-interactive session). Confirm pictures/attachments removed, note banner added, save-to-folder works.
2. Not installed on Albert's other PCs yet (Albert's own action: double-click Install.cmd in Dropbox).
3. "Copy My View to All Folders" was rewritten (View.Copy -> Views.Add + XML) and is untested.

## 5. Tried and failed / gotchas
- v1.0.1 crashed Outlook (AV in clr.dll): IDTExtensibility2 `ref Array custom` must be marshaled as SafeArray of VARIANT. Fixed in ComInterfaces.cs.
- Install.ps1 MessageBox throws when non-interactive; Say() now falls back to output.
- olkexplorer.officeUI was ReadOnly; installer clears/restores the flag when removing old macro buttons (backup *.pre-addin.bak).
- `schtasks /change` over SSH fails silently; use `/create /f` then `/run`.
- Bumping version with a blanket sed also hits the ReferenceAssemblies package version (must stay 1.0.3).
- Creating a GitHub repo from this session was blocked; Albert created u2giants/outlook-addin (PUBLIC).

## 6. How to ship a new version
Bump Version/AssemblyVersion/FileVersion in src/PopOutlookAddin.csproj, `dotnet build src -c Release` (works on Linux), copy PopOutlookAddin.dll (and Install.ps1 if changed) into Dropbox ai/outlook-addin. Each PC installs it after its next Outlook close.

## 7. Verification
On the PC: `%LOCALAPPDATA%\PopOutlookAddin\PopOutlookAddin.dll` FileVersion, HKCU Outlook Addins\PopOutlookAddin.Connect LoadBehavior = 3, Application log has no Outlook Id 1000 events, addin.log.

## 8. Owner / decisions
Albert owns. Classic Outlook on Windows only (Albert confirmed all PCs are classic).

## 9. Next step
Item 4.1 above.
