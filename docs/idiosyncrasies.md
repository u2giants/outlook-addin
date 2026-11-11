---
summary: Deliberate oddities and the one incident — SafeArray marshaling crash in v1.0.1, pinned package version, unused Remove Both handler, installer side effects.
kind: living
status: current
read-when: code or installer behavior looks wrong and you are about to change it, or Outlook crashes on load
---

# Intentional quirks and incidents

## Intentional quirks

### `ref Array custom` is marshaled as a SafeArray of VARIANT

What it looks like: needless marshaling attributes on the `IDTExtensibility2` methods in `src/ComInterfaces.cs`.

Actually: Outlook passes that argument as a SafeArray of VARIANT, and the declaration has to say so.

Why: without it v1.0.1 crashed Outlook on load (see the incident below).

Do not change because: Outlook crashes with an access violation in `clr.dll` as soon as the add-in loads.

### The ReferenceAssemblies package version looks like a stale add-in version

What it looks like: `Microsoft.NETFramework.ReferenceAssemblies` at `1.0.3` in `src/PopOutlookAddin.csproj` looks like an add-in version somebody forgot to bump.

Actually: it is that NuGet package's own version and has nothing to do with the add-in version.

Why: the package is what lets a .NET Framework 4.8 project build on Linux.

Do not change because: a different number may not exist on NuGet and the build fails. A blanket version search-and-replace hits this line; bump the three add-in version lines by hand.

### An unused "Remove Both" handler remains in the code

What it looks like: `StripEverythingSelected` in `src/AttachmentKit.cs` (and its `OnAction` case) is dead code.

Actually: the "Remove Both" ribbon button was removed at the owner's request; the handler was left in place.

Why: not recorded. Only the button was asked to be removed; the handler is a port of the old macro `StripEverything_SelectedEmails`.

Do not change because: nothing breaks if it is removed, but removing it is a decision, not a cleanup. Leave it unless asked.

### The installer renames the old macro file and edits Outlook's ribbon file

What it looks like: `deploy/Install.ps1` reaches outside its own folder — it renames `VbaProject.OTM` and rewrites `olk*.officeUI`.

Actually: both are deliberate, and both are backed up first as `*.pre-addin.bak`.

Why: with the old macros still active every action would run twice, and the old ribbon buttons would point at macros that no longer exist. Outlook's ribbon file can be read-only, so the installer clears that flag, writes, and restores it.

Do not change because: removing either step brings back double-running macros or dead buttons on every PC that had the macros.

### The installer's message box falls back to plain output

What it looks like: the `Say` function in `deploy/Install.ps1` checks for an interactive session and wraps the message box in a try/catch.

Actually: a message box throws when the installer runs without a desktop (the silent self-update, or a remote test).

Why: the same script serves the double-click install and the silent update.

Do not change because: the silent update would fail on every PC.

### The updater also checks at shutdown

What it looks like: `RunPendingUpdate` in `src/Updater.cs` repeats the check that already ran at startup.

Actually: the second check catches a version published while Outlook was open.

Why: without it a PC needed two restarts to pick up a new version.

Do not change because: updates would again lag one restart behind.

## Critical incidents

### 2026-11-11 v1.0.1 crashed Outlook on load

What happened: Outlook crashed as soon as the add-in loaded.

Impact: Outlook unusable on the test PC until the add-in was replaced. No data loss.

Root cause: the `ref Array custom` argument of the `IDTExtensibility2` methods was not declared as a SafeArray of VARIANT, causing an access violation in `clr.dll`.

Recovery: fixed in `src/ComInterfaces.cs` and shipped as v1.0.2; no crash since.

Rule added to prevent recurrence: do not touch that marshaling; after any change to `src/ComInterfaces.cs`, load the add-in in a real Outlook and check the Application log for Outlook event Id 1000 before publishing.
