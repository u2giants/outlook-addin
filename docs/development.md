---
summary: How to build the add-in on Linux or Windows, bump its version, read its log, and test it on a real PC without an interactive session.
kind: living
status: current
read-when: building, version-bumping, debugging or testing the add-in
---

# Development

## Build

```bash
dotnet build src -c Release
```

The project targets .NET Framework 4.8 (`net48`) and builds on Linux as well as Windows, because the `Microsoft.NETFramework.ReferenceAssemblies` package supplies the framework reference assemblies. Output lands in `src/bin/Release/net48/` (git-ignored): `PopOutlookAddin.dll` and `Microsoft.Office.Interop.Outlook.dll`.

There are no automated tests and no linter. Verification is done in a real classic Outlook.

## Bump the version

Edit three lines in `src/PopOutlookAddin.csproj`: `Version`, `AssemblyVersion`, `FileVersion`. Edit them by hand. A blanket search-and-replace also rewrites the `Microsoft.NETFramework.ReferenceAssemblies` package version, which must stay `1.0.3`.

## Where things are

| File | Role |
|---|---|
| `src/Connect.cs` | COM entry point, Outlook event wiring, ribbon callbacks (`OnAction`), the `Log` helper |
| `src/ComInterfaces.cs` | Hand-declared COM interfaces (`IDTExtensibility2`, ribbon) |
| `src/Retention.cs` | Retention column stamping, field registration, copying a view to all folders |
| `src/AttachmentKit.cs` | Remove Pictures / Remove Attachments |
| `src/Updater.cs` | Dropbox version check at startup and shutdown |
| `src/Ribbon.xml` | Ribbon layout, embedded in the DLL |
| `deploy/Install.ps1` | Per-user install, update, uninstall, old-macro cleanup |

`Retention.cs` and `AttachmentKit.cs` are ports of the VBA in `legacy-vba/`; read the VBA when the intended behavior is unclear.

## Debugging

- Log on the PC: `%LOCALAPPDATA%\PopOutlookAddin\addin.log`. Every caught exception is written there with the name of the method that caught it.
- Outlook crash on load: look in the Windows Application event log for Outlook event Id 1000.
- Add-in not loading: check `LoadBehavior` is `3` under `HKCU\Software\Microsoft\Office\Outlook\Addins\PopOutlookAddin.Connect`. Outlook sets it lower when it disables an add-in; re-running the installer clears the "disabled" mark.

## Testing on a PC from a remote session

- An SSH session is non-interactive, so it cannot drive Outlook's window. Run the test through a scheduled task set to run in the logged-on user's console session.
- `schtasks /change` fails silently over SSH. Create the task again with `/create /f`, then `/run`.
- Test Remove Pictures and Remove Attachments on a COPY of an email. They alter the message permanently.

Publishing and per-PC verification: [`deployment.md`](deployment.md). Things that look wrong but are deliberate: [`idiosyncrasies.md`](idiosyncrasies.md).
