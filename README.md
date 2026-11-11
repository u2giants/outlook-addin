# POP Tools for Outlook

Classic Outlook (Windows) COM add-in that replaces the old VBA macros in `legacy-vba/`.
Adds Remove Pictures and Remove Attachments to a **POP Tools** group on the Home tab, and the Retention commands to a separate **POP Tools** tab, and
stamps the Retention column live on new, changed and sent mail.

- Build: `dotnet build src -c Release` (works on Linux). Bump the version in `src/PopOutlookAddin.csproj`.
- Publish: copy the two DLLs plus `deploy/Install.ps1` and `deploy/Install.cmd` to `Dropbox\ai\outlook-addin\`.
- First install on a PC: double-click `Install.cmd` in that Dropbox folder (per-user, no admin). It renames any old
  `VbaProject.OTM` to `VbaProject.OTM.pre-addin.bak` so macros don't run twice.
- Auto-update: on startup the add-in compares its version with the DLL in that Dropbox folder; if newer, it installs
  it silently after Outlook closes.
- Log: `%LOCALAPPDATA%\PopOutlookAddin\addin.log`.
- Classic Outlook only; not "new Outlook", Mac, or web.
