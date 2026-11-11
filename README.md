# POP Tools for Outlook

Classic Outlook (Windows) COM add-in that replaces the old VBA macros in `legacy-vba/`.
Adds a **POP Tools** ribbon tab (Retention stamping, field/view setup, attachment stripping) and
stamps the Retention column live on new, changed and sent mail.

- Build: GitHub Actions (`.github/workflows/build.yml`) produces `PopOutlookAddin-Setup.exe`.
- Install: per-user, no admin. Silent: `PopOutlookAddin-Setup.exe /VERYSILENT /SUPPRESSMSGBOXES`.
  The installer renames any old `VbaProject.OTM` to `VbaProject.OTM.pre-addin.bak` so macros don't run twice.
- Auto-update: on startup the add-in looks for a newer `Dropbox\ai\outlook-addin\PopOutlookAddin-Setup.exe`
  and installs it silently after Outlook closes. Publish a new version = drop the new setup file there.
- Log: `%LOCALAPPDATA%\PopOutlookAddin\addin.log`.
- Classic Outlook only; not "new Outlook", Mac, or web.
