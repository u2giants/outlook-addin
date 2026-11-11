# AGENTS.md — POP Tools for Outlook: operating guide

The complete document index is [`docs/INDEX.md`](docs/INDEX.md), generated from document headers. Search it for your topic's keywords (or read it) before opening or searching any other document. Open a document only when its summary matches your task.

This file is a router: it routes, it does not contain. Detail lives under [`docs/`](docs/INDEX.md).

## Project summary

A classic Outlook (Windows) COM add-in, written in C# for .NET Framework 4.8, that replaces the per-PC VBA macros kept in `legacy-vba/`. Albert Hazan (owner, POP Creations) and his PCs use it. It adds Remove Pictures and Remove Attachments to a POP Tools group on the Home tab, puts the Retention commands on a separate POP Tools tab, and stamps a Retention column on new, changed and sent mail. It installs once per PC, per user, without admin rights, and then updates itself from a shared Dropbox folder. Classic Outlook only: not "new Outlook", Mac or web.

## Multi-model AI note

There is no universal ignore-file standard across AI coding tools.
`.claudeignore` works for Claude Code.
When using any other AI tool, paste this file as the first message and follow the instructions in the "What to ignore" section.

## Standing rules (never skip)

- This repository is PUBLIC. Never commit a secret, a mailbox address, email content or a PC credential.
- Work on a branch and open a pull request into `main`; `main` is the only long-lived branch.
- Test Remove Pictures / Remove Attachments only on a COPY of an email; they change the message permanently.
- Never change the add-in's ProgId or CLSID (see the identifier table); installed PCs are registered under them.
- A new version must carry a higher version number, or no PC will install it ([`docs/deployment.md`](docs/deployment.md)).
- New quirk or incident → [`docs/idiosyncrasies.md`](docs/idiosyncrasies.md), never here.
- Every new Markdown file gets the four-line header (`summary`, `kind`, `status`, `read-when`); then regenerate the index with `ai-doc-health --repo . --fix-index` (tool: popcre/ai-devops `docs/doc-health.md`). Never hand-edit `docs/INDEX.md`.

## Documentation map: what to read for each task

Always start with:

- `AGENTS.md`

Then load additional docs only when relevant:

| Task / question | Read these docs | Usually do not need |
|---|---|---|
| Quick repo orientation | `README.md`, `AGENTS.md` | Anything under `docs/` |
| Change add-in behavior (ribbon, retention, attachments, updater) | `AGENTS.md`, `docs/idiosyncrasies.md` | `docs/deployment.md` unless install or update behavior changes |
| Build, run or debug locally | `docs/development.md` | `docs/deployment.md` |
| Publish a version, install on a PC, roll back, verify a PC | `docs/deployment.md` | `docs/development.md` |
| Code looks wrong or odd; Outlook crashes | `docs/idiosyncrasies.md` | Other docs |
| Continue unfinished work | `HANDOFF.md` (pointer) → OPEN files in `HANDOFF.d/`, newest first | Docs the handoff does not name |
| Claude Code session | `CLAUDE.md`, then `AGENTS.md` | Other docs unless the task needs them |
| Documentation-only cleanup | `AGENTS.md`, `docs/INDEX.md`, the affected docs | Source files except to verify accuracy |

## Repository structure

| Path | What it is |
|---|---|
| `src/` | Project-owned C# add-in (the only code that ships) |
| `deploy/` | Project-owned installer: `Install.ps1` and its double-click wrapper `Install.cmd` |
| `legacy-vba/` | The original VBA macros, kept as reference for the port; not built, not shipped |
| `docs/` | Documentation; `docs/INDEX.md` is generated |
| `HANDOFF.md`, `HANDOFF.d/` | Static pointer and one write-once file per unfinished session |
| `.github/workflows/doc-health.yml` | Advisory documentation check; the only workflow |
| `bin/`, `obj/`, `installer/output/` | Build output, git-ignored |

There is no generated code, vendored code, migration or database. Third-party code arrives only as NuGet packages at build time.

## Prime Directive: custom-code boundary

All custom code lives here:

- `src/`
- `deploy/`
- `docs/`
- `.github/workflows/`

Everything else requires justification before touching. `legacy-vba/` is a frozen reference: read it to understand intended behavior, do not edit it.

## Core modification inventory

None. There is no third-party or framework source in this repository, so nothing outside the custom-code boundary has been modified.

## Task-to-file navigation: what to edit for common changes

| Task | Files to touch | Files not to touch |
|---|---|---|
| Add, move or rename a ribbon button | `src/Ribbon.xml`, the `OnAction` switch in `src/Connect.cs` | `legacy-vba/` |
| Change Remove Pictures / Remove Attachments | `src/AttachmentKit.cs` | `legacy-vba/modAttachInlineKit.bas` |
| Change retention stamping, the Retention column or view copying | `src/Retention.cs` | `legacy-vba/ThisOutlookSession.cls` |
| Change Outlook event wiring or startup/shutdown | `src/Connect.cs`, `src/ComInterfaces.cs` | The `ref Array custom` marshaling in `src/ComInterfaces.cs` (see quirks) |
| Change self-update | `src/Updater.cs`, `deploy/Install.ps1` | — |
| Change install, uninstall or old-macro cleanup | `deploy/Install.ps1` | — |
| Release a version | The three version lines in `src/PopOutlookAddin.csproj` | The `Microsoft.NETFramework.ReferenceAssemblies` package version (stays `1.0.3`) |

## Data model and external identifiers

There is no database. Do not rename, regenerate or replace these:

| Entity/System | Identifier | Where defined | Notes |
|---|---|---|---|
| COM ProgId | `PopOutlookAddin.Connect` | `deploy/Install.ps1`, `src/Connect.cs` | Outlook add-in registry key name |
| COM CLSID | `{AB496E0D-CC70-4F80-854B-0EA3E4C4FBE9}` | `deploy/Install.ps1`, `src/Connect.cs` | Per-user registration under HKCU |
| Install folder | `%LOCALAPPDATA%\PopOutlookAddin` | `deploy/Install.ps1`, `src/Connect.cs` | Holds the two DLLs and `addin.log` |
| Update package folder | `ai\outlook-addin` under the user's Dropbox root | `src/Updater.cs` | Where every PC looks for a newer DLL |
| Package files | `PopOutlookAddin.dll`, `Microsoft.Office.Interop.Outlook.dll`, `Install.ps1`, `Install.cmd` | `src/Updater.cs`, `deploy/` | The updater copies the first three |
| Embedded ribbon | `PopOutlookAddin.Ribbon.xml` | `src/PopOutlookAddin.csproj` | Logical resource name read at runtime |
| GitHub repository | `u2giants/outlook-addin` | GitHub | Public |

## Container and service inventory

None. The add-in runs inside the Outlook process on each PC; there is no server, container or hosted service.

## What to ignore

- `bin/`, `obj/` — build output
- `installer/output/` — packaged output
- `legacy-vba/` — unless you are checking how the original macro behaved

`.claudeignore` and `.cursorignore` list the first three; `legacy-vba/` stays readable on purpose.

## Intentional quirks and non-obvious decisions

Full write-ups: [`docs/idiosyncrasies.md`](docs/idiosyncrasies.md). Titles:

- `ref Array custom` is marshaled as a SafeArray of VARIANT
- The ReferenceAssemblies package version looks like a stale add-in version
- An unused "Remove Both" handler remains in the code
- The installer renames the old macro file and edits Outlook's ribbon file
- The installer's message box falls back to plain output
- The updater also checks at shutdown

## Credentials and environment

No secrets, API keys or configuration files. The add-in reads only standard Windows variables:

| Variable | Purpose | Stored where | Required in dev | Required in prod |
|---|---|---|---|---|
| `LOCALAPPDATA` | Install folder, log, and one place Dropbox keeps `info.json` | Windows | no (build needs none) | yes |
| `APPDATA` | Second place Dropbox keeps `info.json`; old `VbaProject.OTM` location | Windows | no | yes |

## Deployment

Real path, in full in [`docs/deployment.md`](docs/deployment.md): build on any machine, copy the DLL into the Dropbox package folder, and each PC installs it after its next Outlook close. No GitHub Actions workflow builds or deploys; there is no image, tag pattern or hosting platform. Roll back by publishing the older code under a higher version number. SSH to a PC is exceptional and for testing only.

## Critical incidents

One known incident, written up in [`docs/idiosyncrasies.md`](docs/idiosyncrasies.md#critical-incidents): v1.0.1 crashed Outlook on load (fixed in v1.0.2).

## Pending work

In-flight work is tracked only in the open files under `HANDOFF.d/` (read via `HANDOFF.md`). This file keeps no pending-work board, so nothing here can go stale.
