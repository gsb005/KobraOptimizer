# KobraOptimizer

**KobraOptimizer** is a safety-first Windows cleanup and optimization utility built with **PowerShell + WPF**. It is designed for people who want a more transparent alternative to one-click “PC booster” apps: scan first, review what was found, back up important settings, and only then clean or tune the system.

KobraOptimizer focuses on **real maintenance tasks** that users actually care about on **Windows 11** and modern Windows desktops:

- cleaning temporary files and Windows junk
- clearing browser cache without touching bookmarks or saved passwords
- reviewing and cleaning **safe registry traces** with backup options
- reducing startup clutter
- surfacing performance-impacting items without fake miracle claims
- showing results clearly before anything is removed

This project aims to be a **cleaner, more transparent, safer Windows cleaner** with a modern desktop UI, clear sections, and visible safety controls.

---

## Why KobraOptimizer?

Many Windows cleanup tools promise huge performance gains but hide what they change. KobraOptimizer takes a different approach:

- **Preview-first workflow** instead of blind cleaning
- **Dedicated sections** for system cleanup, browser cleanup, registry traces, startup management, and performance review
- **Backup-aware design** for riskier operations
- **Transparent results** with section-based scans and review screens
- **No default deletion of bookmarks or saved passwords**
- **Conservative registry cleanup** focused on user traces, not aggressive registry “repair”

KobraOptimizer is being built as an honest utility for people who want control, visibility, and simplicity.

---

## What KobraOptimizer does

### 1. Quick Scan
Quick Scan is the fast path. It analyzes commonly selected cleanup areas and moves directly into a clean review flow.

Typical quick-scan targets include:

- User Temp
- System Temp
- Windows Update Cache
- Thumbnail / Icon Cache
- Recycle Bin
- selected browser cache targets

Quick Scan is for users who want a simple, guided scan-and-review experience.

### 2. Custom Clean
Custom Clean gives the user more direct control. It breaks cleanup into separate sections so users can scan only what they care about.

Current direction includes:

- **System cleanup**
  - User Temp
  - System Temp
  - Windows Update Cache
  - Thumbnail / Icon Cache
  - DirectX Shader Cache
  - Recycle Bin

- **Browser cleanup**
  - Google Chrome cache
  - Microsoft Edge cache
  - Firefox cache
  - optional cookie cleanup

- **Registry cleanup**
  - safe registry traces only
  - registry backup before cleanup

The goal is to let users scan one area at a time instead of scanning the whole PC every time.

### 3. Results Review
KobraOptimizer is built around a review step. Before cleanup, users can see:

- how many records or items were found
- how much space can be reclaimed
- which categories are selected
- what area is being cleaned

This helps users make informed decisions instead of clicking through a blind “optimize” button.

### 4. Performance Optimizer
Performance Optimizer is being shaped as a **practical control center**, not a fake speed-boost engine.

The focus is on:

- active app review
- startup impact review
- excluded apps
- honest recommendations

KobraOptimizer does **not** aim to be a “miracle RAM booster.” It aims to help users identify clutter, startup load, and resource-heavy apps safely.

### 5. Startup Manager
Startup Manager is intended to help users review and manage startup applications more clearly so Windows boots with less clutter.

### 6. Tools and Backup-Oriented Actions
The Tools area is meant for support actions such as:

- backup bundles
- manifests
- logs
- utility shortcuts
- maintenance helpers

The long-term goal is for advanced details to be available without dominating the main UI.

---

## Safety-first design

KobraOptimizer is not built around risky “clean everything” behavior. Safety is one of the core reasons this project exists.

### Browser safety
By design, browser cleanup is intended to preserve:

- bookmarks
- saved passwords / login data
- autofill data

**Cookies are only removed if the user explicitly selects them.** Removing cookies can sign users out of websites, so that option is kept separate and clearly labeled.

### Registry safety
Registry cleanup is intentionally conservative. The focus is on **safe user traces**, such as:

- Run history
- Typed Paths
- Typed URLs
- Recent Docs traces
- Explorer search / user MRU-style traces

KobraOptimizer is **not** intended to be an aggressive registry cleaner. It does not claim that deleting random registry entries will “make Windows 10x faster.”

### Backup-aware workflow
Where appropriate, the project is moving toward backup-aware behavior such as:

- registry backup before registry cleanup
- restore-point or backup options before deeper maintenance steps
- clear review screens before clean actions

---

## What KobraOptimizer does **not** do

KobraOptimizer is not meant to be:

- an adware-style cleaner
- a fake “AI speed booster”
- a miracle performance app
- an aggressive registry repair suite
- a tool that silently deletes browser data the user wanted to keep

The aim is to reclaim space, reduce clutter, improve visibility, and give Windows users a cleaner, more modern maintenance experience.

---

## Interface and product direction

KobraOptimizer is evolving from a simple PowerShell utility into a more product-style Windows app with:

- a dedicated sidebar
- cleaner page-based sections
- scan, results, and progress views
- a stronger visual identity
- a modern, review-focused workflow

Recent design direction includes:

- red neon theme exploration
- dedicated progress section
- separate Custom Clean sections
- improved Results page
- more product-like navigation and cards

---

## Screenshots

> Replace the filenames below with your actual screenshots in the `Screenshots/` folder.

### Dashboard
![KobraOptimizer Dashboard](Screenshots/dashboard.png)

### Quick Scan
![KobraOptimizer Quick Scan](Screenshots/quick-scan.png)

### Custom Clean
![KobraOptimizer Custom Clean](Screenshots/custom-clean.png)

### Results
![KobraOptimizer Results](Screenshots/results.png)

### Performance Optimizer
![KobraOptimizer Performance Optimizer](Screenshots/performance.png)

### Startup Manager
![KobraOptimizer Startup Manager](Screenshots/startup-manager.png)

---

## Installation

### Option 1: Run from source
1. Download or clone this repository.
2. Open the project folder.
3. Run `Launch_Kobra.cmd`.

### Option 2: PowerShell launch
You can also run the main script directly:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Main.ps1
```

### Requirements
- Windows 11 or modern Windows desktop environment
- PowerShell 5.1 or later
- Administrator mode is recommended for some maintenance actions

---

## How to use KobraOptimizer

### Quick Scan flow
1. Open **Quick Scan**
2. Run the scan
3. Review the results
4. Clean selected items

### Custom Clean flow
1. Open **Custom Clean**
2. Choose **System**, **Browser**, or **Registry** options
3. Run scan for the selected section
4. Review results
5. Clean only what you want

### Performance / Startup flow
1. Review active apps or startup items
2. Check recommendations
3. Exclude anything that should be left alone
4. Apply changes carefully

---

## Repository structure

Typical project structure:

```text
KobraOptimizer/
├── Assets/
├── Modules/
├── Screenshots/
├── Main.ps1
├── Kobra_UI.xaml
├── Launch_Kobra.cmd
├── Build_KobraExe.ps1
└── README.md
```

---

## Project status

KobraOptimizer is in active development.

Current priorities include:

- UI polish
- section-by-section scan and clean flow
- safer backup handling
- stronger results details
- better testing with Pester
- packaging and release workflow
- optional future EXE strategy

---

## Testing direction

The project is moving toward more structured testing with:

- **Pester** for PowerShell tests
- **PSScriptAnalyzer** for code quality and linting
- manual Windows UI testing
- Hyper-V / VM-based testing for safer validation on Windows 11 Pro

Planned testing areas include:

- scan logic
- results generation
- browser cleanup safety
- registry backup behavior
- manifest creation
- startup and performance views
- progress-state behavior

---

## Why this project matters

Windows users still want tools that help them:

- free disk space
- clean temp files
- review browser junk
- manage startup clutter
- keep things simple
- avoid bloated or suspicious cleanup software

KobraOptimizer is meant to fill that gap with a more transparent and safety-first approach.

---

## Support development

If you want to support the project, use the Ko-fi link below:

**Ko-fi:** https://ko-fi.com/kobraoptimizer

If KobraOptimizer helps you, even a small donation helps support more development, more testing, UI improvements, and future features.

---

## Disclaimer

KobraOptimizer is a system utility. Even with safety-first design, system cleanup tools should be used carefully.

Always:

- review what is being scanned
- understand what is selected
- use backup options where available
- avoid using any cleanup tool blindly on important systems without understanding what it will change

Use the software at your own risk.

---

## Open source direction

Open source is part of KobraOptimizer’s trust model. The more transparent the tool is, the easier it is for users to understand what it does and does not do.

That matters in the Windows cleanup space.

---

## Roadmap

Planned or possible future work:

- richer results breakdowns
- better per-section cleaning workflows
- improved backup UX
- stronger test coverage
- packaged release options
- possible C# / Visual Studio host shell in the future
- GitHub Releases with downloadable builds
- expanded documentation and screenshots

---

## Keywords

Windows cleaner, Windows 11 cleanup tool, PC optimizer, startup manager, browser cache cleaner, safe registry trace cleaner, PowerShell cleanup utility, WPF Windows utility, disk space cleanup, temp file cleaner, transparent Windows cleaner, CCleaner alternative, open source Windows maintenance tool.

---

## License

Add your license here. MIT is a simple and practical choice if you want broad reuse and contribution.
