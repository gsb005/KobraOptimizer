# KobraOptimizer v1.5.0

This build focuses on UI refinement and laptop fit while keeping the stable functionality from the v1.4.x line.

## Highlights
- Smaller default window height with screen-aware bounds at launch
- Classic File / Tools / Help menu bar
- Stronger neon accent pass across buttons, group boxes, title, and menu interactions
- Toggleable log panel to free vertical space on smaller screens
- Footer actions for Support development and Disclaimer
- Tooltip coverage on primary actions and Windows tools
- Startup manager, Analyze, backup, manifest generation, and network tools preserved from the working branch
- Font fallback logic: prefers JetBrains Mono when installed, then falls back to Consolas or Segoe UI

## Packaging notes
- Put your logo at `Assets\logo.png`
- Set your real support URL in `Main.ps1` via `$script:DonationUrl`
- Use `Launch_Kobra.cmd` for the cleanest first-run experience
