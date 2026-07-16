# Operations Log

Date: 2026-06-04
Actor: Codex

## Visio Automation Global Skill

- Checked user-level skills directory: `C:\Users\Ommmr\.codex\skills`.
- Checked Visio COM registration with `Visio.Application`: available.
- Staged a `visio-automation` skill under project `.codex/visio-automation-skill`.
- The skill provides a PowerShell COM script that converts JSON diagram specs into `.vsdx` files and optional PDF/PNG exports.
- Added a `.cmd` wrapper so the skill can run on systems where direct `.ps1` execution is blocked by PowerShell execution policy.
- Installed the skill globally at `C:\Users\Ommmr\.codex\skills\visio-automation`.
- Ran an end-to-end smoke test through Visio COM and generated `.codex\visio\ai-visio-demo.vsdx` plus `.codex\visio\ai-visio-demo.png`.
- Confirmed no remaining `VISIO` process after the hidden automation run.
- 2026-06-04: Created `.codex\visio\rd-test-full-flow.json` for a detailed R&D testing process from project initiation to closure.
- 2026-06-04: Visio generation initially failed because Windows PowerShell read UTF-8 Chinese JSON with the default ANSI code page; updated the script to read specs with `-Encoding UTF8`.
- 2026-06-04: Adjusted connector z-order in the Visio automation script so exported previews keep node text readable.
- 2026-06-04: Added visible live-operation parameters `-KeepOpen` and `-StepDelayMs` so Visio can remain open while the user watches the diagram being drawn.
- 2026-06-04: User requested a horizontal flowchart similar to the provided image; added `.codex\visio\rd-test-horizontal-swimlane.json` and a dedicated Visio COM script for a horizontal swimlane layout.
- 2026-06-04: Improved horizontal swimlane rule boxes to use top-left text alignment and text margins.
- 2026-06-04: Moved the horizontal swimlane drawing script from the project `.codex\visio` folder into the staged global `visio-automation` skill as `new_visio_swimlane_diagram.ps1`, with a `.cmd` wrapper.
- 2026-06-04: Enhanced the global swimlane script with canvas-aware coordinate/font scaling and boundary-to-boundary elbow arrows with visible arrowheads.
- 2026-06-04: Added global swimlane incremental-update support: stable Visio shape names, open-document reuse via `-UpdateExisting`, and scoped redraw through `-UpdateIds`.
- 2026-06-04: Added global swimlane layout rules for shape-overlap avoidance, route-channel offsets to reduce line overlap, and text sizing based on each shape's rendered size and text length.
- 2026-06-04: Added global reference-style swimlane controls: `layout.referenceStyle` for thicker Visio-like strokes and `layout.autoAvoidOverlap=false` for exact coordinate reproduction.
- 2026-06-04: Added full-redraw cleanup for existing Visio canvases: `-UpdateExisting` without `-UpdateIds` now keeps Visio open, removes prior `visio_auto_*` generated shapes and flow lines, then redraws the next version on the same page.
- 2026-06-04: Added reference-layout controls for explicit node/rule/lane font sizes and edge `fromPort`/`toPort` routing.
- 2026-06-04: Changed `-UpdateExisting` document selection to prefer the current active Visio document/canvas before opening or creating another document when the output path differs.
- 2026-06-04: Added `layout.clearPageOnFullRedraw=true` so full redraws can remove unnamed legacy artifacts from the active Visio page before drawing the next version.
- 2026-06-10: Replaced the swimlane outside-stub connector strategy with edge-attached connectors plus `layout.layerConnectorsUnderModules=true`; verified syntax, unique ports, unique orthogonal segments, and no connector/module interior crossings before redrawing the active Visio canvas.
