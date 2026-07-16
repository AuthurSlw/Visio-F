# Verification

Date: 2026-06-04
Actor: Codex

## Outcome

Installed a global Codex skill at:

`C:\Users\Ommmr\.codex\skills\visio-automation`

The skill includes:

- `SKILL.md`: trigger instructions and workflow for Visio automation.
- `agents\openai.yaml`: UI metadata for the skill.
- `scripts\new_visio_diagram.ps1`: JSON-to-Visio COM automation script.
- `scripts\new_visio_diagram.cmd`: execution-policy-safe wrapper.
- `scripts\example-flow.json`: runnable example spec.

## Verified Behavior

- Visio COM registration exists as `Visio.Application`.
- The PowerShell script parses successfully.
- The wrapper bypasses local PowerShell script execution policy.
- The script generated a real `.vsdx` file.
- The script exported a non-empty `.png` preview.
- Hidden automation quit Visio after completion.

## Residual Notes

- The generated connector routing is intentionally simple and uses Visio primitive lines. For highly polished diagrams, the next iteration should add orthogonal connector routing or stencil-backed dynamic connectors.
- This skill will be automatically discoverable in new Codex sessions that load user-level skills from `C:\Users\Ommmr\.codex\skills`.

## Horizontal Swimlane Update

Date: 2026-06-04
Actor: Codex

Updated the global `visio-automation` skill with a reusable horizontal swimlane generator:

- `C:\Users\Ommmr\.codex\skills\visio-automation\scripts\new_visio_swimlane_diagram.ps1`
- `C:\Users\Ommmr\.codex\skills\visio-automation\scripts\new_visio_swimlane_diagram.cmd`

The global script now supports:

- canvas-aware coordinate, shape, line, and font scaling
- boundary-to-boundary elbow arrows with arrowheads
- stable generated object names
- `-UpdateExisting` to reuse an open/existing Visio document
- `-UpdateIds` for scoped redraws such as `node:req_review`, `edge:3`, `rule:0`, or `lane:test`
- refusal of partial updates when canvas size changed
- node overlap detection and automatic in-lane nudging before drawing
- text font sizing based on each shape's rendered size and text length
- route-channel offsets so repeated or parallel flow lines do not stack directly on top of each other

The project-only script `.codex\visio\new_horizontal_swimlane.ps1` was removed by moving it into the global skill.

## Reference-Style Output

Date: 2026-06-04
Actor: Codex

Updated global `visio-automation` support for reference-image-style swimlane diagrams:

- `layout.referenceStyle=true` applies thicker strokes and reference-like visual treatment.
- `layout.autoAvoidOverlap=false` preserves exact coordinates for manually routed reference layouts.
- `edges[].waypoints` supports explicit orthogonal flow routing.

Generated:

- `C:\Users\Ommmr\Documents\工作\.codex\visio\rd-test-horizontal-swimlane-reference.vsdx`
- `C:\Users\Ommmr\Documents\工作\.codex\visio\rd-test-horizontal-swimlane-reference.png`

## Same-Canvas Redraw Rule

Date: 2026-06-04
Actor: Codex

Updated global `visio-automation` so `-UpdateExisting` without `-UpdateIds` means full redraw on the same Visio canvas:

- keep Visio open
- reuse the already-open or existing document
- delete prior generated `visio_auto_*` objects from the page
- redraw the next version on the same page
- keep Visio open afterward when `-KeepOpen` is used
## 2026-06-10 Visio Horizontal Swimlane Redraw

Actor: Codex

Result:

- Reused the active Visio canvas with `-UpdateExisting`; no intentional Visio close was performed.
- Redrew the requested horizontal R&D/test process flowchart from `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-final-routed.json`.
- Exported the generated diagram to `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-final-routed.png`.
- Saved the Visio file as `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-final-routed.vsdx`.
- Verified `new_visio_swimlane_diagram.ps1` parses successfully with the PowerShell language parser.

## 2026-06-10 Increased Visio Module Spacing

Actor: Codex

Result:

- Enlarged the logical drawing canvas to create larger module spacing while keeping shape sizes readable.
- Updated the global swimlane script so lane headers follow the canvas height instead of the old fixed `14.8` coordinate.
- Reused the active Visio canvas with `-UpdateExisting`; no intentional Visio close was performed.
- Redrew and exported the spacing-adjusted diagram:
  - `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-spaced.vsdx`
  - `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-spaced.png`
  - `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-spaced.json`

## 2026-06-10 Same-Canvas Module Spacing

Actor: Codex

Result:

- Kept the Visio canvas size unchanged at `27 x 16`.
- Increased visual spacing between modules by reducing module box dimensions and lightly spreading node coordinates inside the original canvas.
- Reused the active Visio canvas with `-UpdateExisting`; no intentional Visio close was performed.
- Generated:
  - `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-same-canvas-spaced.vsdx`
  - `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-same-canvas-spaced.png`
  - `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-same-canvas-spaced.json`

## 2026-06-10 No-Overlap Clean Routing

Actor: Codex

Result:

- Kept the canvas size unchanged at `27 x 16`.
- Moved overlapping document modules away from process modules.
- Re-routed long cross-lane connectors through outer and inter-module channels.
- Programmatic checks reported module overlaps `0` and connector/module intersections `0`.
- Reused the active Visio canvas with `-UpdateExisting`; Visio remained open.
- Generated:
  - `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v3.vsdx`
  - `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v3.png`
  - `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v3.json`
