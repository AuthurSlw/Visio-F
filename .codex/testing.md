# Testing

Date: 2026-06-04
Actor: Codex

## Visio Automation Skill

### Syntax Check

Command:

```powershell
[System.Management.Automation.Language.Parser]::ParseFile(...)
```

Result: passed. `new_visio_diagram.ps1` has valid PowerShell syntax.

### COM Registration Check

Command:

```powershell
[type]::GetTypeFromProgID('Visio.Application') -ne $null
```

Result: passed. Visio COM registration is available.

### End-to-End Smoke Test

Command:

```powershell
& 'C:\Users\Ommmr\.codex\skills\visio-automation\scripts\new_visio_diagram.cmd' `
  -SpecPath '.codex\visio-automation-skill\scripts\example-flow.json' `
  -OutputPath '.codex\visio\ai-visio-demo.vsdx' `
  -ExportPng
```

Result: passed.

Generated:

- `C:\Users\Ommmr\Documents\工作\.codex\visio\ai-visio-demo.vsdx`
- `C:\Users\Ommmr\Documents\工作\.codex\visio\ai-visio-demo.png`

### Process Cleanup

Command:

```powershell
Get-Process VISIO -ErrorAction SilentlyContinue
```

Result: no remaining Visio process after hidden automation run.

## Horizontal Swimlane Global Script

Date: 2026-06-04
Actor: Codex

### Syntax Check

Command:

```powershell
[System.Management.Automation.Language.Parser]::ParseFile('.codex\visio-automation-skill\scripts\new_visio_swimlane_diagram.ps1', ...)
```

Result: passed.

### File-Level Verification

Verified global skill files exist:

- `C:\Users\Ommmr\.codex\skills\visio-automation\scripts\new_visio_swimlane_diagram.ps1`
- `C:\Users\Ommmr\.codex\skills\visio-automation\scripts\new_visio_swimlane_diagram.cmd`
- `C:\Users\Ommmr\.codex\skills\visio-automation\SKILL.md`

Verified global script contains:

- `-UpdateExisting`
- `-UpdateIds`
- stable generated names such as `visio_auto_node_<id>` and `visio_auto_edge_<index>`
- canvas-change guard for partial updates
- shape-overlap avoidance through `Resolve-NodePlacement`
- text sizing by shape size and text length through `Get-AutoFontSize`
- route-channel offsets through `Get-RouteOffset`

No Visio redraw was executed for this verification, per user instruction to avoid reopening/redrawing when preconditions do not require it.

### Reference-Style Generation

Date: 2026-06-04
Actor: Codex

Generated the requested horizontal swimlane diagram with global `visio-automation` code using:

```powershell
C:\Users\Ommmr\.codex\skills\visio-automation\scripts\new_visio_swimlane_diagram.cmd
```

Outputs:

- `C:\Users\Ommmr\Documents\工作\.codex\visio\rd-test-horizontal-swimlane-reference.vsdx`
- `C:\Users\Ommmr\Documents\工作\.codex\visio\rd-test-horizontal-swimlane-reference.png`

Verified the exported PNG renders as a horizontal reference-style swimlane diagram with stage columns, a right-side rules column, green deliverable boxes, decision diamonds, and thick arrowed flow lines.

### Same-Canvas Full Redraw Cleanup

Date: 2026-06-04
Actor: Codex

Verified global `new_visio_swimlane_diagram.ps1` now contains:

- `Clear-GeneratedContent`
- deletion of previous `visio_auto_*` generated shapes
- full-redraw cleanup when `-UpdateExisting` is used without `-UpdateIds`

No Visio redraw was executed for this verification. The change is code-level and documented in the global skill.

### Strict Same-Canvas Visio Redraw

Date: 2026-06-10
Actor: Codex

Executed the global swimlane automation against the active Visio document with `-Visible`, `-KeepOpen`, and `-UpdateExisting`. The script reused the current Visio canvas, cleared the previous generated diagram, and redrew the horizontal R&D/test swimlane process from an explicit routed spec.

Command:

```powershell
C:\Users\Ommmr\.codex\skills\visio-automation\scripts\new_visio_swimlane_diagram.cmd -SpecPath rd-test-horizontal-swimlane-final-routed.json -OutputPath rd-test-horizontal-swimlane-final-routed.vsdx -Visible -KeepOpen -UpdateExisting -ExportPng -StepDelayMs 20
```

Verification:

- PowerShell parser check passed for `.codex\visio-automation-skill\scripts\new_visio_swimlane_diagram.ps1`.
- Exported PNG was visually inspected.
- Visio remained open after generation.

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-final-routed.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-final-routed.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-final-routed.json`

### Increased Module Spacing Redraw

Date: 2026-06-10
Actor: Codex

Updated the global swimlane script so lane headers are positioned from the logical canvas height (`baseHeight - 1.2`) instead of a fixed `14.8` coordinate. This keeps the header row at the top when the canvas is enlarged to increase spacing.

Generated and visually inspected the expanded-spacing version:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-spaced.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-spaced.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-spaced.json`

Verification:

- PowerShell parser check passed for `.codex\visio-automation-skill\scripts\new_visio_swimlane_diagram.ps1`.
- The active Visio canvas was reused with `-UpdateExisting`.
- Visio remained open after redraw with `-KeepOpen`.

### Same-Canvas Module Spacing Redraw

Date: 2026-06-10
Actor: Codex

Corrected the spacing approach after confirming the canvas must remain unchanged. Generated a new same-canvas version from the final routed spec:

- page width: `27`
- page height: `16`
- content adjustment: module boxes reduced by roughly 10 percent and node coordinates lightly spread within the original canvas

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-same-canvas-spaced.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-same-canvas-spaced.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-same-canvas-spaced.json`

Verification:

- Exported PNG was visually inspected.
- Canvas size stayed at `27 x 16`.
- Visio remained open after redraw.

### Red-Box Vertical Spacing Correction

Date: 2026-06-10
Actor: Codex

Corrected the interpretation of "spacing" to mean the vertical distance between modules inside the user's marked red boxes, not canvas enlargement and not only box shrinking.

Canvas:

- page width: `27`
- page height: `16`

Measured center-distance changes:

- `analysis -> req_review`: `1.253 => 1.96`
- `ui_design -> ui_review`: `1.515 => 2.2`
- `tech_design -> tech_review`: `1.494 => 2.05`
- `case_design -> case_review`: `1.223 => 1.63`
- `acceptance -> acceptance_review`: `1.359 => 1.96`

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-redbox-spacing.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-redbox-spacing.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-redbox-spacing.json`

### Full-Height Same-Canvas Spacing

Date: 2026-06-10
Actor: Codex

Expanded node placement vertically across the original `27 x 16` canvas so the process modules fill the available height instead of staying concentrated in the middle.

Verification:

- Canvas remained `27 x 16`.
- Node vertical range expanded to `1.1 .. 13.7`.
- Visio was updated with `-UpdateExisting` and remained open.

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-fullheight-spacing.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-fullheight-spacing.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-fullheight-spacing.json`

### No-Overlap Clean Routing

Date: 2026-06-10
Actor: Codex

Adjusted the same-canvas layout to satisfy:

- modules do not overlap
- connector lines do not pass through non-endpoint modules
- long cross-lane links use outer or inter-module routing channels

Programmatic checks:

- module overlaps: `0`
- connector/module intersections: `0`
- canvas remained `27 x 16`

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v3.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v3.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v3.json`

### Split UI Review Connector Channels

Date: 2026-06-10
Actor: Codex

Separated the two connectors leaving `ui_review` so they no longer share the same vertical segment:

- `ui_review -> ui_design`: right-port return channel at `x=7.42`
- `ui_review -> dev_plan`: right-port cross-stage channel at `x=8.22`

Programmatic checks before redraw:

- module overlaps: `0`
- connector/module intersections: `0`

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v4.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v4.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v4.json`

### Fully Split UI Review Connectors

Date: 2026-06-10
Actor: Codex

Corrected the previous split because the two connectors still shared the initial right-side segment. The final routing separates both the start side and channel:

- `ui_review -> ui_design`: exits from the left side of `ui_review`, routes through the left return channel at `x=5.40`, then connects to the left side of `ui_design`.
- `ui_review -> dev_plan`: exits from the right side of `ui_review`, routes through the cross-stage channel at `x=8.22`, then connects to `dev_plan`.

Programmatic checks before redraw:

- module overlaps: `0`
- connector/module intersections: `0`

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v5.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v5.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v5.json`

### Global Shared-Connector Rule And UI Review Reroute

Date: 2026-06-10
Actor: Codex

Updated the global swimlane automation rule:

- `layout.disallowSharedConnectorSegments=true` by default.
- Reusing the same module input/output port now fails generation.
- Reusing the exact same connector segment now fails generation.

Updated the current diagram:

- `ui_review -> ui_design`: exits from the right side and returns to `UI设计`.
- `ui_review -> dev_plan`: exits from the bottom and connects to `开发计划排期/风险评估`.

Checks:

- PowerShell syntax check passed.
- interface duplicates: `0`
- redraw completed through the global script with `-UpdateExisting`.

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v6.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v6.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-clean-routing-v6.json`

### Strict Edge-Port Connector Rule

Date: 2026-06-10
Actor: Codex

Updated the global swimlane script so connector endpoints and route geometry are validated before drawing:

- every connector must use explicit `fromPort` and `toPort`
- allowed ports are edge centers and corners only
- connectors must be orthogonal
- connector segments must not enter any module interior
- connectors must not reuse module input/output ports or exact same line segments

The first redraw attempt correctly failed on a micro-diagonal segment in `ui_review -> dev_plan`; the waypoint was corrected to align exactly with the `ui_review` bottom-center port.

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-edge-port-strict.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-edge-port-strict.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-edge-port-strict.json`

### Mandatory Self-Test Before Drawing

Date: 2026-06-10
Actor: Codex

Added the global workflow rule: after changing Visio automation code, run complete local self-tests before drawing in Visio.

Self-test checklist executed:

- PowerShell syntax parse: `SCRIPT_SYNTAX_OK`
- explicit port validation: `SPEC_PORTS_OK edges=35`
- geometry validation: `GEOMETRY_OK edges=35 segments=134 stub=0.28`

Then synced the global skill and redrew the diagram.

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-edge-port-visible-v2.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-edge-port-visible-v2.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-edge-port-visible-v2.json`

### Outside-Only Connector Redraw

Date: 2026-06-10
Actor: Codex

Corrected the previous self-test flaw: start/end modules are no longer excluded from connector-interior validation. The drawing logic now supports `layout.keepConnectorsOutsideModules=true`, which draws connectors only to the outside stub points rather than onto the module boundary, preventing arrowheads from covering module interiors.

### Edge-Attached Layered Connector Redraw

Date: 2026-06-10
Actor: Codex

Replaced the outside-stub connector strategy because it produced visually disconnected arrows. The current strict strategy connects each route to the real module edge/corner port, keeps connector segments orthogonal and unique, and brings generated modules/text/rule boxes to the front after routing so connector caps and arrowheads do not visibly cover module interiors.

Self-test checklist executed before redraw:

- PowerShell syntax parse: `SCRIPT_SYNTAX_OK`
- layering code validation: `LAYERING_CODE_OK`
- strict geometry validation: `STRICT_GEOMETRY_OK edges=35 segments=134 interfaces=70 keepOutside=False layered=True`

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-edge-attached-layered.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-edge-attached-layered.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-edge-attached-layered.json`

Self-test checklist executed before redraw:

- PowerShell syntax parse: `SCRIPT_SYNTAX_OK`
- explicit port validation: `SPEC_PORTS_OK edges=35`
- strict geometry validation: `STRICT_GEOMETRY_OK edges=35 segments=64 stub=0.28 keepOutside=True`

Outputs:

- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-outside-connectors.vsdx`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-outside-connectors.png`
- `C:\Users\Ommmr\Documents\工作\rd-test-horizontal-swimlane-outside-connectors.json`
