---
name: visio-automation
description: Use when the user asks Codex to create, edit, automate, generate, export, or visually verify Microsoft Visio diagrams, including .vsdx files, flowcharts, architecture diagrams, process maps, org charts, and "AI operates Visio" workflows on Windows.
metadata:
  short-description: Automate Microsoft Visio diagrams
---

# Visio Automation

Use this skill when the user wants a real Visio diagram, a `.vsdx` file, or an AI-assisted Visio drawing workflow.

## Preferred Workflow

1. Convert the user's request into a diagram spec:
   - nodes: stable `id`, readable `text`, optional `type`, `x`, `y`, `w`, `h`, `fill`, `line`
   - edges: `from`, `to`, optional `label`, `color`
   - page: optional `width`, `height`
2. Save the spec as JSON in the current project, usually under `.codex/visio/`.
3. Run `scripts/new_visio_diagram.ps1` from this skill to create the `.vsdx`.
4. If requested, export PDF or PNG with the same script.
5. If visual quality matters, inspect the export or screenshot and revise the JSON/spec.

After changing this skill's code, always complete local self-tests before drawing in Visio:

- parse the PowerShell script with `System.Management.Automation.Language.Parser`
- validate the target spec for explicit ports, unique connector interfaces, unique connector segments, orthogonal segments, and no connector/module intersections
- only after those checks pass, sync the global skill and run Visio drawing
- inspect the exported PNG before reporting success

Prefer COM automation over mouse/keyboard GUI automation. Use GUI automation only when the task explicitly requires demonstrating visible desktop operation or interacting with existing user-driven Visio UI state.

## Script

For a standard node-and-edge flowchart, use:

```powershell
& "<skill-dir>\scripts\new_visio_diagram.cmd" `
  -SpecPath ".codex\visio\diagram.json" `
  -OutputPath "diagram.vsdx" `
  -Visible `
  -KeepOpen `
  -StepDelayMs 300 `
  -ExportPdf `
  -ExportPng
```

Important:

- `-Visible` launches Visio visibly. Omit it for hidden automation.
- `-KeepOpen` leaves the generated document open after saving; use it when the user wants to watch or continue editing in Visio.
- `-StepDelayMs` slows drawing after each shape/connector so the user can watch real-time operation.
- Use the `.cmd` wrapper by default because many Windows systems block direct `.ps1` execution.
- The script auto-layouts nodes when `x` and `y` are missing.
- Visio uses inches internally. Keep coordinates around a US Letter or A4 page unless the spec sets page size.
- Run the script only on Windows with Microsoft Visio installed and COM registered as `Visio.Application`.

For a horizontal swimlane diagram like a staged R&D/test process chart, use:

```powershell
& "<skill-dir>\scripts\new_visio_swimlane_diagram.cmd" `
  -SpecPath ".codex\visio\swimlane.json" `
  -OutputPath "swimlane.vsdx" `
  -Visible `
  -KeepOpen `
  -UpdateExisting `
  -ExportPng `
  -StepDelayMs 150
```

Swimlane specs support:

- `lanes`: stage headers with `title`, `x1`, `x2`, and `fill`
- `nodes`: regular `process`, `decision`, `terminator`, and green `document` nodes
- `rules`: right-side or auxiliary rule-description boxes
- `edges`: directed links between node ids

The swimlane script treats coordinates as a logical canvas and scales coordinates, shapes, line weights, and font sizes to the actual Visio page size. Text size is calculated from each shape's rendered width, height, and text length. Before drawing, nodes are checked for overlap and automatically nudged within their swimlane to keep shapes separated. Arrows route from shape boundary to shape boundary with arrowheads; diagonal and repeated relationships are drawn through offset elbow channels so flow lines do not stack directly on top of each other.

For incremental changes, prefer:

```powershell
& "<skill-dir>\scripts\new_visio_swimlane_diagram.cmd" `
  -SpecPath ".codex\visio\swimlane.json" `
  -OutputPath "swimlane.vsdx" `
  -Visible `
  -KeepOpen `
  -UpdateExisting `
  -UpdateIds node:req_review,node:analysis,edge:3
```

Incremental update rules:

- If `-UpdateExisting` is set, the script first reuses the already-open document with the same `OutputPath`; if not found, it reuses the current active Visio document/canvas; if no active document exists, it opens the existing file; if no file exists, it creates one.
- When `-UpdateExisting` is used without `-UpdateIds`, the script treats the operation as a full redraw on the same canvas: it keeps Visio open, deletes previous `visio_auto_*` nodes, document boxes, rule boxes, lane headers, separators, and flow lines, then draws the new version on that page.
- If older runs produced unnamed or non-`visio_auto_*` artifacts, set `layout.clearPageOnFullRedraw=true` to clear the whole page before redrawing on the same open canvas.
- Generated objects receive stable names: `visio_auto_node_<id>`, `visio_auto_edge_<index>`, `visio_auto_lane_<id>`, `visio_auto_rule_<index>`.
- `-UpdateIds node:<id>` redraws that node and all related edges.
- `-UpdateIds edge:<index>` redraws only that edge.
- `-UpdateIds rule:<index>` redraws one rule box.
- `-UpdateIds lane:<id>` redraws one swimlane header/separator.
- If the canvas size changed and `-UpdateIds` requests a partial update, the script refuses the partial update and asks for a full redraw because coordinates no longer share the same precondition.
- Layout tuning can be supplied with `layout.shapePadding` in the swimlane JSON; the script uses it as the minimum spacing between node shapes during overlap avoidance.
- For reference-image reproduction, set `layout.referenceStyle=true` and `layout.autoAvoidOverlap=false` when the spec already provides exact coordinates and routed `waypoints`; this preserves the intended canvas composition instead of nudging shapes.
- Reference-style specs may set `node.fontSize`, `lane.fontSize`, `rule.fontSize`, `layout.minFontSize`, and `layout.maxFontSize` to keep text proportional to the designed shape size.
- Font sizing is adaptive: the swimlane script estimates each text block by rendered module width, rendered module height, line count, and visual text length; Chinese/full-width characters are treated wider than ASCII letters and numbers.
- Use `layout.adaptiveFontScale` to enlarge or reduce the calculated fit without changing module coordinates or module sizes.
- Edges may set `fromPort` and `toPort` as `top`, `bottom`, `left`, or `right`; combine these with `waypoints` for Visio-like orthogonal routing that avoids crossing through shapes.
- By default, connector layout is strict:
  - `layout.requireExplicitConnectorPorts=true`: every edge must set `fromPort` and `toPort`.
  - ports must be edge centers (`left`, `right`, `top`, `bottom`) or corners (`top-left`, `top-right`, `bottom-left`, `bottom-right`).
  - `layout.requireOrthogonalConnectors=true`: connectors must be horizontal/vertical segments through `waypoints`.
  - `layout.disallowConnectorIntersections=true`: connector segments must not enter any module interior.
  - `layout.disallowSharedConnectorSegments=true`: two connectors must not reuse the same module input/output port, and must not draw the exact same line segment.
  - `layout.connectorPortStub=0.28`: each connector first draws a clear outward stub from the source module edge and a clear outward stub before entering the target edge, so routes leave/enter from visible edge ports instead of cutting across modules.
  - `layout.keepConnectorsOutsideModules=false`: strict reference diagrams should connect to the real module edge/corner port, not stop at a floating outside point.
  - `layout.targetTerminalArrowOutsideModule=false`: terminal segments use Visio's native arrowhead style and normally route to the target module edge center/corner port.
  - `layout.targetTerminalArrowVisualClearance=0.03`: native Visio arrowheads must keep the terminal line endpoint just outside the target edge by this minimal distance so the rendered arrowhead visually touches the module edge without covering the module interior. This clearance is independent from `layout.connectorPortStub`; do not use the full port stub as terminal clearance because it creates a visible gap.
  - `layout.targetTerminalArrowSize=1.0`: terminal segments use a smaller native Visio arrowhead than the default size so arrows can sit close to modules without visually embedding into module fills.
  - `layout.layerConnectorsUnderModules=true`: after routing, generated modules, text, notes, and rule boxes are brought to the front so connector caps and native arrowhead interiors do not visibly cover module fills; each final connector segment is named `_terminal_arrow` for verification.
  - If a process needs multiple branches, assign each branch a separate `fromPort`/`toPort` and a separate waypoint channel.

## JSON Shape Types

Supported `type` values:

- `process` or omitted: rectangle
- `start`, `end`, `terminator`: oval
- `decision`: rotated square diamond with separate horizontal text
- `data`: parallelogram-like polygon fallback when supported, otherwise rectangle

For maximum reliability across localized Visio installs, the bundled script uses primitive drawing APIs first instead of depending on English stencil names.

## When Editing Existing Diagrams

For existing `.vsdx` files:

1. Open with Visio COM.
2. Enumerate pages, shapes, text, and connectors.
3. Apply targeted edits by shape name, shape text, or page index.
4. Save to a new output file unless the user explicitly wants in-place modification.

## Fallbacks

If Visio COM is unavailable:

- Generate Mermaid, SVG, or draw.io XML in the project.
- Tell the user Visio is not installed or not COM-registered.
- Do not pretend a `.vsdx` was generated.
