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

Prefer COM automation over mouse/keyboard GUI automation. Use GUI automation only when the task explicitly requires demonstrating visible desktop operation or interacting with existing user-driven Visio UI state.

## Script

Use:

```powershell
& "<skill-dir>\scripts\new_visio_diagram.cmd" `
  -SpecPath ".codex\visio\diagram.json" `
  -OutputPath "diagram.vsdx" `
  -Visible `
  -ExportPdf `
  -ExportPng
```

Important:

- `-Visible` launches Visio visibly. Omit it for hidden automation.
- Use the `.cmd` wrapper by default because many Windows systems block direct `.ps1` execution.
- The script auto-layouts nodes when `x` and `y` are missing.
- Visio uses inches internally. Keep coordinates around a US Letter or A4 page unless the spec sets page size.
- Run the script only on Windows with Microsoft Visio installed and COM registered as `Visio.Application`.

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
