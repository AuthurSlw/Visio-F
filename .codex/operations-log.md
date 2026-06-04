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
