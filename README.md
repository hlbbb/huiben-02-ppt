# huiben-02-ppt

Codex skill for turning picture-book source pages or illustrations into a per-spread, editable PPTX.

The default pipeline:

1. generate clean text-free 16:9 artwork with local ComfyUI
2. compose the original Chinese text directly onto the artwork with Python/Pillow
3. build a PPTX with editable text boxes while keeping the image full-bleed
4. export a shot-duration plan for the downstream video skill

## Files

- `SKILL.md` — workflow, layout constraints, and acceptance gates
- `scripts/` — ComfyUI client, text composition, PPTX builder, contact sheets, and self-tests
- `references/layout-spec.md` — layout rules
- `evals/evals.json` — example evaluation cases

## Install

Copy this folder into your Codex skills directory, for example:

```text
~/.codex/skills/huiben-02-ppt
```

or:

```text
~/.agents/skills/huiben-02-ppt
```

The workflow expects a local ComfyUI instance and Python packages such as Pillow and python-pptx. Review the machine-specific paths in `SKILL.md` and `scripts/comfy_ppt/example-config.json` before running.
