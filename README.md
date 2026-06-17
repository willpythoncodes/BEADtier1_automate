# BEAD Tier 1 Assembly

Build BEAD Tier 1 PowerPoint drafts from a **tokenized template** and local PNG/JPEG assets.

Run all commands below from the **project root** (the folder that contains this README). Paths are relative to that folder unless noted.

This is a standalone tool, separate from the `pptx_style` restyling project.

## Quick start

```bash
python3 -m pip install -e ".[dev]"

tier1-assemble assemble \
  --template "./Blank BEAD Tier1 Template - Tokenized Review.pptx" \
  --manifest "./placeholder_assets_manifest.json" \
  --asset-dir "./placeholder_assets" \
  --output-dir "./output" \
  --report "./output/placeholder_review_assembly_audit.json"
```

**Output:** `./output/{run_id}_Tier1_Report_Draft.pptx`  
Example: `./output/placeholder_review_Tier1_Report_Draft.pptx`

## Setup

```bash
python3 -m pip install -e ".[dev]"
```

## Assemble a draft

### With a manifest (recommended)

Use a JSON manifest when assets need paths or map+legend composites. Asset paths in the manifest resolve relative to the manifest file (or `--scratch-dir` if you pass one).

```bash
tier1-assemble assemble \
  --template "./Blank BEAD Tier1 Template - Tokenized Review.pptx" \
  --manifest "./sample_assets_png_only.json" \
  --output-dir "./output" \
  --report "./output/sample_tier1_assembly_audit.json"
```

### Without a manifest (placeholder review)

Match PNG/JPEG files in `./placeholder_assets/` to each `{token}` by filename:

```bash
tier1-assemble assemble \
  --template "./Blank BEAD Tier1 Template - Tokenized Review.pptx" \
  --asset-dir "./placeholder_assets" \
  --run-id placeholder_review \
  --output-dir "./output"
```

## Combine map + legend

Create a composite PNG before assembly, or let the manifest build it automatically:

```bash
tier1-assemble combine-map-legend \
  --map "./new_run_map_Road Closures.png" \
  --legend "./new_run_legend_Road Closures.png" \
  --position bottom-right \
  --output "./placeholder_assets/hazard_overview_map.png"
```

Valid positions: `bottom-left`, `bottom-right`, `top-left`, `top-right`.

## Where files go

| What | Where |
|------|-------|
| Token images | `./placeholder_assets/` or paths in the manifest |
| Combined map+legend PNGs | Same folder as the token images |
| Finished PowerPoint | `./output/` |
| Audit JSON (optional) | `./output/` via `--report` |

## Project layout

| Path | Purpose |
|------|---------|
| `./Blank BEAD Tier1 Template - Tokenized Review.pptx` | Tokenized master template (use this one) |
| `./placeholder_assets/` | PNG/JPEG assets named after tokens |
| `./placeholder_assets_manifest.json` | Manifest for placeholder review runs |
| `./sample_assets_png_only.json` | Example manifest with composite map assets |
| `./output/` | Generated decks and audit JSON (not in git) |
| `./src/tier1_assemble/` | Assembly tool source code |

### Local archive (not in git)

Older templates, experiments, and reference files can live in `./archive/` on your machine. This folder is not tracked by git.

## How it works (short)

- One `{token}` per shape per slide
- Tokens are found in alt text, visible text, or table cells — including inside grouped shapes
- Each tokenized shape is replaced with its matching PNG/JPEG
- Narrative text boxes without tokens stay editable

## Tests

```bash
python3 -m pytest -q
```
