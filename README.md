# BEAD Tier 1 Assembly

Standalone tool for building BEAD Tier 1 PowerPoint drafts from a tokenized
template and local PNG/JPEG assets. This project is separate from the broader
`pptx_style` restyling tool.

## Setup

```bash
python3 -m pip install -e ".[dev]"
```

## Assemble a draft

```bash
tier1-assemble assemble \
  --template "Blank BEAD Tier1 Template - Tokenized Review.pptx" \
  --manifest sample_assets_png_only.json \
  --output-dir output \
  --report output/sample_tier1_assembly_audit.json
```

If no manifest is provided, PNG/JPEG files in `placeholder_assets/` are matched
by filename to each `{token}` in the template:

```bash
tier1-assemble assemble \
  --template "Blank BEAD Tier1 Template - Tokenized Review.pptx" \
  --asset-dir placeholder_assets \
  --run-id placeholder_review \
  --output-dir output
```

## Combine map + legend

```bash
tier1-assemble combine-map-legend \
  --map "new_run_map_Road Closures.png" \
  --legend "new_run_legend_Road Closures.png" \
  --position bottom-right \
  --output placeholder_assets/hazard_overview_map.png
```

## Project layout

| Path | Purpose |
|------|---------|
| `Blank BEAD Tier1 Template - Tokenized Review.pptx` | Tokenized master template |
| `placeholder_assets/` | PNG/JPEG assets named after tokens |
| `sample_assets_png_only.json` | Example manifest with composite map assets |
| `output/` | Generated Tier 1 draft decks and audit JSON |

Older templates, experiments, and reference files live under `archive/`:

| Path | Contents |
|------|----------|
| `archive/templates/` | Superseded PPTX templates |
| `archive/experiments/` | Manual map+legend placement tests |
| `archive/reference/` | Contact sheet and HTML sources |
| `archive/cache/` | Local junk/cache (not tracked in git) |

Composite map assets are written into the same assets folder used to resolve
token images. Completed decks and audit reports go in `output/`.

## Tests

```bash
python3 -m pytest -q
```
