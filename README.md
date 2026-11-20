# academic-poster

A Claude skill that builds academic conference posters in a consistent house style (light-blue panel headers, Arial, check-mark bullets, funder logos) from a short YAML spec. Output: an editable PowerPoint file plus a print-ready PDF and a PNG preview.

Distilled from my Maize Genetics Meeting, PAG, and 2026 posters.

## Install as a skill

Copy the `academic-poster/` folder into your skills directory (e.g. `~/.claude/skills/`), or zip it and upload it in Claude's skill settings. Then ask Claude for a poster and give it your figures, title, authors, and key messages.

## Use the script directly

```bash
pip install python-pptx pyyaml pillow lxml
# LibreOffice (soffice) and poppler (pdftoppm) are needed for PDF/PNG export
python academic-poster/scripts/build_poster.py my_poster.yaml -o out/
```

- Spec format: [`academic-poster/references/spec.md`](academic-poster/references/spec.md)
- Style rules (layout, type sizes, voice): [`academic-poster/references/style.md`](academic-poster/references/style.md)
- Examples (with gray placeholder figures): [`academic-poster/assets/examples/`](academic-poster/assets/examples/)

Two layouts are supported: classic columns (`columns:`) and a claim grid (`rows:`), with presets for 45×40, 33×45, and 42×42 in boards or any custom size.

## Logos

Institution and funder logos are not included in this repository; each organization controls how its marks are used. Put the PNGs you're allowed to use in `academic-poster/assets/logos/` with these names (or pass file paths in the spec): `usda_ars.png`, `nsf.png`, `cornell_seal.png`, `tacc.png`, `biorxiv.png`, `github.png`. Missing logos are skipped with a warning.
