---
name: academic-poster
description: Build Jingjing Zhai's academic conference posters (Maize Meeting, PAG, ASPB, etc.) in her house style — light-blue panel headers, Arial, check-mark bullets, funder logos — as an editable PowerPoint file plus a print-ready PDF. Use whenever she asks for a poster, wants a poster drafted from figures and key points, says "做个poster / 会议海报", or wants an existing poster resized, re-laid-out, or updated for a new meeting, even if she doesn't say "skill" or name a file format.
---

# Academic poster

Jingjing gives you figures and key messages; you return a finished poster (`.pptx` she can still edit + `.pdf` for printing) that looks like her previous posters. The heavy lifting is done by `scripts/build_poster.py`, which turns a YAML spec into a laid-out PowerPoint slide and exports PDF + a PNG preview. Your job is the part a script can't do: deciding the story, writing tight poster text, and placing figures so the result reads well from 2 m away.

Her posters never include her personal photo — leave it out.

## Workflow

1. **Collect inputs.** You need: the meeting and its board size, title, authors with affiliations, the figure files, and the key messages (one per figure is ideal). Funding/acknowledgement text and any QR codes (bioRxiv, GitHub) are optional. If the board size isn't stated, ask once — it changes everything downstream. Default affiliations, email and funders are in `references/style.md`; reuse them unless told otherwise. If she hands over figures as a multi-panel PDF, split/crop them into separate files (pdftoppm -r 300 with -x/-y/-W/-H) so each can be placed on its own.
2. **Look at every figure** (Read the image). Note its aspect ratio and what it shows. Wide figures want full panel width; tall/square figures pair well with text beside them in a `row`.
3. **Plan the story** before writing YAML. Pick the layout (see "Two layouts" below), assign each figure to a panel, and write the headers. Read `references/style.md` for the voice and text rules — the most important one is that result headings are full-sentence claims.
4. **Write the spec** (format: `references/spec.md`; working examples: `assets/examples/`). Put the YAML next to the figures so relative paths work.
5. **Build:** `python <skill>/scripts/build_poster.py poster.yaml -o out/` → `out/<name>.pptx`, `.pdf`, `_preview.png`. Needs `python-pptx pyyaml pillow lxml`; PDF export needs LibreOffice and `pdftoppm` (poppler). If LibreOffice is missing, the .pptx is still built — she can export the PDF from PowerPoint.
6. **Check visually, then iterate.** Read the preview PNG; for detail, re-render parts of the PDF at higher dpi (`pdftoppm -r 40 -x … -y … -W … -H …`). The script prints warnings — treat them as bugs to fix:
   - *fonts shrunk / overflow* → cut words first (poster text is almost always too long), then give the panel more `height`/`width`.
   - *image squeezed* → give it a larger `weight`, move text beside it in a `row`, or enlarge the panel.
   - *low dpi* → ask her for a higher-resolution export (PDF/SVG/≥300-dpi PNG); mention it rather than silently shipping blurry figures.
   Also look for: empty white bands in a panel (rebalance `height`/`width` weights), figures much smaller than neighbours, headers wrapping to three lines (shorten them). Two or three build–look cycles is normal.
7. **Deliver** the PDF and the PPTX (send them with SendUserFile if that tool exists; otherwise give her the output paths). Say in one line that the PPTX is editable (PowerPoint, Keynote, or Google Slides via import) and list anything she still needs to supply (low-res figure, missing QR code). The PDF is exported by LibreOffice with Liberation Sans (metric-identical to Arial); if she wants the exact Arial glyphs she can re-export from PowerPoint.

## Two layouts (both from her real posters)

**Classic columns** (`columns:`) — Maize Meeting 2023/2024, PAG 2025. Left column: *Motivation/Background* → method panel (named after the model, e.g. "PlantCaduceus: a plant DNA LM"). Right column: *Results* (tall), then *Perspectives* (optional), *Acknowledgements* (funding text + logos + QR codes). Panel headers are short nouns; the claims live in bold subheadings (`h:`) inside panels. Use when there's background to explain and 2–4 results.

**Claim grid** (`rows:`) — her 2026 poster, the most recent style. 2×3 (or 2×2) equal panels; every panel header is a 1–2 line claim ("Motif ablation identified cold-related motifs"); minimal text, one figure per panel, funding + logos in a footer strip. Use when the story is a clean sequence of findings, or when she doesn't specify (it's her newest style).

## Size presets

| preset | board | used for |
|---|---|---|
| `landscape` | 45 × 40 in | Maize Genetics Meeting |
| `portrait` | 33 × 45 in | PAG |
| `square` | 42 × 42 in | 2026 poster |

Other boards: set `size_in: [W, H]`; fonts scale with width. Font sizes per preset are in `references/style.md`; don't override them unless a panel genuinely needs it — consistent type sizes are part of what makes her posters look clean.
