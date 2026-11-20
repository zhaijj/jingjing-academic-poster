# House style (measured from 2023–2026 posters)

## Visual

- Font: Arial only. Title bold (older posters: tool name bold, rest regular → `title_bold: false` + `**PlantCAD2**: rest`).
- Panel: white body + header bar filled `#DEEBF7`, both outlined gray `#AFABAB`, ~6 pt on a 45-in board. Square corners.
- Text black. Emphasis = bold; key numbers in a claim may be **red bold** (`**{r:10 out of 12}**`) — once or twice per poster, not more.
- Bullets: ✓ check marks (default). Perspectives/future-work lists use ❖ (`bullet: diamond`).
- Gray box `#E7E6E6` for a "Challenges"/"Key question" list inside a panel (`box:`).
- Rounded callouts for flowchart-ish labels or the one-line takeaway: `yellow_box` `#FFE599`, `green_box` `#B6D7A8`, `blue_box` `#C9DAF8`.
- Citations under borrowed figures: small, right-aligned, "*Author et al.*, Year, Journal".
- Figures should cover roughly 60–70 % of panel area; text fills the gaps, not the other way round.

Font sizes (pt) the script uses per preset:

| | title | authors | affil. | panel header | subheading | body | caption |
|---|---|---|---|---|---|---|---|
| landscape 45×40 | 96 | 44 | 28 | 72 | 54 | 44 | 28 |
| portrait 33×45 | 66 | 36 | 23 | 60 | 44 | 38 | 24 |
| square 42×42 | 80 | 48 | 36 | 70 | 48 | 46 | 28 |

## Text and voice

- **Title**: one line of intent, ≤ 2 lines on the board. Pattern A: "Tool: what it does" (`reelProtein: filtering functional protein annotations with language model guided machine learning`). Pattern B: verb-led study description (`Exploring regulatory rewiring underlying cold tolerance in grasses by fine-tuning a DNA language model`).
- **Result headings are claims**, full sentences with the finding: "PlantCaduceus identifies deleterious mutations", "reelProtein recovers 85% core proteins", "Core cold-regulatory grammar is shared across 12 grasses". Never "Result 1" or "Figure 3".
- **Motivation** subheadings are questions: "Which genes can be translated into functional proteins?", "Why can foundation models prioritize deleterious mutations?"
- **Method panel** pattern: architecture/pipeline figure on the left, a stack on the right of bold labels each followed by 1–2 ✓ bullets — *Outputs:*, *Model architecture:*, *Masked input:*, *Input:* (adapt labels to the method).
- Bullets: ≤ 2 lines each, ≤ 4 per group, no full paragraphs. Fragments are fine. Cut hedging and methods minutiae; the poster is a conversation starter.
- One takeaway sentence beside or under each result figure.
- Perspectives: 2–4 forward-looking ❖ bullets (apply to X, toolkit for community, link to scores/code).
- Acknowledgements: funding sentence + logos; optionally "See the … talk (#T23) on Sunday morning" and QR codes for bioRxiv/GitHub with `bioRxiv`/`github` logos beside them.
- English throughout, US spelling.

## Defaults (reuse unless told otherwise)

- Presenting author: **Jingjing Zhai**, email `jz963@cornell.edu`.
- Common affiliations:
  - Institute for Genomic Diversity, Cornell University, Ithaca, NY USA 14853
  - Section of Plant Breeding and Genetics, Cornell University, Ithaca, NY USA 14853
  - USDA-ARS; Ithaca, NY, USA 14853
  - Department of Computer Science, Cornell University, Ithaca, NY USA 14853
  - Simons Center for Quantitative Biology, Cold Spring Harbor Laboratory, 1 Bungtown Road, Cold Spring Harbor, NY USA 11724
  - Department of Computer Science, Cornell Tech, New York, NY USA 10044
- Funding / acknowledgement line: keep your own current wording in a local copy of this file, e.g.
  - "This project is supported by [FUNDER 1] ([GRANT #]), [FUNDER 2] ([GRANT #]), and [COMPUTING RESOURCE]."
  Confirm which grants apply before each poster — they change.
- Bundled logos (`assets/logos/`): `usda_ars`, `nsf`, `cornell_seal`, `tacc`, `biorxiv`, `github`. Usual order: USDA-ARS, NSF, Cornell seal, TACC. Add others by passing a file path.
