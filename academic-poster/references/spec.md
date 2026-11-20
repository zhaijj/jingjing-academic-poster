# Poster spec (YAML) for build_poster.py

```yaml
preset: square            # landscape | portrait | square   (or size_in: [W, H])
title: "PlantCAD2: ..."   # \n forces a line break
title_bold: true          # false = regular title; use **...** to bold a part
authors:
  - {name: Jingjing Zhai, aff: [1]}          # first author is bolded automatically
  - {name: Edward S. Buckler, aff: [1, 2, 3]}
affiliations: [ "Institute for Genomic Diversity, ...", "..." ]   # numbered in order
email: jz963@cornell.edu
footer:                   # optional bottom strip (grid style)
  text: "This project is supported by ..."
  logos: [usda_ars, nsf, cornell_seal, tacc]   # names in assets/logos or file paths
  logo_height_in: 1.25
output_name: my_poster    # optional; default = spec file name
fonts: {body: 40}         # optional per-role overrides (title, authors, aff, header, sub, body, caption)
margin_in: 0.6            # optional
gap_in: 0.45              # optional

# EITHER classic columns ...
columns:
  - width: 1.1            # relative
    panels:
      - title: Motivation
        height: 1.2       # relative within the column
        blocks: [ ... ]
# ... OR a claim grid (headers aligned per row)
rows:
  - height: 1             # relative
    panels:
      - title: "Claim sentence for this panel"
        width: 1          # relative within the row
        blocks: [ ... ]
```

## Blocks (stacked top to bottom inside a panel)

| block | keys | notes |
|---|---|---|
| subheading | `h` | bold, subheading size; `align` |
| text | `text`, `bold`, `align`, `scale` | blank line (`\n\n`) = new paragraph |
| bullets | `bullets: [..]`, `bullet: check\|diamond\|arrow\|dot\|none` | can be combined with `h`/`text` in one block (heading → text → bullets) |
| image | `image`, `caption`, `caption_align`, `width` (0–1 of panel), `align`, `weight`, `height_in`, `trim` | png/jpg/tif/pdf/svg. Flexible height: images share leftover space by `weight` (default 1) and never exceed their aspect-ratio size. White margins trimmed unless `trim: false`. |
| row | `row: [blocks]`, each child may have `w` (relative width) | side-by-side; children vertically centred. Typical: `[{image, w: 1.4}, {bullets, w: 1}]` |
| stack | `stack: [blocks]` | vertical group, mainly as a `row` child (e.g. several labelled bullet groups beside a figure) |
| gray box | `box: "Box title"` + `bullets`/`text`, `fill` | gray `#E7E6E6` by default |
| callout | `callout: "text"`, `fill`, `bold`, `align` | rounded box; fills `yellow_box`, `green_box`, `blue_box`, `gray_box`, `orange_box`, `pink_box` or hex |
| caption | `caption` alone | small gray right-aligned line (citations) |
| logos | `logos: [..]`, `height_in`, `align` | e.g. inside an Acknowledgements panel; also QR code image paths |
| spacer | `spacer: inches` | |

## Inline markup (any text)

`**bold**`, `*italic*` (species names), `^superscript^`, `~subscript~`, `{r:red}` `{b:blue}` `{g:green}` `{o:orange}` `{gray:...}`, `\*` for a literal asterisk. YAML tip: quote strings containing `:` or starting with `*`/`{`.

## Layout behaviour

- Panel body height is fixed by the column/row weights; text is measured exactly (Arial metrics), images absorb the remaining space.
- If text doesn't fit, the panel's fonts shrink in 4 % steps down to 72 % with a warning — shorten text instead of accepting this.
- Leftover vertical space is spread as gaps between blocks (capped), so lightly filled panels still look balanced.
- Warnings (stderr + JSON `warnings`): font shrink, overflow, image < 2 in tall, image < 100 dpi at print size.
