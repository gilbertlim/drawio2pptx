---
name: drawio2pptx
description: Use when putting a draw.io diagram into PowerPoint, or when a diagram in a deck needs to stay editable. Converts .drawio files into individual native PPT objects - shapes, connectors, text boxes, icons - instead of one flat image. Triggers on "drawio를 ppt에", "구성도 ppt에 넣어줘", "put this diagram in a slide", "each element separately", "요소별로".
---

# drawio2pptx

Converts a draw.io page into individually editable PowerPoint objects.

**Do not hand-roll this conversion.** Pasting a PNG loses editability; rebuilding shapes
by reading the XML by hand gets label placement and edge routing wrong. Use the CLI.

## Use it

```bash
drawio2pptx diagram.drawio                                  # -> diagram.pptx
drawio2pptx diagram.drawio --into deck.pptx --slide 2 --replace
drawio2pptx diagram.drawio --all-pages -o deck.pptx
```

If the command is not installed, run it from this repo without installing anything:

```bash
uv run --with python-pptx --with pillow python -m drawio2pptx.cli diagram.drawio
```

Requires draw.io desktop (`brew install --cask drawio`). It does the shape rendering;
the tool prints install instructions if it is missing.

## Always verify

The conversion is geometric, so mistakes are visual. Run it and look:

```bash
drawio2pptx diagram.drawio --verify check.png
```

`check.png` stacks three panels: draw.io's own export, the generated slide re-rendered
from the saved `.pptx`, and a difference map. Read the bottom panel — **thin outlines
around glyphs and strokes are normal** (the checker's text metrics are not PowerPoint's).
**Solid filled regions mean a shape actually moved** and need investigating.

Then report the object counts the CLI prints, and state that verification was visual.

## Choosing flags

| Situation | Flag |
| --- | --- |
| Korean, Japanese or Chinese labels | `--ea-font "Malgun Gothic"` so text does not reflow on other machines |
| Diagram should not bleed to the slide edge | `--margin 0.04` |
| Deck is 4:3 | `--slide-size 4:3` |
| Multi-page diagram | `--list-pages` first, then `--page N` or `--all-pages` |
| Something looks wrong | `--keep-workdir` and inspect the intermediate renders |

## Tell the user the honest limits

State these rather than letting them discover it:

- **Stencil icons stay raster.** AWS/GCP/Cisco stencils cannot be expressed as PowerPoint
  shapes, so each is placed as its own picture. Movable and resizable, but not recolourable.
- Curved and rounded edges flatten to their anchor points.
- Rotated shapes and swimlanes are not handled.
- Compressed `.drawio` files must be re-saved as uncompressed XML first
  (draw.io: Extras → Edit Diagram, uncheck *Compressed*). The tool says so if it hits one.

## Sensitive diagrams

Architecture diagrams are frequently internal. Never copy a user's real diagram into a
public repo, an example directory, or an artifact without asking first.
