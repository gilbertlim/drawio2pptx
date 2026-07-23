# drawio2pptx

Put a draw.io diagram into PowerPoint as **individually editable objects**, not a flat image.

Boxes stay boxes, arrows stay arrows, text stays text. Move one server icon, recolour a zone,
or fix a typo in a label without going back to draw.io and re-exporting.

![draw.io on the left, the same diagram as separate PowerPoint objects on the right](examples/preview.png)

<sub>Right-hand panel is the tool's own re-render of the saved `.pptx` (see
[Checking the result](#checking-the-result)), not a PowerPoint screenshot.</sub>

There are two ways to run it: ask Claude Code in plain language, or call the Python CLI
yourself. Both do the same work, so pick whichever fits how you're already working.

---

## First, either way: install draw.io desktop

draw.io does the shape rendering, so it has to be on the machine. There's no way around this
one, because PowerPoint cannot express an mxGraph stencil and nothing else draws them right.

```bash
brew install --cask drawio            # macOS
sudo snap install drawio              # Linux
winget install JGraph.Draw            # Windows
```

Auto-detection covers the standard install locations. If yours sits somewhere unusual, set
`DRAWIO_BIN` or pass `--drawio /path/to/drawio`.

---

# Way 1 — prompting

Good when you're already in Claude Code and the diagram is one step of a bigger job, like
assembling a report deck.

### Setup, once

```bash
git clone https://github.com/gilbertlim/drawio2pptx && cd drawio2pptx
uv tool install .                                          # puts drawio2pptx on PATH
ln -s "$PWD/.claude/skills/drawio2pptx" ~/.claude/skills/drawio2pptx
```

The symlink makes the skill visible from every project. Inside this repo it's picked up from
`.claude/skills/` regardless.

### Then say what you want

> 이 구성도 ppt에 넣어줘

> put architecture.drawio on slide 2 of the report deck, replacing what's there

> 페이지별로 슬라이드 하나씩 만들어줘

Claude reads [the skill](.claude/skills/drawio2pptx/SKILL.md), picks the flags, runs the
conversion, and looks at the output before telling you it's done. The skill also tells it to
state the limits up front instead of letting you find them later, and never to copy your
diagram into a public repo or a shared artifact without asking first.

You get back the object counts, plus a `check.png` when verification was asked for:

```
report.pptx  (slide 2, 104 objects: 10 shapes, 22 connectors, 39 text boxes, 33 icons)
```

---

# Way 2 — Python

Good for scripting, batch runs, or when you want to see exactly which flags ran.

### Install

```bash
git clone https://github.com/gilbertlim/drawio2pptx && cd drawio2pptx
pip install .
```

To keep it out of the system Python use `uv tool install .` or `pipx install .`. To run it
from the clone with nothing installed at all:

```bash
uv run --with python-pptx --with pillow python -m drawio2pptx diagram.drawio
```

### Command line

```bash
# diagram.drawio -> diagram.pptx, one 16:9 slide. That's the whole thing.
drawio2pptx diagram.drawio

# name the output
drawio2pptx diagram.drawio -o deck.pptx

# drop it onto slide 2 of a deck you already have, wiping that slide first
drawio2pptx diagram.drawio --into deck.pptx --slide 2 --replace

# every page of a multi-page diagram, one slide each
drawio2pptx diagram.drawio --all-pages -o deck.pptx

# convert it and show me whether it came out right
drawio2pptx diagram.drawio --verify check.png
```

| Flag | Purpose |
| --- | --- |
| `--page N` / `--all-pages` / `--list-pages` | Multi-page diagrams |
| `--into DECK --slide N [--replace]` | Insert into a deck you already have |
| `--slide-size 16:9 \| 4:3 \| 16:10 \| auto \| 13.333x7.5` | Slide dimensions for a new deck |
| `--margin 0.04` | Leave a border instead of filling the slide |
| `--ea-font "Apple SD Gothic Neo"` | Pin the CJK font so Korean or Japanese labels don't reflow elsewhere |
| `--font Arial` | Override the latin font |
| `--scale 8` | Higher-resolution icons (default 6, roughly 370 dpi at slide width) |
| `--drawio PATH` | If auto-detection misses it |
| `--keep-workdir` | Keep the intermediate renders when something looks wrong |

`drawio2pptx --help` has the rest.

### As a library

```python
from drawio2pptx import convert

result = convert("diagram.drawio", "deck.pptx", slide_size="16:9", margin=0.03)
print(result.path, result.counts)
# deck.pptx {'rect': 10, 'picture': 28, 'line': 22, 'text': 39}
```

`convert()` returns a `Result` carrying the output path, the slide it wrote to, the object
counts, and the content bounds in diagram coordinates.

---

## What actually comes out

| draw.io | PowerPoint |
| --- | --- |
| Plain rectangles and containers | Native autoshapes, with fill, border and dash intact |
| Labels | Real text boxes, keeping font, size, weight and colour |
| Edges, including orthogonal routes | Freeform connectors with the original arrowheads |
| AWS / GCP / Cisco / Veeam stencils | One high-resolution PNG per icon, placed to the pixel |
| Embedded images | Extracted at original resolution, crops honoured |

Stencil icons are the one thing that stays raster. Each is a separate picture you can move and
resize; you just can't recolour it in PowerPoint. Everything else is native.

## Checking the result

Conversion errors are visual, so look at them:

```bash
drawio2pptx diagram.drawio --verify check.png
```

`check.png` stacks three panels: draw.io's own export, the generated slide re-rendered from
the saved `.pptx`, and a difference map.

Read the bottom panel with one rule in mind. Thin outlines around glyphs and strokes are
normal, because the checker's text layout isn't PowerPoint's and never will be. **Solid filled
regions mean a shape actually moved.**

## How it works

Three ideas carry the whole thing.

**A frame rectangle pins the coordinate system.** draw.io always crops its exports to the
graph bounds, so two renders of one diagram don't share an origin and you can't reliably crop
a single icon out of a sheet. An invisible rect gets injected into a copy of the diagram
before every export, then read back from the SVG to see where it actually landed. That
reading is the origin, which is why a label overflowing the frame grows the canvas without
breaking anything.

**Label boxes and edge routes are read out of draw.io's SVG export**, keyed by `data-cell-id`.
mxGraph's label placement and orthogonal router aren't worth reimplementing, and the SVG
already holds the answer.

**Non-overlapping stencils share a render pass.** Each draw.io invocation costs a few seconds.
Packing shapes into layers by overlap took a 19-icon diagram from 19 passes to 2, about 75
seconds down to 26.

## Limits

- Curved and rounded edges flatten to their anchor points.
- Stencil icon colours can't be changed in PowerPoint.
- Rotated shapes and swimlane containers aren't handled yet, and they'll place wrongly rather
  than raise an error.
- Compressed `.drawio` files need re-saving as uncompressed XML first (draw.io:
  **Extras → Edit Diagram**, uncheck *Compressed*). The tool tells you when it hits one.

## Development

```bash
pip install -e ".[dev]"
pytest -q
ruff check .
```

The end-to-end tests need draw.io desktop and skip themselves without it, leaving the parsing
and geometry tests to run anywhere. [CLAUDE.md](CLAUDE.md) records the invariants worth
knowing before you change anything.

Regenerate the README image after touching the renderer:

```bash
drawio2pptx examples/sample.drawio -o /tmp/s.pptx --verify /tmp/check.png
python tools/make_preview.py /tmp/check.png examples/preview.png "PowerPoint — 31 separate objects"
```

## License

MIT
