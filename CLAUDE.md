# Working on drawio2pptx

Converts one draw.io page into separate native PowerPoint objects. README explains what it
does for users; this file is about not breaking it.

## Run it

```bash
pip install -e ".[dev]"
pytest -q                 # 34 tests; the end-to-end ones self-skip without draw.io desktop
ruff check .
drawio2pptx examples/sample.drawio -o /tmp/s.pptx --verify /tmp/check.png
```

The installed copy on this machine came from `uv tool install .`, so after changing the
source run `uv tool install --reinstall .` or the `drawio2pptx` on PATH stays stale.

## The one invariant everything rests on

draw.io's PNG and SVG exports always crop to the graph bounds, so the origin shifts
whenever the content changes. Two renders of the same diagram do not share a coordinate
system, which makes cropping a single icon out of a sheet impossible by itself.

The fix: `model.with_frame()` injects an invisible rect (`__d2p_frame`) into a copy of the
diagram before every export. `svgmap.SvgMap.frame_rect()` then reads back where that rect
actually landed, and that reading — not the frame you asked for — is the origin. Reading it
back rather than assuming it means an overflowing label just grows the canvas without
breaking the mapping.

All element geometry lives in **SVG coordinates**. Cell boxes come from the model in
draw.io coordinates and get shifted in `build.collect()`; labels and edge routes already
arrive in SVG space. Mixing the two silently offsets everything by the frame padding, and
the tests will not catch it because they check bounds, not exact placement. Run `--verify`.

## Things that cost real time to find

- **`-p` on the draw.io CLI is 1-based.** Passing 0 silently gives you page 1, so an
  off-by-one shows up as the wrong shape rendered rather than an error.
- **Image crops live in a separate style key**, `clipPath=inset(top% right% bottom% left%)`,
  not inside the `image=data:...` URI. Skip it and cropped logos come out shrunken and
  offset. `build.extract_bitmap()` handles it.
- **Colours arrive as `light-dark(light, dark)`.** Take the first half. `svgmap.normalize_color()`
  is the only place that should parse a colour.
- **In HTML labels the innermost declaration wins.** An outer div saying `font-weight: bold`
  gets overridden by a `<span style="font-weight: normal">` inside it, so `_label_from_foreign_object()`
  takes the *last* match for weight, size and colour. Same for `font-family`.
- **python-pptx autoshapes and freeforms inherit `effectRef` from the theme**, which paints a
  drop shadow on every box and connector. `_Emitter._no_theme_effects()` kills it and
  `test_no_shape_inherits_a_theme_shadow` guards it; anything new that calls `add_shape` or
  `build_freeform` needs the same treatment.
- **`a:ln` children are order-sensitive.** solidFill, then prstDash, then headEnd, then
  tailEnd. Appending out of order produces a file PowerPoint offers to repair.
- **PowerPoint and Keynote AppleScript automation are blocked on this machine** — `open`
  succeeds and leaves zero documents. Don't burn time on it; `verify.py` exists because of it.

## Layer packing

`stencils.plan_layers()` packs shapes that don't overlap into one draw.io invocation, since
each invocation costs a few seconds. The invariant is that no cell's crop box may touch
another cell's painted pixels in the same layer, or the crop picks up its neighbour.

`aws4.group` is the awkward case: the cell box is the whole container, but only a ~25px badge
in the top-left corner gets painted (the border is drawn natively instead, with `strokeColor`
forced to `none` in the render copy). `crop_box()` shrinks group cells to `GROUP_ICON_EXTENT`
so shapes sitting inside the container don't collide with it. Widening that constant will
start capturing children.

## verify.py is deliberately approximate

It redraws the saved `.pptx` with PIL to catch shapes in the wrong place. It is not trying to
be PowerPoint, and its text metrics differ — thin outlines around glyphs and strokes in the
difference panel are expected and fine. Solid filled regions mean something moved. Don't
"fix" the glyph-edge noise; that direction leads to reimplementing a text layout engine.

## Adding support for a new shape type

`stencils.needs_render()` decides what draw.io has to draw versus what becomes a native
object. Default to native only when PowerPoint can express the thing exactly: a rectangle
with a fill, a border and a dash pattern. Everything else renders. Rotated shapes and
swimlanes are unhandled and will place wrongly rather than fail loudly, so if you touch
those, add a case to `collect()` and a test.

## Never commit someone's real diagram

`examples/sample.drawio` is deliberately generic (DMZ, WAS, DBMS). Architecture diagrams from
actual work are internal, sometimes regulator-facing, and must not land in this repo, in a
test fixture, or in a published artifact. If you need a real diagram to reproduce a bug, keep
it under `fixtures-private/`, which is gitignored.
