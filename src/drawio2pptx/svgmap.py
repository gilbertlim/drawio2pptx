"""Read exact label boxes and edge routes out of draw.io's own SVG export.

Re-implementing mxGraph's label placement and orthogonal edge router would be a
losing game. The SVG export already contains the answer, tagged per cell with
`data-cell-id`, so we read the geometry back out instead of recomputing it.
"""
from __future__ import annotations

import html
import re
from dataclasses import dataclass, field
from html.parser import HTMLParser


@dataclass
class Run:
    """A stretch of label text sharing one set of character properties."""
    text: str
    size: float                 # px, as drawn
    bold: bool
    color: str | None           # '#rrggbb'


@dataclass
class Label:
    # one entry per rendered line; a line can mix sizes, e.g. a 20px emoji then 12px text
    lines: list[list[Run]]
    x: float
    y: float
    w: float
    h: float
    size: float                 # tallest run, used for line spacing
    bold: bool
    color: str | None
    bg: str | None
    align: str                  # 'left' | 'center'
    family: str = "Helvetica"

    @property
    def text(self) -> str:
        return "\n".join("".join(r.text for r in line) for line in self.lines)


@dataclass
class Paint:
    """Resolved fill and stroke for a plain shape, straight from the SVG."""
    fill: str | None            # '#rrggbb', or None for no fill
    stroke: str | None
    stroke_width: float
    dashed: bool
    radius: float               # corner radius in px; 0 for a square rectangle


@dataclass
class EdgeRoute:
    points: list[tuple[float, float]]
    color: str
    width: float
    # draw.io draws arrowheads at a fixed size. PowerPoint scales its own with the line
    # width, so a hairline connector gets a tip too small to see and a short one loses it
    # entirely. Carry draw.io's actual arrowhead outlines and emit them as filled subpaths.
    arrowheads: list[list[tuple[float, float]]] = field(default_factory=list)


_CELL_SPLIT = re.compile(r'(?=<g data-cell-id=")')
_CELL_ID = re.compile(r'<g data-cell-id="([^"]+)"')
_VIEWBOX = re.compile(r'<svg[^>]*\bviewBox="([\d.\- ]+)"')
_SVG_SIZE = re.compile(r'<svg[^>]*\bwidth="([\d.]+)px"[^>]*\bheight="([\d.]+)px"')


class SvgMap:
    """Per-cell view of a draw.io SVG export."""

    def __init__(self, svg_text: str):
        self.text = svg_text
        self.blocks: dict[str, str] = {}
        for part in _CELL_SPLIT.split(svg_text):
            m = _CELL_ID.match(part)
            if m:
                self.blocks[m.group(1)] = part

    @classmethod
    def load(cls, path) -> "SvgMap":
        return cls(open(path, encoding="utf-8").read())

    @property
    def size(self) -> tuple[float, float]:
        m = _SVG_SIZE.search(self.text)
        if m:
            return float(m.group(1)), float(m.group(2))
        m = _VIEWBOX.search(self.text)
        if m:
            _, _, w, h = (float(v) for v in m.group(1).split())
            return w, h
        raise ValueError("could not read the SVG canvas size")

    def paint(self, cell_id: str) -> "Paint | None":
        """How draw.io actually painted a plain shape.

        The style dict is not enough: draw.io falls back to a white fill and a black
        border when `fillColor` / `strokeColor` are absent, so reading "missing" as
        "none" erases the shape. The SVG carries the resolved values, and the corner
        radius with them.
        """
        block = self.blocks.get(cell_id)
        if not block:
            return None
        m = re.search(r"<rect\s([^>]*?)/>", block)
        if not m:
            return None
        attrs = m.group(1)

        def get(key):
            found = re.search(rf'\b{key}="([^"]*)"', attrs)
            return found.group(1) if found else None

        width = get("stroke-width")
        radius = get("rx")
        return Paint(
            fill=normalize_color(get("fill")),
            stroke=normalize_color(get("stroke")),
            stroke_width=float(width) if width else 1.0,
            dashed=bool(get("stroke-dasharray")),
            radius=float(radius) if radius else 0.0,
        )

    def frame_rect(self, frame_id: str) -> tuple[float, float, float, float]:
        """Where the injected frame rect actually landed in this export.

        Everything else is measured against it, so the mapping stays correct even if
        draw.io grew the canvas because a label spilled past the frame.
        """
        block = self.blocks.get(frame_id)
        if not block:
            raise ValueError(f"frame cell {frame_id} is missing from the SVG export")
        m = re.search(r'<rect\s+x="([\d.\-]+)"\s+y="([\d.\-]+)"\s+'
                      r'width="([\d.]+)"\s+height="([\d.]+)"', block)
        if not m:
            raise ValueError(f"frame cell {frame_id} did not render a rect")
        return tuple(float(v) for v in m.groups())

    # ------------------------------------------------------------------ labels
    def label(self, cell_id: str) -> Label | None:
        block = self.blocks.get(cell_id)
        if not block:
            return None
        return _label_from_text(block) or _label_from_foreign_object(block)

    # ------------------------------------------------------------------- edges
    def edge(self, cell_id: str) -> EdgeRoute | None:
        block = self.blocks.get(cell_id)
        if not block:
            return None
        line = None
        heads = []
        for d, attrs in re.findall(r'<path d="([^"]+)"((?:(?!/>).)*)/>', block, re.S):
            fill = re.search(r'\bfill="([^"]+)"', attrs)
            pts = _path_points(d)
            if fill and fill.group(1) == "none":
                if line is None and len(pts) >= 2:
                    line = (pts, attrs)
            elif len(pts) >= 3:
                heads.append(pts)
        if line is None:
            return None
        pts, attrs = line
        stroke = (re.search(r'stroke="([^"]+)"', attrs) or [None, "#000000"])[1]
        sw = re.search(r'stroke-width="([\d.]+)"', attrs)
        return EdgeRoute(points=pts, color=normalize_color(stroke) or "#000000",
                         width=float(sw.group(1)) if sw else 1.0, arrowheads=heads)


# ------------------------------------------------------------------- internals
def _path_points(d: str) -> list[tuple[float, float]]:
    """Flatten an M/L/C path to its anchor points (curves keep only their endpoint)."""
    toks = d.replace(",", " ").split()
    pts, i = [], 0
    while i < len(toks):
        t = toks[i]
        if t in ("M", "L"):
            pts.append((float(toks[i + 1]), float(toks[i + 2])))
            i += 3
        elif t == "C":
            pts.append((float(toks[i + 5]), float(toks[i + 6])))
            i += 7
        else:
            i += 1
    return pts


def normalize_color(value: str | None) -> str | None:
    """'#abc123' / 'rgb(1,2,3)' / 'light-dark(a, b)' -> '#abc123' (light half wins)."""
    if not value:
        return None
    v = value.strip()
    m = re.match(r"light-dark\((.*),(.*)\)$", v)
    if m:
        v = m.group(1).strip()
    if re.match(r"#[0-9a-fA-F]{6}$", v):
        return v.upper()
    m = re.match(r"rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", v)
    if m:
        return "#%02X%02X%02X" % tuple(int(g) for g in m.groups())
    return None


def strip_markup(s: str) -> str:
    s = re.sub(r"<br\s*/?>", "\n", s, flags=re.I)
    s = re.sub(r"</div>", "", s, flags=re.I)
    s = re.sub(r"<div[^>]*>", "\n", s, flags=re.I)
    s = re.sub(r"<[^>]+>", "", s)
    return html.unescape(s).replace("\xa0", " ")


_TEXT_GROUP = re.compile(r'<g fill="([^"]*)" font-family="([^"]*)"([^>]*)>(.*?)</g>', re.S)
_BG_RECT = re.compile(r'<rect\s+fill="([^"]*)"[^>]*?x="([\d.\-]+)"\s+y="([\d.\-]+)"\s+'
                      r'width="([\d.]+)"\s+height="([\d.]+)"')


def _label_from_text(block: str) -> Label | None:
    """Plain <text> labels. draw.io emits a background <rect> that is the exact block box."""
    m = _TEXT_GROUP.search(block)
    if not m or "<text" not in m.group(4):
        return None
    attrs, inner = m.group(3), m.group(4)
    fs_m = re.search(r'font-size="([\d.]+)px"', attrs)
    if not fs_m:
        return None
    size = float(fs_m.group(1))
    texts = [strip_markup(t) for t in re.findall(r"<text[^>]*>(.*?)</text>", inner, re.S)]
    if not texts:
        return None
    anchor = (re.search(r'text-anchor="([^"]+)"', attrs) or [None, "start"])[1]
    rect = _BG_RECT.search(inner)
    if rect:
        bg = rect.group(1)
        x, y, w, h = (float(v) for v in rect.groups()[1:])
    else:
        bg = None
        xs = [float(re.search(r'x="([\d.\-]+)"', t).group(1))
              for t in re.findall(r"<text[^>]*>", inner)]
        ys = [float(re.search(r'y="([\d.\-]+)"', t).group(1))
              for t in re.findall(r"<text[^>]*>", inner)]
        w = max(len(t) for t in texts) * size * 0.62
        x = min(xs) - w / 2 if anchor == "middle" else min(xs)
        y, h = min(ys) - size, len(texts) * size * 1.25
    bold = 'font-weight="bold"' in attrs
    colour = normalize_color(m.group(1))
    lines = [[Run(t, size, bold, colour)] for t in texts]
    return Label(lines=lines, x=x, y=y, w=w, h=h, size=size,
                 bold=bold,
                 color=normalize_color(m.group(1)), bg=normalize_color(bg),
                 align="center" if anchor == "middle" else "left",
                 family=m.group(2).split(",")[0].strip() or "Helvetica")


_FO = re.compile(r"<foreignObject.*?</foreignObject>", re.S)
_FO_IMG = re.compile(r'</foreignObject>\s*<image\s+x="([\d.\-]+)"\s+y="([\d.\-]+)"\s+'
                     r'width="([\d.]+)"\s+height="([\d.]+)"')
_INNER = re.compile(r'<div style="display: inline-block;[^"]*">(.*)</div>\s*</div>\s*</div>', re.S)
# CSS colours: skip `background-color`, `text-decoration-color`, ... but keep bare `color`
_COLOR = re.compile(r'(?<![-\w])color:\s*(light-dark\([^)]*\)[^;"]*\)?|#[0-9a-fA-F]{6}'
                    r'|rgb\([^)]*\))')


class _RunParser(HTMLParser):
    """Split a draw.io HTML label into runs, tracking nested character styling.

    Labels commonly mix sizes inside one line — `<font style="font-size: 20px">emoji</font>
    Text` — so collapsing a line to a single size blows up the smaller half.
    """

    BREAKS = {"div", "p", "li"}

    def __init__(self, size: float, bold: bool, color: str | None):
        super().__init__(convert_charrefs=True)
        self.stack = [(size, bold, color)]
        self.lines: list[list[Run]] = [[]]

    def _push(self, tag, attrs):
        size, bold, color = self.stack[-1]
        style = dict(attrs).get("style") or ""
        m = re.search(r"font-size:\s*([\d.]+)px", style)
        if m:
            size = float(m.group(1))
        m = re.search(r"font-weight:\s*(\w+)", style)
        if m:
            bold = m.group(1) in ("bold", "700", "800", "900")
        if tag in ("b", "strong"):
            bold = True
        m = _COLOR.search(style)
        if m:
            color = normalize_color(m.group(1)) or color
        self.stack.append((size, bold, color))

    def handle_starttag(self, tag, attrs):
        if tag == "br":
            self.lines.append([])
            return
        if tag in self.BREAKS and self.lines[-1]:
            self.lines.append([])
        self._push(tag, attrs)

    def handle_startendtag(self, tag, attrs):
        if tag == "br":
            self.lines.append([])

    def handle_endtag(self, tag):
        if tag != "br" and len(self.stack) > 1:
            self.stack.pop()

    def handle_data(self, data):
        text = data.replace("\xa0", " ")
        if not text:
            return
        if not text.strip():
            if self.lines[-1]:          # keep separators between runs, drop layout whitespace
                self.lines[-1].append(Run(" ", *self.stack[-1]))
            return
        self.lines[-1].append(Run(text, *self.stack[-1]))

    def result(self) -> list[list[Run]]:
        out = []
        for line in self.lines:
            merged: list[Run] = []
            for run in line:
                prev = merged[-1] if merged else None
                if prev and (prev.size, prev.bold, prev.color) == (run.size, run.bold, run.color):
                    merged[-1] = Run(prev.text + run.text, run.size, run.bold, run.color)
                else:
                    merged.append(run)
            if not merged or not "".join(r.text for r in merged).strip():
                continue
            first, last = merged[0], merged[-1]
            merged[0] = Run(first.text.lstrip(), first.size, first.bold, first.color)
            merged[-1] = Run(merged[-1].text.rstrip(), last.size, last.bold, last.color)
            out.append(merged)
        return out


def _label_from_foreign_object(block: str) -> Label | None:
    """HTML labels. The <image> fallback right after the foreignObject is the exact box."""
    fo = _FO.search(block)
    img = _FO_IMG.search(block)
    if not fo or not img:
        return None
    fob = fo.group(0)
    x, y, w, h = (float(v) for v in img.groups())
    body = _INNER.search(fob)
    if not body:
        return None

    # the inline-block wrapper carries the base styling; nested runs override it
    wrapper = re.search(r'<div style="display: inline-block;([^"]*)"', fob)
    base = wrapper.group(1) if wrapper else ""
    m = re.search(r"font-size:\s*([\d.]+)px", base)
    size = float(m.group(1)) if m else 12.0
    m = re.search(r"font-weight:\s*(\w+)", base)
    bold = bool(m) and m.group(1) in ("bold", "700")
    m = _COLOR.search(base)
    color = normalize_color(m.group(1)) if m else None

    parser = _RunParser(size, bold, color)
    parser.feed(body.group(1))
    lines = parser.result()
    if not lines:
        return None

    families = re.findall(r"font-family:\s*([^;\"]+)", fob)
    bg = None
    for c in re.findall(r"background-color:\s*(#[0-9a-fA-F]{6})", fob):
        bg = normalize_color(c) or bg
    return Label(
        lines=lines, x=x, y=y, w=w, h=h,
        size=max(r.size for line in lines for r in line),
        bold=bold, color=color, bg=bg,
        align="left" if "justify-content: unsafe flex-start" in fob else "center",
        family=(families[-1].split(",")[0].strip() if families else "Helvetica"),
    )
