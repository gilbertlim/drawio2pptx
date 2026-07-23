"""Parse a .drawio file into cells, preserving document (paint) order."""
from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path


def parse_style(style: str) -> dict:
    """`a=1;b;c=2` -> {'a': '1', 'b': True, 'c': '2'}. Values may contain '=' and '('."""
    out = {}
    for part in (style or "").split(";"):
        if not part:
            continue
        key, eq, val = part.partition("=")
        out[key] = val if eq else True
    return out


@dataclass
class Cell:
    id: str
    value: str
    style: str
    st: dict
    is_edge: bool
    source: str | None = None
    target: str | None = None
    x: float | None = None
    y: float | None = None
    w: float | None = None
    h: float | None = None
    points: list[tuple[float, float]] = field(default_factory=list)

    @property
    def shape(self) -> str | None:
        s = self.st.get("shape")
        return s if isinstance(s, str) else None

    @property
    def is_group(self) -> bool:
        return self.shape == "mxgraph.aws4.group"

    @property
    def is_text_only(self) -> bool:
        return self.st.get("text") is True and self.st.get("fillColor") in (None, "none")

    @property
    def box(self) -> tuple[float, float, float, float] | None:
        if None in (self.x, self.y, self.w, self.h):
            return None
        return (self.x, self.y, self.w, self.h)


@dataclass
class Page:
    id: str
    name: str
    index: int          # 1-based, matches the drawio CLI `-p` flag
    cells: list[Cell]

    def by_id(self, cid: str) -> Cell | None:
        for c in self.cells:
            if c.id == cid:
                return c
        return None

    def content_bbox(self) -> tuple[float, float, float, float]:
        """Geometric bounds over vertices and edge waypoints (labels not included)."""
        xs, ys = [], []
        for c in self.cells:
            if c.box:
                xs += [c.x, c.x + c.w]
                ys += [c.y, c.y + c.h]
            for px, py in c.points:
                xs.append(px)
                ys.append(py)
        if not xs:
            raise ValueError("page has no positioned cells")
        return min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys)


def _floats(el, *names):
    return [float(el.get(n)) if el is not None and el.get(n) else None for n in names]


def parse(path: str | Path) -> list[Page]:
    """Read every <diagram> page of a .drawio file.

    Compressed (deflate-encoded) diagrams are rejected with an actionable message
    rather than silently producing an empty page.
    """
    root = ET.parse(str(path)).getroot()
    diagrams = root.findall(".//diagram") if root.tag == "mxfile" else []
    if not diagrams and root.tag == "mxGraphModel":
        diagrams = [root]

    pages: list[Page] = []
    for i, dia in enumerate(diagrams, start=1):
        model = dia.find(".//mxGraphModel") if dia.tag == "diagram" else dia
        if model is None:
            if (dia.text or "").strip():
                raise ValueError(
                    f"page {i} of {path} is stored compressed. Open it in draw.io and turn off "
                    "Extras > Edit Diagram > Compressed, or re-save with 'Uncompressed XML'."
                )
            continue
        cells: list[Cell] = []
        for mx in model.findall(".//mxCell"):
            cid = mx.get("id")
            if cid in ("0", "1"):
                continue
            geo = mx.find("mxGeometry")
            x, y, w, h = _floats(geo, "x", "y", "width", "height")
            pts = []
            if geo is not None:
                for p in geo.findall("./Array[@as='points']/mxPoint"):
                    pts.append((float(p.get("x", 0)), float(p.get("y", 0))))
                for p in geo.findall("./mxPoint"):
                    if p.get("as") in ("sourcePoint", "targetPoint"):
                        pts.append((float(p.get("x", 0)), float(p.get("y", 0))))
            style = mx.get("style") or ""
            cells.append(Cell(
                id=cid, value=mx.get("value") or "", style=style, st=parse_style(style),
                is_edge=mx.get("edge") == "1", source=mx.get("source"), target=mx.get("target"),
                x=x, y=y, w=w, h=h, points=pts,
            ))
        pages.append(Page(id=dia.get("id") or f"p{i}", name=dia.get("name") or f"Page-{i}",
                          index=i, cells=cells))
    if not pages:
        raise ValueError(f"no diagram pages found in {path}")
    return pages


FRAME_ID = "__d2p_frame"
_FRAME_TMPL = (
    '<mxCell id="{fid}" value="" style="fillColor=none;strokeColor=none;" vertex="1" parent="1">'
    '<mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/></mxCell>'
)


def frame_cell_xml(x: float, y: float, w: float, h: float, fid: str = FRAME_ID) -> str:
    """An invisible rect that pins the export bounds so every render shares one origin."""
    return _FRAME_TMPL.format(fid=fid, x=x, y=y, w=w, h=h)


_ROOT_OPEN = re.compile(r'(<mxCell\s+id="1"[^>]*/>)')


def with_frame(src_text: str, page_index: int, frame: tuple[float, float, float, float]) -> str:
    """Insert the frame rect into one page of a .drawio document (raw text edit).

    Kept textual on purpose: re-serialising the XML would reorder attributes and
    churn the embedded base64 image payloads.
    """
    x, y, w, h = frame
    inject = frame_cell_xml(x, y, w, h)
    seen = 0

    def repl(m):
        nonlocal seen
        seen += 1
        return m.group(1) + inject if seen == page_index else m.group(1)

    out, n = _ROOT_OPEN.subn(repl, src_text)
    if seen < page_index:
        raise ValueError(f"could not locate page {page_index} root cell to insert the frame")
    return out
