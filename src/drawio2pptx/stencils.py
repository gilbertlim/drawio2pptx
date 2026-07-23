"""Locate the draw.io desktop binary and render individual cells through it.

draw.io has no per-shape export, so shapes that are not plain rectangles or embedded
bitmaps (AWS/GCP/Cisco stencils and the like) are rendered by draw.io itself and placed
as pictures. Every render is done on the *same* framed canvas so a crop at known
coordinates lands exactly where the shape belongs.
"""
from __future__ import annotations

import os
import platform
import re
import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path

from .model import Cell, Page, frame_cell_xml

_MAC = [
    "/Applications/draw.io.app/Contents/MacOS/draw.io",
    "/Applications/drawio.app/Contents/MacOS/drawio",
    os.path.expanduser("~/Applications/draw.io.app/Contents/MacOS/draw.io"),
]
_LINUX = ["/usr/bin/drawio", "/usr/local/bin/drawio", "/opt/drawio/drawio",
          "/snap/bin/drawio"]
_WIN = [r"C:\Program Files\draw.io\draw.io.exe",
        r"C:\Program Files (x86)\draw.io\draw.io.exe"]

INSTALL_HINT = {
    "Darwin": "brew install --cask drawio",
    "Linux": "sudo snap install drawio   (or grab the .deb/.AppImage from "
             "https://github.com/jgraph/drawio-desktop/releases)",
    "Windows": "winget install JGraph.Draw",
}


class DrawioNotFound(RuntimeError):
    pass


def find_drawio(explicit: str | None = None) -> str:
    """Return a usable draw.io desktop executable path, or raise with install advice."""
    if explicit:
        if Path(explicit).exists():
            return explicit
        raise DrawioNotFound(f"--drawio {explicit} does not exist")
    env = os.environ.get("DRAWIO_BIN")
    if env and Path(env).exists():
        return env
    on_path = shutil.which("drawio") or shutil.which("draw.io")
    if on_path:
        return on_path
    for cand in {"Darwin": _MAC, "Linux": _LINUX, "Windows": _WIN}.get(platform.system(), []):
        if Path(cand).exists():
            return cand
    hint = INSTALL_HINT.get(platform.system(), "see https://www.drawio.com/blog/diagrams-offline")
    raise DrawioNotFound(
        "draw.io desktop was not found — it does the shape rendering, so it is required.\n"
        f"  install:  {hint}\n"
        "  or point at it directly:  --drawio /path/to/drawio  (or set DRAWIO_BIN)"
    )


def run(binary: str, args: list[str], timeout: int = 300) -> None:
    cmd = [binary, "-x", *args, "--no-sandbox", "--disable-gpu"]
    env = {**os.environ, "ELECTRON_DISABLE_SECURITY_WARNINGS": "1"}
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, env=env)
    except subprocess.TimeoutExpired:
        # a draw.io window already holding the user-data-dir makes CLI exports queue forever
        raise RuntimeError(
            f"draw.io did not finish within {timeout}s.\n"
            "  if the draw.io app is open, close it and run this again — a running window "
            "can block CLI exports"
        ) from None
    if proc.returncode != 0:
        raise RuntimeError(f"draw.io export failed ({proc.returncode}):\n"
                           f"{proc.stderr.strip() or proc.stdout.strip()}")


def export_svg(binary: str, src: Path, out: Path, page: int) -> Path:
    # NOTE: the draw.io CLI numbers pages from 1, not 0.
    run(binary, ["-f", "svg", "--crop", "-b", "0", "-p", str(page), "-o", str(out), str(src)])
    if not out.exists():
        raise RuntimeError(f"draw.io produced no SVG for page {page} of {src}")
    return out


def export_png(binary: str, src: Path, out: Path, page: int, scale: int) -> Path:
    run(binary, ["-f", "png", "-s", str(scale), "-t", "-b", "0", "-p", str(page),
                 "-o", str(out), str(src)])
    if not out.exists():
        raise RuntimeError(f"draw.io produced no PNG for page {page} of {src}")
    return out


# --------------------------------------------------------------------------- planning
GROUP_ICON_EXTENT = 64.0   # aws4.group paints its badge in the top-left corner only
BLEED = 4.0                # strokes may spill a hair outside the declared geometry


def needs_render(cell: Cell) -> bool:
    """True when the cell is a stencil that only draw.io itself knows how to draw."""
    shape = cell.shape
    if not shape or cell.box is None:
        return False
    if shape == "image":
        img = cell.st.get("image", "")
        return isinstance(img, str) and img.startswith("data:image/svg")
    return True


def crop_box(cell: Cell) -> tuple[float, float, float, float]:
    x, y, w, h = cell.box
    if cell.is_group:
        w, h = min(w, GROUP_ICON_EXTENT), min(h, GROUP_ICON_EXTENT)
    return (x - BLEED, y - BLEED, w + 2 * BLEED, h + 2 * BLEED)


def _overlap(a, b) -> bool:
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return ax < bx + bw and bx < ax + aw and ay < by + bh and by < ay + ah


def plan_layers(cells: list[Cell]) -> list[list[Cell]]:
    """Group cells so that members of a layer never overlap.

    Each layer costs one draw.io invocation (a few seconds), so packing matters:
    a typical diagram collapses to one or two layers instead of one render per shape.
    """
    layers: list[list[Cell]] = []
    boxes: list[list[tuple]] = []
    for cell in sorted(cells, key=lambda c: (-(c.w * c.h), c.id)):
        cb = crop_box(cell)
        for layer, taken in zip(layers, boxes):
            if not any(_overlap(cb, t) for t in taken):
                layer.append(cell)
                taken.append(cb)
                break
        else:
            layers.append([cell])
            boxes.append([cb])
    return layers


# --------------------------------------------------------------------------- rendering
@dataclass
class Rendered:
    """A cropped PNG plus the diagram-space box it must occupy."""
    cell_id: str
    path: Path
    x: float
    y: float
    w: float
    h: float


_MINI = (
    '<mxfile host="drawio2pptx">{pages}</mxfile>'
)
_PAGE = (
    '<diagram id="l{i}" name="layer{i}"><mxGraphModel grid="0" page="1" '
    'pageWidth="850" pageHeight="1100" math="0" shadow="0"><root>'
    '<mxCell id="0"/><mxCell id="1" parent="0"/>{frame}{cells}'
    "</root></mxGraphModel></diagram>"
)


def _esc(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def _cell_xml(cell: Cell) -> str:
    style = cell.style
    if cell.is_group:
        # drop the border so the crop yields the bare badge; the border is drawn natively
        style = re.sub(r"strokeColor=[^;]*", "strokeColor=none", style)
    return (f'<mxCell id="{_esc(cell.id)}" value="" style="{_esc(style)}" vertex="1" parent="1">'
            f'<mxGeometry x="{cell.x}" y="{cell.y}" width="{cell.w}" height="{cell.h}" '
            f'as="geometry"/></mxCell>')


def render_cells(binary: str, page: Page, frame: tuple[float, float, float, float],
                 workdir: Path, scale: int = 6, progress=lambda *_: None) -> dict[str, Rendered]:
    """Render every stencil cell of `page` and tighten each to its painted pixels."""
    from PIL import Image

    targets = [c for c in page.cells if needs_render(c)]
    if not targets:
        return {}
    layers = plan_layers(targets)
    fx, fy, fw, fh = frame

    mini = workdir / "stencils.drawio"
    mini.write_text(_MINI.format(pages="".join(
        _PAGE.format(i=i, frame=frame_cell_xml(fx, fy, fw, fh),
                     cells="".join(_cell_xml(c) for c in layer))
        for i, layer in enumerate(layers)
    )), encoding="utf-8")

    out: dict[str, Rendered] = {}
    for i, layer in enumerate(layers):
        progress(f"rendering shapes, pass {i + 1}/{len(layers)} ({len(layer)} shapes)")
        sheet = export_png(binary, mini, workdir / f"layer{i}.png", page=i + 1, scale=scale)
        with Image.open(sheet) as im:
            im = im.convert("RGBA")
            # the frame fixes the canvas; draw.io adds a symmetric sub-pixel border
            pad_x = (im.width - fw * scale) / 2
            pad_y = (im.height - fh * scale) / 2
            for cell in layer:
                cx, cy, cw, ch = crop_box(cell)
                box = im.crop((round(pad_x + (cx - fx) * scale), round(pad_y + (cy - fy) * scale),
                               round(pad_x + (cx + cw - fx) * scale),
                               round(pad_y + (cy + ch - fy) * scale)))
                bb = box.getbbox()
                if bb is None:                      # shape rendered nothing (empty stencil)
                    continue
                tight = box.crop(bb)
                dest = workdir / f"cell_{len(out)}.png"
                tight.save(dest, optimize=True)
                out[cell.id] = Rendered(
                    cell_id=cell.id, path=dest,
                    x=cx + bb[0] / scale, y=cy + bb[1] / scale,
                    w=(bb[2] - bb[0]) / scale, h=(bb[3] - bb[1]) / scale,
                )
    return out
