"""Unit tests for the parts that need no draw.io install."""
from pathlib import Path

import pytest

from drawio2pptx import model, stencils
from drawio2pptx.svgmap import SvgMap, normalize_color, strip_markup

SAMPLE = Path(__file__).parent.parent / "examples" / "sample.drawio"


# ------------------------------------------------------------------ style parsing
def test_parse_style_handles_flags_and_values():
    st = model.parse_style("shape=image;html=1;dashed;fillColor=none;")
    assert st["shape"] == "image"
    assert st["dashed"] is True
    assert st["fillColor"] == "none"


def test_parse_style_keeps_parentheses_in_values():
    st = model.parse_style("clipPath=inset(20.41% 20% 20.41% 22.67%);html=1")
    assert st["clipPath"] == "inset(20.41% 20% 20.41% 22.67%)"


# ------------------------------------------------------------------ colours / text
@pytest.mark.parametrize("raw,expected", [
    ("#dae8fc", "#DAE8FC"),
    ("rgb(0, 0, 0)", "#000000"),
    ("light-dark(#999999, #6a6a6a)", "#999999"),
    ("light-dark(rgb(0, 0, 0), rgb(237, 237, 237))", "#000000"),
    ("none", None),
    (None, None),
])
def test_normalize_color(raw, expected):
    assert normalize_color(raw) == expected


def test_strip_markup_turns_divs_into_lines():
    assert strip_markup("LLM<div>call</div>") == "LLM\ncall"
    assert strip_markup("a<br/>b") == "a\nb"
    assert strip_markup("&amp;nbsp;") == "&nbsp;"


# ------------------------------------------------------------------ model
def test_parse_sample_page():
    pages = model.parse(SAMPLE)
    assert len(pages) == 1
    page = pages[0]
    assert page.index == 1
    assert page.cells
    x, y, w, h = page.content_bbox()
    assert w > 0 and h > 0


def test_group_and_text_classification():
    page = model.parse(SAMPLE)[0]
    assert any(c.is_group for c in page.cells), "sample should contain an aws4.group"
    assert all(not (c.is_group and c.is_text_only) for c in page.cells)


def test_with_frame_injects_exactly_one_cell():
    text = SAMPLE.read_text(encoding="utf-8")
    out = model.with_frame(text, 1, (0, 0, 100, 100))
    assert out.count(model.FRAME_ID) == 1
    assert 'width="100" height="100"' in out


def test_with_frame_rejects_missing_page():
    with pytest.raises(ValueError):
        model.with_frame(SAMPLE.read_text(encoding="utf-8"), 5, (0, 0, 10, 10))


# ------------------------------------------------------------------ layer packing
def _cell(cid, x, y, w, h, style="shape=mxgraph.aws4.resourceIcon"):
    return model.Cell(id=cid, value="", style=style, st=model.parse_style(style),
                      is_edge=False, x=x, y=y, w=w, h=h)


def test_disjoint_shapes_share_one_render_pass():
    cells = [_cell("a", 0, 0, 40, 40), _cell("b", 200, 0, 40, 40), _cell("c", 0, 200, 40, 40)]
    assert len(stencils.plan_layers(cells)) == 1


def test_overlapping_shapes_are_split():
    cells = [_cell("a", 0, 0, 40, 40), _cell("b", 10, 10, 40, 40)]
    assert len(stencils.plan_layers(cells)) == 2


def test_group_crop_box_is_just_the_corner_badge():
    group = _cell("g", 0, 0, 400, 300, style="shape=mxgraph.aws4.group;strokeColor=#232F3E")
    x, y, w, h = stencils.crop_box(group)
    assert w <= stencils.GROUP_ICON_EXTENT + 2 * stencils.BLEED
    # a shape sitting inside the group must therefore not collide with it
    inner = _cell("i", 150, 120, 60, 60)
    assert len(stencils.plan_layers([group, inner])) == 1


def test_needs_render_skips_bitmaps_but_keeps_svg_data_uris():
    png = _cell("p", 0, 0, 10, 10, style="shape=image;image=data:image/png,AAAA")
    svg = _cell("s", 0, 0, 10, 10, style="shape=image;image=data:image/svg+xml,PHN2")
    stencil = _cell("t", 0, 0, 10, 10, style="shape=mxgraph.cisco19.rect")
    assert not stencils.needs_render(png)
    assert stencils.needs_render(svg)
    assert stencils.needs_render(stencil)


# ------------------------------------------------------------------ svg reading
TEXT_SVG = """<svg width="100px" height="50px" viewBox="0 0 100 50">
<g data-cell-id="n1"><g fill="#000000" font-family="Helvetica" text-anchor="middle"
 font-size="12px"><rect fill="#eef0f3" stroke="none" x="60" y="290" width="37" height="15"
 stroke-width="0"/><text x="78" y="300">DBMS</text></g></g>
<g data-cell-id="e1"><g transform="translate(0.5,0.5)"><path d="M 10 20 L 40 20 L 40 60"
 fill="none" stroke="#000000" stroke-miterlimit="10"/></g></g>
<g data-cell-id="fr"><g transform="translate(0.5,0.5)"><rect x="3" y="4" width="100"
 height="50" fill="none" stroke="none"/></g></g></svg>"""


def test_label_uses_the_background_rect_as_the_exact_box():
    lab = SvgMap(TEXT_SVG).label("n1")
    assert lab.lines == ["DBMS"]
    assert (lab.x, lab.y, lab.w, lab.h) == (60, 290, 37, 15)
    assert lab.size == 12 and lab.align == "center" and lab.bg == "#EEF0F3"


def test_edge_route_keeps_every_bend():
    route = SvgMap(TEXT_SVG).edge("e1")
    assert route.points == [(10, 20), (40, 20), (40, 60)]
    assert route.color == "#000000"


def test_frame_rect_reports_where_the_frame_landed():
    assert SvgMap(TEXT_SVG).frame_rect("fr") == (3.0, 4.0, 100.0, 50.0)


def test_svg_size():
    assert SvgMap(TEXT_SVG).size == (100.0, 50.0)


# ------------------------------------------------------------------ error paths
def test_find_drawio_reports_how_to_install():
    with pytest.raises(stencils.DrawioNotFound) as exc:
        stencils.find_drawio("/definitely/not/here")
    assert "/definitely/not/here" in str(exc.value)
