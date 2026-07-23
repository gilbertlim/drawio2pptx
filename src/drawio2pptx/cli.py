"""Command line entry point."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__, build, model, stencils

USAGE = """examples:
  drawio2pptx diagram.drawio                      -> diagram.pptx, one slide, 16:9
  drawio2pptx diagram.drawio -o deck.pptx
  drawio2pptx diagram.drawio --into deck.pptx --slide 2 --replace
  drawio2pptx diagram.drawio --all-pages -o deck.pptx
  drawio2pptx diagram.drawio --verify check.png   -> stacked before/after/diff image
"""


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="drawio2pptx", epilog=USAGE,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Convert a draw.io diagram into individually editable PowerPoint objects.")
    p.add_argument("source", nargs="?", help="path to a .drawio / .xml diagram")
    p.add_argument("-o", "--output", help="output .pptx (default: alongside the source)")
    p.add_argument("--into", metavar="DECK.pptx",
                   help="insert into an existing deck instead of creating one")
    p.add_argument("--slide", type=int, metavar="N",
                   help="target slide number in --into (default: append a new slide)")
    p.add_argument("--replace", action="store_true",
                   help="clear the target slide before inserting")
    p.add_argument("--page", type=int, default=1, metavar="N",
                   help="which diagram page to convert (default: 1)")
    p.add_argument("--all-pages", action="store_true",
                   help="convert every page, one slide each")
    p.add_argument("--list-pages", action="store_true", help="print the page names and exit")
    p.add_argument("--slide-size", default="16:9", metavar="SIZE",
                   help="16:9 (default), 4:3, 16:10, auto, or WxH in inches e.g. 13.333x7.5")
    p.add_argument("--margin", type=float, default=0.0, metavar="F",
                   help="blank border as a fraction of the slide, e.g. 0.04 (default: 0)")
    p.add_argument("--scale", type=int, default=6, metavar="N",
                   help="raster scale for stencil icons (default: 6)")
    p.add_argument("--font", metavar="NAME",
                   help="override the latin font for all labels (default: as drawn)")
    p.add_argument("--ea-font", metavar="NAME",
                   help="East Asian font for CJK labels, e.g. 'Malgun Gothic'")
    p.add_argument("--drawio", metavar="PATH", help="path to the draw.io desktop binary")
    p.add_argument("--verify", metavar="OUT.png",
                   help="also write a reference/result/diff image for a visual check")
    p.add_argument("--keep-workdir", action="store_true",
                   help="keep intermediate renders for debugging")
    p.add_argument("-q", "--quiet", action="store_true")
    p.add_argument("-V", "--version", action="version", version=f"drawio2pptx {__version__}")
    return p


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if not args.source:
        _parser().print_help()
        return 2

    source = Path(args.source).expanduser()
    say = (lambda *_: None) if args.quiet else (lambda m: print(f"  {m}", file=sys.stderr))

    if not args.into and (args.slide is not None or args.replace):
        print("drawio2pptx: --slide and --replace apply to --into DECK.pptx. "
              "A newly created deck has no slide to target.", file=sys.stderr)
        return 2

    if not source.exists():
        print(f"drawio2pptx: no such diagram: {source}", file=sys.stderr)
        sample = Path(__file__).resolve().parents[2] / "examples" / "sample.drawio"
        if source.name == "diagram.drawio":
            print("  'diagram.drawio' is the placeholder used in the docs — "
                  "replace it with your own file.", file=sys.stderr)
        if sample.exists():
            print(f"  to try the bundled example:  drawio2pptx {sample}", file=sys.stderr)
        return 1

    try:
        pages = model.parse(source)
        if args.list_pages:
            for pg in pages:
                print(f"{pg.index}\t{pg.name}\t{len(pg.cells)} cells")
            return 0

        targets = [p.index for p in pages] if args.all_pages else [args.page]
        deck = args.into
        out = args.output
        result = None

        for n, idx in enumerate(targets):
            if not args.quiet and len(targets) > 1:
                print(f"page {idx}/{len(targets)}: {pages[idx - 1].name}", file=sys.stderr)
            result = build.convert(
                source, out, into=deck,
                slide=args.slide if n == 0 else None,
                replace=args.replace if n == 0 else False,
                page=idx, slide_size=args.slide_size, margin=args.margin, scale=args.scale,
                drawio=args.drawio, font=args.font, ea_font=args.ea_font,
                keep_workdir=args.keep_workdir, progress=say,
            )
            # subsequent pages append to the deck we just wrote
            deck, out = str(result.path), str(result.path)

        c = result.counts
        total = sum(c.values())
        print(f"{result.path}  (slide {result.slide_index}, {total} objects: "
              f"{c['rect']} shapes, {c['line']} connectors, {c['text']} text boxes, "
              f"{c['picture']} icons)")
        if result.workdir:
            print(f"workdir kept: {result.workdir}")

        if args.verify:
            from . import verify
            binary = stencils.find_drawio(args.drawio)
            ref = Path(args.verify).with_suffix(".reference.png")
            stencils.run(binary, ["-f", "png", "-s", "2", "--crop", "-b", "0",
                                  "-p", str(targets[-1]), "-o", str(ref),
                                  str(source.expanduser().resolve())])
            out_png = verify.compare(result.path, result.slide_index, ref, args.verify)
            print(f"{out_png}  (top: draw.io, middle: this slide, bottom: difference)")
        return 0

    except stencils.DrawioNotFound as e:
        print(f"drawio2pptx: {e}", file=sys.stderr)
        return 3
    except (FileNotFoundError, ValueError, RuntimeError) as e:
        print(f"drawio2pptx: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
