"""Regenerate examples/preview.png from a --verify run. Not part of the package."""
import sys
from PIL import Image, ImageDraw, ImageFont

check, out, caption = sys.argv[1], sys.argv[2], sys.argv[3]
im = Image.open(check).convert("RGB")
gap = 10
h = (im.height - 2 * gap) // 3
panels = [im.crop((0, 0, im.width, h)), im.crop((0, h + gap, im.width, 2 * h + gap))]


def trim(p, pad=14):
    bb = p.convert("L").point(lambda v: 0 if v > 248 else 255).getbbox()
    return p.crop((max(0, bb[0] - pad), max(0, bb[1] - pad),
                   min(p.width, bb[2] + pad), min(p.height, bb[3] + pad)))


H = 620
panels = [trim(p) for p in panels]
panels = [p.resize((round(p.width * H / p.height), H), Image.LANCZOS) for p in panels]
left, right = panels
gapx, lab = 56, 52
canvas = Image.new("RGB", (left.width + gapx + right.width, H + lab), "white")
canvas.paste(left, (0, lab))
canvas.paste(right, (left.width + gapx, lab))
d = ImageDraw.Draw(canvas)
try:
    f = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 24)
except Exception:
    f = ImageFont.load_default()
d.text((2, 12), "draw.io", font=f, fill=(90, 108, 134))
d.text((left.width + gapx + 2, 12), caption, font=f, fill=(90, 108, 134))
d.line([(left.width + gapx // 2, lab), (left.width + gapx // 2, H + lab)],
       fill=(224, 226, 230), width=2)
canvas.thumbnail((1500, 1500), Image.LANCZOS)
canvas.save(out)
print(out, canvas.size)
