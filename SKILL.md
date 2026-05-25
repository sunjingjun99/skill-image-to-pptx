---
name: image-to-pptx
description: >
  Convert a framework/architecture diagram PNG into a fully editable PPTX where
  every element (text, icon, shape, connector) is a separate, moveable object.
  Use this skill whenever asked to turn a screenshot or diagram into an editable
  PowerPoint slide.
---

# Image → Editable PPTX Skill

## Core Principle (non-negotiable)

**TEXT PRINCIPLE**: `rect()` / `circle()` = pure geometry, NO embedded text.
`txb()` = ALL text as standalone editable textboxes.
`pic()` / `icon_px()` = image objects only, NO text.

Every element must be independently selectable and moveable in PowerPoint.

---

## Phase 1 — Analyse the Source Image

```bash
# Get exact pixel dimensions
python - <<'EOF'
from PIL import Image
img = Image.open("framework.png")
print(img.size)   # (W, H)
EOF
```

1. Open the image visually and identify **zones** (colour regions, header bars, sub-boxes).
2. Note for each zone:
   - Pixel bounding box `(x1, y1, x2, y2)`
   - Background hex colour (sample with a colour picker or numpy)
   - Header bar colour
   - Accent / text colour
3. List every distinct element type:
   - Zone backgrounds, header bars
   - Sub-boxes (risk boxes, oracle boxes, bid-curve boxes, etc.)
   - Icons (one per logical concept)
   - Text labels (every line that must be independently editable)
   - Connectors / arrows

```python
# Sample a pixel colour from the original
from PIL import Image
img = Image.open("framework.png").convert("RGB")
print(img.getpixel((x, y)))   # → (R, G, B)
```

---

## Phase 2 — Script Skeleton

```python
#!/usr/bin/env python3
"""
TEXT PRINCIPLE: rect()/circle() = pure geometry, NO text.
                txb() = all text as standalone textboxes.
                icon_px() = icon images only, NO text.
"""
import io, os, numpy as np
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_CONNECTOR_TYPE
from pptx.oxml.ns import qn
from lxml import etree

IMG_PATH  = r"path/to/framework.png"
ICON_DIR  = r"path/to/icons"          # 128×128 RGBA PNGs go here
OUT_PATH  = r"path/to/output.pptx"

IMG_W, IMG_H = 1408, 768             # ← actual pixel dims of your image
SLD_W, SLD_H = 13.33, 7.5           # standard widescreen inches

XS = SLD_W / IMG_W    # pixels → inches (x)
YS = SLD_H / IMG_H    # pixels → inches (y)

def xi(px): return px * XS           # x-position
def yi(py): return py * YS           # y-position
def wi(pw): return pw * XS           # width
def hi(ph): return ph * YS           # height
def C(h):   return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))
```

---

## Phase 3 — Helper Functions

### rect — pure shape, never contains text
```python
def rect(sl, x, y, w, h, fill=None, lc=None, lw=0.5):
    shp = sl.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill: shp.fill.solid(); shp.fill.fore_color.rgb = C(fill)
    else:    shp.fill.background()
    if lc:   shp.line.color.rgb = C(lc); shp.line.width = Pt(lw)
    else:    shp.line.fill.background()
    return shp
```

### circle — pure oval, never contains text
```python
def circle(sl, cx, cy, r, fill=None, lc=None, lw=0.75):
    shp = sl.shapes.add_shape(9, Inches(cx-r), Inches(cy-r),
                               Inches(2*r), Inches(2*r))
    if fill: shp.fill.solid(); shp.fill.fore_color.rgb = C(fill)
    else:    shp.fill.background()
    if lc:   shp.line.color.rgb = C(lc); shp.line.width = Pt(lw)
    else:    shp.line.fill.background()
    return shp
```

### txb — ALL text lives here, standalone textbox
```python
def txb(sl, text, x, y, w, h, size=7, bold=False, color="111111",
        align=PP_ALIGN.LEFT, italic=False):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, ln in enumerate(str(text).split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = C(color)
    return tb
```

### conn — arrow / line connector
```python
def conn(sl, x1, y1, x2, y2, color="444444", lw=1.5, arrow=True, dbl=False):
    c = sl.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT,
                                 Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = C(color)
    c.line.width = Pt(lw)
    if arrow or dbl:
        try:
            sp = c._element
            spPr = sp.find(qn('p:spPr'))
            if spPr is None: return c
            ln_el = spPr.find(qn('a:ln'))
            if ln_el is None: ln_el = etree.SubElement(spPr, qn('a:ln'))
            for ch in list(ln_el):
                if ch.tag in (qn('a:headEnd'), qn('a:tailEnd')):
                    ln_el.remove(ch)
            if dbl:
                t = etree.SubElement(ln_el, qn('a:tailEnd'))
                t.set('type','arrow'); t.set('w','sm'); t.set('len','sm')
            else:
                etree.SubElement(ln_el, qn('a:tailEnd')).set('type','none')
            h_el = etree.SubElement(ln_el, qn('a:headEnd'))
            h_el.set('type','arrow'); h_el.set('w','med'); h_el.set('len','med')
        except Exception:
            pass
    return c
```

### tint_rgba + icon_px — clean icon placement
```python
def tint_rgba(img, hex_color):
    """Recolor all non-transparent pixels of an RGBA icon to hex_color."""
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    arr = np.array(img.convert("RGBA"), dtype=np.uint8).copy()
    mask = arr[:, :, 3] > 10
    arr[mask, 0] = r; arr[mask, 1] = g; arr[mask, 2] = b
    return Image.fromarray(arr, "RGBA")

def icon_px(sl, slot, x1_px, y1_px, x2_px, y2_px, color="111111", pad=0.82):
    """Place a named icon (tinted, square) centered in a pixel bounding box."""
    cx_px = (x1_px + x2_px) / 2
    cy_px = (y1_px + y2_px) / 2
    side_px = min(x2_px - x1_px, y2_px - y1_px) * pad
    path = os.path.join(ICON_DIR, f"{slot}.png")
    img = Image.open(path).convert("RGBA")
    img = tint_rgba(img, color)
    buf = io.BytesIO(); img.save(buf, "PNG"); buf.seek(0)
    sl.shapes.add_picture(buf,
        Inches(xi(cx_px - side_px / 2)),
        Inches(yi(cy_px - side_px / 2)),
        Inches(wi(side_px)),
        Inches(hi(side_px)))
```

---

## Phase 4 — Zone Layout Pattern

```python
# 1. Zone backgrounds — rect only, no text
rect(sl, xi(x1), yi(y1), wi(w), hi(h), fill="BG_HEX")

# 2. Header bar — rect + SEPARATE txb
def header(x1, x2, y1, hb_color, label):
    rect(sl, xi(x1), yi(y1), wi(x2-x1), hi(HDR_H), fill=hb_color)
    txb(sl, label, xi(x1)+0.06, yi(y1)+0.01, wi(x2-x1)-0.1, hi(HDR_H)-0.02,
        size=8, bold=True, color="FFFFFF", align=PP_ALIGN.LEFT)

# 3. Sub-box — rect + SEPARATE txb for title + SEPARATE icon_px
rect(sl, xi(bx1), yi(by1), wi(bw), hi(bh), fill="BOX_HEX", lc="BORDER_HEX")
txb(sl, "Box title", xi(bx1+2), yi(by1+2), wi(bw-4), hi(18),
    size=6, bold=True, color="TITLE_HEX")
icon_px(sl, "slot_name", bx1+4, by1+20, bx1+bw-4, by1+bh-4, color="ICON_HEX")

# 4. Icon with label — icon_px + SEPARATE txb below
icon_px(sl, "slot_name", ix1, iy1, ix2, iy2, color="ACCENT_HEX")
txb(sl, "Label text", xi(ix1), yi(iy2+3), wi(ix2-ix1), hi(32),
    size=6.5, color="ACCENT_HEX", align=PP_ALIGN.CENTER)
```

---

## Phase 5 — Icon Sourcing

### Option A — Icons8 CDN (preferred, 128×128 RGBA PNG)

```python
# fetch_icons.py
import io, os, requests
from PIL import Image

OUT = r"path/to/icons"
os.makedirs(OUT, exist_ok=True)
SZ  = 128
HDR = {"User-Agent": "Mozilla/5.0"}
BASE = "https://img.icons8.com/{style}/{sz}/{name}.png"

# style options: ios | material-outlined | ios-filled | fluent
ICONS = {
    "slot_name": ("ios", "icon-name-on-icons8"),
    # examples:
    "forecast":  ("ios", "combo-chart"),
    "green":     ("ios", "wind-turbine"),
    "database":  ("ios", "database"),
    "lock":      ("ios", "lock-2"),
    "tree":      ("ios", "tree-structure"),
    "scales":    ("ios", "scales"),
    "calculator":("ios", "calculator"),
    "lightning": ("ios", "lightning-bolt"),
}

for slot, (style, name) in ICONS.items():
    url = BASE.format(style=style, sz=SZ, name=name)
    try:
        r = requests.get(url, timeout=12, headers=HDR)
        img = Image.open(io.BytesIO(r.content)).convert("RGBA")
        if img.size != (SZ, SZ):
            img = img.resize((SZ, SZ), Image.LANCZOS)
        img.save(os.path.join(OUT, f"{slot}.png"))
        print(f"  OK  {slot:20s}  {name}")
    except Exception as e:
        print(f"  FAIL {slot}: {e}")
```

**Verify the icon is valid** (a bad URL still returns an HTTP 200 with a watermarked placeholder):
```python
from PIL import Image
img = Image.open("icons/slot.png").convert("RGBA")
import numpy as np
arr = np.array(img)
print("size:", img.size)                        # expect (128,128)
print("corner RGBA:", arr[0, 0])                # expect [0,0,0,0] transparent
print("non-transparent px:", (arr[:,:,3]>10).sum())  # expect > 500
```

### Option B — Custom PIL drawing (when no clean download exists)

```python
# gen_custom_icons.py
from PIL import Image, ImageDraw

SZ = 128

def save(img, name, out_dir):
    img.save(os.path.join(out_dir, f"{name}.png"))

# ── Staircase / bid-curve icon ─────────────────────────────────────────────
img = Image.new("RGBA", (SZ, SZ), (0,0,0,0))
draw = ImageDraw.Draw(img)
steps=4; sw=27; sh=24; mx=10; my=10; lw=7
draw.line([(mx, SZ-my), (SZ-mx, SZ-my)], fill=(0,0,0,255), width=4)  # x-axis
draw.line([(mx, my), (mx, SZ-my)], fill=(0,0,0,255), width=4)         # y-axis
pts = [(mx, SZ-my)]
for i in range(steps):
    x_now  = mx + i*sw + 3;     y_now  = SZ-my-i*sh
    x_next = mx + (i+1)*sw + 3; y_next = SZ-my-(i+1)*sh
    pts += [(x_now, y_now), (x_next, y_now), (x_next, y_next)]
draw.line(pts, fill=(0,0,0,255), width=lw)
save(img, "staircase", OUT)

# ── Queue box with outgoing arrow ──────────────────────────────────────────
img = Image.new("RGBA", (SZ, SZ), (0,0,0,0))
draw = ImageDraw.Draw(img)
bx, by, bw, bh = 14, 30, 80, 68
draw.rectangle([bx, by, bx+bw, by+bh], outline=(0,0,0,255), width=5)
for i in range(4):
    y = by + 13 + i*14
    draw.line([(bx+8, y), (bx+bw-8, y)], fill=(0,0,0,255), width=5)
draw.line([(bx+bw, by+bh//2), (bx+bw+26, by+bh//2)], fill=(0,0,0,255), width=5)
draw.polygon([(bx+bw+22, by+bh//2-8), (bx+bw+36, by+bh//2),
              (bx+bw+22, by+bh//2+8)], fill=(0,0,0,255))
save(img, "queue", OUT)

# ── Bidirectional-arrows flow (MCN matching) ───────────────────────────────
img = Image.new("RGBA", (SZ, SZ), (0,0,0,0))
draw = ImageDraw.Draw(img)
for i, y in enumerate([30, 63, 96]):
    x1, x2 = 14, 114
    draw.line([(x1, y), (x2, y)], fill=(0,0,0,255), width=5)
    if i % 2 == 0:
        draw.polygon([(x2-12, y-8),(x2+2, y),(x2-12, y+8)], fill=(0,0,0,255))
    else:
        draw.polygon([(x1+12, y-8),(x1-2, y),(x1+12, y+8)], fill=(0,0,0,255))
save(img, "mcn_arrows", OUT)

# ── Queue-box + cylinder combo (rolling update) ────────────────────────────
img = Image.new("RGBA", (SZ, SZ), (0,0,0,0))
draw = ImageDraw.Draw(img)
qx, qy, qw, qh = 8, 35, 48, 58
draw.rectangle([qx, qy, qx+qw, qy+qh], outline=(0,0,0,255), width=5)
for i in range(3):
    y = qy + 13 + i*13
    draw.line([(qx+7, y), (qx+qw-7, y)], fill=(0,0,0,255), width=4)
ax = qx+qw; ay = qy+qh//2
draw.line([(ax, ay), (ax+16, ay)], fill=(0,0,0,255), width=5)
draw.polygon([(ax+12, ay-7),(ax+22, ay),(ax+12, ay+7)], fill=(0,0,0,255))
cx, cy, cw, ch = 78, 30, 44, 68
draw.ellipse([cx, cy, cx+cw, cy+16], outline=(0,0,0,255), width=5)
draw.line([(cx, cy+8), (cx, cy+ch)], fill=(0,0,0,255), width=5)
draw.line([(cx+cw, cy+8), (cx+cw, cy+ch)], fill=(0,0,0,255), width=5)
draw.ellipse([cx, cy+ch-8, cx+cw, cy+ch+8], outline=(0,0,0,255), width=5)
save(img, "queue_cyl", OUT)

# ── Lightning bolt ─────────────────────────────────────────────────────────
img = Image.new("RGBA", (SZ, SZ), (0,0,0,0))
draw = ImageDraw.Draw(img)
draw.polygon([(72,10),(38,66),(60,66),(40,118),(90,56),(65,56)],
             fill=(0,0,0,255))
save(img, "lightning", OUT)
```

---

## Phase 6 — Pixel-Scanning Diagnostics

When you're unsure where an element actually sits in the source image:

```python
# Find non-white rows in a horizontal band
import numpy as np
from PIL import Image
img = np.array(Image.open("framework.png").convert("RGB"))

band = img[y1:y2, x1:x2]           # crop to suspected region
row_diff = band.max(axis=(1,2)) - band.min(axis=(1,2))
content_rows = np.where(row_diff > 30)[0] + y1
print("content rows:", content_rows[[0,-1]])   # first / last row with content

# Spot-check a pixel
print(img[row, col])                # (R, G, B)
```

Use this to locate icon strips, icon bounds, and confirm that a text label
does NOT extend into what you thought was icon-only territory.

---

## Phase 7 — QA Export & Inspection

### Export via PowerPoint COM (Windows)

```python
# qa_export.py
import win32com.client, os, time

pptx_path = r"path\to\output.pptx"
out_dir   = r"path\to\qa_images"
os.makedirs(out_dir, exist_ok=True)

pptx_abs = os.path.abspath(pptx_path)
out_abs  = os.path.join(os.path.abspath(out_dir), "slide1.png")

ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = False
prs = ppt.Presentations.Open(pptx_abs, ReadOnly=True, WithWindow=False)
slide = prs.Slides(1)
slide.Export(out_abs, "PNG", 1920, 1080)
prs.Close()
ppt.Quit()
print("Exported:", out_abs)
```

### Export via LibreOffice (cross-platform)

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

### Zone crops for targeted QA

```python
# qa_crop.py
from PIL import Image

img = Image.open("qa_images/slide1.png")
W, H = img.size   # e.g. 1920×1080
sx = W / 1408; sy = H / 768   # scale factors

def crop_zone(name, px1, py1, px2, py2, pad=10):
    box = (int(px1*sx)-pad, int(py1*sy)-pad,
           int(px2*sx)+pad, int(py2*sy)+pad)
    img.crop(box).save(f"qa_images/{name}.png")

crop_zone("z1", 10,  9,  585, 255)
crop_zone("z2", 10, 258, 585, 757)
crop_zone("z3", 612, 9,  782, 757)
crop_zone("z4", 809, 9, 1140, 757)
crop_zone("z5",1166, 9, 1400, 757)
```

### Visual check prompts (for subagent or manual review)

```
Look for:
- Text overflowing its textbox bounds
- Icons rendered too small (< 20px effective side) or clipped
- Overlapping elements (text through a shape)
- Missing icons (blank area where icon should be)
- Wrong tint colour (icon in wrong zone accent colour)
- Connector arrows pointing in the wrong direction
```

---

## Phase 8 — Common Pitfalls & Fixes

| Problem | Symptom | Fix |
|---------|---------|-----|
| Cropping title text instead of icon | Crop shows text, not icon | Pixel-scan to find true icon row range |
| Icons8 URL returns watermarked placeholder | `non-transparent px` count is huge / icon has text | Try different `style` or `name`; use PIL fallback |
| cairosvg fails on Windows | `no library called "cairo-2"` | Use Icons8 PNG CDN instead of SVG sources |
| Icon too small in narrow bounding box | Icon appears as a tiny dot | Expand the bounding box or increase `pad` in `icon_px()` |
| Text wraps unexpectedly | Label overflows or wraps to 3+ lines | Increase text box width `wi()` or reduce font size |
| Connector arrow direction reversed | Arrow head at wrong end | `conn()` uses HEAD-end arrow; swap `(x1,y1)` and `(x2,y2)` |
| Shape type 5 is rounded rect, not triangle | Code uses `add_shape(5,...)` | Use `conn()` for arrows; use `polygon` via PIL or custom XML for triangles |
| Settlement icon shows border artifact | Thin `\|` line at icon edge | Inset `x1_px`/`x2_px` by ~6–8px to avoid the sub-box border |

---

## Phase 9 — Complete Workflow Checklist

```
[ ] 1. Measure image dimensions (PIL)
[ ] 2. Identify zones — pixel bounds, background colours, header colours, accent colours
[ ] 3. List every icon slot needed (one slot per concept)
[ ] 4. Download icons via Icons8 CDN; validate each (size + transparency check)
[ ] 5. Generate custom PIL icons for anything not cleanly available online
[ ] 6. Write build_pptx.py using the helpers above
[ ] 7. For each zone: rect bg → header (rect+txb) → sub-boxes (rect+txb+icon_px) → connectors
[ ] 8. Run build_pptx.py
[ ] 9. Export to PNG (qa_export.py)
[ ] 10. Crop zone views (qa_crop.py) and inspect each zone
[ ] 11. Fix any overlaps, overflow, or wrong-position issues
[ ] 12. Re-export and verify
```

---

## Appendix — Full Icon Slot Convention

Name icons as `{zone}_{concept}`, e.g.:

| Slot | Concept |
|------|---------|
| `z1_forecast` | workload forecast chart |
| `z1_green` | renewable / wind energy |
| `z1_uncertain` | probability / uncertainty |
| `z1_price` | electricity / price signal |
| `z2_svc` | service-delay / time management |
| `z2_gmr` | green mismatch / balance scale |
| `z2_nab` | optimization / calculator |
| `z2_lock` | privacy lock |
| `z3_baseline` | baseline plan / clipboard |
| `z3_cost` | cost chart / bar chart |
| `z3_oracle` | oracle labels / data config |
| `z4_green` | real-time green power / lightning |
| `z4_queue` | rolling queue / conveyor |
| `z4_storage` | storage state / database |
| `z4_lock` | private oracle lock |
| `z4_tree` | XGBoost / decision tree |
| `z4_sigmoid` | surrogate model / graph |
| `z4_staircase` | bid curve / staircase |
| `z5_mcn` | market-clearing flow / arrows |
| `z5_vcg` | VCG payment / coins |
| `z5_risk` | risk dividend / receive-cash |
| `z5_profit` | profit balancing / scales |
| `z5_queue_cyl` | rolling update queue+cylinder |
