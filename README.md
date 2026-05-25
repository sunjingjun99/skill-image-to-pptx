# image-to-pptx — Claude Code Skill

A reusable skill for **Claude Code** that converts any framework / architecture diagram PNG into a fully editable PowerPoint file, where every element (text, icon, shape, connector) is a separate, independently moveable object.

## What it does

Given a screenshot or diagram image, this skill guides Claude to:

1. Analyse pixel bounds, zone colours, and element layout
2. Generate a `build_pptx.py` script using `python-pptx`
3. Download clean icons from Icons8 CDN (or draw them with PIL)
4. Produce a `.pptx` where:
   - All **text** lives in standalone editable textboxes (`txb()`)
   - All **shapes** are pure geometry with no embedded text (`rect()` / `circle()`)
   - All **icons** are tinted RGBA images placed as separate picture objects (`icon_px()`)
   - **Connectors** and arrows are independent line objects (`conn()`)
5. QA-export the result and crop zone views for visual inspection

## Usage in Claude Code

Install the skill by placing `image-to-pptx.md` in your Claude Code skills directory, then invoke it with:

```
/image-to-pptx
```

Or reference it directly by asking Claude:

> "Use the image-to-pptx skill to convert this diagram to an editable PPTX."

## Key design principle

```
rect() / circle()  →  pure geometry, NEVER contains text
txb()              →  ALL text as standalone textboxes
icon_px()          →  tinted icon images, NEVER contains text
conn()             →  arrows and connectors
```

## Dependencies

```
pip install python-pptx Pillow numpy requests
```

For QA export on Windows (optional):
```
pip install pywin32   # PowerPoint COM automation
```

## File

| File | Description |
|------|-------------|
| `SKILL.md` | The full skill — phases 1–9 with all helper code (standard entry point) |

## License

MIT
