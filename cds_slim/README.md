# Comcast Design System

Code-as-source-of-truth for the **Comcast Corporate 2026 brand template**.
Drop this folder into your repo and reference it whenever you need to
generate or style a deck programmatically.

## What you get

```
comcast-design-system/
├── assets/
│   └── comcast_template.pptx     # the canonical .pptx with embedded fonts + 94 layouts
├── tokens/
│   ├── comcast_brand.json        # palette, fonts, dimensions, grid coords
│   └── layouts.json              # all 94 layouts indexed with placeholder geometry
├── src/
│   ├── brand.py                  # typed Python access to the JSON tokens
│   ├── deck.py                   # template-driven deck builder (recommended)
│   └── fallback.py               # build from scratch when the .pptx isn't shippable
└── examples/
    ├── example_01_template_driven.py
    └── example_02_fallback.py
```

## Install

```bash
pip install python-pptx
```

That's it. No build step, no external services.

## Two ways to use it

### 1. Template-driven (recommended)

Opens `assets/comcast_template.pptx` and adds slides by layout name.
Output keeps the embedded **Comcast New Vision** font family and every
layout the brand team designed.

```python
from src.deck import Deck

deck = Deck()
deck.add('presentation_title_slide',
         title='Q3 Operating Review',
         subhead='Revenue, product, and the path into Q4',
         date='OCTOBER 2026')
deck.add('comcast_standard',
         banner='OPERATING REVIEW',
         title='Where we landed in Q3',
         body='Revenue grew 14% YoY to $2.4B...')
deck.add('data_3',
         title='Three numbers from Q3',
         stat_values=['14%', '$2.4B', '380K'],
         stat_captions=['Growth YoY', 'Total', 'Net adds'])
deck.add('closing')
deck.save('q3.pptx')
```

### 2. Fallback (no .pptx required)

Builds slides shape-by-shape using brand tokens. Output is on-palette but
falls back to **Aptos** (no Comcast New Vision) and has no peacock graphics.
Use this when shipping the .pptx isn't an option (locked-down CI, etc).

```python
from src.fallback import BlankDeck

d = BlankDeck()
d.add_title_slide(title='Q3 Review', subhead='...', date='OCTOBER 2026')
d.add_data_slide(title='Three numbers',
                 stats=[('14%','Growth'), ('$2.4B','Revenue'), ('380K','Net adds')])
d.save('q3.pptx')
```

Fallback covers 7 patterns: title, section divider, standard, data, key
points, closing, plus the chrome (banner / page number / confidentiality).
Anything beyond that — bento, mockups, gantts — use the template-driven
path or extend `fallback.py`.

## Reference: brand tokens

```python
from src import brand

brand.COLORS.yellow            # '#FBCC11'
brand.COLORS.orange            # '#FF7011'
brand.COLORS.red               # '#EF1541'
brand.COLORS.purple            # '#6D54DB'
brand.COLORS.blue              # '#059CDF'
brand.COLORS.green             # '#05AB3E'

brand.FONTS.heading            # 'Comcast New Vision Medium'
brand.FONTS.body               # 'Comcast New Vision'

brand.SLIDE_WIDTH_IN           # 13.333
brand.SLIDE_HEIGHT_IN          # 7.5

brand.layout('comcast_standard')          # LayoutSpec object
brand.find_layouts(family='bento')        # all bento layouts
brand.list_families()                     # family -> [layout names]
brand.print_inventory()                   # full catalog to stdout
```

## How to find the right layout

Layouts have human-readable slugs derived from the original PowerPoint names.
Useful mappings:

| Slug                          | Slot | Original name                  |
|-------------------------------|------|--------------------------------|
| `presentation_title_slide`    | 0    | Presentation Title Slide       |
| `section_title`               | 3    | Section Title                  |
| `comcast_standard`            | 16   | 7_Comcast Standard             |
| `key_points_slide`            | 20   | Key Points Slide               |
| `data_3` / `data_4` / `data_6`| 52–55| Big-stat callout slides        |
| `bento_layout`                | 72   | Bento Layout                   |
| `chart_layout`                | 74   | Chart Layout                   |
| `single_image`                | 82   | Single image                   |
| `closing`                     | 90   | CLOSING                        |
| `blank`                       | 28   | 1_BLANK                        |

For the full list, run `brand.print_inventory()` or grep `tokens/layouts.json`.

## Discovering placeholder schemas

Each layout has different placeholders. To find what kwargs to pass:

```python
from src.deck import Deck
deck = Deck()
print(deck.describe_layout('data_3'))
```

Prints every placeholder, its idx, position, role, and default sample text.

## Recognised content kwargs

`Deck.add(layout_id, **content)` accepts:

| Kwarg            | What it fills                                              |
|------------------|------------------------------------------------------------|
| `title`          | The title placeholder                                      |
| `banner`         | Top-left BANNER TEXT placeholder                           |
| `subhead`        | Subhead under the title                                    |
| `body`           | Main body paragraph                                        |
| `subtitle`/`date`| Date or subtitle in title-style layouts                    |
| `subheaders`     | List, for layouts with multiple headers (key points, etc.) |
| `bodies`         | List, for layouts with multiple body blocks                |
| `stat_values`    | List of big numbers for `data_*` layouts                   |
| `stat_captions`  | List of captions under those numbers                       |
| `placeholders`   | `{idx: 'text'}` raw override for anything not above        |

Items in list-valued kwargs are mapped top-to-bottom, left-to-right based on
the layout's placeholder positions.

The returned `Slide` object is the standard python-pptx Slide, so for
anything custom (charts, tables, images, custom positioning), just keep
working on it directly:

```python
slide = deck.add('chart_layout', title='Revenue by region')
# now use slide.shapes / slide.placeholders / etc. directly
```

## Known issue: LibreOffice rendering of date placeholders

If you preview output with LibreOffice (or any soffice-based PDF
converter), the **date placeholder on the title slide** may show
ghosted/overlaid text. The saved `.pptx` is correct — PowerPoint and the
Microsoft web viewer render it cleanly. This is a LibreOffice rendering
quirk where it composites both the layout's prompt text and the slide's
override text.

## Updating the design system

1. **Brand colors / fonts / dimensions changed** → edit `tokens/comcast_brand.json`
2. **New layout added in PowerPoint** → drop the new `.pptx` into `assets/`,
   re-run the extractor (the script that produced `tokens/layouts.json` is
   in version history; rerun it against the new template).
3. **New family classification** → edit `_FAMILY_RULES` in `src/brand.py`.

The JSON files are the single source of truth. `src/brand.py` is a thin
wrapper, so any token change propagates to both `deck.py` and `fallback.py`
automatically.

## Caveats and design choices

- **Why open the template instead of building from XML?**
  Comcast New Vision is embedded inside the .pptx. Recreating layouts in
  code would lose the font unless every consuming machine has it installed.
  Opening the template ships the fonts inside the output deck.

- **Why not codify all 94 layouts in `fallback.py`?**
  Recreating that many layouts shape-by-shape doubles the maintenance
  surface for marginal value. The 7 fallback layouts cover ~80% of routine
  decks. For everything else, ship the .pptx.

- **What about charts and tables?**
  Use `Deck.add('chart_layout', title='...')` and then add a chart via
  python-pptx's standard `slide.shapes.add_chart(...)` API on the returned
  slide.
