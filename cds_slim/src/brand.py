"""
brand.py — typed access to the Comcast design tokens.

Loads from tokens/comcast_brand.json and tokens/layouts.json so the JSON
files remain the single source of truth. If you edit a token, edit the JSON.

Usage:
    from comcast_design_system.src import brand
    brand.COLORS.blue                # '#059CDF'
    brand.FONTS.heading              # 'Comcast New Vision Medium'
    brand.SLIDE_WIDTH_IN             # 13.333
    brand.layout_slot('bento_layout')  # -> int slot index
    brand.find_layouts(family='bento')  # -> list of LayoutSpec
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

_TOKENS_DIR = Path(__file__).resolve().parent.parent / "tokens"

with open(_TOKENS_DIR / "comcast_brand.json") as f:
    _BRAND: dict[str, Any] = json.load(f)
with open(_TOKENS_DIR / "layouts.json") as f:
    _LAYOUTS_RAW: dict[str, Any] = json.load(f)


# ---------------------------------------------------------------------------
# Colors
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class _Colors:
    # 2024 Comcast accent palette
    yellow: str = _BRAND["colors"]["accents"]["yellow"]["hex"]
    orange: str = _BRAND["colors"]["accents"]["orange"]["hex"]
    red: str = _BRAND["colors"]["accents"]["red"]["hex"]
    purple: str = _BRAND["colors"]["accents"]["purple"]["hex"]
    blue: str = _BRAND["colors"]["accents"]["blue"]["hex"]
    green: str = _BRAND["colors"]["accents"]["green"]["hex"]
    black: str = _BRAND["colors"]["neutral"]["black"]
    white: str = _BRAND["colors"]["neutral"]["white"]
    hyperlink: str = _BRAND["colors"]["semantic"]["hyperlink"]

    def as_rgb(self, hex_str: str) -> tuple[int, int, int]:
        h = hex_str.lstrip("#")
        return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


COLORS = _Colors()

# ordered list for cycling — yellow first because it's the dominant brand color
ACCENT_CYCLE = [
    COLORS.yellow, COLORS.orange, COLORS.red,
    COLORS.purple, COLORS.blue, COLORS.green,
]


# ---------------------------------------------------------------------------
# Fonts
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class _Fonts:
    heading: str = _BRAND["typography"]["roles"]["heading"]["family"]
    heading_fallback: str = _BRAND["typography"]["roles"]["heading"]["fallback"]
    body: str = _BRAND["typography"]["roles"]["body"]["family"]
    body_fallback: str = _BRAND["typography"]["roles"]["body"]["fallback"]
    subhead: str = _BRAND["typography"]["roles"]["subhead"]["family"]
    banner: str = _BRAND["typography"]["roles"]["banner"]["family"]
    data_stat: str = _BRAND["typography"]["roles"]["data_stat"]["family"]


FONTS = _Fonts()

# pt sizes from the master text style
SIZE_TITLE_PT = _BRAND["typography"]["sizes_pt"]["title"]
SIZE_BODY_PT = _BRAND["typography"]["sizes_pt"]["body"]
SIZE_SECTION_HEADER_PT = _BRAND["typography"]["sizes_pt"]["section_header"]
SIZE_STAT_CALLOUT_PT = _BRAND["typography"]["sizes_pt"]["stat_callout"]
SIZE_STAT_CAPTION_PT = _BRAND["typography"]["sizes_pt"]["stat_caption"]
SIZE_BANNER_PT = _BRAND["typography"]["sizes_pt"]["banner_text"]


# ---------------------------------------------------------------------------
# Slide geometry
# ---------------------------------------------------------------------------

SLIDE_WIDTH_IN: float = _BRAND["slide"]["width_in"]
SLIDE_HEIGHT_IN: float = _BRAND["slide"]["height_in"]
SLIDE_WIDTH_EMU: int = _BRAND["slide"]["width_emu"]
SLIDE_HEIGHT_EMU: int = _BRAND["slide"]["height_emu"]

MARGIN_LEFT_IN: float = _BRAND["grid"]["margins_in"]["left"]
MARGIN_RIGHT_IN: float = _BRAND["grid"]["margins_in"]["right"]
MARGIN_TOP_IN: float = _BRAND["grid"]["margins_in"]["top"]
MARGIN_BOTTOM_IN: float = _BRAND["grid"]["margins_in"]["bottom"]


@dataclass(frozen=True)
class Region:
    x: float
    y: float
    w: float
    h: float


def _region(name: str) -> Region:
    r = _BRAND["grid"]["common_regions"][name]
    return Region(r["x"], r["y"], r["w"], r["h"])


REGION_BANNER = _region("banner")
REGION_TITLE = _region("title")
REGION_SUBHEAD = _region("subhead")
REGION_PAGE_NUMBER = _region("page_number")
REGION_CONFIDENTIALITY = _region("confidentiality")
REGION_CONTENT_FULL = _region("content_area_full")


# ---------------------------------------------------------------------------
# Layouts
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Placeholder:
    is_placeholder: bool
    ph_type: str | None
    ph_idx: str | None
    name: str
    x_in: float | None
    y_in: float | None
    w_in: float | None
    h_in: float | None
    sample_text: str
    kind: str = "shape"  # 'shape' or 'pic'


@dataclass(frozen=True)
class LayoutSpec:
    slot_index: int          # use as prs.slide_layouts[slot_index]
    layout_file: str
    name: str                # original PowerPoint name, e.g. '7_Comcast Standard'
    slug: str                # snake_case, unique
    placeholders: tuple[Placeholder, ...] = field(default_factory=tuple)
    family: str = "uncategorized"


# family classification — pattern-matched against names. Matches in order.
_FAMILY_RULES = [
    ("title_section",     ["presentation title", "title slide", "section title", "sub-section"]),
    ("agenda",            ["agenda"]),
    ("categories",        ["categories"]),
    ("standard_content",  ["comcast standard"]),
    ("impact_headlines",  ["impact headline", "impact head"]),
    ("key_points",        ["key points", "main idea with dividers"]),
    ("strategic_pillars", ["strategic pillars", "from and to"]),
    ("timeline",          ["timeline", "gantt"]),
    ("data_callouts",     ["data_"]),
    ("statements",        ["statement", "image and impact"]),
    ("multi_column_text", ["column bullet", "column text", "headline, text, bullet"]),
    ("bento",             ["bento"]),
    ("headshots",         ["headshots"]),
    ("tables_charts",     ["table layout", "chart"]),
    ("device_mockups",    ["mockup", "laptop"]),
    ("image_layouts",     ["image", "images & content", "images and content"]),
    ("closing_blank",     ["closing", "blank"]),
]


def _classify(name: str) -> str:
    n = name.lower()
    for family, patterns in _FAMILY_RULES:
        for p in patterns:
            if p in n:
                return family
    return "uncategorized"


def _build_layouts() -> tuple[LayoutSpec, ...]:
    out: list[LayoutSpec] = []
    for raw in _LAYOUTS_RAW["layouts"]:
        phs = tuple(
            Placeholder(
                is_placeholder=p["is_placeholder"],
                ph_type=p["ph_type"],
                ph_idx=p["ph_idx"],
                name=p["name"],
                x_in=p["x_in"], y_in=p["y_in"], w_in=p["w_in"], h_in=p["h_in"],
                sample_text=p["sample_text"],
                kind=p.get("kind", "shape"),
            )
            for p in raw["placeholders"]
        )
        out.append(LayoutSpec(
            slot_index=raw["slot_index"],
            layout_file=raw["layout_file"],
            name=raw["name"],
            slug=raw["slug"],
            placeholders=phs,
            family=_classify(raw["name"]),
        ))
    return tuple(out)


LAYOUTS: tuple[LayoutSpec, ...] = _build_layouts()
LAYOUT_BY_SLUG: dict[str, LayoutSpec] = {l.slug: l for l in LAYOUTS}
LAYOUT_BY_NAME: dict[str, LayoutSpec] = {l.name: l for l in LAYOUTS}


# ---------------------------------------------------------------------------
# Lookup helpers
# ---------------------------------------------------------------------------

def layout(identifier: str | int) -> LayoutSpec:
    """Fetch a layout by slug, exact name, or slot index. Raises KeyError if missing."""
    if isinstance(identifier, int):
        if 0 <= identifier < len(LAYOUTS):
            return LAYOUTS[identifier]
        raise KeyError(f"slot {identifier} out of range (0..{len(LAYOUTS) - 1})")
    if identifier in LAYOUT_BY_SLUG:
        return LAYOUT_BY_SLUG[identifier]
    if identifier in LAYOUT_BY_NAME:
        return LAYOUT_BY_NAME[identifier]
    # last-ditch: case-insensitive name match
    for spec in LAYOUTS:
        if spec.name.lower() == identifier.lower():
            return spec
    raise KeyError(f"no layout named {identifier!r}")


def layout_slot(identifier: str | int) -> int:
    return layout(identifier).slot_index


def find_layouts(family: str | None = None, contains: str | None = None) -> list[LayoutSpec]:
    """Filter layouts by family and/or substring match on name."""
    out = list(LAYOUTS)
    if family:
        out = [l for l in out if l.family == family]
    if contains:
        c = contains.lower()
        out = [l for l in out if c in l.name.lower()]
    return out


def list_families() -> dict[str, list[str]]:
    """Family -> list of layout names. Useful for printing to console."""
    families: dict[str, list[str]] = {}
    for l in LAYOUTS:
        families.setdefault(l.family, []).append(l.name)
    return families


# ---------------------------------------------------------------------------
# Convenience for callers
# ---------------------------------------------------------------------------

def print_inventory() -> None:
    """Print the full layout catalog grouped by family."""
    fams = list_families()
    for fam in sorted(fams):
        print(f"\n[{fam}]  ({len(fams[fam])})")
        for name in fams[fam]:
            spec = LAYOUT_BY_NAME[name]
            print(f"  slot={spec.slot_index:3d}  slug={spec.slug:40s}  {spec.name}")
