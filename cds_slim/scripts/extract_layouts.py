"""
Re-extract tokens/layouts.json from a Comcast template .pptx.

Run this when the brand team ships a new version of the template:

    python scripts/extract_layouts.py assets/comcast_template.pptx tokens/layouts.json

It unpacks the .pptx (it's a zip), walks every slideLayout XML, and writes
the index json that brand.py consumes.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
EMU_PER_IN = 914400


def emu_to_in(emu):
    return round(int(emu) / EMU_PER_IN, 4) if emu is not None else None


def slugify(name):
    s = name.lower().strip()
    # drop leading 'N_' artifacts (e.g. '7_Comcast Standard' -> 'comcast_standard')
    s = re.sub(r'^\d+_', '', s)
    s = s.replace('w/', 'with ').replace('&', 'and')
    s = re.sub(r'[^a-z0-9]+', '_', s)
    s = re.sub(r'_+', '_', s).strip('_')
    return s


def parse_layout(path):
    tree = ET.parse(path)
    root = tree.getroot()
    csld = root.find('p:cSld', NS)
    name = csld.attrib.get('name', '') if csld is not None else ''

    placeholders = []
    spTree = csld.find('p:spTree', NS)
    if spTree is None:
        return name, placeholders

    for sp in spTree.findall('p:sp', NS):
        ph = sp.find('p:nvSpPr/p:nvPr/p:ph', NS)
        cNvPr = sp.find('p:nvSpPr/p:cNvPr', NS)
        spPr = sp.find('p:spPr', NS)

        is_placeholder = ph is not None
        ph_type = ph.attrib.get('type', 'body') if is_placeholder else None
        ph_idx = ph.attrib.get('idx', None) if is_placeholder else None
        shape_name = cNvPr.attrib.get('name', '') if cNvPr is not None else ''

        x = y = w = h = None
        xfrm = spPr.find('a:xfrm', NS) if spPr is not None else None
        if xfrm is not None:
            off = xfrm.find('a:off', NS)
            ext = xfrm.find('a:ext', NS)
            if off is not None:
                x = emu_to_in(off.attrib.get('x'))
                y = emu_to_in(off.attrib.get('y'))
            if ext is not None:
                w = emu_to_in(ext.attrib.get('cx'))
                h = emu_to_in(ext.attrib.get('cy'))

        sample = ''
        txBody = sp.find('p:txBody', NS)
        if txBody is not None:
            parts = [t.text for t in txBody.findall('.//a:t', NS) if t.text]
            sample = ' | '.join(parts)[:140]

        if not is_placeholder and not sample and (x is None):
            continue

        placeholders.append({
            'kind': 'shape',
            'is_placeholder': is_placeholder,
            'ph_type': ph_type,
            'ph_idx': ph_idx,
            'name': shape_name,
            'x_in': x, 'y_in': y, 'w_in': w, 'h_in': h,
            'sample_text': sample,
        })

    # picture placeholders live under p:pic
    for pic in spTree.findall('p:pic', NS):
        ph = pic.find('p:nvPicPr/p:nvPr/p:ph', NS)
        cNvPr = pic.find('p:nvPicPr/p:cNvPr', NS)
        spPr = pic.find('p:spPr', NS)
        if ph is None:
            continue
        ph_type = ph.attrib.get('type', 'pic')
        ph_idx = ph.attrib.get('idx', None)
        shape_name = cNvPr.attrib.get('name', '') if cNvPr is not None else ''
        x = y = w = h = None
        xfrm = spPr.find('a:xfrm', NS) if spPr is not None else None
        if xfrm is not None:
            off = xfrm.find('a:off', NS)
            ext = xfrm.find('a:ext', NS)
            if off is not None:
                x = emu_to_in(off.attrib.get('x'))
                y = emu_to_in(off.attrib.get('y'))
            if ext is not None:
                w = emu_to_in(ext.attrib.get('cx'))
                h = emu_to_in(ext.attrib.get('cy'))
        placeholders.append({
            'kind': 'pic',
            'is_placeholder': True,
            'ph_type': ph_type,
            'ph_idx': ph_idx,
            'name': shape_name,
            'x_in': x, 'y_in': y, 'w_in': w, 'h_in': h,
            'sample_text': '',
        })

    return name, placeholders


def extract(pptx_path: Path, out_path: Path) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmpd = Path(tmp)
        with zipfile.ZipFile(pptx_path) as zf:
            zf.extractall(tmpd)

        layouts_dir = tmpd / 'ppt' / 'slideLayouts'
        master_rels_path = tmpd / 'ppt' / 'slideMasters' / '_rels' / 'slideMaster1.xml.rels'
        master_path = tmpd / 'ppt' / 'slideMasters' / 'slideMaster1.xml'

        # build rId -> layout filename map
        rels_root = ET.parse(master_rels_path).getroot()
        rel_target_by_id = {}
        for rel in rels_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            rel_target_by_id[rel.attrib['Id']] = rel.attrib['Target']

        # presentation order comes from the master's sldLayoutIdLst
        master_root = ET.parse(master_path).getroot()
        layout_id_lst = master_root.find('p:sldLayoutIdLst', NS)
        rid_attr = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        ordered_files = []
        for sld_layout_id in layout_id_lst.findall('p:sldLayoutId', NS):
            rid = sld_layout_id.attrib[rid_attr]
            target = rel_target_by_id[rid]
            ordered_files.append(os.path.basename(target))

        layouts = []
        for slot_idx, fname in enumerate(ordered_files):
            name, placeholders = parse_layout(layouts_dir / fname)
            layouts.append({
                'slot_index': slot_idx,
                'layout_file': fname,
                'name': name,
                'slug': slugify(name),
                'placeholders': placeholders,
            })

        # de-dup slugs: first occurrence wins, rest get _v2 / _v3
        seen = {}
        for lay in layouts:
            base = lay['slug']
            if base not in seen:
                seen[base] = 1
                continue
            seen[base] += 1
            lay['slug'] = f"{base}_v{seen[base]}"

        out = {
            'count': len(layouts),
            'master_name': master_root.attrib.get('name', 'Comcast'),
            'layouts': layouts,
        }

        out_path.parent.mkdir(parents=True, exist_ok=True)
        with open(out_path, 'w') as f:
            json.dump(out, f, indent=2)

        print(f'wrote {out_path} with {len(layouts)} layouts')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pptx', type=Path, help='path to the Comcast template .pptx')
    ap.add_argument('output', type=Path, nargs='?',
                    default=Path('tokens/layouts.json'),
                    help='where to write layouts.json (default: tokens/layouts.json)')
    args = ap.parse_args()

    if not args.pptx.exists():
        print(f'ERROR: template not found at {args.pptx}', file=sys.stderr)
        sys.exit(1)

    extract(args.pptx, args.output)


if __name__ == '__main__':
    main()
