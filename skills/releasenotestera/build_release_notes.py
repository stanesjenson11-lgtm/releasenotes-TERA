"""Build a Regal release-notes .docx from a JSON item list.

    python build_release_notes.py items.json [-o "Release Notes 07-15-2026.docx"]
    python build_release_notes.py --selfcheck

Formatting comes entirely from template.docx (derived from Release Notes
07-01-2026.docx), so the three sections keep their own colours, fonts, cell
shading and column widths. This script only pours text in.
"""
import argparse
import copy
import json
import re
import tempfile
from pathlib import Path

import docx
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from PIL import Image, ImageDraw, ImageFont

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
HERE = Path(__file__).parent
TEMPLATE = HERE / 'template.docx'

RELEASE_TITLE = 'Regal Wealth Advisors'
SECTIONS = [('new_features', 'New Features'),
            ('improvements', 'Improvements'),
            ('bug_fixes', 'Bug Fixes')]

# placeholder box, sized to the median real screenshot in the reference doc
PH_W_IN, PH_H_IN = 6.9, 3.2
PH_SCALE = 2                      # render at 2x, Word scales down -> crisp
PH_BG, PH_BORDER, PH_FG = '#F2F2F2', '#BFBFBF', '#808080'


# ---------------------------------------------------------------- text helpers

def set_text(p, text):
    """Replace a paragraph's text, keeping run 0's formatting."""
    runs = p._p.findall(f'{W}r')
    if not runs:
        raise ValueError(f'paragraph has no run to write into: {p.text!r}')
    for r in runs[1:]:
        p._p.remove(r)
    p.runs[0].text = text


def set_benefit(p, text):
    """Benefit paragraph keeps two runs: bold 'Benefit' then a newline + body."""
    runs = p._p.findall(f'{W}r')
    for r in runs[2:]:
        p._p.remove(r)
    p.runs[1].text = '\n' + text


def find_body_paragraph(doc, text):
    """The body-level <w:p> element whose text is exactly `text`."""
    for child in doc.element.body.iterchildren():
        if child.tag == f'{W}p':
            joined = ''.join(t.text or '' for t in child.iter(f'{W}t'))
            if joined.strip() == text:
                return child
    raise LookupError(f'body paragraph {text!r} not found in template')


def para_starting(doc, prefix):
    """Body-level paragraph whose text starts with `prefix`. Indexing
    doc.paragraphs by number is not safe here - it skips the four tables."""
    for p in doc.paragraphs:
        if p.text.strip().startswith(prefix):
            return p
    raise LookupError(f'no body paragraph starting {prefix!r}')


def para_after(doc, heading):
    """The paragraph immediately following a body-level heading paragraph."""
    nxt = find_body_paragraph(doc, heading).getnext()
    if nxt is None or nxt.tag != f'{W}p':
        raise LookupError(f'no paragraph after {heading!r}')
    return Paragraph(nxt, doc)


# ------------------------------------------------------------------ the picture

def _font(px):
    for name in ('segoeui.ttf', 'arial.ttf', 'calibri.ttf'):
        try:
            return ImageFont.truetype(name, px)
        except OSError:
            continue
    return ImageFont.load_default()


def placeholder_png(title, path):
    w, h = int(PH_W_IN * 96 * PH_SCALE), int(PH_H_IN * 96 * PH_SCALE)
    img = Image.new('RGB', (w, h), PH_BG)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w - 1, h - 1], outline=PH_BORDER, width=2 * PH_SCALE)

    font = _font(15 * PH_SCALE)
    lines, cur = [], ''
    for word in f'Screenshot: {title}'.split():
        trial = f'{cur} {word}'.strip()
        if cur and d.textlength(trial, font=font) > w * 0.8:
            lines.append(cur)
            cur = word
        else:
            cur = trial
    lines.append(cur)

    lh = int(22 * PH_SCALE)
    y = (h - lh * len(lines)) // 2
    for line in lines:
        d.text(((w - d.textlength(line, font=font)) / 2, y), line, fill=PH_FG, font=font)
        y += lh
    img.save(path, 'PNG')
    return path


# ------------------------------------------------------------------ table fill

def fill_cell(cell, item, png):
    """Pour one item into a cloned prototype cell.

    Prototype paragraphs: 0 title, 1 summary, 2 blank, 3 'Details',
    4-6 bullets, 7 blank, 8 'Benefit', 9 blank, 10 image, 11 blank.
    References are captured before mutating so they survive bullet cloning.
    """
    ps = cell.paragraphs
    if len(ps) != 12:
        raise ValueError(f'prototype cell has {len(ps)} paragraphs, expected 12')
    p_title, p_summary, p_details = ps[0], ps[1], ps[3]
    proto_bullet = copy.deepcopy(ps[4]._p)
    old_bullets = [ps[4], ps[5], ps[6]]
    p_benefit, p_img = ps[8], ps[10]

    set_text(p_title, item['title'])
    set_text(p_summary, item['summary'])
    set_benefit(p_benefit, item['benefit'])

    details = item.get('details') or []
    if not details:
        raise ValueError(f'item {item["title"]!r} has no details bullets')
    for b in old_bullets:
        b._p.getparent().remove(b._p)
    prev = p_details._p
    for text in details:
        el = copy.deepcopy(proto_bullet)
        prev.addnext(el)
        prev = el
        set_text(Paragraph(el, cell), text)

    p_img.add_run().add_picture(str(png), width=Inches(PH_W_IN))


def fill_table(table, items, tmpdir, tag):
    proto = table.rows[0]._tr
    for i, item in enumerate(items, 1):
        table._tbl.append(copy.deepcopy(proto))
        row = table.rows[-1]
        set_text(row.cells[0].paragraphs[0], str(i))
        png = placeholder_png(item['title'], tmpdir / f'{tag}{i}.png')
        fill_cell(row.cells[1], item, png)
    table._tbl.remove(proto)


def drop_section(doc, heading):
    """Remove a heading paragraph, its table, and the blank line after it."""
    p = find_body_paragraph(doc, heading)
    tbl = p.getnext()
    trailing = tbl.getnext() if tbl is not None else None
    if trailing is not None and trailing.tag != f'{W}p':
        trailing = None
    for el in (p, tbl, trailing):
        if el is not None:
            el.getparent().remove(el)


# ----------------------------------------------------------------------- build

def build(data, out_path):
    for key in ('version', 'date'):
        if not data.get(key):
            raise ValueError(f'missing "{key}" in input JSON')
    if not re.fullmatch(r'\d{2}/\d{2}/\d{4}', data['date']):
        raise ValueError(f'date must be MM/DD/YYYY, got {data["date"]!r}')
    buckets = {key: data.get(key) or [] for key, _ in SECTIONS}
    if not any(buckets.values()):
        raise ValueError('no items in any section')

    doc = docx.Document(TEMPLATE)
    tables = {name: doc.tables[i + 1] for i, (_, name) in enumerate(SECTIONS)}

    v = doc.tables[0].rows[1].cells
    set_text(v[0].paragraphs[0], data['version'])
    set_text(v[1].paragraphs[0], data['date'])
    set_text(v[2].paragraphs[0], data.get('release_title', RELEASE_TITLE))

    intro = para_starting(doc, 'We are thrilled')
    set_text(intro, re.sub(r'Version\s*[\d.]+\s*:', f'Version {data["version"]}:', intro.text))

    counts = [f'{len(buckets[k])} {name}' for k, name in SECTIONS if buckets[k]]
    set_text(para_after(doc, 'Highlights'), ' | '.join(counts))

    with tempfile.TemporaryDirectory() as td:
        tmpdir = Path(td)
        for key, name in SECTIONS:
            if buckets[key]:
                fill_table(tables[name], buckets[key], tmpdir, key)
            else:
                drop_section(doc, name)
        doc.save(out_path)
    return out_path


def default_name(date):
    return f'Release Notes {date.replace("/", "-")}.docx'


# ------------------------------------------------------------------- selfcheck

def selfcheck():
    data = json.loads((HERE / 'example.json').read_text(encoding='utf8'))
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / 'check.docx'
        build(data, out)
        d = docx.Document(out)

        v = d.tables[0].rows[1].cells
        assert v[0].text == data['version'], v[0].text
        assert v[1].text == data['date'], v[1].text
        assert v[2].text == RELEASE_TITLE, v[2].text

        want = ' | '.join(f'{len(data[k])} {n}' for k, n in SECTIONS if data.get(k))
        got = para_after(d, 'Highlights').text
        assert got == want, got
        intro = para_starting(d, 'We are thrilled').text
        assert f'Version {data["version"]}:' in intro, intro

        for i, (key, name) in enumerate(SECTIONS):
            t = d.tables[i + 1]
            items = data[key]
            assert len(t.rows) == len(items), f'{name}: {len(t.rows)} rows != {len(items)}'
            for n, (row, item) in enumerate(zip(t.rows, items), 1):
                assert row.cells[0].text.strip() == str(n), row.cells[0].text
                ps = row.cells[1].paragraphs
                assert ps[0].text == item['title'], ps[0].text
                assert ps[1].text == item['summary'], ps[1].text
                assert ps[3].text == 'Details', ps[3].text
                bullets = [p.text for p in ps
                           if p._p.find(f'{W}pPr/{W}numPr') is not None]
                assert bullets == item['details'], bullets
                ben = next(p for p in ps if p.text.startswith('Benefit'))
                assert ben.text == 'Benefit\n' + item['benefit'], ben.text
                assert row.cells[1]._tc.xml.count('graphicData') == 2, \
                    f'{name} row {n}: expected exactly one image'

        # a section with no items disappears entirely
        data['bug_fixes'] = []
        empty_out = Path(td) / 'empty.docx'
        build(data, empty_out)
        e = docx.Document(empty_out)
        assert 'Bug Fixes' not in [p.text.strip() for p in e.paragraphs]
        assert len(e.tables) == 3, len(e.tables)
        line = para_after(e, 'Highlights').text
        assert 'Bug Fixes' not in line, line
    print('selfcheck OK')


# ------------------------------------------------------------------------ main

def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('items', nargs='?', help='path to the items JSON')
    ap.add_argument('-o', '--out',
                    help='output .docx (default: "Release Notes MM-DD-YYYY.docx")')
    ap.add_argument('--selfcheck', action='store_true',
                    help='run built-in assertions and exit')
    a = ap.parse_args()

    if a.selfcheck:
        selfcheck()
        return
    if not a.items:
        ap.error('items JSON required (or use --selfcheck)')

    data = json.loads(Path(a.items).read_text(encoding='utf8'))
    out = Path(a.out) if a.out else Path.cwd() / default_name(data['date'])
    build(data, out)
    print(out)


if __name__ == '__main__':
    main()
