#!/usr/bin/env python3
"""
Write a FortuneSheet model to legacy BIFF8 .xls with xlwt.

The plugin calls this only for "另存为 xls" because SheetJS CE's BIFF writer
does not retain formulas. Formulas stay formulas; Excel/WPS/LibreOffice
recalculate them on open. Styles are approximated to the BIFF palette.
"""
import json
import math
import re
import sys

sys.path.insert(0, __file__.rsplit('/', 3)[0] + '/vendor/python')
import xlwt  # noqa: E402


def rgb_of_hex(value):
    if not isinstance(value, str):
        return None
    text = value.strip().lstrip('#')
    if len(text) == 8:
        text = text[2:]
    if len(text) != 6:
        return None
    try:
        return int(text, 16)
    except ValueError:
        return None


PALETTE = []
for _name, _value in getattr(xlwt.Style, 'colour_map', {}).items():
    if isinstance(_value, tuple):
        _value = (_value[0] << 16) | (_value[1] << 8) | _value[2]
    if isinstance(_value, int):
        PALETTE.append((_name, _value))


def nearest_colour(hex_value):
    rgb = rgb_of_hex(hex_value)
    if rgb is None:
        return None
    if not PALETTE:
        return None
    best = min(PALETTE, key=lambda item: (
        ((item[1] >> 16) & 0xFF) - ((rgb >> 16) & 0xFF)) ** 2
        + (((item[1] >> 8) & 0xFF) - ((rgb >> 8) & 0xFF)) ** 2
        + ((item[1] & 0xFF) - (rgb & 0xFF)) ** 2)
    return best[0]


def make_style(data):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.bold = bool(data.get('bl'))
    font.italic = bool(data.get('it'))
    font.underline = xlwt.Font.UNDERLINE_SINGLE if data.get('un') else xlwt.Font.UNDERLINE_NONE
    font.struck_out = bool(data.get('cl'))
    size = data.get('fs')
    if isinstance(size, (int, float)) and size > 0:
        font.height = int(round(float(size) * 20))
    if data.get('ff'):
        font.name = str(data['ff'])
    colour = data.get('fc')
    if colour:
        name = nearest_colour(colour)
        if name:
            font.colour_index = xlwt.Style.colour_map[name]
    style.font = font

    if data.get('bg'):
        name = nearest_colour(data['bg'])
        if name:
            pattern = xlwt.Pattern()
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN
            pattern.pattern_fore_colour = xlwt.Style.colour_map[name]
            style.pattern = pattern

    align = xlwt.Alignment()
    if data.get('ht') == 0:
        align.horz = xlwt.Alignment.HORZ_CENTER
    elif data.get('ht') == 2:
        align.horz = xlwt.Alignment.HORZ_RIGHT
    elif data.get('ht') == 1:
        align.horz = xlwt.Alignment.HORZ_LEFT
    if data.get('vt') == 0:
        align.vert = xlwt.Alignment.VERT_CENTER
    elif data.get('vt') == 2:
        align.vert = xlwt.Alignment.VERT_BOTTOM
    elif data.get('vt') == 1:
        align.vert = xlwt.Alignment.VERT_TOP
    if str(data.get('tb')) == '2':
        align.wrap = 1
    style.alignment = align

    fmt = (data.get('ct') or {}).get('fa')
    if fmt and fmt != 'General':
        style.num_format_str = str(fmt)
    return style


def main():
    if len(sys.argv) != 3:
        raise SystemExit('usage: xls-writer.py workbook.json output.xls')
    with open(sys.argv[1], 'r', encoding='utf-8') as handle:
        payload = json.load(handle)
    workbook = xlwt.Workbook(encoding='utf-8')
    for sheet in payload.get('sheets') or []:
        name = str(sheet.get('name') or 'Sheet1')[:31]
        worksheet = workbook.add_sheet(name)
        style_cache = {}
        config = sheet.get('config') or {}
        merges = []
        covered = set()
        for merge in (config.get('merge') or {}).values():
            if not isinstance(merge, dict):
                continue
            r = int(merge.get('r', 0))
            c = int(merge.get('c', 0))
            rs = int(merge.get('rs') or 1)
            cs = int(merge.get('cs') or 1)
            if rs > 1 or cs > 1:
                merges.append((r, c, rs, cs))
                for rr in range(r, r + rs):
                    for cc in range(c, c + cs):
                        if (rr, cc) != (r, c):
                            covered.add((rr, cc))
        for entry in sheet.get('celldata') or []:
            data = entry.get('v') if isinstance(entry, dict) else None
            if not isinstance(data, dict):
                continue
            if data.get('v') is None and not data.get('f'):
                runs = (data.get('ct') or {}).get('s')
                if isinstance(runs, list):
                    data = {**data, 'v': ''.join(str(run.get('v', '')) for run in runs if isinstance(run, dict))}
            row = int(entry.get('r', 0))
            col = int(entry.get('c', 0))
            if row < 0 or col < 0 or row > 65535 or col > 255:
                continue
            if (row, col) in covered:
                continue
            key = json.dumps({k: data.get(k) for k in ('bl', 'it', 'un', 'cl', 'fs', 'ff', 'fc', 'bg', 'ht', 'vt', 'tb', 'ct')}, sort_keys=True)
            style = style_cache.get(key)
            if style is None:
                style = make_style(data)
                style_cache[key] = style
            formula = data.get('f')
            value = data.get('v')
            if isinstance(formula, str) and formula.strip():
                xls_formula = formula.lstrip('=')
                # xlwt's formula parser accepts bare TRUE/FALSE but not TRUE()/FALSE().
                xls_formula = re.sub(r'\b(TRUE|FALSE)\(\)', r'\1', xls_formula, flags=re.IGNORECASE)
                worksheet.write(row, col, xlwt.Formula(xls_formula), style)
            elif value is not None:
                if isinstance(value, bool):
                    worksheet.write(row, col, bool(value), style)
                elif isinstance(value, (int, float)) and not isinstance(value, bool):
                    worksheet.write(row, col, float(value), style)
                else:
                    worksheet.write(row, col, str(value), style)
        for r, c, rs, cs in merges:
            worksheet.merge(r, r + rs - 1, c, c + cs - 1)
        for index, width in (config.get('columnlen') or {}).items():
            try:
                idx = int(index)
                worksheet.col(idx).width = max(256, min(65535, int(float(width) * 256 / 7)))
            except (TypeError, ValueError):
                pass
        for index, height in (config.get('rowlen') or {}).items():
            try:
                idx = int(index)
                row = worksheet.row(idx)
                row.height_mismatch = True
                row.height = max(1, int(float(height) * 15))
            except (TypeError, ValueError):
                pass
        for index in (config.get('colhidden') or {}):
            try:
                worksheet.col(int(index)).hidden = True
            except (TypeError, ValueError):
                pass
        for index in (config.get('rowhidden') or {}):
            try:
                worksheet.row(int(index)).hidden = True
            except (TypeError, ValueError):
                pass
        rows = int(sheet.get('frozenRows') or 0)
        cols = int(sheet.get('frozenColumns') or 0)
        if rows or cols:
            worksheet.set_panes_frozen(True)
            if rows:
                worksheet.set_horz_split_pos(rows)
            if cols:
                worksheet.set_vert_split_pos(cols)
            worksheet.set_remove_splits(True)
    workbook.save(sys.argv[2])


if __name__ == '__main__':
    main()
