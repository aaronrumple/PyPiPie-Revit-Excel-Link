# -*- coding: utf-8 -*-
"""Turn a worksheet range into drawable primitives measured in paper points
(1 pt = 1/72 in), origin at the range's top-left corner, +y DOWN.

Revit-agnostic so it can be tested (and previewed) outside Revit.
"""
from __future__ import division

import math

from . import numfmt
from .xlsx_reader import parse_range

PX = 0.75  # points per pixel at 96 dpi
WRAP_FIT = 0.995  # wrap a hair early: Excel breaks lines that only just fit

# Approximate digit-advance (em fraction) for common default fonts; used to
# derive Excel's "maximum digit width" which drives column widths.
DIGIT_EM = {
    'calibri': 0.507, 'calibri light': 0.507, 'aptos': 0.552, 'aptos narrow': 0.47,
    'arial': 0.556, 'arial narrow': 0.456, 'helvetica': 0.556, 'liberation sans': 0.556,
    'times new roman': 0.5, 'cambria': 0.555, 'verdana': 0.636, 'tahoma': 0.546,
    'segoe ui': 0.559, 'consolas': 0.55, 'courier new': 0.6, 'georgia': 0.614,
    'century gothic': 0.552, 'franklin gothic book': 0.54,
}

# Higher wins when two cells disagree about a shared edge.
BORDER_RANK = {
    'hair': 1, 'dotted': 2, 'dashDotDot': 3, 'dashDot': 4, 'dashed': 5, 'thin': 6,
    'mediumDashDotDot': 7, 'slantDashDot': 8, 'mediumDashDot': 9,
    'mediumDashed': 10, 'medium': 11, 'thick': 12, 'double': 13,
}


def max_digit_width(font_name, size_pt):
    em = DIGIT_EM.get((font_name or '').lower(), 0.52)
    return max(1, int(round(size_pt * 96.0 / 72.0 * em)))


class TextItem(object):
    __slots__ = ('text', 'x', 'y', 'halign', 'valign', 'width', 'rotation',
                 'font', 'runs', 'cell', 'rect', 'wrap', 'shrink')

    def __init__(self, **kw):
        for k in self.__slots__:
            setattr(self, k, kw.get(k))

    def __repr__(self):
        return 'Text(%r @%.1f,%.1f %s/%s w=%s)' % (
            self.text, self.x, self.y, self.halign, self.valign, self.width)


class Segment(object):
    __slots__ = ('x1', 'y1', 'x2', 'y2', 'style', 'color')

    def __init__(self, x1, y1, x2, y2, style, color):
        self.x1, self.y1, self.x2, self.y2 = x1, y1, x2, y2
        self.style = style
        self.color = color or '000000'

    def __repr__(self):
        return 'Seg(%.1f,%.1f-%.1f,%.1f %s #%s)' % (
            self.x1, self.y1, self.x2, self.y2, self.style, self.color)


class FillRect(object):
    __slots__ = ('x', 'y', 'w', 'h', 'color')

    def __init__(self, x, y, w, h, color):
        self.x, self.y, self.w, self.h, self.color = x, y, w, h, color


class Layout(object):
    def __init__(self):
        self.width = 0.0
        self.height = 0.0
        self.fills = []
        self.segments = []
        self.texts = []
        self.warnings = []
        self.range = None
        self.mdw = 7


class Options(object):
    def __init__(self, include_hidden=False, gridlines=False, fills=True,
                 borders=True, text=True, mdw=None):
        self.include_hidden = include_hidden
        self.gridlines = gridlines
        self.fills = fills
        self.borders = borders
        self.text = text
        self.mdw = mdw


# --------------------------------------------------------------------------
def build_layout(sheet, rng=None, opts=None):
    opts = opts or Options()
    wb = sheet.wb
    st = wb.styles
    lay = Layout()

    if rng:
        r1, c1, r2, c2 = parse_range(rng)
        ur1, uc1, ur2, uc2 = sheet.used_range()
        r2 = min(r2, max(ur2, r1))
        c2 = min(c2, max(uc2, c1))
    else:
        r1, c1, r2, c2 = sheet.default_range()
    lay.range = (r1, c1, r2, c2)

    base_font = st.fonts[0] if st.fonts else None
    mdw = opts.mdw or max_digit_width(base_font.name if base_font else 'Calibri',
                                      base_font.size if base_font else 11)
    lay.mdw = mdw

    # ---- column & row geometry ----------------------------------------
    def col_px(c):
        w = sheet.col_widths.get(c)
        if w is None:
            if sheet.default_col_width is not None:
                w = sheet.default_col_width
            else:
                px = 8 * int(math.ceil((sheet.base_col_width * mdw + 5) / 8.0))
                return px
        return int(((256 * w + int(128 / mdw)) / 256.0) * mdw)

    default_rh = sheet.default_row_height or (15.0 if not base_font else
                                              max(15.0, round(base_font.size * 1.36 * 4) / 4))

    col_x, col_w = {}, {}
    x = 0.0
    for c in range(c1, c2 + 1):
        hidden = c in sheet.col_hidden and not opts.include_hidden
        w = 0.0 if hidden else col_px(c) * PX
        if c in sheet.col_widths and sheet.col_widths[c] == 0:
            w = 0.0
        col_x[c], col_w[c] = x, w
        x += w
    col_x[c2 + 1] = x
    lay.width = x

    # Rows saved without a height (files written by tools other than Excel)
    # get Excel's auto-fit height for their largest font.
    row_font_max = {}
    for (cr, cc), cell in sheet.cells.items():
        if r1 <= cr <= r2 and cell.value not in (None, u'') and cr not in sheet.row_heights:
            xf = sheet.xf(cr, cc)
            if 0 <= xf.font_id < len(st.fonts):
                sz = st.fonts[xf.font_id].size
                if sz > row_font_max.get(cr, 0):
                    row_font_max[cr] = sz

    def auto_height(r):
        sz = row_font_max.get(r)
        if not sz:
            return default_rh
        return max(default_rh, math.ceil(sz * 1.32 / PX) * PX)

    row_y, row_h = {}, {}
    y = 0.0
    for r in range(r1, r2 + 1):
        hidden = r in sheet.row_hidden and not opts.include_hidden
        h = 0.0 if hidden else (sheet.row_heights[r] if r in sheet.row_heights else auto_height(r))
        row_y[r], row_h[r] = y, h
        y += h
    row_y[r2 + 1] = y
    lay.height = y

    def visible(r, c):
        return row_h[r] > 0 and col_w[c] > 0

    # ---- merges ------------------------------------------------------------
    merge_root = {}
    merge_of = {}
    for m in sheet.merges:
        mr1, mc1, mr2, mc2 = m
        if mr2 < r1 or mr1 > r2 or mc2 < c1 or mc1 > c2:
            continue
        cm = (max(mr1, r1), max(mc1, c1), min(mr2, r2), min(mc2, c2))
        for rr in range(cm[0], cm[2] + 1):
            for cc in range(cm[1], cm[3] + 1):
                merge_root[(rr, cc)] = (mr1, mc1)
                merge_of[(rr, cc)] = cm

    def rect_of(r, c):
        if (r, c) in merge_of:
            a, b, cc2, d = merge_of[(r, c)]
            return col_x[b], row_y[a], col_x[d + 1] - col_x[b], row_y[cc2 + 1] - row_y[a]
        return col_x[c], row_y[r], col_w[c], row_h[r]

    # ---- fills -------------------------------------------------------------
    fill_color = {}
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            key = merge_root.get((r, c), (r, c))
            xf = sheet.xf(*key)
            f = st.fills[xf.fill_id] if 0 <= xf.fill_id < len(st.fills) else None
            col = f.display_color() if f else None
            if col:
                fill_color[(r, c)] = col

    if opts.fills:
        done = set()
        for (r, c), m in merge_of.items():
            root = merge_root[(r, c)]
            if root in done or (r, c) not in fill_color:
                continue
            done.add(root)
            x0, y0, w, h = rect_of(r, c)
            if w > 0 and h > 0:
                lay.fills.append(FillRect(x0, y0, w, h, fill_color[(r, c)]))
        # greedy rectangles for the rest: horizontal runs over visible
        # columns, then stack identical runs on consecutive visible rows
        vis_cols = [c for c in range(c1, c2 + 1) if col_w[c] > 0]
        vis_rows = [r for r in range(r1, r2 + 1) if row_h[r] > 0]
        open_runs = {}  # (start_col, end_col, colour) -> (first_row, last_row)

        def flush(key, span):
            s, e, col = key
            x0, y0 = col_x[s], row_y[span[0]]
            lay.fills.append(FillRect(x0, y0, col_x[e] + col_w[e] - x0,
                                      row_y[span[1]] + row_h[span[1]] - y0, col))

        for r in vis_rows:
            runs = []
            cur = None
            for c in vis_cols:
                col = None if (r, c) in merge_of else fill_color.get((r, c))
                if cur and col == cur[2]:
                    cur[1] = c
                else:
                    if cur and cur[2]:
                        runs.append(tuple(cur))
                    cur = [c, c, col]
            if cur and cur[2]:
                runs.append(tuple(cur))
            new_open = {}
            for run in runs:
                if run in open_runs:
                    new_open[run] = (open_runs.pop(run)[0], r)
                else:
                    new_open[run] = (r, r)
            for key, span in open_runs.items():
                flush(key, span)
            open_runs = new_open
        for key, span in open_runs.items():
            flush(key, span)

    # ---- borders -----------------------------------------------------------
    edges = {}  # ('h'|'v', fixed coord, cell index along) -> (rank, style, color, a, b)

    def offer(kind, fixed, along, a, b, side):
        if not side or a >= b:
            return
        key = (kind, round(fixed, 3), along)
        rank = BORDER_RANK.get(side.style, 6)
        cur = edges.get(key)
        if cur is None or rank > cur[0]:
            edges[key] = (rank, side.style, side.color, a, b)

    diagonals = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            if not visible(r, c):
                continue
            xf = sheet.xf(r, c)
            bd = st.borders[xf.border_id] if 0 <= xf.border_id < len(st.borders) else None
            if bd is None:
                continue
            m = merge_of.get((r, c))
            x0, y0, w, h = col_x[c], row_y[r], col_w[c], row_h[r]
            if m is None or r == m[0]:
                offer('h', y0, c, x0, x0 + w, bd.top)
            if m is None or r == m[2]:
                offer('h', y0 + h, c, x0, x0 + w, bd.bottom)
            if m is None or c == m[1]:
                offer('v', x0, r, y0, y0 + h, bd.left)
            if m is None or c == m[3]:
                offer('v', x0 + w, r, y0, y0 + h, bd.right)
            if bd.diagonal and (bd.diag_up or bd.diag_down) and (m is None or (r, c) == (m[0], m[1])):
                rx, ry, rw, rh = rect_of(r, c)
                if bd.diag_down:
                    diagonals.append(Segment(rx, ry, rx + rw, ry + rh, bd.diagonal.style, bd.diagonal.color))
                if bd.diag_up:
                    diagonals.append(Segment(rx, ry + rh, rx + rw, ry, bd.diagonal.style, bd.diagonal.color))

    if opts.gridlines:
        grid_side = _Side('grid', 'C0C0C0')
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if not visible(r, c):
                    continue
                x0, y0, w, h = col_x[c], row_y[r], col_w[c], row_h[r]
                filled = (r, c) in fill_color
                m = merge_of.get((r, c))
                for kind, fixed, along, a, b, edge_ok, nb in (
                        ('h', y0, c, x0, x0 + w, m is None or r == m[0], (r - 1, c)),
                        ('h', y0 + h, c, x0, x0 + w, m is None or r == m[2], (r + 1, c)),
                        ('v', x0, r, y0, y0 + h, m is None or c == m[1], (r, c - 1)),
                        ('v', x0 + w, r, y0, y0 + h, m is None or c == m[3], (r, c + 1))):
                    if not edge_ok or filled or nb in fill_color:
                        continue
                    key = (kind, round(fixed, 3), along)
                    if key not in edges:
                        edges[key] = (0, 'grid', 'C0C0C0', a, b)

    if opts.borders or opts.gridlines:
        groups = {}
        for (kind, fixed, along), (rank, style, color, a, b) in edges.items():
            if not opts.borders and style != 'grid':
                continue
            groups.setdefault((kind, fixed, style, color or '000000'), []).append((a, b))
        for (kind, fixed, style, color), spans in groups.items():
            spans.sort()
            merged = []
            for a, b in spans:
                if merged and a <= merged[-1][1] + 1e-6:
                    merged[-1][1] = max(merged[-1][1], b)
                else:
                    merged.append([a, b])
            for a, b in merged:
                offs = (-PX, PX) if style == 'double' else (0.0,)
                for o in offs:
                    if kind == 'h':
                        lay.segments.append(Segment(a, fixed + o, b, fixed + o, style, color))
                    else:
                        lay.segments.append(Segment(fixed + o, a, fixed + o, b, style, color))
        if opts.borders:
            lay.segments.extend(diagonals)

    # ---- text --------------------------------------------------------------
    if opts.text:
        _layout_text(sheet, lay, r1, c1, r2, c2, col_x, col_w, row_y, row_h,
                     merge_root, merge_of, rect_of, visible, mdw)
    return lay


class _Side(object):
    def __init__(self, style, color):
        self.style, self.color = style, color

    def __bool__(self):
        return True
    __nonzero__ = __bool__


def _resolve_runs(st, cell_font, runs):
    """Return (base_font, [(start, length, Font)]) for rich text."""
    out = []
    pos = 0
    for run in runs:
        f = st.font_from(run.font, base=cell_font) if run.font is not None else cell_font
        n = len(run.text)
        if n:
            out.append((pos, n, f))
        pos += n
    base = out[0][2] if out else cell_font
    return base, out


def _layout_text(sheet, lay, r1, c1, r2, c2, col_x, col_w, row_y, row_h,
                 merge_root, merge_of, rect_of, visible, mdw):
    st = sheet.wb.styles
    pad = 2 * PX
    vpad = 1 * PX
    for (r, c), cell in sorted(sheet.cells.items()):
        if not (r1 <= r <= r2 and c1 <= c <= c2):
            continue
        if merge_root.get((r, c), (r, c)) != (r, c):
            continue  # only the merge's top-left shows text
        if cell.value is None or cell.value == u'':
            continue
        rx, ry, rw, rh = rect_of(r, c)
        if rw <= 0 or rh <= 0:
            continue
        xf = sheet.xf(r, c)
        code = st.num_fmt_code(xf.num_fmt_id)
        text, fmt_color = numfmt.format_value(cell.value, cell.type, code, sheet.wb.date1904)
        if not text.strip(u' \x01'):
            continue
        font = (st.fonts[xf.font_id] if 0 <= xf.font_id < len(st.fonts) else st.fonts[0]).copy()
        runs = None
        if cell.runs:
            font, rr = _resolve_runs(st, font, cell.runs)
            font = font.copy()
            runs = [(s, n, f.bold, f.italic, bool(f.underline), f.vert_align) for s, n, f in rr]
            if all(rn[2:] == (font.bold, font.italic, bool(font.underline), font.vert_align) for rn in runs):
                runs = None
            diff = set((f.name, f.size, f.color) for _, _, f in rr)
            if len(diff) > 1:
                lay.warnings.append('%s: mixed font/size/colour inside one cell; using the first run\'s.'
                                    % _addr(r, c))
        if fmt_color:
            font.color = fmt_color

        h = xf.horizontal or 'general'
        is_num = cell.type in ('n', 'd') and not isinstance(cell.value, str)
        if h == 'general':
            h = 'right' if is_num else ('center' if cell.type in ('b', 'e') else 'left')
        elif h in ('fill', 'justify'):
            h = 'left'
        elif h == 'distributed':
            h = 'center'
        elif h == 'centerContinuous':
            cc = c + 1
            while cc <= c2:
                nxt = sheet.cells.get((r, cc))
                if (nxt is not None and nxt.value not in (None, u'')) or \
                        sheet.xf(r, cc).horizontal != 'centerContinuous':
                    break
                cc += 1
            rw = col_x[cc] - rx
            h = 'center'
        v = xf.vertical or 'bottom'
        v = {'center': 'middle', 'justify': 'top', 'distributed': 'middle'}.get(v, v)
        wrap = xf.wrap or xf.horizontal in ('justify', 'distributed') or (xf.vertical in ('justify', 'distributed'))

        def add(item, _shrink=xf.shrink):
            item.shrink = bool(_shrink)
            lay.texts.append(item)

        rot = xf.rotation or 0
        if rot == 255:
            text = u'\n'.join(ch for ch in text.replace(u'\x01', u'') if ch != u'\n')
            rot, h, v, wrap = 0, 'center', v, False
        elif rot > 90:
            rot = -(rot - 90)

        indent_pt = xf.indent * mdw * 9.0 / 7.0 * PX if xf.indent else 0.0
        left = rx + pad + (indent_pt if h == 'left' else 0)
        right = rx + rw - pad - (indent_pt if h == 'right' else 0)
        # Lines that only just fit are a toss-up between Excel and Revit; wrap
        # a hair early so near-ties break like Excel without adding lines.
        width = max((right - left) * WRAP_FIT, 1.0)

        def anchor_y():
            if v == 'top':
                return ry + vpad
            if v == 'middle':
                return ry + rh / 2.0
            return ry + rh - vpad

        if rot:
            add(TextItem(text=text.replace(u'\x01', u' '), x=rx + rw / 2.0, y=ry + rh / 2.0,
                                      halign='center', valign='middle', width=None, rotation=rot,
                                      font=font, runs=None, cell=(r, c), rect=(rx, ry, rw, rh),
                                      wrap=False))
            continue

        if u'\x01' in text and not wrap:
            # accounting-style: left part hugs left edge, right part right edge
            lpart, rpart = text.split(u'\x01', 1)
            rpart = rpart.replace(u'\x01', u'')
            if lpart.strip():
                add(TextItem(text=lpart.strip(), x=rx + pad + (len(lpart) - len(lpart.lstrip())) * mdw * 0.5 * PX,
                                          y=anchor_y(), halign='left', valign=v, width=None,
                                          rotation=0, font=font, runs=None, cell=(r, c),
                                          rect=(rx, ry, rw, rh), wrap=False))
            if rpart.strip():
                trail = len(rpart) - len(rpart.rstrip())
                add(TextItem(text=rpart.strip(), x=rx + rw - pad - trail * mdw * 0.5 * PX,
                                          y=anchor_y(), halign='right', valign=v, width=None,
                                          rotation=0, font=font, runs=None, cell=(r, c),
                                          rect=(rx, ry, rw, rh), wrap=False))
            continue
        text = text.replace(u'\x01', u' ')
        # trailing "_)" padding spaces shift right-aligned numbers left
        if h == 'right':
            trail = len(text) - len(text.rstrip(u' '))
            right -= trail * mdw * 0.5 * PX
            text = text.rstrip(u' ')
        if h == 'left':
            lead = len(text) - len(text.lstrip(u' '))
            left += lead * mdw * 0.5 * PX
            text = text.lstrip(u' ')

        if wrap:
            ax = {'left': left, 'center': (left + right) / 2.0, 'right': right}[h]
            add(TextItem(text=text, x=ax, y=anchor_y(), halign=h, valign=v,
                                      width=width, rotation=0, font=font, runs=runs,
                                      cell=(r, c), rect=(rx, ry, rw, rh), wrap=True))
        else:
            ax = {'left': left, 'center': rx + rw / 2.0, 'right': right}[h]
            add(TextItem(text=text, x=ax, y=anchor_y(), halign=h, valign=v,
                                      width=None, rotation=0, font=font, runs=runs,
                                      cell=(r, c), rect=(rx, ry, rw, rh), wrap=False))


def _addr(r, c):
    from .xlsx_reader import index_to_col
    return '%s%d' % (index_to_col(c), r)
