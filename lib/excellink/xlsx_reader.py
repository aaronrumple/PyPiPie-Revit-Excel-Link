# -*- coding: utf-8 -*-
"""Minimal, dependency-free .xlsx reader focused on *visual* fidelity.

Reads cell values, shared/rich strings, number formats, fonts, fills,
borders, alignment, theme/indexed colours with tint, column widths, row
heights, hidden rows/columns, merged ranges and print areas.

Pure Python (zipfile + ElementTree) so it runs under IronPython 2.7,
IronPython 3 and CPython 3 (pyRevit engines) and can be unit-tested
outside Revit.
"""
from __future__ import division

import colorsys
import io
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

try:
    text_type = unicode  # noqa: F821  (IronPython 2.7)
except NameError:
    text_type = str


# --------------------------------------------------------------------------
# File access
# --------------------------------------------------------------------------
def read_file_bytes(path):
    """Read a file even when Excel has it open (falls back to a shared
    read/write .NET stream if a plain open is refused)."""
    try:
        with open(path, 'rb') as f:
            return f.read()
    except (IOError, OSError):
        pass
    from System.IO import FileStream, FileMode, FileAccess, FileShare, MemoryStream  # noqa
    fs = FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)
    try:
        ms = MemoryStream()
        fs.CopyTo(ms)
        arr = ms.ToArray()
    finally:
        fs.Close()
    try:
        return bytes(arr)  # pythonnet
    except Exception:
        import System
        return System.Text.Encoding.GetEncoding(28591).GetString(arr)  # IronPython 2 str


def _byte_stream(data):
    """File-like object over raw bytes that zipfile accepts on every engine."""
    if sys.version_info[0] < 3:
        import StringIO
        if not isinstance(data, str):
            data = data.decode('latin-1')  # IronPython 2: 1 char per byte
        return StringIO.StringIO(data)
    return io.BytesIO(data)


def _strip_ns(root):
    for el in root.iter():
        if isinstance(el.tag, str) and '}' in el.tag:
            el.tag = el.tag.split('}', 1)[1]
        if el.attrib:
            new = {}
            for k, v in el.attrib.items():
                if '}' in k:
                    ns, local = k[1:].split('}', 1)
                    if 'relationships' in ns:
                        local = 'r:' + local
                    elif ns.endswith('/XML/1998/namespace'):
                        local = 'xml:' + local
                    new[local] = v
                else:
                    new[k] = v
            el.attrib.clear()
            el.attrib.update(new)
    return root


def _child(el, name):
    if el is None:
        return None
    for c in el:
        if c.tag == name:
            return c
    return None


def _children(el, name):
    if el is None:
        return []
    return [c for c in el if c.tag == name]


def _bool(v, default=False):
    if v is None:
        return default
    return v.lower() in ('1', 'true', 'on')


_XESC = re.compile(r'_x([0-9A-Fa-f]{4})_')


def _unescape(s):
    if not s:
        return s or u''
    return _XESC.sub(lambda m: unichr_(int(m.group(1), 16)), s)


def unichr_(i):
    try:
        return unichr(i)  # noqa: F821
    except NameError:
        return chr(i)


# --------------------------------------------------------------------------
# Cell references
# --------------------------------------------------------------------------
_REF = re.compile(r'^\$?([A-Za-z]{1,3})\$?(\d+)$')


def col_to_index(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def index_to_col(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_ref(ref):
    m = _REF.match(ref.strip())
    if not m:
        raise ValueError('Bad cell reference: %s' % ref)
    return int(m.group(2)), col_to_index(m.group(1))


def parse_range(rng):
    """'A1:D10' -> (r1, c1, r2, c2); also accepts single cell, 'Sheet'!A1:B2,
    and whole column/row refs like A:D or 1:5."""
    rng = rng.strip()
    if '!' in rng:
        rng = rng.split('!')[-1]
    rng = rng.replace('$', '')
    if ':' not in rng:
        r, c = parse_ref(rng)
        return r, c, r, c
    a, b = rng.split(':', 1)
    if a.isalpha() and b.isalpha():
        return 1, col_to_index(a), 1048576, col_to_index(b)
    if a.isdigit() and b.isdigit():
        return int(a), 1, int(b), 16384
    r1, c1 = parse_ref(a)
    r2, c2 = parse_ref(b)
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def format_range(r1, c1, r2, c2):
    return '%s%d:%s%d' % (index_to_col(c1), r1, index_to_col(c2), r2)


# --------------------------------------------------------------------------
# Colours
# --------------------------------------------------------------------------
INDEXED_COLORS = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
]

DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31',
                 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']


def apply_tint(hex6, tint):
    if not tint:
        return hex6
    r, g, b = [int(hex6[i:i + 2], 16) / 255.0 for i in (0, 2, 4)]
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if tint < 0:
        l = l * (1.0 + tint)
    else:
        l = l * (1.0 - tint) + tint
    l = max(0.0, min(1.0, l))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return '%02X%02X%02X' % (int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))


def mix(c1, c2, t):
    """Blend hex colour c1 toward c2 by t (0..1)."""
    a = [int(c1[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(c2[i:i + 2], 16) for i in (0, 2, 4)]
    return '%02X%02X%02X' % tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3))


# --------------------------------------------------------------------------
# Style model
# --------------------------------------------------------------------------
class Font(object):
    __slots__ = ('name', 'size', 'bold', 'italic', 'underline', 'strike',
                 'color', 'vert_align')

    def __init__(self, name='Calibri', size=11.0, bold=False, italic=False,
                 underline=None, strike=False, color=None, vert_align=None):
        self.name = name
        self.size = size
        self.bold = bold
        self.italic = italic
        self.underline = underline
        self.strike = strike
        self.color = color
        self.vert_align = vert_align

    def copy(self):
        return Font(self.name, self.size, self.bold, self.italic,
                    self.underline, self.strike, self.color, self.vert_align)

    def key(self):
        return (self.name, round(self.size, 2), self.bold, self.italic,
                bool(self.underline), self.color or '000000')

    def __repr__(self):
        return 'Font(%s %s%s%s%s #%s)' % (
            self.name, self.size, ' B' if self.bold else '', ' I' if self.italic else '',
            ' U' if self.underline else '', self.color)


class Fill(object):
    __slots__ = ('pattern', 'fg', 'bg', 'stops')

    def __init__(self, pattern='none', fg=None, bg=None, stops=None):
        self.pattern = pattern
        self.fg = fg
        self.bg = bg
        self.stops = stops

    _DENSITY = {
        'gray0625': 0.0625, 'gray125': 0.125, 'lightGray': 0.25,
        'mediumGray': 0.5, 'darkGray': 0.75,
        'lightHorizontal': 0.25, 'lightVertical': 0.25, 'lightDown': 0.25,
        'lightUp': 0.25, 'lightGrid': 0.44, 'lightTrellis': 0.44,
        'darkHorizontal': 0.5, 'darkVertical': 0.5, 'darkDown': 0.5,
        'darkUp': 0.5, 'darkGrid': 0.75, 'darkTrellis': 0.75,
    }

    def display_color(self):
        """Flat colour that best represents this fill, or None."""
        if self.stops:
            cols = [c for _, c in self.stops if c]
            if not cols:
                return None
            out = cols[0]
            for i, c in enumerate(cols[1:], 2):
                out = mix(out, c, 1.0 / i)
            return out
        if not self.pattern or self.pattern == 'none':
            return None
        fg = self.fg or '000000'
        if self.pattern == 'solid':
            return fg
        bg = self.bg or 'FFFFFF'
        d = self._DENSITY.get(self.pattern, 0.5)
        return mix(bg, fg, d)


class BorderSide(object):
    __slots__ = ('style', 'color')

    def __init__(self, style=None, color=None):
        self.style = style
        self.color = color

    def __bool__(self):
        return bool(self.style) and self.style != 'none'
    __nonzero__ = __bool__

    def __repr__(self):
        return 'Side(%s,%s)' % (self.style, self.color)


class Border(object):
    __slots__ = ('left', 'right', 'top', 'bottom', 'diagonal', 'diag_up', 'diag_down')

    def __init__(self):
        self.left = BorderSide()
        self.right = BorderSide()
        self.top = BorderSide()
        self.bottom = BorderSide()
        self.diagonal = BorderSide()
        self.diag_up = False
        self.diag_down = False


class Xf(object):
    __slots__ = ('num_fmt_id', 'font_id', 'fill_id', 'border_id', 'horizontal',
                 'vertical', 'wrap', 'indent', 'rotation', 'shrink')

    def __init__(self):
        self.num_fmt_id = 0
        self.font_id = 0
        self.fill_id = 0
        self.border_id = 0
        self.horizontal = 'general'
        self.vertical = 'bottom'
        self.wrap = False
        self.indent = 0
        self.rotation = 0
        self.shrink = False


class Styles(object):
    def __init__(self, theme_colors):
        self.theme = theme_colors
        self.indexed = list(INDEXED_COLORS)
        self.num_fmts = {}
        self.fonts = [Font()]
        self.fills = [Fill(), Fill('gray125')]
        self.borders = [Border()]
        self.xfs = [Xf()]

    # colours ------------------------------------------------------------
    def color(self, el):
        if el is None:
            return None
        a = el.attrib
        if _bool(a.get('auto')):
            return None
        hexv = None
        if 'rgb' in a:
            v = a['rgb'].strip()
            hexv = v[-6:].upper() if len(v) >= 6 else None
        elif 'theme' in a:
            idx = int(a['theme'])
            if 0 <= idx < len(self.theme):
                hexv = self.theme[idx]
        elif 'indexed' in a:
            idx = int(a['indexed'])
            if idx == 64:
                hexv = '000000'
            elif idx == 65:
                hexv = 'FFFFFF'
            elif 0 <= idx < len(self.indexed):
                hexv = self.indexed[idx]
        if hexv is None:
            return None
        tint = float(a.get('tint', 0) or 0)
        return apply_tint(hexv, tint)

    def font_from(self, el, base=None):
        f = base.copy() if base else Font()
        if el is None:
            return f
        for c in el:
            t = c.tag
            v = c.attrib.get('val')
            if t in ('name', 'rFont'):
                f.name = v or f.name
            elif t == 'sz':
                f.size = float(v)
            elif t == 'b':
                f.bold = _bool(v, True)
            elif t == 'i':
                f.italic = _bool(v, True)
            elif t == 'u':
                f.underline = (v or 'single') if v != 'none' else None
            elif t == 'strike':
                f.strike = _bool(v, True)
            elif t == 'color':
                f.color = self.color(c)
            elif t == 'vertAlign':
                f.vert_align = v if v in ('superscript', 'subscript') else None
        return f

    def num_fmt_code(self, num_fmt_id):
        from . import numfmt
        if num_fmt_id in self.num_fmts:
            return self.num_fmts[num_fmt_id]
        return numfmt.BUILTIN_FORMATS.get(num_fmt_id, 'General')


def _parse_theme(xml_bytes):
    colors = list(DEFAULT_THEME)
    if not xml_bytes:
        return colors
    root = _strip_ns(ET.fromstring(xml_bytes))
    scheme = None
    for el in root.iter('clrScheme'):
        scheme = el
        break
    if scheme is None:
        return colors
    order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3',
             'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
    found = {}
    for el in scheme:
        for c in el:
            if c.tag == 'srgbClr':
                found[el.tag] = c.attrib.get('val', '000000').upper()
            elif c.tag == 'sysClr':
                found[el.tag] = c.attrib.get('lastClr', '000000').upper()
    for i, name in enumerate(order):
        if name in found:
            colors[i] = found[name]
    return colors


def _parse_styles(xml_bytes, theme):
    st = Styles(theme)
    if not xml_bytes:
        return st
    root = _strip_ns(ET.fromstring(xml_bytes))

    colors = _child(root, 'colors')
    idx = _child(colors, 'indexedColors')
    if idx is not None:
        custom = [c.attrib.get('rgb', 'FF000000')[-6:].upper() for c in _children(idx, 'rgbColor')]
        for i, c in enumerate(custom):
            if i < len(st.indexed):
                st.indexed[i] = c
            else:
                st.indexed.append(c)

    for nf in _children(_child(root, 'numFmts'), 'numFmt'):
        st.num_fmts[int(nf.attrib['numFmtId'])] = nf.attrib.get('formatCode', 'General')

    fonts = _children(_child(root, 'fonts'), 'font')
    if fonts:
        st.fonts = [st.font_from(f) for f in fonts]

    fills = []
    for f in _children(_child(root, 'fills'), 'fill'):
        pf = _child(f, 'patternFill')
        gf = _child(f, 'gradientFill')
        if gf is not None:
            stops = []
            for s in _children(gf, 'stop'):
                stops.append((float(s.attrib.get('position', 0)), st.color(_child(s, 'color'))))
            fills.append(Fill('gradient', stops=stops))
        elif pf is not None:
            pattern = pf.attrib.get('patternType')
            fg = st.color(_child(pf, 'fgColor'))
            bg = st.color(_child(pf, 'bgColor'))
            if pattern is None:
                # Excel treats a patternFill with colours but no type as solid
                pattern = 'solid' if (fg or bg) else 'none'
                if pattern == 'solid' and not fg:
                    fg = bg
            fills.append(Fill(pattern, fg, bg))
        else:
            fills.append(Fill())
    if fills:
        st.fills = fills

    borders = []
    for b in _children(_child(root, 'borders'), 'border'):
        bd = Border()
        bd.diag_up = _bool(b.attrib.get('diagonalUp'))
        bd.diag_down = _bool(b.attrib.get('diagonalDown'))
        for side in ('left', 'right', 'top', 'bottom', 'diagonal'):
            el = _child(b, side)
            if el is None:
                el = _child(b, {'left': 'start', 'right': 'end'}.get(side, side))
            if el is not None and el.attrib.get('style'):
                setattr(bd, side, BorderSide(el.attrib['style'], st.color(_child(el, 'color'))))
        borders.append(bd)
    if borders:
        st.borders = borders

    xfs = []
    for x in _children(_child(root, 'cellXfs'), 'xf'):
        xf = Xf()
        a = x.attrib
        xf.num_fmt_id = int(a.get('numFmtId', 0))
        xf.font_id = int(a.get('fontId', 0))
        xf.fill_id = int(a.get('fillId', 0))
        xf.border_id = int(a.get('borderId', 0))
        al = _child(x, 'alignment')
        if al is not None:
            aa = al.attrib
            xf.horizontal = aa.get('horizontal', 'general')
            xf.vertical = aa.get('vertical', 'bottom')
            xf.wrap = _bool(aa.get('wrapText'))
            xf.indent = int(aa.get('indent', 0) or 0)
            xf.rotation = int(aa.get('textRotation', 0) or 0)
            xf.shrink = _bool(aa.get('shrinkToFit'))
        xfs.append(xf)
    if xfs:
        st.xfs = xfs
    return st


# --------------------------------------------------------------------------
# Strings
# --------------------------------------------------------------------------
class Run(object):
    __slots__ = ('text', 'font')

    def __init__(self, text, font=None):
        self.text = text
        self.font = font  # Font or None (= cell font)


def _parse_si(si, styles):
    """Return (plain_text, runs or None). Run fonts are partial overlays."""
    runs = _children(si, 'r')
    if not runs:
        t = _child(si, 't')
        return _unescape(t.text if t is not None else u''), None
    out = []
    for r in runs:
        t = _child(r, 't')
        text = _unescape(t.text if t is not None and t.text else u'')
        rpr = _child(r, 'rPr')
        out.append(Run(text, rpr))  # resolved later against cell font
    return u''.join(r.text for r in out), out


# --------------------------------------------------------------------------
# Sheet model
# --------------------------------------------------------------------------
class Cell(object):
    __slots__ = ('row', 'col', 'value', 'type', 'style', 'runs', 'formula')

    def __init__(self, row, col):
        self.row = row
        self.col = col
        self.value = None
        self.type = 'n'
        self.style = 0
        self.runs = None
        self.formula = False


class Sheet(object):
    def __init__(self, workbook, name):
        self.wb = workbook
        self.name = name
        self.cells = {}
        self.col_widths = {}   # col -> width (chars)
        self.col_hidden = set()
        self.col_styles = {}
        self.row_heights = {}  # row -> points
        self.row_hidden = set()
        self.row_styles = {}
        self.default_row_height = None
        self.default_col_width = None
        self.base_col_width = 8
        self.merges = []       # (r1, c1, r2, c2)
        self.dimension = None
        self.show_gridlines = True
        self.print_area = None

    # styles ---------------------------------------------------------------
    def style_index(self, r, c):
        cell = self.cells.get((r, c))
        if cell is not None:
            return cell.style
        if r in self.row_styles:
            return self.row_styles[r]
        return self.col_styles.get(c, 0)

    def xf(self, r, c):
        xfs = self.wb.styles.xfs
        i = self.style_index(r, c)
        return xfs[i] if 0 <= i < len(xfs) else xfs[0]

    def used_range(self):
        rows = [k[0] for k in self.cells]
        cols = [k[1] for k in self.cells]
        for r1, c1, r2, c2 in self.merges:
            rows += [r1, r2]
            cols += [c1, c2]
        if self.dimension:
            try:
                r1, c1, r2, c2 = parse_range(self.dimension)
                rows += [r1, r2]
                cols += [c1, c2]
            except ValueError:
                pass
        if not rows:
            return 1, 1, 1, 1
        return min(rows), min(cols), max(rows), max(cols)

    def default_range(self):
        if self.print_area:
            try:
                return parse_range(self.print_area)
            except ValueError:
                pass
        return self.used_range()


class Workbook(object):
    def __init__(self, path=None, data=None):
        self.path = path
        if data is None:
            data = read_file_bytes(path)
        self._zip = zipfile.ZipFile(_byte_stream(data))
        self._names = set(self._zip.namelist())
        self.date1904 = False
        self.sheet_entries = []  # (name, part path, state)
        self.defined_print_areas = {}
        self._load_workbook()
        self.theme = _parse_theme(self._read('xl/theme/theme1.xml'))
        self.styles = _parse_styles(self._read('xl/styles.xml'), self.theme)
        self.shared = self._load_shared()
        self._sheets = {}

    def _read(self, name):
        name = name.lstrip('/')
        if name in self._names:
            return self._zip.read(name)
        # case-insensitive fallback
        low = name.lower()
        for n in self._names:
            if n.lower() == low:
                return self._zip.read(n)
        return None

    def _load_workbook(self):
        root = _strip_ns(ET.fromstring(self._read('xl/workbook.xml')))
        pr = _child(root, 'workbookPr')
        if pr is not None:
            self.date1904 = _bool(pr.attrib.get('date1904'))
        rels = {}
        rel_xml = self._read('xl/_rels/workbook.xml.rels')
        if rel_xml:
            rroot = _strip_ns(ET.fromstring(rel_xml))
            for r in rroot:
                tgt = r.attrib.get('Target', '')
                if tgt.startswith('/'):
                    tgt = tgt[1:]
                elif not tgt.startswith('xl/'):
                    tgt = 'xl/' + tgt
                rels[r.attrib.get('Id')] = tgt
        names = []
        for i, s in enumerate(_children(_child(root, 'sheets'), 'sheet')):
            rid = s.attrib.get('r:id')
            part = rels.get(rid, 'xl/worksheets/sheet%d.xml' % (i + 1))
            names.append(s.attrib.get('name'))
            self.sheet_entries.append((s.attrib.get('name'), part, s.attrib.get('state', 'visible')))
        for dn in _children(_child(root, 'definedNames'), 'definedName'):
            if dn.attrib.get('name') == '_xlnm.Print_Area' and dn.text:
                sid = dn.attrib.get('localSheetId')
                if sid is not None and int(sid) < len(names):
                    area = dn.text.split(',')[0]
                    self.defined_print_areas[names[int(sid)]] = area

    def _load_shared(self):
        data = self._read('xl/sharedStrings.xml')
        if not data:
            return []
        root = _strip_ns(ET.fromstring(data))
        return [_parse_si(si, self.styles) for si in _children(root, 'si')]

    @property
    def sheet_names(self):
        return [n for n, _, _ in self.sheet_entries]

    @property
    def visible_sheet_names(self):
        return [n for n, _, st in self.sheet_entries if st == 'visible']

    def sheet(self, name):
        if name in self._sheets:
            return self._sheets[name]
        entry = None
        for e in self.sheet_entries:
            if e[0] == name:
                entry = e
        if entry is None:
            raise KeyError('Worksheet not found: %s' % name)
        sh = self._load_sheet(name, entry[1])
        self._sheets[name] = sh
        return sh

    def _load_sheet(self, name, part):
        sh = Sheet(self, name)
        sh.print_area = self.defined_print_areas.get(name)
        root = _strip_ns(ET.fromstring(self._read(part)))
        dim = _child(root, 'dimension')
        if dim is not None:
            sh.dimension = dim.attrib.get('ref')
        for sv in _children(_child(root, 'sheetViews'), 'sheetView'):
            sh.show_gridlines = _bool(sv.attrib.get('showGridLines'), True)
            break
        fp = _child(root, 'sheetFormatPr')
        if fp is not None:
            a = fp.attrib
            if 'defaultRowHeight' in a:
                sh.default_row_height = float(a['defaultRowHeight'])
            if 'defaultColWidth' in a:
                sh.default_col_width = float(a['defaultColWidth'])
            if 'baseColWidth' in a:
                sh.base_col_width = int(a['baseColWidth'])
        for col in _children(_child(root, 'cols'), 'col'):
            a = col.attrib
            lo, hi = int(a['min']), int(a['max'])
            hi = min(hi, lo + 16384)
            for c in range(lo, hi + 1):
                if 'width' in a:
                    sh.col_widths[c] = float(a['width'])
                if _bool(a.get('hidden')):
                    sh.col_hidden.add(c)
                if 'style' in a:
                    sh.col_styles[c] = int(a['style'])
        for m in _children(_child(root, 'mergeCells'), 'mergeCell'):
            try:
                sh.merges.append(parse_range(m.attrib['ref']))
            except (KeyError, ValueError):
                pass

        shared = self.shared
        for row in _children(_child(root, 'sheetData'), 'row'):
            a = row.attrib
            rnum = int(a['r']) if 'r' in a else None
            if rnum is None:
                rnum = (max([k[0] for k in sh.cells]) + 1) if sh.cells else 1
            if 'ht' in a:
                sh.row_heights[rnum] = float(a['ht'])
            if _bool(a.get('hidden')):
                sh.row_hidden.add(rnum)
            if _bool(a.get('customFormat')) and 's' in a:
                sh.row_styles[rnum] = int(a['s'])
            next_col = 1
            for c in _children(row, 'c'):
                ca = c.attrib
                if 'r' in ca:
                    rr, cc = parse_ref(ca['r'])
                else:
                    rr, cc = rnum, next_col
                next_col = cc + 1
                cell = Cell(rr, cc)
                cell.style = int(ca.get('s', 0))
                t = ca.get('t', 'n')
                v = _child(c, 'v')
                f = _child(c, 'f')
                cell.formula = f is not None
                vt = v.text if v is not None else None
                if t == 's':
                    if vt is not None:
                        txt, runs = shared[int(vt)]
                        cell.value, cell.runs, cell.type = txt, runs, 's'
                elif t == 'inlineStr':
                    txt, runs = _parse_si(_child(c, 'is'), self.styles) if _child(c, 'is') is not None else (u'', None)
                    cell.value, cell.runs, cell.type = txt, runs, 's'
                elif t == 'str':
                    cell.value, cell.type = _unescape(vt or u''), 's'
                elif t == 'b':
                    cell.value, cell.type = (vt == '1') if vt is not None else None, 'b'
                elif t == 'e':
                    cell.value, cell.type = vt, 'e'
                elif t == 'd':
                    cell.value, cell.type = vt, 'd'
                else:
                    if vt is not None:
                        try:
                            cell.value = float(vt)
                        except ValueError:
                            cell.value = vt
                            t = 's'
                    cell.type = 'n' if t == 'n' else t
                sh.cells[(rr, cc)] = cell
        return sh

    def missing_cached_values(self, sheet):
        """Formula cells saved without a cached result (non-Excel writers)."""
        return [c for c in sheet.cells.values() if c.formula and c.value is None]
