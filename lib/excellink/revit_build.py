# -*- coding: utf-8 -*-
"""Create Revit elements from an excellink Layout.

Works under IronPython 2.7 and CPython 3 (pythonnet).
All layout units are paper points; model length = pt / 864 * view.Scale.
"""
from __future__ import division

import math
import re

import clr
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Collections.Generic import List
from System.Windows import FontStyles, FontWeights, FontStretches
from System.Windows.Media import Typeface, FontFamily, Fonts

from Autodesk.Revit import DB

PT_FT = 1.0 / 864.0  # 1 point in feet (paper)
MIN_PATTERN_FT = 1.0 / 32.0 / 12.0  # Revit's shortest line pattern segment

DEFAULT_WEIGHTS = {'hair': 1, 'thin': 1, 'medium': 3, 'thick': 5, 'grid': 1}

# Excel dash patterns (in px) -> segments; scaled so that the shortest
# segment respects Revit's 1/32" minimum.
_PATTERNS = {
    'dotted': ('XL Dotted', [('dot', 0), ('space', 1)]),
    'hair': (None, None),
    'dashed': ('XL Dashed', [('dash', 2), ('space', 1)]),
    'mediumDashed': ('XL Medium Dashed', [('dash', 3), ('space', 1)]),
    'dashDot': ('XL Dash Dot', [('dash', 3), ('space', 1), ('dot', 0), ('space', 1)]),
    'mediumDashDot': ('XL Dash Dot', [('dash', 3), ('space', 1), ('dot', 0), ('space', 1)]),
    'slantDashDot': ('XL Dash Dot', [('dash', 3), ('space', 1), ('dot', 0), ('space', 1)]),
    'dashDotDot': ('XL Dash Dot Dot', [('dash', 3), ('space', 1), ('dot', 0), ('space', 1), ('dot', 0), ('space', 1)]),
    'mediumDashDotDot': ('XL Dash Dot Dot', [('dash', 3), ('space', 1), ('dot', 0), ('space', 1), ('dot', 0), ('space', 1)]),
}
_WEIGHT_CLASS = {
    'hair': 'hair', 'dotted': 'thin', 'dashed': 'thin', 'dashDot': 'thin',
    'dashDotDot': 'thin', 'thin': 'thin', 'double': 'thin', 'grid': 'grid',
    'medium': 'medium', 'mediumDashed': 'medium', 'mediumDashDot': 'medium',
    'mediumDashDotDot': 'medium', 'slantDashDot': 'medium', 'thick': 'thick',
}

_BAD_NAME = re.compile(r'[\\:{}\[\]|;<>?`~]')


def safe_name(s, maxlen=120):
    s = _BAD_NAME.sub('-', s).strip()
    return s[:maxlen]


def to_list(net_type, items):
    lst = List[net_type]()
    for i in items:
        lst.Add(i)
    return lst


def display_hex(hex6):
    """Revit draws anything exactly matching the view background (white) in
    the inverse colour, so pure white text/lines/fills would show as black.
    Near-white 254,254,254 displays and prints as white."""
    hex6 = (hex6 or '000000')[-6:].upper()
    return 'FEFEFE' if hex6 == 'FFFFFF' else hex6


def rgb(hex6):
    hex6 = display_hex(hex6)
    return DB.Color(int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16))


def color_int(hex6):
    c = rgb(hex6)
    return int(c.Red) + int(c.Green) * 256 + int(c.Blue) * 65536


# --------------------------------------------------------------------------
# Font metrics (WPF)
# --------------------------------------------------------------------------
class FontMetrics(object):
    def __init__(self):
        self._installed = None
        self._cache = {}

    def installed(self, name):
        if self._installed is None:
            names = set()
            for ff in Fonts.SystemFontFamilies:
                try:
                    for v in ff.FamilyNames.Values:
                        names.add(v.lower())
                except Exception:
                    names.add(ff.Source.lower())
            self._installed = names
        return (name or '').lower() in self._installed

    def glyphs(self, name, bold=False, italic=False):
        key = (name, bool(bold), bool(italic))
        if key in self._cache:
            return self._cache[key]
        gt = None
        try:
            tf = Typeface(FontFamily(name),
                          FontStyles.Italic if italic else FontStyles.Normal,
                          FontWeights.Bold if bold else FontWeights.Normal,
                          FontStretches.Normal)
            res = tf.TryGetGlyphTypeface()
            # pythonnet returns (bool, out) ; IronPython returns the same tuple
            if isinstance(res, tuple):
                ok, gt = res
            else:
                ok, gt = res, None
            if not ok:
                gt = None
        except Exception:
            gt = None
        self._cache[key] = gt
        return gt

    def caps_ratio(self, name):
        gt = self.glyphs(name) if self.installed(name) else None
        if gt is None:
            return 0.716
        try:
            c = float(gt.CapsHeight)
            return c if 0.3 < c < 1.0 else 0.716
        except Exception:
            return 0.716

    def measure(self, text, name, size_pt, bold=False, italic=False):
        """Advance width in points of a single line."""
        gt = self.glyphs(name, bold, italic) or self.glyphs('Arial', bold, italic)
        if gt is None:
            return len(text) * size_pt * 0.55
        cmap, adv = gt.CharacterToGlyphMap, gt.AdvanceWidths
        w = 0.0
        for ch in text:
            code = ord(ch)
            try:
                if cmap.ContainsKey(code):
                    w += adv[cmap[code]]
                else:
                    w += 0.55
            except Exception:
                w += 0.55
        return w * size_pt


METRICS = FontMetrics()


# --------------------------------------------------------------------------
# Style factories (cached per document)
# --------------------------------------------------------------------------
class Styles(object):
    def __init__(self, doc, weights=None):
        self.doc = doc
        self.weights = dict(DEFAULT_WEIGHTS)
        if weights:
            self.weights.update(weights)
        self._text = {}
        self._fill = {}
        self._line = {}
        self._pattern = {}
        self.warnings = []
        self._missing_fonts = set()
        self._solid = None
        self._invisible = None

    # ---- text ---------------------------------------------------------
    def text_type(self, font, size_pt=None):
        size_pt = size_pt or font.size
        color = font.color or '000000'
        name = safe_name('XL %s %gpt%s%s%s #%s' % (
            font.name, round(size_pt, 2), ' B' if font.bold else '', ' I' if font.italic else '',
            ' U' if font.underline else '', color))
        if name in self._text:
            return self._text[name]
        existing = None
        for t in DB.FilteredElementCollector(self.doc).OfClass(DB.TextNoteType):
            if DB.Element.Name.GetValue(t) == name:
                existing = t
                break
        if existing is not None:
            p = existing.get_Parameter(DB.BuiltInParameter.LINE_COLOR)
            if p is not None and not p.IsReadOnly and p.AsInteger() != color_int(color):
                p.Set(color_int(color))
        if existing is None:
            if not METRICS.installed(font.name) and font.name not in self._missing_fonts:
                self._missing_fonts.add(font.name)
                self.warnings.append('Font "%s" is not installed on this computer; '
                                     'Revit will substitute it.' % font.name)
            base = DB.FilteredElementCollector(self.doc).OfClass(DB.TextNoteType).FirstElement()
            existing = base.Duplicate(name)
            P = DB.BuiltInParameter
            caps = METRICS.caps_ratio(font.name)
            sets = [
                (P.TEXT_FONT, font.name),
                (P.TEXT_SIZE, size_pt * caps * PT_FT),
                (P.TEXT_STYLE_BOLD, 1 if font.bold else 0),
                (P.TEXT_STYLE_ITALIC, 1 if font.italic else 0),
                (P.TEXT_STYLE_UNDERLINE, 1 if font.underline else 0),
                (P.LINE_COLOR, color_int(color)),
                (P.TEXT_BACKGROUND, 1),          # transparent
                (P.TEXT_BOX_VISIBILITY, 0),
                (P.TEXT_WIDTH_SCALE, 1.0),
                (P.LINE_PEN, 1),
            ]
            for bip, val in sets:
                p = existing.get_Parameter(bip)
                if p is not None and not p.IsReadOnly:
                    try:
                        p.Set(val)
                    except Exception:
                        pass
        self._text[name] = existing
        return existing

    # ---- fills --------------------------------------------------------
    def solid_fill(self):
        if self._solid is None:
            for fp in DB.FilteredElementCollector(self.doc).OfClass(DB.FillPatternElement):
                pat = fp.GetFillPattern()
                if pat.IsSolidFill and pat.Target == DB.FillPatternTarget.Drafting:
                    self._solid = fp.Id
                    break
            if self._solid is None:
                for fp in DB.FilteredElementCollector(self.doc).OfClass(DB.FillPatternElement):
                    if fp.GetFillPattern().IsSolidFill:
                        self._solid = fp.Id
                        break
        return self._solid

    def invisible_lines(self):
        if self._invisible is None:
            for gs in DB.FilteredElementCollector(self.doc).OfClass(DB.GraphicsStyle):
                cat = gs.GraphicsStyleCategory
                if cat is not None and cat.Id == DB.ElementId(DB.BuiltInCategory.OST_InvisibleLines):
                    self._invisible = gs.Id
                    break
            if self._invisible is None:
                for gs in DB.FilteredElementCollector(self.doc).OfClass(DB.GraphicsStyle):
                    if gs.Name == '<Invisible lines>':
                        self._invisible = gs.Id
                        break
        return self._invisible

    def fill_type(self, hex6):
        name = 'XL Fill #%s' % hex6
        if name in self._fill:
            return self._fill[name]
        found = None
        for t in DB.FilteredElementCollector(self.doc).OfClass(DB.FilledRegionType):
            if DB.Element.Name.GetValue(t) == name:
                found = t
                break
        if found is None:
            base = DB.FilteredElementCollector(self.doc).OfClass(DB.FilledRegionType).FirstElement()
            found = base.Duplicate(name)
            found.ForegroundPatternId = self.solid_fill()
            found.ForegroundPatternColor = rgb(hex6)
            try:
                found.BackgroundPatternId = DB.ElementId.InvalidElementId
            except Exception:
                pass
            found.IsMasking = False
            try:
                found.LineWeight = 1
            except Exception:
                pass
        else:
            c = rgb(hex6)
            cur = found.ForegroundPatternColor
            if (cur.Red, cur.Green, cur.Blue) != (c.Red, c.Green, c.Blue):
                found.ForegroundPatternColor = c
        self._fill[name] = found
        return found

    # ---- lines --------------------------------------------------------
    def line_pattern(self, style):
        name, segs = _PATTERNS.get(style, (None, None))
        if not name:
            return None
        if name in self._pattern:
            return self._pattern[name]
        lpe = DB.LinePatternElement.GetLinePatternElementByName(self.doc, name)
        if lpe is None:
            lp = DB.LinePattern(name)
            unit = MIN_PATTERN_FT  # 1 Excel "unit" -> 1/32"
            segments = []
            for kind, n in segs:
                if kind == 'dot':
                    segments.append(DB.LinePatternSegment(DB.LinePatternSegmentType.Dot, 0.0))
                elif kind == 'dash':
                    segments.append(DB.LinePatternSegment(DB.LinePatternSegmentType.Dash, n * unit))
                else:
                    segments.append(DB.LinePatternSegment(DB.LinePatternSegmentType.Space, n * unit))
            lp.SetSegments(to_list(DB.LinePatternSegment, segments))
            lpe = DB.LinePatternElement.Create(self.doc, lp)
        self._pattern[name] = lpe.Id
        return lpe.Id

    def line_style(self, style, hex6):
        wclass = _WEIGHT_CLASS.get(style, 'thin')
        weight = int(self.weights.get(wclass, 1))
        pat_name = _PATTERNS.get(style, (None, None))[0]
        name = safe_name('XL %s%s #%s' % (
            'Pen%d' % weight, (' ' + pat_name[3:]) if pat_name else '', hex6 or '000000'))
        if name in self._line:
            return self._line[name]
        lines_cat = self.doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
        sub = None
        if lines_cat.SubCategories.Contains(name):
            sub = lines_cat.SubCategories.get_Item(name)
            c = rgb(hex6)
            cur = sub.LineColor
            if (cur.Red, cur.Green, cur.Blue) != (c.Red, c.Green, c.Blue):
                sub.LineColor = c
        else:
            sub = self.doc.Settings.Categories.NewSubcategory(lines_cat, name)
            sub.LineColor = rgb(hex6)
            sub.SetLineWeight(weight, DB.GraphicsStyleType.Projection)
            pid = self.line_pattern(style)
            if pid is not None:
                sub.SetLinePatternId(pid, DB.GraphicsStyleType.Projection)
        gs = sub.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
        self._line[name] = gs
        return gs


# --------------------------------------------------------------------------
# Builder
# --------------------------------------------------------------------------
class BuildResult(object):
    def __init__(self):
        self.element_ids = []
        self.group = None
        self.warnings = []
        self.counts = {'fills': 0, 'lines': 0, 'texts': 0}


def build(doc, view, layout, origin, styles=None, weights=None):
    """Draw *layout* in *view* with its top-left corner at *origin* (XYZ).
    Must be called inside an open Transaction. Returns BuildResult."""
    styles = styles or Styles(doc, weights)
    warn_start = len(styles.warnings)
    res = BuildResult()
    scale = float(view.Scale) if view.Scale else 1.0
    k = PT_FT * scale
    ox, oy, oz = origin.X, origin.Y, origin.Z

    # view's 2D basis (drafting/legend: X right, Y up; works for plan-like
    # views too since we use RightDirection/UpDirection)
    right = view.RightDirection
    up = view.UpDirection

    def P(x, y):
        return DB.XYZ(ox + right.X * x * k - up.X * y * k,
                      oy + right.Y * x * k - up.Y * y * k,
                      oz + right.Z * x * k - up.Z * y * k)

    min_len = doc.Application.ShortCurveTolerance

    # fills ---------------------------------------------------------------
    inv = styles.invisible_lines()
    for f in layout.fills:
        if f.w * k < min_len or f.h * k < min_len:
            continue
        pts = [P(f.x, f.y), P(f.x + f.w, f.y), P(f.x + f.w, f.y + f.h), P(f.x, f.y + f.h)]
        loop = DB.CurveLoop()
        for i in range(4):
            loop.Append(DB.Line.CreateBound(pts[i], pts[(i + 1) % 4]))
        frt = styles.fill_type(f.color)
        fr = DB.FilledRegion.Create(doc, frt.Id, view.Id, to_list(DB.CurveLoop, [loop]))
        if inv is not None:
            try:
                fr.SetLineStyleId(inv)
            except Exception:
                pass
        res.element_ids.append(fr.Id)
        res.counts['fills'] += 1

    # lines ---------------------------------------------------------------
    for s in layout.segments:
        a, b = P(s.x1, s.y1), P(s.x2, s.y2)
        if a.DistanceTo(b) < min_len:
            continue
        dc = doc.Create.NewDetailCurve(view, DB.Line.CreateBound(a, b))
        dc.LineStyle = styles.line_style(s.style, s.color)
        res.element_ids.append(dc.Id)
        res.counts['lines'] += 1

    # text ----------------------------------------------------------------
    H = {'left': DB.HorizontalTextAlignment.Left, 'center': DB.HorizontalTextAlignment.Center,
         'right': DB.HorizontalTextAlignment.Right}
    V = {'top': DB.VerticalTextAlignment.Top, 'middle': DB.VerticalTextAlignment.Middle,
         'bottom': DB.VerticalTextAlignment.Bottom}
    min_w = None
    for t in layout.texts:
        font = t.font
        size = font.size
        text = t.text.replace(u'\r\n', u'\n').replace(u'\n', u'\r')
        # shrink-to-fit / overflow measurements
        if getattr(t, 'shrink', False) and not t.wrap and t.rect:
            avail = t.rect[2] - 3.0
            w = max(METRICS.measure(line, font.name, size, font.bold, font.italic)
                    for line in text.split(u'\r'))
            if w > avail > 0:
                size = max(1.0, size * avail / w)
        ttype = styles.text_type(font, size)
        opts = DB.TextNoteOptions(ttype.Id)
        opts.HorizontalAlignment = H.get(t.halign, DB.HorizontalTextAlignment.Left)
        opts.VerticalAlignment = V.get(t.valign, DB.VerticalTextAlignment.Bottom)
        if t.rotation:
            opts.Rotation = math.radians(t.rotation)
            opts.KeepRotatedTextReadable = False
        pt = P(t.x, t.y)
        if t.wrap and t.width:
            if min_w is None:
                min_w = DB.TextNote.GetMinimumAllowedWidth(doc, ttype.Id)
            width_ft = max(t.width * PT_FT, min_w)
            # anchor for a wrapped note is the side of its box given by halign
            note = DB.TextNote.Create(doc, view.Id, pt, width_ft, text, opts)
        else:
            note = DB.TextNote.Create(doc, view.Id, pt, text, opts)
        if t.runs:
            try:
                ft = note.GetFormattedText()
                for start, length, b, i, u, va in t.runs:
                    rng = DB.TextRange(start, length)
                    if b != font.bold:
                        ft.SetBoldStatus(rng, bool(b))
                    if i != font.italic:
                        ft.SetItalicStatus(rng, bool(i))
                    if u != bool(font.underline):
                        ft.SetUnderlineStatus(rng, bool(u))
                    if va == 'superscript':
                        ft.SetSuperscriptStatus(rng, True)
                    elif va == 'subscript':
                        ft.SetSubscriptStatus(rng, True)
                note.SetFormattedText(ft)
            except Exception as ex:
                res.warnings.append('Rich text on %s not applied: %s' % (t.cell, ex))
        elif font.vert_align in ('superscript', 'subscript'):
            ft = note.GetFormattedText()
            rng = ft.AsTextRange() if hasattr(ft, 'AsTextRange') else DB.TextRange(0, len(text))
            if font.vert_align == 'superscript':
                ft.SetSuperscriptStatus(rng, True)
            else:
                ft.SetSubscriptStatus(rng, True)
            note.SetFormattedText(ft)
        res.element_ids.append(note.Id)
        res.counts['texts'] += 1

    res.warnings = list(layout.warnings) + styles.warnings[warn_start:] + res.warnings
    return res


def make_group(doc, element_ids, name):
    if not element_ids:
        return None
    grp = doc.Create.NewGroup(to_list(DB.ElementId, element_ids))
    gt = grp.GroupType
    existing = set(DB.Element.Name.GetValue(g) for g in
                   DB.FilteredElementCollector(doc).OfClass(DB.GroupType))
    base = safe_name(name)
    nm, n = base, 2
    while nm in existing:
        nm = '%s (%d)' % (base, n)
        n += 1
    try:
        gt.Name = nm
    except Exception:
        pass
    return grp
