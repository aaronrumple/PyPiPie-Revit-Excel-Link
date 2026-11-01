# -*- coding: utf-8 -*-
"""Excel number-format renderer (en-US).

format_value(value, cell_type, code, date1904) -> (text, color_name_or_None)

Covers General, fixed/thousands, percent, scientific, fractions, currency
and accounting (incl. [$..-409] locale tags), literal text, _ padding,
* fill, [Red]/[Color n] colours, [>=100] conditions, 1–4 section formats,
dates, times, elapsed [h]:mm:ss and AM/PM.
"""
from __future__ import division

import datetime
import math
import re
from decimal import Decimal, ROUND_HALF_UP

try:
    text_type = unicode  # noqa: F821
except NameError:
    text_type = str

BUILTIN_FORMATS = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
    5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
    7: '"$"#,##0.00_);\\("$"#,##0.00\\)', 8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
    22: 'm/d/yyyy h:mm',
    37: '#,##0_);(#,##0)', 38: '#,##0_);[Red](#,##0)',
    39: '#,##0.00_);(#,##0.00)', 40: '#,##0.00_);[Red](#,##0.00)',
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
}

COLOR_NAMES = {
    'black': '000000', 'blue': '0000FF', 'cyan': '00FFFF', 'green': '00FF00',
    'magenta': 'FF00FF', 'red': 'FF0000', 'white': 'FFFFFF', 'yellow': 'FFFF00',
}

MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
          'August', 'September', 'October', 'November', 'December']
# Marks where a '*x' repeat-fill sits; layout splits text there so the left
# part hugs the left cell edge and the right part the right edge (accounting).
FILL_MARK = u'\x01'

DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']


# --------------------------------------------------------------------------
def round_half_up(x, nd):
    q = Decimal(1).scaleb(-nd) if nd > 0 else Decimal(1)
    d = Decimal(repr(float(x))).quantize(q, rounding=ROUND_HALF_UP)
    return d


def _split_sections(code):
    sections, cur, i, n = [], [], 0, len(code)
    while i < n:
        ch = code[i]
        if ch == '"':
            j = code.find('"', i + 1)
            j = n - 1 if j < 0 else j
            cur.append(code[i:j + 1])
            i = j + 1
            continue
        if ch == '\\' and i + 1 < n:
            cur.append(code[i:i + 2])
            i += 2
            continue
        if ch == '[':
            j = code.find(']', i)
            j = n - 1 if j < 0 else j
            cur.append(code[i:j + 1])
            i = j + 1
            continue
        if ch == ';':
            sections.append(''.join(cur))
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    sections.append(''.join(cur))
    return sections


_COND = re.compile(r'^\[(<=|>=|<>|<|>|=)(-?[\d.]+)\]$')


class Section(object):
    def __init__(self, raw):
        self.raw = raw
        self.color = None
        self.cond = None
        self.tokens = []
        self._parse(raw)

    def _parse(self, raw):
        toks = []
        i, n = 0, len(raw)
        while i < n:
            ch = raw[i]
            if ch == '"':
                j = raw.find('"', i + 1)
                j = n if j < 0 else j
                toks.append(('lit', raw[i + 1:j]))
                i = j + 1
            elif ch == '\\' and i + 1 < n:
                toks.append(('lit', raw[i + 1]))
                i += 2
            elif ch == '_' and i + 1 < n:
                toks.append(('lit', ' '))
                i += 2
            elif ch == '*' and i + 1 < n:
                toks.append(('fill', raw[i + 1]))
                i += 2
            elif ch == '[':
                j = raw.find(']', i)
                j = n - 1 if j < 0 else j
                body = raw[i + 1:j]
                low = body.lower()
                m = _COND.match('[' + body + ']')
                if low in COLOR_NAMES:
                    self.color = COLOR_NAMES[low]
                elif low.startswith('color'):
                    try:
                        from .xlsx_reader import INDEXED_COLORS
                        k = int(low[5:]) - 1
                        if 0 <= k < 56:
                            self.color = INDEXED_COLORS[k + 8]
                    except ValueError:
                        pass
                elif m:
                    self.cond = (m.group(1), float(m.group(2)))
                elif body.startswith('$'):
                    sym = body[1:].split('-')[0]
                    if sym:
                        toks.append(('lit', sym))
                elif low in ('h', 'hh', 'm', 'mm', 's', 'ss'):
                    toks.append(('elapsed', low))
                i = j + 1
            else:
                toks.append(('ch', ch))
                i += 1
        self.tokens = toks

    def plain(self):
        return ''.join(v for k, v in self.tokens if k == 'ch')

    def is_text(self):
        return '@' in self.plain()

    def is_date(self):
        if any(k == 'elapsed' for k, _ in self.tokens):
            return True
        p = self.plain().lower()
        if 'general' in p:
            return False
        return bool(re.search(r'[ymdhs]', p))

    def matches(self, v):
        if not self.cond:
            return None
        op, x = self.cond
        return {'<': v < x, '>': v > x, '=': v == x, '<=': v <= x,
                '>=': v >= x, '<>': v != x}[op]


# --------------------------------------------------------------------------
def format_general(v):
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if v == 0:
        return '0'
    if v == int(v) and abs(v) < 1e11:
        return '%d' % int(v)
    av = abs(v)
    if av >= 1e11 or av < 1e-9:
        s = '%.5E' % v
        mant, exp = s.split('E')
        mant = mant.rstrip('0').rstrip('.')
        e = int(exp)
        return '%sE%s%02d' % (mant, '+' if e >= 0 else '-', abs(e))
    digits_int = len(str(int(av))) if av >= 1 else 1
    dec = max(0, 10 - digits_int)
    s = '%s' % round_half_up(av, dec)
    if '.' in s:
        s = s.rstrip('0').rstrip('.')
    return ('-' if v < 0 else '') + s


def _excel_date(serial, date1904):
    if date1904:
        base = datetime.datetime(1904, 1, 1)
        return base + datetime.timedelta(days=serial)
    if serial < 60:
        base = datetime.datetime(1899, 12, 31)
    else:
        base = datetime.datetime(1899, 12, 30)
    return base + datetime.timedelta(days=serial)


def _format_date(sec, v, date1904):
    toks = sec.tokens
    has_ampm = 'am/pm' in sec.plain().lower() or 'a/p' in sec.plain().lower()
    # determine sub-second precision
    plain = sec.plain()
    msec_digits = 0
    m = re.search(r's+\.(0+)', plain.lower())
    if m:
        msec_digits = len(m.group(1))
    frac_day = round(v * 86400.0, msec_digits) / 86400.0
    try:
        dt = _excel_date(frac_day, date1904)
    except (OverflowError, ValueError):
        return '#' * 8
    # round to seconds (Excel rounds)
    us = dt.microsecond
    if msec_digits == 0 and us >= 500000:
        dt = dt + datetime.timedelta(microseconds=1000000 - us)
        us = 0

    # tokenize date parts from char tokens
    parts = []
    i = 0
    chars = toks
    while i < len(chars):
        k, val = chars[i]
        if k != 'ch':
            parts.append((k, val))
            i += 1
            continue
        low = val.lower()
        rest = ''.join(v2 if k2 == 'ch' else '\x00' for k2, v2 in chars[i:])
        rl = rest.lower()
        if rl.startswith('am/pm'):
            parts.append(('ampm', rest[:5]))
            i += 5
            continue
        if rl.startswith('a/p'):
            parts.append(('ap', rest[:3]))
            i += 3
            continue
        if low in 'ymdhs':
            j = i
            while j < len(chars) and chars[j][0] == 'ch' and chars[j][1].lower() == low:
                j += 1
            parts.append((low, j - i))
            i = j
            continue
        if low == '0' and parts and parts[-1] == ('ch', '.'):
            j = i
            while j < len(chars) and chars[j] == ('ch', '0'):
                j += 1
            parts.append(('frac', j - i))
            i = j
            continue
        parts.append(('ch', val))
        i += 1

    # resolve m as minutes when adjacent to h or s
    idx_dt = [n for n, p in enumerate(parts) if p[0] in ('y', 'm', 'd', 'h', 's', 'elapsed')]
    for pos, n in enumerate(idx_dt):
        if parts[n][0] != 'm':
            continue
        prev = parts[idx_dt[pos - 1]] if pos > 0 else None
        nxt = parts[idx_dt[pos + 1]] if pos + 1 < len(idx_dt) else None
        if (prev and prev[0] in ('h', 'elapsed')) or (nxt and nxt[0] == 's'):
            parts[n] = ('min', parts[n][1])

    out = []
    total_seconds = v * 86400.0
    for k, val in parts:
        if k == 'y':
            out.append('%02d' % (dt.year % 100) if val <= 2 else '%04d' % dt.year)
        elif k == 'm':
            if val == 1:
                out.append('%d' % dt.month)
            elif val == 2:
                out.append('%02d' % dt.month)
            elif val == 3:
                out.append(MONTHS[dt.month - 1][:3])
            elif val == 5:
                out.append(MONTHS[dt.month - 1][0])
            else:
                out.append(MONTHS[dt.month - 1])
        elif k == 'd':
            if val == 1:
                out.append('%d' % dt.day)
            elif val == 2:
                out.append('%02d' % dt.day)
            elif val == 3:
                out.append(DAYS[dt.weekday()][:3])
            else:
                out.append(DAYS[dt.weekday()])
        elif k == 'h':
            h = dt.hour
            if has_ampm:
                h = h % 12 or 12
            out.append('%d' % h if val == 1 else '%02d' % h)
        elif k == 'min':
            out.append('%d' % dt.minute if val == 1 else '%02d' % dt.minute)
        elif k == 's':
            out.append('%d' % dt.second if val == 1 else '%02d' % dt.second)
        elif k == 'frac':
            out.append(('%0' + str(val) + 'd') % int(us / 10 ** (6 - val)))
        elif k == 'elapsed':
            unit = val[0]
            ts = int(round(total_seconds))
            num = {'h': ts // 3600, 'm': ts // 60, 's': ts}[unit]
            out.append(('%0' + str(len(val)) + 'd') % num)
        elif k == 'ampm':
            s = 'AM' if dt.hour < 12 else 'PM'
            out.append(s if val[0].isupper() else s.lower())
        elif k == 'ap':
            s = 'A' if dt.hour < 12 else 'P'
            out.append(s if val[0].isupper() else s.lower())
        elif k in ('ch', 'lit'):
            out.append(val)
    return ''.join(out)


def _group(int_str):
    out = []
    while len(int_str) > 3:
        out.insert(0, int_str[-3:])
        int_str = int_str[:-3]
    out.insert(0, int_str)
    return ','.join(out)


def _apply_mask(digits, mask, align_right=True):
    """Fill integer mask (#0?) right-to-left with digits string."""
    res = []
    di = len(digits) - 1
    for mc in reversed(mask):
        if di >= 0:
            res.append(digits[di])
            di -= 1
        elif mc == '0':
            res.append('0')
        elif mc == '?':
            res.append(' ')
    if di >= 0:
        res.append(digits[:di + 1])
    return ''.join(reversed(res))


def _format_number(sec, v, auto_minus):
    toks = sec.tokens
    plain = sec.plain()

    # scale by trailing commas before decimal/end
    num_chars = [n for n, (k, ch) in enumerate(toks) if k == 'ch' and ch in '#0?.,']
    scale = 0
    if num_chars:
        last = num_chars[-1]
        n = last
        while n >= 0 and toks[n] == ('ch', ','):
            scale += 1
            n -= 1
        if scale:
            for k in range(n + 1, last + 1):
                toks = toks[:k] + [('lit', '')] + toks[k + 1:]
    percent = plain.count('%')
    x = abs(v) * (100 ** percent) / (1000 ** scale)
    neg = v < 0 and auto_minus

    if 'E' in plain.upper() and re.search(r'[0#?]\.?[0#?]*E[+-]', plain, re.I):
        return _format_sci(sec, abs(v), neg)
    if '/' in plain and re.search(r'[#0?]\s*/\s*[#0?\d]', plain):
        return _format_fraction(sec, abs(v), neg)

    # identify integer & fraction masks
    seq = []
    for k, ch in toks:
        if k == 'ch' and ch in '#0?,.':
            seq.append(ch)
    s = ''.join(seq)
    if '.' in s:
        int_mask, frac_mask = s.split('.', 1)
    else:
        int_mask, frac_mask = s, ''
    thousands = ',' in int_mask
    int_mask = int_mask.replace(',', '')
    frac_mask = frac_mask.replace(',', '')
    nd = len(frac_mask)
    d = round_half_up(x, nd)
    ds = '%s' % d
    if '.' in ds:
        ip, fp = ds.split('.')
    else:
        ip, fp = ds, ''
    fp = (fp + '0' * nd)[:nd]
    if ip == '0' and int_mask and '0' not in int_mask:
        ip = ''
    int_str = _apply_mask(ip, int_mask)
    if thousands:
        core = int_str.lstrip(' ')
        pad = len(int_str) - len(core)
        int_str = ' ' * pad + _group(core) if core else int_str
    # fraction mask: trailing # dropped, ? -> space
    fr = list(fp)
    for i in range(nd - 1, -1, -1):
        if frac_mask[i] == '0':
            break
        if fr[i] == '0':
            fr[i] = '' if frac_mask[i] == '#' else ' '
        else:
            break
    frac_str = ''.join(fr)

    # Integer placeholders separated by literals (e.g. 0"'-"0\" for feet-inches,
    # 000-00-0000): Excel spreads the digits right-to-left across them.
    int_pos = []
    for n, (k, ch) in enumerate(toks):
        if k == 'ch' and ch == '.':
            break
        if k == 'ch' and ch in '#0?':
            int_pos.append(n)
    split = (not thousands and len(int_pos) > 1 and
             any(toks[n][0] in ('lit', 'ch') and toks[n][1] not in '#0?,'
                 for n in range(int_pos[0] + 1, int_pos[-1])))
    if split:
        digits = list(ip)
        assigned = {}
        for n in reversed(int_pos):
            ch = toks[n][1]
            if digits:
                assigned[n] = digits.pop()
            else:
                assigned[n] = {'0': '0', '?': ' '}.get(ch, '')
        assigned[int_pos[0]] = ''.join(digits) + assigned[int_pos[0]]
        out = []
        in_frac = False
        for n, (k, ch) in enumerate(toks):
            if n in assigned:
                out.append(assigned[n])
            elif k == 'ch' and ch == '.':
                out.append('.')
                out.append(frac_str)
                in_frac = True
            elif k == 'ch' and ch in '#0?,':
                continue
            elif k in ('lit', 'ch'):
                out.append(ch)
            elif k == 'fill':
                out.append(FILL_MARK)
        text = ''.join(out)
        return ('-' + text) if (neg and float(d) != 0) else text

    # rebuild output walking tokens; emit number at first digit token
    out = []
    emitted_int = False
    in_frac = False
    for k, ch in toks:
        if k == 'ch' and ch in '#0?,':
            if not in_frac and not emitted_int:
                out.append(int_str)
                emitted_int = True
            continue
        if k == 'ch' and ch == '.':
            if not emitted_int:
                out.append(int_str)
                emitted_int = True
            out.append('.')
            out.append(frac_str)
            in_frac = True
            continue
        if k in ('lit', 'ch'):
            out.append(ch)
        elif k == 'fill':
            out.append(FILL_MARK)
    if not emitted_int and (int_mask or not frac_mask) and seq:
        out.append(int_str)
    text = ''.join(out)
    if neg and float(d) != 0:
        text = '-' + text
    return text


def _format_sci(sec, x, neg):
    plain = sec.plain()
    m = re.search(r'([#0?]*)(?:\.([#0?]*))?E([+-])([#0?]+)', plain, re.I)
    int_mask, frac_mask, sign, exp_mask = m.group(1), m.group(2) or '', m.group(3), m.group(4)
    nint = max(1, len(int_mask))
    if x == 0:
        exp = 0
        mant = 0.0
    else:
        exp = int(math.floor(math.log10(x)))
        if len(int_mask) > 1:
            # engineering style: exponent multiple of int digits
            exp = int(math.floor(exp / len(int_mask))) * len(int_mask)
        mant = x / (10 ** exp)
    nd = len(frac_mask)
    d = round_half_up(mant, nd)
    if float(d) >= 10 ** nint and x != 0:
        exp += nint if len(int_mask) > 1 else 1
        d = round_half_up(x / (10 ** exp), nd)
    ms = '%s' % d
    if '.' in ms:
        ip, fp = ms.split('.')
        fp = fp.rstrip('0') if frac_mask and '0' not in frac_mask else fp
    else:
        ip, fp = ms, ''
    es = ('%0' + str(len(exp_mask)) + 'd') % abs(exp)
    esign = '-' if exp < 0 else ('+' if sign == '+' else '')
    body = ip + ('.' + fp if frac_mask else '') + 'E' + esign + es
    prefix = plain[:m.start()]
    suffix = plain[m.end():]
    return ('-' if neg else '') + prefix + body + suffix


def _format_fraction(sec, x, neg):
    plain = sec.plain()
    m = re.search(r'(?:([#0?]+)\s+)?([#0?]+)\s*/\s*([#0?]+|\d+)', plain)
    whole_mask, num_mask, den_mask = m.group(1), m.group(2), m.group(3)
    whole = int(math.floor(x)) if whole_mask else 0
    frac = x - whole
    if den_mask.isdigit():
        den = int(den_mask)
        num = int(round(frac * den))
    else:
        maxd = 10 ** len(den_mask) - 1
        best = (abs(frac), 0, 1)
        for dd in range(1, maxd + 1):
            nn = int(round(frac * dd))
            err = abs(frac - nn / dd)
            if err < best[0] - 1e-12:
                best = (err, nn, dd)
        _, num, den = best
    if num == den:
        whole += 1
        num = 0
    parts = []
    if whole_mask:
        parts.append(str(whole) if whole or num == 0 else '')
    if num:
        w = len(den_mask) if not den_mask.isdigit() else len(den_mask)
        parts.append(('%' + str(len(num_mask)) + 'd') % num + '/' + ('%-' + str(w) + 'd') % den)
    elif not whole_mask:
        parts.append('0')
    text = ' '.join(p for p in parts if p) or '0'
    return ('-' if neg else '') + text


# --------------------------------------------------------------------------
def format_value(value, cell_type, code, date1904=False):
    """Return (display_text, colour hex or None)."""
    if value is None:
        return u'', None
    if cell_type == 'e':
        return text_type(value), None
    if cell_type == 'b':
        return ('TRUE' if value else 'FALSE'), None
    code = code or 'General'
    sections = [Section(s) for s in _split_sections(code)]
    if cell_type == 's' or not isinstance(value, (int, float)):
        txt = value if isinstance(value, text_type) else text_type(value)
        sec = None
        for s in sections:
            if s.is_text():
                sec = s
        if sec is None and len(sections) >= 4:
            sec = sections[3]
        if sec is None:
            return txt, None
        out = []
        for k, ch in sec.tokens:
            if k == 'ch' and ch == '@':
                out.append(txt)
            elif k in ('lit', 'ch'):
                out.append(ch)
            elif k == 'fill':
                out.append(FILL_MARK)
        return u''.join(out), sec.color

    v = float(value)
    # pick section
    sec, auto_minus = sections[0], True
    if any(s.cond for s in sections[:2]):
        chosen = None
        for s in sections[:3]:
            r = s.matches(v)
            if r:
                chosen = s
                break
            if r is None and chosen is None and s is not sections[0] and not s.cond:
                chosen = s
                break
        sec = chosen or sections[-1]
        auto_minus = not (sec is not sections[0] and v < 0)
    else:
        num_secs = [s for s in sections if not (s.is_text() and len(sections) > 1 and s is sections[-1] and len(sections) == 4)]
        if v > 0 or len(num_secs) == 1:
            sec = num_secs[0]
        elif v < 0:
            sec, auto_minus = num_secs[1], False
        else:
            sec = num_secs[2] if len(num_secs) > 2 else num_secs[0]
    if len(sections) == 1:
        auto_minus = True

    if sec.plain().strip().lower() == 'general' or sec.raw.strip().lower() == 'general' or \
            ('general' in sec.plain().lower()):
        txt = format_general(v if auto_minus else abs(v))
        out = []
        done = False
        for k, ch in sec.tokens:
            if k in ('lit',):
                out.append(ch)
            elif k == 'ch' and not done and ch.lower() == 'g':
                out.append(txt)
                done = True
        if not done:
            out.append(txt)
        return ''.join(out), sec.color
    if sec.is_date():
        if v < 0:
            return '#' * 8, sec.color
        return _format_date(sec, v, date1904), sec.color
    has_digit = any(k == 'ch' and ch in '#0?' for k, ch in sec.tokens)
    if not has_digit:
        out = []
        for k, ch in sec.tokens:
            if k in ('lit', 'ch'):
                out.append(ch)
            elif k == 'fill':
                out.append(FILL_MARK)
        txt = ''.join(out)
        return (('-' + txt) if (v < 0 and auto_minus and txt) else txt), sec.color
    return _format_number(sec, v, auto_minus), sec.color
