# -*- coding: utf-8 -*-
"""Light / dark theming for Excel Link windows, following Revit's UI theme."""
import os

import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.IO import FileStream, FileMode, FileAccess
from System.Windows import Window, SystemColors
from System.Windows.Markup import XamlReader
from System.Windows.Media import SolidColorBrush, Color

HERE = os.path.dirname(__file__)

LIGHT = {
    'Bg': 'F0F0F0', 'Panel': 'E4E7EB', 'Input': 'FFFFFF', 'Border': 'B9C0CA',
    'Fg': '1E1E1E', 'SubFg': '5F6670', 'Button': 'FAFAFA', 'Hover': 'DCEAF7',
    'Select': 'C5DDF3', 'Accent': '0696D7', 'Thumb': 'B0B6BF', 'Warn': 'A0522D',
    'TabStrip': 'E1E1E1', 'Tab': 'EBEBEB', 'TabSel': 'FFFFFF', 'TabText': '333333',
    'TabAccent': '217346', 'Canvas': 'D9DCE0',
}
DARK = {
    'Bg': '3B4453', 'Panel': '323A47', 'Input': '2A303B', 'Border': '566173',
    'Fg': 'EEF1F5', 'SubFg': 'A9B3C2', 'Button': '465163', 'Hover': '4F5B6E',
    'Select': '2F5F86', 'Accent': '1F8FD1', 'Thumb': '6A7587', 'Warn': 'F0B467',
    'TabStrip': '2A303B', 'Tab': '343C4A', 'TabSel': '4A5568', 'TabText': 'D7DCE4',
    'TabAccent': '6CC28F', 'Canvas': '252A33',
}


def is_dark():
    """True when Revit's UI theme is Dark (Revit 2024+); False otherwise."""
    try:
        from Autodesk.Revit.UI import UIThemeManager, UITheme
        return UIThemeManager.CurrentTheme == UITheme.Dark
    except Exception:
        return False


def palette():
    return DARK if is_dark() else LIGHT


def brush(hex6):
    hex6 = (hex6 or '000000')[-6:]
    b = SolidColorBrush(Color.FromRgb(int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)))
    b.Freeze()
    return b


def _load_dictionary():
    fs = FileStream(os.path.join(HERE, 'Theme.xaml'), FileMode.Open, FileAccess.Read)
    try:
        return XamlReader.Load(fs)
    finally:
        fs.Close()


def fill(res, pal=None):
    """Put the palette brushes and themed control styles into a ResourceDictionary."""
    pal = pal or palette()
    for k, v in pal.items():
        res['XL.' + k] = brush(v)
    for key, name in ((SystemColors.ControlBrushKey, 'Panel'), (SystemColors.WindowBrushKey, 'Input'),
                      (SystemColors.ControlTextBrushKey, 'Fg'), (SystemColors.WindowTextBrushKey, 'Fg'),
                      (SystemColors.HighlightBrushKey, 'Select'), (SystemColors.HighlightTextBrushKey, 'Fg'),
                      (SystemColors.MenuBrushKey, 'Input'), (SystemColors.MenuTextBrushKey, 'Fg')):
        res[key] = brush(pal[name])
    try:
        res.MergedDictionaries.Add(_load_dictionary())
    except Exception:
        pass
    return pal


def apply(window):
    """Theme *window*: merge the control styles, inject the palette and set its
    colours. Returns the palette dict (hex strings) for code that paints itself."""
    pal = fill(window.Resources)
    window.Background = brush(pal['Bg'])
    window.Foreground = brush(pal['Fg'])
    return pal
