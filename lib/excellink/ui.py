# -*- coding: utf-8 -*-
"""WPF windows and button entry points for Excel Link (IronPython / pyRevit)."""
from __future__ import division

import os
import traceback

import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Windows import (FontWeights, FontStyles, TextWrapping, TextAlignment,
                            TextDecorations, Size, Point, Thickness, BaselineAlignment,
                            Visibility, VerticalAlignment, CornerRadius)
from System.Windows.Controls import Border, StackPanel, CheckBox, Orientation
from System.Windows.Input import Cursors
from System.Windows.Controls import Canvas, TextBlock
from System.Windows.Documents import Run as WpfRun
from System.Windows.Shapes import Rectangle, Line
from System.Windows.Media import (SolidColorBrush, Color, FontFamily, RotateTransform,
                                  ScaleTransform, DoubleCollection)

from pyrevit import forms, script, revit, DB

from . import link as xlink
from .xlsx_reader import Workbook, format_range, parse_range
from .layout import build_layout
from . import revit_build
from . import theme

HERE = os.path.dirname(__file__)
HELP_FILE = os.path.normpath(os.path.join(HERE, '..', '..', 'help', 'ExcelLink_Help.html'))
PX = 96.0 / 72.0
CFG_SECTION = 'ExcelLink'

ALLOWED_VIEWS = (DB.ViewType.DraftingView, DB.ViewType.Legend, DB.ViewType.Detail,
                 DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.EngineeringPlan,
                 DB.ViewType.AreaPlan, DB.ViewType.Section, DB.ViewType.Elevation)


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------
def brand(window, help_section):
    """PyPiPie branding shared with Rover: app icon, drawn title bar with
    help / maximize / close, and the 'Crafted by' credit in the status bar."""
    from System import Uri
    from System.Windows import WindowState
    from System.Windows.Media.Imaging import BitmapImage
    try:
        img = BitmapImage(Uri(os.path.join(HERE, 'excellink.png')))
        window.Icon = img
        window.TitleIcon.Source = img
    except Exception:
        pass
    try:
        from . import __version__
        window.VersionLabel.Text = 'Excel Link %s' % __version__
    except Exception:
        pass

    def on_state(sender, args):
        maxed = window.WindowState == WindowState.Maximized
        window.RootGrid.Margin = Thickness(7) if maxed else Thickness(0)  # maximized chrome overhangs the screen
        window.CapMax.Content = u'\ue923' if maxed else u'\ue922'
        window.CapMax.ToolTip = 'Restore' if maxed else 'Maximize'

    def toggle_max(sender, args):
        window.WindowState = WindowState.Normal if window.WindowState == WindowState.Maximized \
            else WindowState.Maximized

    def open_link(sender, args):
        try:
            from System.Diagnostics import Process
            Process.Start(args.Uri.AbsoluteUri)
        except Exception:
            pass
        args.Handled = True

    window.CapHelp.Click += lambda s_, a_: open_help(help_section)
    window.CapMax.Click += toggle_max
    window.CapClose.Click += lambda s_, a_: window.Close()
    window.StateChanged += on_state
    window.CreditLink.RequestNavigate += open_link


def open_help(section=None):
    """Open the help page in the default browser at *section*, themed like Revit.

    Windows drops the #fragment when it opens a local file, so a tiny redirect
    page in %TEMP% carries the section and theme instead."""
    import tempfile
    try:
        from urllib import pathname2url  # IronPython 2
    except ImportError:
        from urllib.request import pathname2url
    if not os.path.isfile(HELP_FILE):
        forms.alert('Help file not found:\n%s' % HELP_FILE)
        return
    url = 'file:' + pathname2url(HELP_FILE) + '?theme=' + ('dark' if theme.is_dark() else 'light')
    if section:
        url += '#' + section
    redirect = os.path.join(tempfile.gettempdir(), 'excellink_help.html')
    try:
        with open(redirect, 'w') as f:
            f.write('<!DOCTYPE html><meta charset="utf-8"><meta http-equiv="refresh" content="0;url=%s">'
                    '<title>Excel Link Help</title><a href="%s">Excel Link Help</a>' % (url, url))
        os.startfile(redirect)
    except Exception:
        os.startfile(HELP_FILE)

def cfg():
    return script.get_config(CFG_SECTION)


def cfg_get(name, default):
    try:
        return cfg().get_option(name, default)
    except Exception:
        return default


def cfg_set(**kw):
    try:
        c = cfg()
        for k, v in kw.items():
            c.set_option(k, v)
        script.save_config()
    except Exception:
        pass


def brush(hex6):
    hex6 = (hex6 or '000000')[-6:]
    return SolidColorBrush(Color.FromRgb(int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)))


def view_ok(view):
    return view is not None and not view.IsTemplate and view.ViewType in ALLOWED_VIEWS


def scale_text(view):
    try:
        p = view.get_Parameter(DB.BuiltInParameter.VIEW_SCALE_PULLDOWN_IMPERIAL) or \
            view.get_Parameter(DB.BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC)
        s = p.AsValueString() if p else None
        return s or '1 : %d' % view.Scale
    except Exception:
        return '1 : %d' % view.Scale


# --------------------------------------------------------------------------
# Preview drawing
# --------------------------------------------------------------------------
_DASH = {'dotted': [1, 1], 'hair': [1, 1], 'dashed': [3, 1], 'mediumDashed': [3, 1],
         'dashDot': [3, 1, 1, 1], 'mediumDashDot': [3, 1, 1, 1], 'slantDashDot': [3, 1, 1, 1],
         'dashDotDot': [3, 1, 1, 1, 1, 1], 'mediumDashDotDot': [3, 1, 1, 1, 1, 1]}
_THICK = {'medium': 2, 'mediumDashed': 2, 'mediumDashDot': 2, 'mediumDashDotDot': 2,
          'slantDashDot': 2, 'thick': 3}


def draw_preview(canvas, lay):
    canvas.Children.Clear()
    canvas.Width = max(lay.width * PX, 10)
    canvas.Height = max(lay.height * PX, 10)
    for f in lay.fills:
        r = Rectangle()
        r.Width, r.Height = f.w * PX, f.h * PX
        r.Fill = brush(f.color)
        Canvas.SetLeft(r, f.x * PX)
        Canvas.SetTop(r, f.y * PX)
        canvas.Children.Add(r)
    for s in lay.segments:
        ln = Line()
        ln.X1, ln.Y1, ln.X2, ln.Y2 = s.x1 * PX, s.y1 * PX, s.x2 * PX, s.y2 * PX
        ln.Stroke = brush(s.color if s.style != 'grid' else 'D0D0D0')
        ln.StrokeThickness = _THICK.get(s.style, 1)
        if s.style in _DASH:
            dc = DoubleCollection()
            for d in _DASH[s.style]:
                dc.Add(float(d))
            ln.StrokeDashArray = dc
        canvas.Children.Add(ln)
    for t in lay.texts:
        tb = TextBlock()
        f = t.font
        tb.FontFamily = FontFamily(f.name)
        tb.FontSize = f.size * PX
        tb.Foreground = brush(f.color)
        if t.runs:
            for start, length, b, i, u, va in t.runs:
                run = WpfRun(t.text[start:start + length])
                run.FontWeight = FontWeights.Bold if b else FontWeights.Normal
                run.FontStyle = FontStyles.Italic if i else FontStyles.Normal
                if u:
                    run.TextDecorations = TextDecorations.Underline
                if va == 'superscript':
                    run.BaselineAlignment = BaselineAlignment.Superscript
                elif va == 'subscript':
                    run.BaselineAlignment = BaselineAlignment.Subscript
                tb.Inlines.Add(run)
        else:
            tb.Text = t.text
            tb.FontWeight = FontWeights.Bold if f.bold else FontWeights.Normal
            tb.FontStyle = FontStyles.Italic if f.italic else FontStyles.Normal
            if f.underline:
                tb.TextDecorations = TextDecorations.Underline
        if t.wrap and t.width:
            tb.Width = t.width * PX
            tb.TextWrapping = TextWrapping.Wrap
            tb.TextAlignment = {'left': TextAlignment.Left, 'center': TextAlignment.Center,
                                'right': TextAlignment.Right}[t.halign]
        tb.Measure(Size(float('inf'), float('inf')))
        w, h = tb.DesiredSize.Width, tb.DesiredSize.Height
        x, y = t.x * PX, t.y * PX
        if t.rotation:
            tb.RenderTransformOrigin = Point(0.5, 0.5)
            tb.RenderTransform = RotateTransform(-t.rotation)
            left, top = x - w / 2.0, y - h / 2.0
        else:
            left = x if t.halign == 'left' else (x - w / 2.0 if t.halign == 'center' else x - w)
            top = y if t.valign == 'top' else (y - h / 2.0 if t.valign == 'middle' else y - h)
        Canvas.SetLeft(tb, left)
        Canvas.SetTop(tb, top)
        canvas.Children.Add(tb)


# --------------------------------------------------------------------------
# Import / edit window
# --------------------------------------------------------------------------


class _SheetTab(object):
    """One Excel-style tab: a checkbox (import it) plus the sheet name (view it)."""

    def __init__(self, owner, name, show_check):
        self.owner = owner
        self.name = name
        self.border = Border()
        self.border.BorderThickness = Thickness(1, 0, 1, 3)
        self.border.CornerRadius = CornerRadius(0, 0, 3, 3)
        self.border.Padding = Thickness(6, 4, 8, 4)
        self.border.Margin = Thickness(0, 0, 2, 2)
        self.border.Cursor = Cursors.Hand
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        self.check = CheckBox()
        self.check.VerticalAlignment = VerticalAlignment.Center
        self.check.Margin = Thickness(0, 0, 5, 0)
        self.check.ToolTip = 'Import this sheet'
        self.check.Click += self._checked
        if not show_check:
            self.check.Visibility = Visibility.Collapsed
        self.label = TextBlock()
        self.label.Text = name
        self.label.VerticalAlignment = VerticalAlignment.Center
        sp.Children.Add(self.check)
        sp.Children.Add(self.label)
        self.border.Child = sp
        self.border.MouseLeftButtonUp += self._clicked
        self.set_selected(False)

    def _clicked(self, sender, args):
        self.owner.select_sheet(self.name)

    def _checked(self, sender, args):
        self.owner.sheet_checked()

    @property
    def included(self):
        return bool(self.check.IsChecked)

    def set_selected(self, sel):
        pal = self.owner.pal
        self.border.Background = brush(pal['TabSel'] if sel else pal['Tab'])
        self.border.BorderBrush = brush(pal['TabAccent'] if sel else pal['Border'])
        self.border.BorderThickness = Thickness(1, 0, 1, 3 if sel else 1)
        self.label.FontWeight = FontWeights.SemiBold if sel else FontWeights.Normal
        self.label.Foreground = brush(pal['TabAccent'] if sel else pal['TabText'])


class ImportWindow(forms.WPFWindow):
    """Import (new drafting view per sheet) or, with link_data, edit one link."""

    def __init__(self, link_data=None, view=None):
        forms.WPFWindow.__init__(self, os.path.join(HERE, 'ImportWindow.xaml'))
        self.pal = theme.apply(self)
        brand(self, 'import')
        self.PreviewKeyDown += self._key_down
        self.edit = link_data is not None
        self.view = view
        self.result = None
        self.wb = None
        self.layout = None
        self.tabs = []
        self.current = None
        self.ranges = {}      # sheet -> fixed range ('' = auto)
        self._busy = True

        for cb, key, default in ((self.w_thin, 'thin', 1), (self.w_medium, 'medium', 3),
                                 (self.w_thick, 'thick', 5)):
            for n in range(1, 17):
                cb.Items.Add(str(n))
            saved = (link_data or {}).get('weights', {}).get(key) if link_data else None
            cb.SelectedItem = str(saved or cfg_get('weight_' + key, default))
            cb.SelectionChanged += self.opt_changed
        for z in ('50%', '75%', '100%', '125%', '150%', '200%'):
            self.zoom_cb.Items.Add(z)
        self.zoom_cb.SelectedItem = cfg_get('zoom', '100%')

        self._vtypes, self._unusable = drafting_view_types(revit.doc)
        self._legends = sorted([v for v in DB.FilteredElementCollector(revit.doc).OfClass(DB.View)
                                if v.ViewType == DB.ViewType.Legend and not v.IsTemplate],
                               key=lambda v: v.Name)
        self._pick = {'drafting': cfg_get('view_type', ''), 'legend': cfg_get('legend_source', '')}
        self._kind = None
        self.vname_tb.Text = cfg_get('view_name', '{sheet}')
        if cfg_get('view_kind', 'drafting') == 'legend' and self._legends:
            self.kind_legend_rb.IsChecked = True
        self._set_kind('legend' if self.kind_legend_rb.IsChecked else 'drafting')

        if self.edit:
            self.Title = 'Excel Link - Edit link'
            self.TitleSub.Text = 'Edit link'
            self.ok_b.Content = 'Update link'
            self.sheets_box.Visibility = Visibility.Collapsed
            self.views_panel.Visibility = Visibility.Collapsed
            o = link_data.get('opts', {})
            self.fills_cb.IsChecked = o.get('fills', True)
            self.borders_cb.IsChecked = o.get('borders', True)
            self.text_cb.IsChecked = o.get('text', True)
            self.grid_cb.IsChecked = o.get('gridlines', False)
            self.hidden_cb.IsChecked = o.get('include_hidden', False)
            path = xlink.resolve_path(revit.doc, link_data) or link_data.get('path')
            sheet = link_data.get('sheet')
            self._load(path, sheet)
            if sheet:
                self.ranges[sheet] = link_data.get('range') or ''
                self._show_range()
        else:
            self.grid_cb.IsChecked = bool(cfg_get('gridlines', False))
            self.hidden_cb.IsChecked = bool(cfg_get('include_hidden', False))
            last = cfg_get('last_file', '')
            if last and os.path.isfile(last):
                self._load(last)
        self._busy = False
        self.refresh()

    # ---- workbook & tabs ---------------------------------------------
    def _load(self, path, sheet=None):
        try:
            wb = Workbook(path)
        except Exception as ex:
            self.warn_lbl.Text = 'Could not read workbook:\n%s' % ex
            return
        self.wb = wb
        self.path_tb.Text = path
        self.path_tb.ToolTip = path
        self.ranges = {}
        self.tabs = []
        self.tabs_sp.Children.Clear()
        names = wb.visible_sheet_names or wb.sheet_names
        for n in names:
            tab = _SheetTab(self, n, not self.edit)
            self.tabs.append(tab)
            self.tabs_sp.Children.Add(tab.border)
        self.current = None
        self.select_sheet(sheet if sheet in names else names[0], refresh=False)
        self.sheet_checked()

    def select_sheet(self, name, refresh=True):
        self._store_range()
        self.current = name
        for t in self.tabs:
            t.set_selected(t.name == name)
            if t.name == name:
                try:
                    t.border.BringIntoView()
                except Exception:
                    pass
        was = self._busy
        self._busy = True
        self._show_range()
        self._busy = was
        try:
            idx = [t.name for t in self.tabs].index(name) + 1
            self.tabs_count.Text = '%d / %d' % (idx, len(self.tabs))
        except ValueError:
            self.tabs_count.Text = ''
        if refresh:
            self.refresh()
        self.sheet_checked()

    def _step_sheet(self, delta):
        names = [t.name for t in self.tabs]
        if not names or self.current not in names:
            return
        i = max(0, min(len(names) - 1, names.index(self.current) + delta))
        if names[i] != self.current:
            self.select_sheet(names[i])

    def help_click(self, sender, args):
        open_help('import')

    def _key_down(self, sender, args):
        from System.Windows.Input import Key, Keyboard, ModifierKeys
        if args.Key == Key.F1:
            open_help('import')
            args.Handled = True
            return
        if Keyboard.Modifiers & ModifierKeys.Control:
            if args.Key == Key.PageDown:
                self._step_sheet(1)
                args.Handled = True
            elif args.Key == Key.PageUp:
                self._step_sheet(-1)
                args.Handled = True

    def tabs_left_click(self, sender, args):
        self.tabs_sv.ScrollToHorizontalOffset(max(0, self.tabs_sv.HorizontalOffset - 240))

    def tabs_right_click(self, sender, args):
        self.tabs_sv.ScrollToHorizontalOffset(self.tabs_sv.HorizontalOffset + 240)

    def tabs_scrolled(self, sender, args):
        from System.Windows import Visibility as Vis
        sv, sb = self.tabs_sv, self.tabs_sb
        span = max(0.0, sv.ExtentWidth - sv.ViewportWidth)
        sb.Minimum = 0.0
        sb.Maximum = span
        sb.ViewportSize = sv.ViewportWidth
        sb.LargeChange = max(40.0, sv.ViewportWidth * 0.8)
        sb.Value = sv.HorizontalOffset
        sb.Visibility = Vis.Visible if span > 0.5 else Vis.Collapsed

    def tabs_scroll(self, sender, args):
        self.tabs_sv.ScrollToHorizontalOffset(self.tabs_sb.Value)

    def tabs_wheel(self, sender, args):
        self.tabs_sv.ScrollToHorizontalOffset(max(0, self.tabs_sv.HorizontalOffset - args.Delta))
        args.Handled = True

    def tabs_list_click(self, sender, args):
        from System.Windows.Controls import ContextMenu, MenuItem
        from System.Windows.Controls.Primitives import PlacementMode
        cm = ContextMenu()
        theme.fill(cm.Resources, self.pal)
        for t in self.tabs:
            mi = MenuItem()
            mi.Header = t.name + (u'   [import]' if t.included else u'')
            mi.IsChecked = (t.name == self.current)
            mi.Click += (lambda n: (lambda s, a: self.select_sheet(n)))(t.name)
            cm.Items.Add(mi)
        cm.PlacementTarget = sender
        cm.Placement = PlacementMode.Top
        cm.IsOpen = True

    def _sheet(self):
        if self.wb is None or self.current is None:
            return None
        return self.wb.sheet(self.current)

    def _show_range(self):
        sh = self._sheet()
        if sh is None:
            return
        self.range_hdr.Text = u'Range — %s' % self.current
        if sh.print_area:
            r = parse_range(sh.print_area)
            self.auto_lbl.Text = 'Print area (%s), follows the sheet' % format_range(*r)
        else:
            r = sh.used_range()
            self.auto_lbl.Text = 'Used range (%s), follows the sheet' % format_range(*r)
        fixed = self.ranges.get(self.current, '')
        if fixed:
            self.fixed_rb.IsChecked = True
            self.range_tb.Text = fixed
        else:
            self.auto_rb.IsChecked = True
            self.range_tb.Text = format_range(*r)
        self.range_tb.IsEnabled = bool(fixed)

    def _store_range(self):
        if self.current is None:
            return
        if self.fixed_rb.IsChecked:
            self.ranges[self.current] = (self.range_tb.Text or '').strip().upper()
        else:
            self.ranges[self.current] = ''

    def checked_sheets(self):
        return [t.name for t in self.tabs if t.included]

    def targets(self):
        """Sheets the Import button will bring in."""
        if self.edit:
            return [self.current] if self.current else []
        return self.checked_sheets() or ([self.current] if self.current else [])

    def sheet_checked(self):
        if self.edit:
            return
        chk = self.checked_sheets()
        if chk:
            self.sheets_lbl.Text = '%d of %d sheets ticked: %s' % (
                len(chk), len(self.tabs), ', '.join(chk[:6]) + (u' …' if len(chk) > 6 else ''))
            self.ok_b.Content = 'Import %d sheet%s' % (len(chk), '' if len(chk) == 1 else 's')
        else:
            self.sheets_lbl.Text = ('No tabs ticked: the tab you are viewing (%s) will be imported. '
                                    'Tick tabs below the preview to import several.' % (self.current or '-'))
            self.ok_b.Content = 'Import'
        self._update_view_names()

    def _update_view_names(self):
        if self.edit or not self.wb:
            return
        names = [self.view_name(s) for s in self.targets()]
        self.vname_tb.ToolTip = 'Will create: ' + ', '.join(names[:8])

    def view_name(self, sheet):
        stem = os.path.splitext(os.path.basename(self.path_tb.Text or ''))[0]
        pat = self.vname_tb.Text or '{sheet}'
        return revit_build.safe_name(pat.replace('{sheet}', sheet).replace('{workbook}', stem)) or sheet

    # ---- view kind -----------------------------------------------------
    def view_kind(self):
        return 'legend' if self.kind_legend_rb.IsChecked else 'drafting'

    def _set_kind(self, kind):
        if self._kind and self.vtype_cb.SelectedItem is not None:
            self._pick[self._kind] = str(self.vtype_cb.SelectedItem)
        self._kind = kind
        self.vtype_cb.Items.Clear()
        if kind == 'legend':
            self.vtype_lbl.Text = 'From'
            names = [v.Name for v in self._legends]
            if names:
                self.kind_note.Text = ('Revit can\'t create legends from scratch, so each one is a copy '
                                       'of this legend with its contents removed.')
            else:
                self.kind_note.Text = ('This model has no legend yet. Create one (View > Legends > Legend), '
                                       'then import again.')
        else:
            self.vtype_lbl.Text = 'Type'
            names = [_type_label(t, self._vtypes) for t in self._vtypes]
            self.kind_note.Text = ''
            if self._unusable:
                self.kind_note.Text = (u'Not listed: %s \u2014 callout (Detail view) type%s, which drafting '
                                       u'views can\'t use. Click + to add a drafting type with that name.'
                                       % (', '.join(self._unusable[:4]), '' if len(self._unusable) == 1 else 's'))
        self.vtype_new_b.Visibility = Visibility.Visible if kind == 'drafting' else Visibility.Collapsed
        for n in names:
            self.vtype_cb.Items.Add(n)
        want = self._pick.get(kind)
        self.vtype_cb.SelectedItem = want if want in names else (names[0] if names else None)
        self.vtype_cb.IsEnabled = bool(names)
        self.ok_b.IsEnabled = bool(names) and self.wb is not None

    def new_type_click(self, sender, args):
        suggest = self._unusable[0] if self._unusable else 'Excel Sheets'
        taken = set(DB.Element.Name.GetValue(t) for t in self._vtypes)
        name = forms.ask_for_string(default=suggest if suggest not in taken else '',
                                    prompt='Name for the new drafting view type:',
                                    title='New drafting view type')
        if not name:
            return
        name = name.strip()
        if name in taken:
            self.vtype_cb.SelectedItem = name
            return
        try:
            t = new_drafting_type(revit.doc, revit_build.safe_name(name))
        except Exception as ex:
            self.warn_lbl.Text = 'Could not create view type: %s' % ex
            return
        self._vtypes, self._unusable = drafting_view_types(revit.doc)
        self._pick['drafting'] = DB.Element.Name.GetValue(t)
        self._kind = None
        self._set_kind('drafting')

    def kind_changed(self, sender, args):
        if getattr(self, '_kind', None) is None:
            return
        self._set_kind(self.view_kind())

    # ---- options -----------------------------------------------------
    def opts(self):
        return {'fills': bool(self.fills_cb.IsChecked), 'borders': bool(self.borders_cb.IsChecked),
                'text': bool(self.text_cb.IsChecked), 'gridlines': bool(self.grid_cb.IsChecked),
                'include_hidden': bool(self.hidden_cb.IsChecked)}

    def weights(self):
        w = {}
        for cb, key in ((self.w_thin, 'thin'), (self.w_medium, 'medium'), (self.w_thick, 'thick')):
            w[key] = int(cb.SelectedItem or 1)
        w['hair'] = w['thin']
        w['grid'] = w['thin']
        return w

    def range_value(self):
        self._store_range()
        return self.ranges.get(self.current, '')

    def refresh(self):
        if self._busy:
            return
        sh = self._sheet()
        if sh is None:
            self.ok_b.IsEnabled = False
            return
        try:
            rng = self.range_value()
            if rng:
                parse_range(rng)
            self.layout = build_layout(sh, rng or None, xlink.options_from(self.opts()))
        except Exception as ex:
            self.layout = None
            self.ok_b.IsEnabled = False
            self.warn_lbl.Text = 'Range error: %s' % ex
            return
        lay = self.layout
        draw_preview(self.preview, lay)
        self._apply_zoom()
        w_in, h_in = lay.width / 72.0, lay.height / 72.0
        self.size_lbl.Text = (u'%s!%s  →  %.2f" × %.2f" printed (%.0f × %.0f mm)\n'
                              u'%d fills, %d line segments, %d text notes' % (
                                  self.current, format_range(*lay.range), w_in, h_in, w_in * 25.4, h_in * 25.4,
                                  len(lay.fills), len(lay.segments), len(lay.texts)))
        warns = list(lay.warnings)
        missing = self.wb.missing_cached_values(sh)
        if missing:
            warns.append('%d formula cell(s) have no saved result; save the workbook in Excel first.' % len(missing))
        miss = [f for f in set(t.font.name for t in lay.texts) if not revit_build.METRICS.installed(f)]
        if miss:
            warns.append('Not installed (Revit will substitute): %s' % ', '.join(sorted(miss)))
        self.warn_lbl.Text = '\n'.join(warns[:8])
        self.ok_b.IsEnabled = self.edit or self.vtype_cb.SelectedIndex >= 0

    def _apply_zoom(self):
        try:
            z = float(str(self.zoom_cb.SelectedItem).rstrip('%')) / 100.0
        except Exception:
            z = 1.0
        self.preview.LayoutTransform = ScaleTransform(z, z)

    # ---- events ------------------------------------------------------
    def browse_click(self, sender, args):
        init = os.path.dirname(self.path_tb.Text) if self.path_tb.Text else cfg_get('last_dir', '')
        path = forms.pick_file(files_filter='Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm',
                               init_dir=init if init and os.path.isdir(init) else None)
        if path:
            self._busy = True
            self._load(path)
            self._busy = False
            self.refresh()

    def check_all_click(self, sender, args):
        for t in self.tabs:
            t.check.IsChecked = True
        self.sheet_checked()

    def check_none_click(self, sender, args):
        for t in self.tabs:
            t.check.IsChecked = False
        self.sheet_checked()

    def range_mode_changed(self, sender, args):
        try:
            self.range_tb.IsEnabled = bool(self.fixed_rb.IsChecked)
        except Exception:
            pass
        self.refresh()

    def range_edited(self, sender, args):
        self.refresh()

    def range_key(self, sender, args):
        if str(args.Key) == 'Return':
            self.refresh()

    def opt_changed(self, sender, args):
        self.refresh()

    def vname_changed(self, sender, args):
        if not self._busy:
            self._update_view_names()

    def zoom_changed(self, sender, args):
        self._apply_zoom()

    def ok_click(self, sender, args):
        if not self.wb or not self.targets():
            return
        self._store_range()
        bad = []
        for s in self.targets():
            r = self.ranges.get(s, '')
            if r:
                try:
                    parse_range(r)
                except Exception:
                    bad.append('%s (%s)' % (s, r))
        if bad:
            self.warn_lbl.Text = 'Fix the range for: ' + ', '.join(bad)
            return
        path = self.path_tb.Text
        kind = self.view_kind()
        idx = self.vtype_cb.SelectedIndex
        pool = self._legends if kind == 'legend' else self._vtypes
        if not self.edit and (idx < 0 or idx >= len(pool)):
            self.warn_lbl.Text = ('Create a legend in the model first, or choose Drafting view.'
                                  if kind == 'legend' else 'No drafting view type available.')
            return
        source = pool[idx] if 0 <= idx < len(pool) else None
        self.result = {
            'path': path,
            'sheets': [(s, self.ranges.get(s, ''), self.view_name(s)) for s in self.targets()],
            'opts': self.opts(), 'weights': self.weights(),
            'view_kind': kind,
            'view_type_id': source.Id if (source is not None and kind == 'drafting') else None,
            'legend_source_id': source.Id if (source is not None and kind == 'legend') else None,
        }
        w = self.result['weights']
        cfg_set(last_file=path, last_dir=os.path.dirname(path),
                gridlines=self.result['opts']['gridlines'],
                include_hidden=self.result['opts']['include_hidden'],
                weight_thin=w['thin'], weight_medium=w['medium'], weight_thick=w['thick'],
                zoom=str(self.zoom_cb.SelectedItem), view_name=self.vname_tb.Text or '{sheet}',
                view_kind=kind,
                **({'legend_source': str(self.vtype_cb.SelectedItem or '')} if kind == 'legend'
                   else {'view_type': str(self.vtype_cb.SelectedItem or '')}))
        self.Close()

    def cancel_click(self, sender, args):
        self.result = None
        self.Close()


def _report(title, results, errors):
    warns = []
    for r in results:
        for w in r.warnings:
            if w not in warns:
                warns.append(w)
    if not warns and not errors:
        return
    out = script.get_output()
    out.set_title('Excel Link')
    out.print_md('### %s' % title)
    for e in errors:
        out.print_md('- **Error:** %s' % e)
    for w in warns:
        out.print_md('- %s' % w)


# --------------------------------------------------------------------------
# Commands
# --------------------------------------------------------------------------
def unique_view_name(doc, wanted, taken=None):
    taken = taken if taken is not None else set(
        v.Name for v in DB.FilteredElementCollector(doc).OfClass(DB.View) if not v.IsTemplate)
    name, n = wanted, 2
    while name in taken:
        name = '%s (%d)' % (wanted, n)
        n += 1
    taken.add(name)
    return name


def drafting_view_types(doc):
    """(drafting view types, names of callout Detail types). Drafting views only
    accept the Drafting family; callout (Detail view) types are reported so the
    UI can explain why they aren't offered."""
    allvft = list(DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType))
    usable = sorted([t for t in allvft if t.ViewFamily == DB.ViewFamily.Drafting],
                    key=lambda t: DB.Element.Name.GetValue(t).lower())
    names = set(DB.Element.Name.GetValue(t) for t in usable)
    unusable = sorted(set(DB.Element.Name.GetValue(t) for t in allvft
                          if t.ViewFamily == DB.ViewFamily.Detail) - names)
    return usable, unusable


def new_drafting_type(doc, name):
    """Duplicate the first drafting view type as *name*; returns the new type."""
    base = None
    for t in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
        if t.ViewFamily == DB.ViewFamily.Drafting:
            base = t
            break
    if base is None:
        raise ValueError('This model has no drafting view type to copy.')
    with xlink.Tx(doc, 'Excel Link: new drafting view type'):
        new = base.Duplicate(name)
    return doc.GetElement(new.Id) if hasattr(new, 'Id') else new


def _type_label(t, pool):
    return DB.Element.Name.GetValue(t)


def create_drafting_view(doc, name, type_id=None, taken=None):
    """New drafting view of *type_id* (any type drafting views accept)."""
    base = None
    for t in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
        if t.ViewFamily == DB.ViewFamily.Drafting:
            base = t.Id
            break
    want = doc.GetElement(type_id) if type_id is not None else None
    if want is not None and want.ViewFamily == DB.ViewFamily.Drafting:
        base = want.Id
    if base is None:
        raise ValueError('This model has no drafting view type.')
    view = DB.ViewDrafting.Create(doc, base)
    if want is not None and want.Id != base:
        view.ChangeTypeId(want.Id)
    view.Scale = 1
    view.Name = unique_view_name(doc, name, taken)
    return view


def create_legend_view(doc, name, source_id, taken=None):
    """Revit has no API to make a blank legend, so duplicate *source_id*
    (an existing legend) and delete whatever it contains."""
    src = doc.GetElement(source_id)
    if src is None or src.ViewType != DB.ViewType.Legend:
        raise ValueError('Choose an existing legend to copy.')
    new_id = src.Duplicate(DB.ViewDuplicateOption.Duplicate)
    view = doc.GetElement(new_id)
    owned = [e.Id for e in DB.FilteredElementCollector(doc, new_id).WhereElementIsNotElementType()
             if e.OwnerViewId == new_id and e.Category is not None and e.Id != new_id]
    if owned:
        doc.Delete(revit_build.to_list(DB.ElementId, owned))
    view.Scale = 1
    view.Name = unique_view_name(doc, name, taken)
    return view


def import_sheets(doc, path, sheets, opts, weights, view_type_id=None, view_kind='drafting',
                  legend_source_id=None):
    """Create one new drafting view or legend per (sheet, range, view_name)
    and build a linked group in each. Returns (views, results, errors)."""
    views, results, errors = [], [], []
    styles = revit_build.Styles(doc, weights)
    taken = set(v.Name for v in DB.FilteredElementCollector(doc).OfClass(DB.View) if not v.IsTemplate)
    tg = DB.TransactionGroup(doc, 'Excel Link: import')
    tg.Start()
    try:
        for sheet, rng, vname in sheets:
            try:
                with xlink.Tx(doc, 'Excel Link: import %s' % sheet):
                    if view_kind == 'legend':
                        view = create_legend_view(doc, vname, legend_source_id, taken)
                    else:
                        view = create_drafting_view(doc, vname, view_type_id, taken)
                    grp, res = xlink.create(doc, view, path, sheet, rng, opts, DB.XYZ(0, 0, 0),
                                            weights=weights, styles=styles)
                views.append(view)
                results.append(res)
            except Exception as ex:
                styles = revit_build.Styles(doc, weights)
                errors.append('%s: %s' % (sheet, ex))
    finally:
        if views:
            tg.Assimilate()
        else:
            tg.RollBack()
    return views, results, errors


def run_import():
    doc, uidoc = revit.doc, revit.uidoc
    win = ImportWindow()
    win.show_dialog()
    r = win.result
    if not r:
        return
    try:
        views, results, errors = import_sheets(doc, r['path'], r['sheets'], r['opts'], r['weights'],
                                               r.get('view_type_id'), r.get('view_kind', 'drafting'),
                                               r.get('legend_source_id'))
    except Exception as ex:
        forms.alert('Import failed:\n%s' % ex, expanded=traceback.format_exc())
        return
    if views:
        try:
            uidoc.ActiveView = views[0]
        except Exception:
            pass
    _report('Import notes', results, errors)
    if len(r['sheets']) > 1 or errors:
        forms.alert('Created %d %s:\n%s%s' % (
            len(views), 'legend(s)' if r.get('view_kind') == 'legend' else 'drafting view(s)', '\n'.join('  ' + v.Name for v in views),
            ('\n\nErrors:\n' + '\n'.join(errors)) if errors else ''), warn_icon=bool(errors))


def _update_groups(doc, groups):
    results, errors, done = [], [], 0
    styles = revit_build.Styles(doc)
    tg = DB.TransactionGroup(doc, 'Excel Link: update')
    tg.Start()
    try:
        for g in groups:
            try:
                with xlink.Tx(doc, 'Excel Link: update'):
                    styles.weights = dict(revit_build.DEFAULT_WEIGHTS)
                    data = xlink.read_link(g) or {}
                    styles.weights.update(data.get('weights') or {})
                    _, res = xlink.update(doc, g, styles=styles)
                results.append(res)
                done += 1
            except Exception as ex:
                styles = revit_build.Styles(doc)  # cached styles may have been rolled back
                try:
                    name = DB.Element.Name.GetValue(g.GroupType)
                except Exception:
                    name = 'Excel Link'
                errors.append('%s: %s' % (name, ex))
    finally:
        if done:
            tg.Assimilate()
        else:
            tg.RollBack()
    return done, results, errors


def run_update():
    doc = revit.doc
    sel = [e for e in revit.get_selection() if isinstance(e, DB.Group) and xlink.read_link(e)]
    if sel:
        targets = sel
    else:
        links = xlink.find_links(doc)
        if not links:
            forms.alert('No Excel Link groups in this model.', exitscript=True)
        stale = []
        for g in links:
            st, _ = xlink.status(doc, g)
            if st in (xlink.CHANGED, xlink.SCALE):
                stale.append(g)
        missing = [g for g in links if xlink.status(doc, g)[0] == xlink.MISSING]
        if not stale:
            msg = 'All %d Excel Link(s) are up to date.' % len(links)
            if missing:
                msg += '\n%d link(s) point to a missing workbook; use Excel Links > Edit to repath.' % len(missing)
            forms.alert(msg, exitscript=True)
        targets = stale
    done, results, errors = _update_groups(doc, targets)
    _report('Update notes', results, errors)
    forms.alert('Updated %d of %d Excel Link(s).%s' % (
        done, len(targets), '\nSee the output window for errors.' if errors else ''), warn_icon=bool(errors))


class LinkRow(object):
    def __init__(self, doc, group):
        self.group = group
        self.data = xlink.read_link(group) or {}
        st, path = xlink.status(doc, group, self.data)
        self.Status = st
        v = doc.GetElement(group.OwnerViewId)
        self.ViewName = v.Name if v is not None else '?'
        self.Path = path or self.data.get('path', '')
        self.FileName = os.path.basename(self.Path)
        self.Source = '%s!%s%s' % (self.data.get('sheet', '?'), self.data.get('resolved_range', ''),
                                   '' if self.data.get('range') else ' (auto)')
        self.Imported = self.data.get('imported', '')


class ManagerWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, os.path.join(HERE, 'ManagerWindow.xaml'))
        self.pal = theme.apply(self)
        brand(self, 'manager')
        self.PreviewKeyDown += self._key_down
        self.reload()

    def _key_down(self, sender, args):
        from System.Windows.Input import Key
        if args.Key == Key.F1:
            open_help('manager')
            args.Handled = True

    def help_click(self, sender, args):
        open_help('manager')

    def reload(self, keep_ids=None):
        doc = revit.doc
        rows = [LinkRow(doc, g) for g in xlink.find_links(doc)]
        rows.sort(key=lambda r: (r.ViewName, r.FileName))
        self.links_lv.ItemsSource = rows
        n = len([r for r in rows if r.Status != xlink.OK])
        self.msg_lbl.Text = '%d link(s), %d need attention.' % (len(rows), n)

    def selected(self):
        return list(self.links_lv.SelectedItems)

    def _update(self, rows):
        if not rows:
            self.msg_lbl.Text = 'Select one or more links first.'
            return
        done, results, errors = _update_groups(revit.doc, [r.group for r in rows])
        self.reload()
        self.msg_lbl.Text = 'Updated %d of %d.%s' % (done, len(rows), ('  Errors: ' + '; '.join(errors)) if errors else '')
        _report('Update notes', results, errors)

    def update_click(self, sender, args):
        self._update(self.selected())

    def update_changed_click(self, sender, args):
        rows = [r for r in self.links_lv.ItemsSource if r.Status in (xlink.CHANGED, xlink.SCALE)]
        if not rows:
            self.msg_lbl.Text = 'Nothing to update.'
            return
        self._update(rows)

    def edit_click(self, sender, args):
        rows = self.selected()
        if len(rows) != 1:
            self.msg_lbl.Text = 'Select exactly one link to edit.'
            return
        row = rows[0]
        doc = revit.doc
        view = doc.GetElement(row.group.OwnerViewId)
        win = ImportWindow(link_data=row.data, view=view)
        win.show_dialog()
        r = win.result
        if not r:
            return
        try:
            with xlink.Tx(doc, 'Excel Link: edit'):
                data = dict(row.data)
                sheet, rng, _ = r['sheets'][0]
                data.update({'path': r['path'], 'rel': xlink.rel_path(doc, r['path']),
                             'cloud': xlink.cloud_location(r['path']), 'sheet': sheet,
                             'range': rng, 'opts': r['opts'], 'weights': r['weights']})
                xlink.write_link(row.group, data)
                styles = revit_build.Styles(doc, r['weights'])
                _, res = xlink.update(doc, row.group, styles=styles)
            _report('Edit notes', [res], [])
        except Exception as ex:
            forms.alert('Edit failed:\n%s' % ex, expanded=traceback.format_exc())
        self.reload()

    def open_click(self, sender, args):
        for r in self.selected():
            if r.Path and os.path.isfile(r.Path):
                try:
                    open_in_excel(r.Path, r.data.get('sheet'),
                                  r.data.get('resolved_range') or r.data.get('range'))
                except Exception as ex:
                    self.msg_lbl.Text = 'Could not open: %s' % ex
            else:
                self.msg_lbl.Text = 'Workbook not found: %s' % r.Path

    def show_click(self, sender, args):
        rows = self.selected()
        if not rows:
            return
        ids = revit_build.to_list(DB.ElementId, [r.group.Id for r in rows])
        revit.uidoc.Selection.SetElementIds(ids)
        try:
            v = revit.doc.GetElement(rows[0].group.OwnerViewId)
            if v is not None:
                revit.uidoc.RequestViewChange(v)
            revit.uidoc.ShowElements(ids)
        except Exception:
            pass
        self.Close()

    def unlink_click(self, sender, args):
        rows = self.selected()
        if not rows:
            return
        if not forms.alert('Remove the Excel link from %d group(s)?\n'
                           'The drawing stays; it just won\'t update any more.' % len(rows),
                           yes=True, no=True):
            return
        sc = xlink.schema()
        with xlink.Tx(revit.doc, 'Excel Link: unlink'):
            for r in rows:
                r.group.DeleteEntity(sc)
        self.reload()

    def close_click(self, sender, args):
        self.Close()


def open_in_excel(path, sheet=None, rng=None):
    """Open *path* in Excel and go to *sheet*!*rng*. Uses Excel automation when
    Excel is installed (reusing a running Excel and an already-open copy of the
    workbook); otherwise just opens the file with its default program."""
    if not path or not os.path.isfile(path):
        raise IOError('Workbook not found: %s' % path)
    try:
        import System
        from System.Runtime.InteropServices import Marshal
        if not hasattr(Marshal, 'GetActiveObject'):
            # .NET 8 (Revit 2025+) has no way to attach to a running Excel;
            # starting a second Excel would open the workbook read-only.
            raise NotImplementedError
        try:
            xl = Marshal.GetActiveObject('Excel.Application')
        except Exception:
            xl = System.Activator.CreateInstance(System.Type.GetTypeFromProgID('Excel.Application'))
        xl.Visible = True
        wb = None
        target = os.path.normcase(os.path.abspath(path))
        for i in range(1, xl.Workbooks.Count + 1):
            w = xl.Workbooks.Item(i)
            try:
                if os.path.normcase(os.path.abspath(w.FullName)) == target:
                    wb = w
                    break
            except Exception:
                pass
        if wb is None:
            wb = xl.Workbooks.Open(path)
        wb.Activate()
        if sheet:
            try:
                ws = wb.Worksheets.Item(sheet)
                ws.Activate()
                if rng:
                    ws.Range(rng).Select()
            except Exception:
                pass
        try:
            xl.WindowState = -4143 if xl.WindowState == -4140 else xl.WindowState  # restore if minimised
        except Exception:
            pass
        return True
    except Exception:
        os.startfile(path)
        return False


def _links_in_view(doc, view):
    """Linked groups in *view*, or in every view placed on it when it's a sheet."""
    if isinstance(view, DB.ViewSheet):
        groups = []
        for vid in view.GetAllPlacedViews():
            v = doc.GetElement(vid)
            if v is not None:
                groups.extend(xlink.find_links(doc, v))
        return groups
    return xlink.find_links(doc, view)


def run_open_excel():
    doc = revit.doc
    view = doc.ActiveView
    sel = [e for e in revit.get_selection() if isinstance(e, DB.Group) and xlink.read_link(e)]
    groups = sel or _links_in_view(doc, view)
    if not groups:
        forms.alert('No Excel Link tables in "%s".\n\nOpen a view (or sheet) that has an imported '
                    'Excel table, or select one, then try again.' % view.Name, exitscript=True)
    choices = {}
    for g in groups:
        d = xlink.read_link(g) or {}
        path = xlink.resolve_path(doc, d) or d.get('path', '')
        key = os.path.normcase(path)
        if key not in choices:
            choices[key] = (path, d.get('sheet'), d.get('resolved_range') or d.get('range'), [])
        choices[key][3].append('%s!%s' % (d.get('sheet', '?'), d.get('resolved_range', '')))
    items = list(choices.values())
    if len(items) > 1:
        labels = ['%s   (%s)' % (os.path.basename(p), ', '.join(srcs[:3]) + (u' \u2026' if len(srcs) > 3 else ''))
                  for p, _, _, srcs in items]
        pick = forms.SelectFromList.show(labels, title='Open which workbook?', button_name='Open in Excel')
        if not pick:
            return
        items = [items[labels.index(pick)]]
    path, sheet, rng, _ = items[0]
    if not path or not os.path.isfile(path):
        forms.alert('Workbook not found:\n%s\n\nUse Excel Links > Edit to point the link at its new location.'
                    % (path or '(no path stored)'), exitscript=True)
    try:
        open_in_excel(path, sheet, rng)
    except Exception as ex:
        forms.alert('Could not open the workbook:\n%s' % ex)


def run_manager():
    if not xlink.find_links(revit.doc):
        forms.alert('No Excel Link groups in this model yet.\nUse "Import Excel" first.', exitscript=True)
    ManagerWindow().show_dialog()
