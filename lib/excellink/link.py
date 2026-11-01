# -*- coding: utf-8 -*-
"""Import / update orchestration and the link data stored on each group.

Link data lives in Extensible Storage on the detail Group instance, as a
JSON string. It travels with the group when it is copied or pasted, so each
copy can be updated independently.
"""
from __future__ import division

import datetime
import hashlib
import json
import os

import System
from Autodesk.Revit import DB
from Autodesk.Revit.DB.ExtensibleStorage import (SchemaBuilder, Schema, Entity,
                                                 AccessLevel, ExtensibleStorageFilter)

from .xlsx_reader import Workbook, read_file_bytes, format_range
from .layout import build_layout, Options
from . import revit_build

SCHEMA_GUID = System.Guid('6d3b1f0a-2c8e-4b71-9a55-e1c0b7a4f3d2')
SCHEMA_NAME = 'ExcelLinkData'
VERSION = 1

# status codes
OK, CHANGED, MISSING, SCALE, ERROR = 'Up to date', 'Excel changed', 'File missing', 'View scale changed', 'Error'


# --------------------------------------------------------------------------
# Transactions that hide the harmless "group changed outside edit mode"
# warning Revit raises while a freshly made group gets its link data.
# --------------------------------------------------------------------------
class _QuietGroups(DB.IFailuresPreprocessor):
    __namespace__ = 'ExcelLink'

    def PreprocessFailures(self, accessor):
        try:
            quiet = DB.BuiltInFailures.GroupFailures.AtomViolationWhenOnePerGroup
        except Exception:
            quiet = None
        for f in list(accessor.GetFailureMessages()):
            if f.GetSeverity() != DB.FailureSeverity.Warning:
                continue
            if (quiet is not None and f.GetFailureDefinitionId() == quiet) or \
                    'outside group edit mode' in (f.GetDescriptionText() or ''):
                accessor.DeleteWarning(f)
        return DB.FailureProcessingResult.Continue


class Tx(object):
    """with Tx(doc, 'name'): ...  -- commits on success, rolls back on error."""

    def __init__(self, doc, name):
        self.t = DB.Transaction(doc, name)

    def __enter__(self):
        self.t.Start()
        opts = self.t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(_QuietGroups())
        self.t.SetFailureHandlingOptions(opts)
        return self.t

    def __exit__(self, exc_type, exc, tb):
        if exc_type is None:
            self.t.Commit()
        elif self.t.HasStarted() and not self.t.HasEnded():
            self.t.RollBack()
        return False


# --------------------------------------------------------------------------
# Extensible storage
# --------------------------------------------------------------------------
def schema():
    sc = Schema.Lookup(SCHEMA_GUID)
    if sc is not None:
        return sc
    sb = SchemaBuilder(SCHEMA_GUID)
    sb.SetSchemaName(SCHEMA_NAME)
    sb.SetReadAccessLevel(AccessLevel.Public)
    sb.SetWriteAccessLevel(AccessLevel.Public)
    sb.SetDocumentation('Excel Link (pyRevit): source workbook, range and import options.')
    sb.AddSimpleField('Json', System.String)
    return sb.Finish()


def read_link(elem):
    sc = Schema.Lookup(SCHEMA_GUID)
    if sc is None or elem is None:
        return None
    try:
        ent = elem.GetEntity(sc)
    except Exception:
        return None
    if ent is None or not ent.IsValid():
        return None
    try:
        raw = ent.Get[System.String]('Json')
        return json.loads(raw) if raw else None
    except Exception:
        return None


def write_link(elem, data):
    sc = schema()
    ent = Entity(sc)
    ent.Set[System.String]('Json', json.dumps(data, sort_keys=True))
    elem.SetEntity(ent)


def find_links(doc, view=None):
    sc = Schema.Lookup(SCHEMA_GUID)
    if sc is None:
        return []
    col = DB.FilteredElementCollector(doc) if view is None else DB.FilteredElementCollector(doc, view.Id)
    col = col.OfClass(DB.Group).WherePasses(ExtensibleStorageFilter(SCHEMA_GUID))
    return list(col)


# --------------------------------------------------------------------------
# File fingerprinting & path resolution
# --------------------------------------------------------------------------
def fingerprint(path):
    data = read_file_bytes(path)
    return {
        'mtime': os.path.getmtime(path),
        'size': len(data),
        'sha1': hashlib.sha1(data).hexdigest(),
    }, data


def doc_folder(doc):
    try:
        p = doc.PathName
        if p and os.path.isfile(p):
            return os.path.dirname(p)
    except Exception:
        pass
    return None


def rel_path(doc, path):
    folder = doc_folder(doc)
    if not folder:
        return None
    try:
        if os.path.splitdrive(folder)[0].lower() != os.path.splitdrive(path)[0].lower():
            return None
        return os.path.relpath(path, folder)
    except Exception:
        return None


# --------------------------------------------------------------------------
# Autodesk Desktop Connector (ACC / BIM 360 / Autodesk Docs)
# --------------------------------------------------------------------------
# Desktop Connector mirrors cloud projects under a per-user workspace, e.g.
#   C:\Users\<user>\DC\ACCDocs\<Hub>\<Project>\Project Files\...
# The part from the connector folder (ACCDocs) down is the same for everyone
# on the project, so links store it and each machine resolves it against its
# own workspace.
DC_CONNECTORS = ('ACCDocs', 'BIM 360', 'Autodesk Docs')


def cloud_location(path):
    """'ACCDocs\\Hub\\Project\\...\\file.xlsx' if *path* is inside a Desktop
    Connector workspace, else None."""
    if not path:
        return None
    parts = os.path.normpath(path).split(os.sep)
    low = [x.lower() for x in parts]
    for name in DC_CONNECTORS:
        if name.lower() in low:
            i = low.index(name.lower())
            if i + 2 < len(parts):  # connector \ hub \ ... \ file
                return os.sep.join(parts[i:])
    return None


def dc_workspaces(stored_path=None):
    """Candidate Desktop Connector workspace roots on this machine."""
    roots = []
    home = os.environ.get('USERPROFILE') or os.path.expanduser('~')
    roots.append(os.path.join(home, 'DC'))
    # same folder layout as where the link was made, under this user's profile
    # (covers a workspace moved off the default, e.g. D:\DC stays D:\DC)
    if stored_path:
        parts = os.path.normpath(stored_path).split(os.sep)
        low = [x.lower() for x in parts]
        for name in DC_CONNECTORS:
            if name.lower() in low:
                root = os.sep.join(parts[:low.index(name.lower())])
                if root and root not in roots:
                    roots.append(root)
                if 'users' in low:
                    u = low.index('users')
                    if u + 1 < low.index(name.lower()):
                        swapped = os.sep.join([home] + parts[u + 2:low.index(name.lower())])
                        if swapped not in roots:
                            roots.append(swapped)
                break
    env = os.environ.get('EXCELLINK_DC_ROOT')
    if env:
        roots.insert(0, env)
    return roots


def resolve_path(doc, data):
    """Find the workbook on this machine: the stored path, then the
    Desktop Connector location under this user's workspace, then the path
    relative to the Revit model."""
    p = data.get('path')
    if p and os.path.isfile(p):
        return p
    cloud = data.get('cloud') or cloud_location(p)
    if cloud:
        for root in dc_workspaces(p):
            cand = os.path.join(root, cloud)
            if os.path.isfile(cand):
                return cand
    rel = data.get('rel')
    folder = doc_folder(doc)
    if rel and folder:
        cand = os.path.normpath(os.path.join(folder, rel))
        if os.path.isfile(cand):
            return cand
    return None


def status(doc, group, data=None):
    data = data or read_link(group)
    if not data:
        return ERROR, None
    path = resolve_path(doc, data)
    if not path:
        return MISSING, None
    view = doc.GetElement(group.OwnerViewId)
    if view is not None and int(view.Scale) != int(data.get('scale', view.Scale)):
        return SCALE, path
    try:
        size = os.path.getsize(path)
        mtime = os.path.getmtime(path)
    except Exception:
        return MISSING, None
    if size == data.get('size') and abs(mtime - data.get('mtime', 0)) < 1e-3:
        return OK, path
    # timestamp moved; confirm via hash (touch / sync clients change mtime)
    try:
        fp, _ = fingerprint(path)
        if fp['sha1'] == data.get('sha1'):
            return OK, path
    except Exception:
        return ERROR, path
    return CHANGED, path


# --------------------------------------------------------------------------
# Import / update
# --------------------------------------------------------------------------
def options_from(d):
    d = d or {}
    return Options(include_hidden=d.get('include_hidden', False),
                   gridlines=d.get('gridlines', False),
                   fills=d.get('fills', True), borders=d.get('borders', True),
                   text=d.get('text', True))


def layout_for(path, sheet_name, rng, opts_dict, data=None):
    wb = Workbook(path, data=data)
    if sheet_name not in wb.sheet_names:
        raise KeyError('Worksheet "%s" no longer exists in %s' % (sheet_name, os.path.basename(path)))
    sh = wb.sheet(sheet_name)
    lay = build_layout(sh, rng or None, options_from(opts_dict))
    missing = wb.missing_cached_values(sh)
    if missing:
        lay.warnings.append('%d formula cell(s) have no saved result (file not saved by Excel?). '
                            'Open and save the workbook in Excel to refresh them.' % len(missing))
    return lay


def group_anchor(group, view):
    loc = group.Location
    if isinstance(loc, DB.LocationPoint):
        p = loc.Point
        return [p.X, p.Y, p.Z]
    bb = group.get_BoundingBox(view)
    if bb is not None:
        return [bb.Min.X, bb.Max.Y, bb.Min.Z]
    return [0.0, 0.0, 0.0]


def group_name(path, sheet, rng):
    stem = os.path.splitext(os.path.basename(path))[0]
    return revit_build.safe_name('Excel - %s - %s - %s' % (stem, sheet, (rng or 'auto').replace(':', '-')))


def create(doc, view, path, sheet, rng, opts_dict, origin, weights=None, styles=None, name=None):
    """Build and group a linked table. Call inside a Transaction.
    Returns (group, BuildResult)."""
    fp, data = fingerprint(path)
    lay = layout_for(path, sheet, rng, opts_dict, data)
    res = revit_build.build(doc, view, lay, origin, styles=styles, weights=weights)
    actual_range = format_range(*lay.range)
    grp = revit_build.make_group(doc, res.element_ids, name or group_name(path, sheet, actual_range))
    res.group = grp
    if grp is None:
        return None, res
    link = {
        'v': VERSION,
        'path': path,
        'rel': rel_path(doc, path),
        'cloud': cloud_location(path),
        'sheet': sheet,
        'range': rng or '',
        'resolved_range': actual_range,
        'opts': opts_dict or {},
        'weights': weights or {},
        'origin': [origin.X, origin.Y, origin.Z],
        'anchor': group_anchor(grp, view),
        'scale': int(view.Scale),
        'size_pt': [round(lay.width, 3), round(lay.height, 3)],
        'imported': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }
    link.update(fp)
    write_link(grp, link)
    p = grp.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if p is not None and not p.IsReadOnly:
        try:
            p.Set('Excel Link: %s [%s!%s]' % (os.path.basename(path), sheet, actual_range))
        except Exception:
            pass
    return grp, res


def update(doc, group, new_path=None, styles=None):
    """Rebuild *group* from its workbook at its current position.
    Call inside a Transaction. Returns (new_group, BuildResult)."""
    data = read_link(group)
    if not data:
        raise ValueError('Selected group carries no Excel Link data.')
    path = new_path or resolve_path(doc, data)
    if not path:
        raise IOError('Workbook not found: %s' % data.get('path'))
    view = doc.GetElement(group.OwnerViewId)
    if view is None:
        raise ValueError('Excel Link groups must be view-specific (detail) groups.')

    # follow the group if the user moved it
    now = group_anchor(group, view)
    was = data.get('anchor', now)
    o = data.get('origin', now)
    origin = DB.XYZ(o[0] + now[0] - was[0], o[1] + now[1] - was[1], o[2] + now[2] - was[2])

    old_type = group.GroupType
    old_name = DB.Element.Name.GetValue(old_type)
    pinned = group.Pinned
    siblings = [i for i in old_type.Groups if i.Id != group.Id] if hasattr(old_type, 'Groups') else []

    if pinned:
        group.Pinned = False
    members = group.UngroupMembers()
    if members is not None and members.Count:
        doc.Delete(members)
    if not siblings:
        try:
            doc.Delete(old_type.Id)
        except Exception:
            pass

    grp, res = create(doc, view, path, data['sheet'], data.get('range') or '',
                      data.get('opts'), origin, weights=data.get('weights') or None,
                      styles=styles, name=old_name if not siblings else None)
    if grp is not None:
        if pinned:
            grp.Pinned = True
    return grp, res
