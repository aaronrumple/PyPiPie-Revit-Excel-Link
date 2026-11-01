# -*- coding: utf-8 -*-
"""Notify when Excel Link groups in the opened model are out of date."""
try:
    from pyrevit import EXEC_PARAMS, forms
    from excellink import link as xlink

    _doc = EXEC_PARAMS.event_args.Document
    if _doc is not None and not _doc.IsFamilyDocument:
        _links = xlink.find_links(_doc)
        if _links:
            _stale = [g for g in _links if xlink.status(_doc, g)[0] != xlink.OK]
            if _stale:
                forms.toast('%d of %d Excel Link(s) in %s are out of date or missing. '
                            'Use Excel Link > Update Links.' % (len(_stale), len(_links), _doc.Title),
                            title='Excel Link', appid='Excel Link')
except Exception:
    pass
