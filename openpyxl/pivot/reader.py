from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

from zipfile import ZipFile

from openpyxl.xml.constants import ARC_CONTENT_TYPES
from openpyxl.xml.functions import fromstring
from openpyxl.packaging.manifest import Manifest
from openpyxl.packaging.relationship import get_dependents, get_rels_path

from .pivot import PivotTableDefinition
from .cache import PivotCacheDefinition
from .record import RecordList


def get_rel(archive, deps, id=None, cls=None):
    """
    Get related object based on id or rel_type
    """
    if not any([id, cls]):
        raise ValueError("Either the id or the content type are required")
    if id is not None:
        rel = deps[id]
    else:
        rel = next(deps.find(cls.rel_type))

    path = rel.target
    src = archive.read(path)
    tree = fromstring(src)
    obj = cls.from_tree(tree)

    rels_path = get_rels_path(path)
    try:
        obj.deps = get_dependents(archive, rels_path)
    except KeyError:
        obj.deps = []

    return obj


def read_pivot(file):
    archive = ZipFile(file)

    src = archive.read(ARC_CONTENT_TYPES)
    root = fromstring(src)
    package = Manifest.from_tree(root)

    tables = package.findall(PivotTableDefinition.mime_type)
    table = list(tables)[0]
    path = table.PartName[1:]
    src = archive.read(path)
    tree = fromstring(src)
    table = PivotTableDefinition.from_tree(tree)

    rels_path = get_rels_path(path)
    deps = get_dependents(archive, rels_path)

    cache = get_rel(archive, deps, table.id, PivotCacheDefinition)
    table.cache = cache

    records = get_rel(archive, cache.deps, cache.id, RecordList)
    cache.records = records

    return table
