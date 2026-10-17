"""Repair a known signed character-grid export defect without rewriting XML.

Only documents freshly exported by the converter use this path. Existing DOCX
inputs continue to use the byte-preserving path.
"""
from __future__ import annotations

import io
import re
from xml.parsers import expat
from zipfile import ZipFile

WORD_NAMESPACES = {
    'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'http://purl.oclc.org/ooxml/wordprocessingml/main',
}
START_TAG = re.compile(rb'<(?:[^>\'"]|\'[^\']*\'|"[^"]*")*>')


def _name(name):
    fields = name.split('|')
    return (fields + [''] * 3)[:3]


def _repair_grid_xml(data):
    edits = []
    parser = expat.ParserCreate(namespace_separator='|')
    parser.namespace_prefixes = True

    def start(name, attributes):
        namespace, local, _ = _name(name)
        if namespace not in WORD_NAMESPACES or local != 'docGrid':
            return
        for attribute, value in attributes.items():
            attr_namespace, attr_local, prefix = _name(attribute)
            if attr_namespace != namespace or attr_local != 'charSpace':
                continue
            if not re.fullmatch(r'[+-]?[0-9]+', value):
                raise ValueError('转换引擎输出的中文字符网格字距无效。')
            number = int(value)
            if not -(2 ** 31) <= number < 2 ** 32:
                raise ValueError('转换引擎输出的中文字符网格字距越界。')
            if number < 2 ** 31:
                continue
            # Word's charSpace is an Int32Value. Some exports write a negative
            # two's-complement value as uint32, changing Chinese line wrapping.
            # https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.docgrid.characterspace
            offset = parser.CurrentByteIndex
            tag = START_TAG.match(data, offset)
            if tag is None or not prefix:
                raise ValueError('无法可靠修复中文字符网格字距。')
            qname = re.escape((prefix + ':charSpace').encode('utf-8'))
            attribute_match = re.search(
                rb'(?<![\w:.-])' + qname + rb'\s*=\s*([\'"])([+-]?[0-9]+)\1',
                tag.group())
            if attribute_match is None or int(attribute_match[2]) != number:
                raise ValueError('无法可靠定位中文字符网格字距。')
            first, last = attribute_match.span(2)
            edits.append((offset + first, offset + last, str(number - 2 ** 32).encode('ascii')))

    parser.StartElementHandler = start
    parser.Parse(data, True)
    for first, last, replacement in reversed(edits):
        data = data[:first] + replacement + data[last:]
    return data, len(edits)


def repair_exported_layout(data):
    """Return exported DOCX bytes and a record of exact layout repairs."""
    with ZipFile(io.BytesIO(data)) as source:
        original = source.read('word/document.xml')
        repaired, count = _repair_grid_xml(original)
        if not count:
            return data, []
        buffer = io.BytesIO()
        with ZipFile(buffer, 'w') as output:
            output.comment = source.comment
            for info in source.infolist():
                output.writestr(info, repaired if info.filename == 'word/document.xml' else source.read(info))
    return buffer.getvalue(), [{'kind': 'signed_character_grid', 'sections': count}]
