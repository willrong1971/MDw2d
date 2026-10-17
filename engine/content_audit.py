"""Fail-closed text audit for document stories outside the main body.

RTF uses the supplied ODT import baseline, not an independent RTF parser.
DOC also checks the independently extracted binary story text. Paragraph and
line separators and PAGE/NUMPAGES cached values are ignored; spaces are not.
"""
from __future__ import annotations

import hashlib
import posixpath
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
TEXT = '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}'
OFFICE = '{urn:oasis:names:tc:opendocument:xmlns:office:1.0}'
STYLE = '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}'
DRAW = '{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}'
DC = '{http://purl.org/dc/elements/1.1/}'
MAX_XML = 64 * 1024 * 1024
KINDS = ('footnote', 'endnote', 'annotation', 'header', 'footer',
         'textbox', 'header_textbox')
DYNAMIC = re.compile(r'^\s*(?:PAGE|NUMPAGES)(?=\s|\\|$)', re.I)
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
SEPARATORS = ('footnote_separator', 'footnote_continuationSeparator',
              'footnote_continuationNotice', 'endnote_separator',
              'endnote_continuationSeparator', 'endnote_continuationNotice')


def _normalized(text):
    # Generated reference/anchor characters are structural, not visible text.
    return ''.join(c for c in text if c not in '\r\n\x01\x02\x05\x07\x08\x0b\x0c\x0e')


def _binary_visible(text):
    output = []
    fields = []
    for char in text:
        if char == '\x13':
            fields.append({'code': '', 'result': False, 'dynamic': False})
        elif char == '\x14':
            if not fields:
                raise ValueError('旧文档非正文包含无法核验的域分隔符。')
            fields[-1]['result'] = True
            fields[-1]['dynamic'] = bool(DYNAMIC.match(fields[-1]['code']))
        elif char == '\x15':
            if not fields:
                raise ValueError('旧文档非正文包含无法核验的域结束符。')
            fields.pop()
        elif fields and not fields[-1]['result']:
            fields[-1]['code'] += char
        elif not any(not f['result'] or f['dynamic'] for f in fields):
            output.append(char)
    if fields:
        raise ValueError('旧文档非正文域未闭合，无法可靠核验。')
    return _normalized(''.join(output))


def _read_xml(z, name):
    info = z.getinfo(name)
    if info.file_size > MAX_XML:
        raise ValueError('非正文 XML 过大，无法可靠核验。')
    return ET.fromstring(z.read(info))


def _logical_children(node):
    """Read one OOXML AlternateContent representation, never both copies."""
    if node.tag == MC + 'AlternateContent':
        choices = [e for e in node if e.tag == MC + 'Choice']
        selected = choices[0] if choices else node.find(MC + 'Fallback')
        return list(selected) if selected is not None else []
    return list(node)


def _logical_nodes(node):
    yield node
    for child in _logical_children(node):
        yield from _logical_nodes(child)


def _word_visible(node, skip_textboxes=False):
    out = []
    fields = []

    def visit(el):
        if skip_textboxes and el.tag == W + 'txbxContent':
            return
        if el.tag == W + 'fldSimple':
            if DYNAMIC.match(el.get(W + 'instr', '')):
                return
        if el.tag == W + 'fldChar':
            action = el.get(W + 'fldCharType')
            if action == 'begin':
                fields.append({'code': '', 'result': False, 'dynamic': False})
            elif action == 'separate':
                if not fields:
                    raise ValueError('DOCX 非正文域分隔符没有起始域。')
                fields[-1]['result'] = True
                fields[-1]['dynamic'] = bool(DYNAMIC.match(fields[-1]['code']))
            elif action == 'end':
                if not fields:
                    raise ValueError('DOCX 非正文域结束符没有起始域。')
                fields.pop()
            return
        if el.tag == W + 'instrText':
            if fields:
                fields[-1]['code'] += el.text or ''
            return
        if el.tag in (W + 't', W + 'delText'):
            if not any(not f['result'] or f['dynamic'] for f in fields):
                out.append(el.text or '')
            return
        if el.tag == W + 'tab':
            if not any(not f['result'] or f['dynamic'] for f in fields):
                out.append('\t')
            return
        for child in _logical_children(el):
            visit(child)

    visit(node)
    if fields:
        raise ValueError('DOCX 非正文域未闭合，无法可靠核验。')
    return _normalized(''.join(out))


def _odt_visible(node, skip_textboxes=False):
    def flatten(el):
        if el.tag.startswith(DC) or el.tag in (TEXT + 'page-number', TEXT + 'page-count',
                TEXT + 'note-citation', TEXT + 'note', OFFICE + 'annotation'):
            return ''
        if skip_textboxes and el.tag == DRAW + 'text-box':
            return ''
        if el.tag == TEXT + 's':
            count = int(el.get(TEXT + 'c', '1'))
            if not 0 <= count <= 1_000_000:
                raise ValueError('ODT 非正文空白长度异常。')
            return ' ' * count
        if el.tag == TEXT + 'tab':
            return '\t'
        if el.tag == TEXT + 'line-break':
            return ''
        result = el.text or ''
        for child in el:
            result += flatten(child) + (child.tail or '')
        return result

    # Ignore XML indentation outside paragraphs while preserving literal spaces
    # inside each text paragraph, including leading/trailing spaces and NBSP.
    out = []
    def paragraphs(el):
        if skip_textboxes and el.tag == DRAW + 'text-box':
            return
        if el.tag in (TEXT + 'p', TEXT + 'h'):
            out.append(flatten(el))
            return
        for child in el:
            paragraphs(child)
    paragraphs(node)
    return _normalized(''.join(out))


def _odt_extras(path):
    result = {kind: [] for kind in KINDS}
    with ZipFile(path) as z:
        content = _read_xml(z, 'content.xml')
        styles = _read_xml(z, 'styles.xml')
        body = content.find(OFFICE + 'body')
        if body is None:
            raise ValueError('ODT 基准缺少正文。')
        for node in body.iter():
            if node.tag == TEXT + 'note':
                kind = node.get(TEXT + 'note-class', 'footnote')
                if kind not in ('footnote', 'endnote'):
                    raise ValueError('ODT 含无法识别的注释类别。')
                note_body = node.find(TEXT + 'note-body')
                if note_body is None:
                    raise ValueError('ODT 脚注或尾注缺少正文。')
                result[kind].append(_odt_visible(note_body))
            elif node.tag == OFFICE + 'annotation':
                result['annotation'].append(_odt_visible(node))
            elif node.tag == DRAW + 'text-box':
                result['textbox'].append(_odt_visible(node))
        header_tags = {STYLE + 'header': 'header', STYLE + 'header-left': 'header',
            STYLE + 'header-first': 'header', STYLE + 'footer': 'footer',
            STYLE + 'footer-left': 'footer', STYLE + 'footer-first': 'footer'}
        for node in styles.iter():
            kind = header_tags.get(node.tag)
            if kind:
                result[kind].append(_odt_visible(node, skip_textboxes=True))
                for box in node.iter(DRAW + 'text-box'):
                    result['header_textbox'].append(_odt_visible(box))
        result['embedded_objects'] = sum(1 for root in (content, styles)
            for node in root.iter() if node.tag in (DRAW + 'object', DRAW + 'object-ole'))
    return result


def _docx_extras(path):
    result = {kind: [] for kind in KINDS}
    result['note_separators'] = dict.fromkeys(SEPARATORS, '')
    with ZipFile(path) as z:
        names = set(z.namelist())
        document = _read_xml(z, 'word/document.xml')
        for node in _logical_nodes(document):
            if node.tag == W + 'txbxContent':
                result['textbox'].append(_word_visible(node))
        for part, kind, tag in [('word/footnotes.xml', 'footnote', 'footnote'),
                ('word/endnotes.xml', 'endnote', 'endnote'),
                ('word/comments.xml', 'annotation', 'comment')]:
            if part in names:
                root = _read_xml(z, part)
                for node in root.findall(W + tag):
                    if node.get(W + 'type') in ('separator', 'continuationSeparator', 'continuationNotice'):
                        result['note_separators'][kind + '_' + node.get(W + 'type')] = _word_visible(node)
                        continue
                    result[kind].append(_word_visible(node))
        payloads = {}
        for kind in ('header', 'footer'):
            pattern = re.compile(r'^word/' + kind + r'\d*\.xml$')
            for part in sorted(name for name in names if pattern.match(name)):
                root = _read_xml(z, part)
                payload = _word_visible(root, skip_textboxes=True)
                payloads[part] = (kind, payload)
                result[kind].append(payload)
                for box in _logical_nodes(root):
                    if box.tag == W + 'txbxContent':
                        result['header_textbox'].append(_word_visible(box))
        result['header_sections'] = _docx_header_sections(z, document, payloads)
        result['embedded_objects'] = len([name for name in names
            if name.startswith('word/embeddings/') and not name.endswith('/')])
    return result


def _on_off(node):
    return node is not None and node.get(W + 'val', 'true').lower() not in ('0', 'false', 'off')


def _docx_header_sections(z, document, payloads):
    """Resolve first/default/even references and links to previous sections."""
    relationships = {}
    names = set(z.namelist())
    if 'word/_rels/document.xml.rels' in names:
        for rel in _read_xml(z, 'word/_rels/document.xml.rels'):
            if rel.get('Type', '').rsplit('/', 1)[-1] not in ('header', 'footer'):
                continue
            if rel.get('TargetMode') == 'External':
                raise ValueError('DOCX 页眉页脚引用了外部文件。')
            target = rel.get('Target', '').replace('\\', '/')
            part = posixpath.normpath(target.lstrip('/') if target.startswith('/')
                                      else posixpath.join('word', target))
            relationships[rel.get('Id')] = part
    facing = False
    if 'word/settings.xml' in names:
        facing = _on_off(_read_xml(z, 'word/settings.xml').find(W + 'evenAndOddHeaders'))
    previous = {kind: dict.fromkeys(('default', 'even', 'first'), '') for kind in ('header', 'footer')}
    sections = []
    body = document.find(W + 'body')
    if body is None:
        raise ValueError('DOCX 缺少正文，无法核验页眉页脚分节。')
    for section in _logical_nodes(body):
        if section.tag != W + 'sectPr':
            continue
        current = {kind: previous[kind].copy() for kind in previous}
        for kind in ('header', 'footer'):
            seen = set()
            for ref in section.findall(W + kind + 'Reference'):
                variant = ref.get(W + 'type')
                if variant not in ('default', 'even', 'first') or variant in seen:
                    raise ValueError('DOCX 页眉页脚分节引用类型无效。')
                seen.add(variant)
                part = relationships.get(ref.get(REL + 'id'))
                if part not in payloads or payloads[part][0] != kind:
                    raise ValueError('DOCX 页眉页脚分节引用缺失。')
                current[kind][variant] = payloads[part][1]
        current['first_page'] = _on_off(section.find(W + 'titlePg'))
        sections.append(current)
        previous = {kind: current[kind] for kind in ('header', 'footer')}
    return {'sections': sections, 'even_and_odd_headers': facing}


def _combined(items):
    return ''.join(items)


def _shared_header_parts(items):
    # Separate first/even/default parts can contain precisely the same whole
    # header/footer. Deduplicate only entire part payloads, never paragraphs
    # or substrings, so literal repeated text within a part remains intact.
    # Package part names/order and copies used by section variants are not
    # content. Compare complete nonempty payloads; never merge or split them.
    return sorted(set(item for item in items if item))


def _binary_header_sections(metadata):
    previous = {kind: dict.fromkeys(('default', 'even', 'first'), '') for kind in ('header', 'footer')}
    sections = []
    for section in metadata['sections']:
        current = {'first_page': section['first_page']}
        for kind in previous:
            current[kind] = {}
            for variant in previous[kind]:
                raw = section[kind][variant]
                # A zero-length story links to the same type in the previous
                # section. An explicit empty paragraph clears that type.
                current[kind][variant] = _binary_visible(raw) if raw else previous[kind][variant]
        sections.append(current)
        previous = {kind: current[kind] for kind in previous}
    return sections


def _effective_headers(sections, facing):
    effective = []
    for section in sections:
        current = {'first_page': section['first_page']}
        for kind in ('header', 'footer'):
            normal = section[kind]['default']
            current[kind] = {'default': normal,
                             'even': section[kind]['even'] if facing else normal,
                             'first': section[kind]['first'] if section['first_page'] else normal}
        effective.append(current)
    return effective


def _payload_report(items):
    return [_fingerprint(text) for text in items]


def _mismatch_message(mismatches):
    labels = {'footnote': '脚注', 'endnote': '尾注', 'annotation': '批注',
              'header': '页眉', 'footer': '页脚', 'textbox': '正文文本框',
              'header_textbox': '页眉页脚文本框', 'header_and_footer': '页眉页脚',
              'header_sections': '页眉页脚的分节及首页、奇偶页设置',
              'note_separators': '脚注尾注分隔内容',
              'embedded_objects_unverifiable': '嵌入对象'}
    return '、'.join(dict.fromkeys(labels.get(name.removeprefix('binary_'), '非正文元素')
                                  for name in sorted(set(mismatches))))


def _fingerprint(text):
    return {'characters': len(text), 'sha256_utf8': hashlib.sha256(text.encode('utf-8')).hexdigest()}


def audit_extra(source_kind, binary_stories, source_odt_path, output_docx_path):
    """Return a report; a nonempty unsupported/mismatched story fails the gate."""
    details = {'source_reader': '独立 Word 二进制故事 + ODT 导入基准' if source_kind == 'doc'
               else 'ODT 原件或转换引擎导入基准；RTF 未独立解析'}
    try:
        imported = _odt_extras(Path(source_odt_path))
        output = _docx_extras(Path(output_docx_path))
        mismatches = []
        checks = {}
        for kind in KINDS:
            src_items, out_items = imported[kind], output[kind]
            if kind in ('header', 'footer'):
                src_items, out_items = _shared_header_parts(src_items), _shared_header_parts(out_items)
            before, after = _combined(src_items), _combined(out_items)
            equal = src_items == out_items if kind in ('header', 'footer') else before == after
            checks[kind] = {'passed': equal, 'source': _fingerprint(before), 'output': _fingerprint(after)}
            if kind in ('header', 'footer'):
                checks[kind]['source_parts'] = _payload_report(src_items)
                checks[kind]['output_parts'] = _payload_report(out_items)
                checks[kind]['source_shared_part_copies'] = len(imported[kind]) - len(src_items)
                checks[kind]['output_shared_part_copies'] = len(output[kind]) - len(out_items)
            if not equal:
                mismatches.append(kind)
        details['import_baseline_checks'] = checks
        details['embedded_objects'] = {'source': imported['embedded_objects'],
                                      'output': output['embedded_objects']}
        if imported['embedded_objects'] or output['embedded_objects']:
            mismatches.append('embedded_objects_unverifiable')

        if source_kind == 'doc':
            if not isinstance(binary_stories, dict):
                raise ValueError('缺少旧 Word 非正文的独立读取基准。')
            required = {'main', 'footnote', 'endnote', 'annotation', 'header', 'textbox', 'header_textbox'}
            if not required.issubset(binary_stories):
                raise ValueError('旧 Word 独立读取基准缺少非正文故事。')
            unknown = [k for k, v in binary_stories.items() if k not in required | {'header_parts'} and v]
            if unknown:
                raise ValueError('旧 Word 含无法核验的非正文故事：' + '、'.join(unknown))
            independent = {}
            for kind in ('footnote', 'endnote', 'annotation', 'textbox', 'header_textbox'):
                before = _binary_visible(binary_stories[kind])
                after = _combined(output[kind])
                equal = before == after
                independent[kind] = {'passed': equal, 'source': _fingerprint(before), 'output': _fingerprint(after)}
                if not equal:
                    mismatches.append('binary_' + kind)
            metadata = binary_stories.get('header_parts')
            if not isinstance(metadata, dict):
                if binary_stories['header']:
                    raise ValueError('缺少旧 Word 页眉页脚分区的独立读取基准。')
                metadata = {'header': [], 'footer': [], 'separators': [''] * 6,
                            'sections': [], 'even_and_odd_headers': False}
            # MS-DOC Header Document starts with six note-separator stories;
            # it is not a concatenation of only header/footer text. Structural
            # 0x03/0x04 are ignored only in their corresponding separator slot.
            separator_checks = {}
            if len(metadata['separators']) != 6:
                raise ValueError('旧 Word 脚注尾注分隔索引不完整。')
            for index, name in enumerate(SEPARATORS):
                text = metadata['separators'][index]
                structural = ('\x03', '\x04', '', '\x03', '\x04', '')[index]
                before = _binary_visible(text.replace(structural, '') if structural else text)
                after = output['note_separators'][name]
                equal = before == after
                separator_checks[name] = {'passed': equal, 'source': _fingerprint(before),
                                          'output': _fingerprint(after)}
                if not equal:
                    mismatches.append('binary_note_separators')
            independent['note_separators'] = separator_checks
            for kind in ('header', 'footer'):
                before = _shared_header_parts([_binary_visible(text) for text in metadata[kind]])
                after = _shared_header_parts(output[kind])
                equal = before == after
                independent[kind] = {'passed': equal, 'source_parts': _payload_report(before),
                                     'output_parts': _payload_report(after)}
                if not equal:
                    mismatches.append('binary_' + kind)
            before_sections = _effective_headers(_binary_header_sections(metadata), metadata['even_and_odd_headers'])
            out_sections = output['header_sections']
            after_sections = _effective_headers(out_sections['sections'], out_sections['even_and_odd_headers'])
            # Documents without a Header Document contain no independent
            # section/header assignments to compare; their empty payload check
            # above still detects any newly introduced header/footer text.
            equal = not before_sections or (before_sections == after_sections and
                metadata['even_and_odd_headers'] == out_sections['even_and_odd_headers'])
            section_checks = []
            for index, (before, after) in enumerate(zip(before_sections, after_sections)):
                for kind in ('header', 'footer'):
                    for variant in ('default', 'even', 'first'):
                        section_checks.append({'section': index + 1, 'kind': kind, 'variant': variant,
                            'passed': before[kind][variant] == after[kind][variant],
                            'source': _fingerprint(before[kind][variant]), 'output': _fingerprint(after[kind][variant])})
            independent['header_sections'] = {'passed': equal, 'source_sections': len(before_sections),
                'output_sections': len(after_sections),
                'source_even_and_odd_headers': metadata['even_and_odd_headers'],
                'output_even_and_odd_headers': out_sections['even_and_odd_headers'],
                'source_first_page_headers': [section['first_page'] for section in before_sections],
                'output_first_page_headers': [section['first_page'] for section in after_sections],
                'checks': section_checks}
            if not equal:
                mismatches.append('binary_header_sections')
            details['independent_binary_checks'] = independent

        details['mismatches'] = sorted(set(mismatches))
        if mismatches:
            return {'passed': False, 'message': '非正文内容无法确认完整保留：' + _mismatch_message(mismatches), 'details': details}
        return {'passed': True, 'message': '脚注、尾注、批注、页眉页脚和文本框的文字核验通过（保留空格；忽略动态页码）。', 'details': details}
    except Exception as error:
        details['error'] = str(error)
        return {'passed': False, 'message': '非正文内容无法可靠核验，停止输出；' + str(error), 'details': details}
