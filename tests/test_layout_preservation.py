#!/usr/bin/env python3
"""Regression checks for exact, namespace-aware character-grid repair."""
import io
from pathlib import Path
import sys
import unittest
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile, ZipInfo

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'engine'))
from layout_preservation import repair_exported_layout

WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
STRICT_WORD = 'http://purl.oclc.org/ooxml/wordprocessingml/main'


def package(document):
    """Use nontrivial ZIP metadata and opaque sibling parts to catch broad edits."""
    out = io.BytesIO()
    parts = [
        ('[Content_Types].xml', b'<Types/>', ZIP_DEFLATED),
        ('word/document.xml', document, ZIP_DEFLATED),
        ('word/styles.xml', b'<?xml version="1.0"?><styles><!-- 4294966271 --></styles>', ZIP_STORED),
        ('word/header1.xml', b'<w:hdr xmlns:w="' + WORD.encode() +
         b'"><w:docGrid w:charSpace="4294966271"/></w:hdr>', ZIP_DEFLATED),
        ('word/media/image1.png', bytes(range(256)), ZIP_STORED),
    ]
    with ZipFile(out, 'w') as archive:
        archive.comment = b'archive metadata must survive'
        for index, (name, content, compression) in enumerate(parts):
            info = ZipInfo(name, (2024, 2, 3, 4, 5, 6))
            info.compress_type = compression
            info.external_attr = (0o100640 << 16)
            info.comment = ('part-%d' % index).encode()
            archive.writestr(info, content)
    return out.getvalue()


def document(grid, prefix='w', namespace=WORD):
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<{p}:document xmlns:{p}="{ns}"><{p}:body>'
            '<{p}:p><{p}:r><{p}:t>原文 4294966271</{p}:t></{p}:r></{p}:p>'
            '<{p}:sectPr>{grid}</{p}:sectPr></{p}:body></{p}:document>'
            .format(p=prefix, ns=namespace, grid=grid)).encode('utf-8')


class LayoutPreservationTests(unittest.TestCase):
    def assert_package_unchanged(self, xml):
        original = package(xml)
        repaired, changes = repair_exported_layout(original)
        self.assertEqual(changes, [])
        self.assertEqual(repaired, original)

    def test_valid_signed_values_and_absent_grid_preserve_entire_package(self):
        for value in ('0', '1025', '+1025', '2147483647', '-1025', '-2147483648'):
            with self.subTest(value=value):
                self.assert_package_unchanged(document('<w:docGrid w:charSpace="%s"/>' % value))
        self.assert_package_unchanged(document(''))
        self.assert_package_unchanged(document('<w:docGrid w:linePitch="579"/>'))

    def test_repairs_only_target_attribute_and_preserves_other_bytes(self):
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<w:document xmlns:w="' + WORD + '" '
               'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
               'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
               'mc:Ignorable="w14">\n<w:body><!-- 4294966271 -->'
               '<w:p w14:paraId="1234"><w:r><w:t>原文 4294966271</w:t></w:r></w:p>'
               '<mc:AlternateContent><mc:Choice Requires="w14"><w:p/></mc:Choice>'
               '<mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent>'
               '<w:sectPr><w:docGrid w:linePitch="579" w:charSpace = "4294966271"/>'
               '</w:sectPr></w:body></w:document>').encode('utf-8')
        original = package(xml)
        repaired, changes = repair_exported_layout(original)
        self.assertEqual(changes, [{'kind': 'signed_character_grid', 'sections': 1}])
        with ZipFile(io.BytesIO(original)) as before, ZipFile(io.BytesIO(repaired)) as after:
            self.assertEqual(before.namelist(), after.namelist())
            self.assertEqual(before.comment, after.comment)
            changed_parts = []
            for info in before.infolist():
                copied = after.getinfo(info.filename)
                for field in ('date_time', 'compress_type', 'external_attr', 'comment'):
                    self.assertEqual(getattr(info, field), getattr(copied, field))
                if before.read(info.filename) != after.read(info.filename):
                    changed_parts.append(info.filename)
            self.assertEqual(changed_parts, ['word/document.xml'])
            self.assertEqual(after.read('word/document.xml'), xml.replace(
                b'w:charSpace = "4294966271"', b'w:charSpace = "-1025"'))

    def test_prefix_namespace_and_quote_variants(self):
        for namespace in (WORD, STRICT_WORD):
            for prefix in ('word', 'otherPrefix'):
                for quote in ('"', "'"):
                    with self.subTest(namespace=namespace, prefix=prefix, quote=quote):
                        grid = '<{p}:docGrid {p}:charSpace = {q}4294967295{q}/>'.format(
                            p=prefix, q=quote)
                        xml = document(grid, prefix, namespace)
                        repaired, changes = repair_exported_layout(package(xml))
                        with ZipFile(io.BytesIO(repaired)) as archive:
                            self.assertEqual(archive.read('word/document.xml'),
                                             xml.replace(b'4294967295', b'-1'))
                        self.assertEqual(changes[0]['sections'], 1)

    def test_default_element_namespace_and_separate_attribute_prefix(self):
        xml = ('<document xmlns="' + WORD + '" xmlns:attribute="' + WORD + '">'
               '<body><sectPr><docGrid attribute:charSpace="2147483648"/>'
               '</sectPr></body></document>').encode()
        repaired, changes = repair_exported_layout(package(xml))
        with ZipFile(io.BytesIO(repaired)) as archive:
            self.assertEqual(archive.read('word/document.xml'),
                             xml.replace(b'2147483648', b'-2147483648'))
        self.assertEqual(changes[0]['sections'], 1)

    def test_other_namespaces_attributes_and_elements_are_untouched(self):
        self.assert_package_unchanged(document(
            '<w:docGrid w:charSpace="4294966271"/>', namespace='urn:unrelated'))
        for grid in (
            '<w:docGrid xmlns:x="urn:unrelated" x:charSpace="4294966271"/>',
            '<w:docGrid charSpace="4294966271"/>',
            '<w:p w:charSpace="4294966271"/>',
            '<x:docGrid xmlns:x="urn:unrelated" w:charSpace="4294966271"/>',
        ):
            with self.subTest(grid=grid):
                self.assert_package_unchanged(document(grid))

    def test_multiple_sections_and_idempotence(self):
        xml = document('<w:docGrid w:charSpace="4294966271"/>'
                       '<w:docGrid w:charSpace="7"/>'
                       '<w:docGrid w:charSpace=\'2147483648\'/>')
        repaired, changes = repair_exported_layout(package(xml))
        self.assertEqual(changes, [{'kind': 'signed_character_grid', 'sections': 2}])
        with ZipFile(io.BytesIO(repaired)) as archive:
            self.assertEqual(archive.read('word/document.xml'),
                             xml.replace(b'charSpace="4294966271"', b'charSpace="-1025"')
                             .replace(b"charSpace='2147483648'", b"charSpace='-2147483648'"))
        repeated, repeated_changes = repair_exported_layout(repaired)
        self.assertEqual(repeated, repaired)
        self.assertEqual(repeated_changes, [])

    def test_out_of_range_and_non_integer_values_are_rejected(self):
        for value in ('4294967296', '-2147483649', '99999999999999999999999999999',
                      '1.5', 'NaN', ''):
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    repair_exported_layout(package(document(
                        '<w:docGrid w:charSpace="%s"/>' % value)))


if __name__ == '__main__':
    unittest.main()
