"""Regression cases for modern Word properties and complex package relationships."""
from io import BytesIO
import base64

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
import pytest

from scripts.generate_fixtures import PASSWORD, highlighted, saved
from tests.test_core import simple, text
from wordshield.core import Package, SafetyError, analyze, redact, restore

W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid'
CUSTOM_RELATIONSHIP = 'https://example.org/relationships/review-metadata'


def complex_document(term):
    doc = Document()
    paragraph = doc.add_paragraph()
    highlighted(paragraph, term)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph._p.set(f'{{{W14}}}paraId', '3A4F1020')
    paragraph._p.set(f'{{{W14}}}textId', '77777777')
    properties = paragraph._p.get_or_add_pPr()
    border = OxmlElement('w:top')
    border.set(qn('w:val'), 'single')
    borders = OxmlElement('w:pBdr')
    borders.append(border)
    properties.append(borders)
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), '360')
    spacing.set(qn('w:lineRule'), 'exact')
    properties.append(spacing)
    language = OxmlElement('w:lang')
    for key, value in (('val', 'en-US'), ('eastAsia', 'zh-CN'), ('bidi', 'ar-SA')):
        language.set(qn('w:' + key), value)
    paragraph.runs[0]._r.get_or_add_rPr().append(language)
    doc.sections[0]._sectPr.find(qn('w:docGrid')).set(qn('w:type'), 'lines')
    doc.sections[0].header.paragraphs[0].text = 'Public header'
    doc.sections[0].footer.paragraphs[0].text = 'Public footer'
    table = doc.add_table(rows=2, cols=2)
    table.style = 'Table Grid'
    table.cell(0, 0).text = 'Public table'
    table.cell(1, 1).add_table(rows=1, cols=1).cell(0, 0).text = 'Nested public table'
    return doc


@pytest.mark.parametrize('term', [
    '1', '6', '10', '20', '77', 'id', 'default', 'center', 'single',
    'dxa', 'en-US', 'zh-CN', 'ar-SA', 'lines', 'exact',
])
def test_complex_word_formatting_does_not_become_sensitive_residue(term):
    source = saved(complex_document(term))
    out, mapping, report = redact(analyze(source), PASSWORD)
    restored, restored_report = restore(out, mapping, PASSWORD)
    assert text(restored) == text(source)
    assert report['replacement_count'] == 1
    assert restored_report['status'] == '通过'
    assert len(Document(BytesIO(out)).tables) == 1
    assert Document(BytesIO(out)).paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER


def test_picture_and_hyperlink_relationship_references_are_structural():
    doc = complex_document('rId')
    image = base64.b64decode(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4nGP4DwQACfsD/fteaysAAAAASUVORK5CYII='
    )
    doc.add_picture(BytesIO(image))
    relationship_id = doc.part.relate_to('https://example.org/resource', RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), relationship_id)
    run, node = OxmlElement('w:r'), OxmlElement('w:t')
    node.text = 'Public resource'
    run.append(node)
    hyperlink.append(run)
    doc.add_paragraph()._p.append(hyperlink)
    source = saved(doc)
    out, mapping, _ = redact(analyze(source), PASSWORD, acknowledged=True)
    restored, report = restore(out, mapping, PASSWORD, acknowledged=True)
    assert text(restored) == text(source)
    assert report['placeholder_status'] == '通过'
    original = Package(source)
    exported = Package(out)
    media = [name for name in original.parts if name.startswith('word/media/')]
    assert media and all(exported.parts[name] == original.parts[name] for name in media)


@pytest.mark.parametrize('namespace,attribute', [
    (W14, 'paraId'), (W14, 'textId'), (W15, 'paraIdParent'),
    (W16CID, 'paraId'), (W16CID, 'durableId'),
])
def test_modern_word_identifiers_are_not_content(namespace, attribute):
    doc = Document()
    paragraph = doc.add_paragraph()
    highlighted(paragraph, '77')
    paragraph._p.set(f'{{{namespace}}}{attribute}', '77777777')
    source = saved(doc)
    out, mapping, _ = redact(analyze(source), PASSWORD)
    assert text(restore(out, mapping, PASSWORD)[0]) == text(source)


@pytest.mark.parametrize('location', ['relationship', 'identifier', 'alignment', 'unknown_attribute', 'image_description'])
def test_content_and_invalid_format_values_still_block_residuals(location):
    value = 'FictionalPerson'
    doc = Document()
    paragraph = doc.add_paragraph()
    highlighted(paragraph, value)
    if location == 'relationship':
        paragraph._p.set(qn('r:id'), value)
    elif location == 'identifier':
        paragraph._p.set(f'{{{W14}}}paraId', value)
    elif location in {'alignment', 'unknown_attribute'}:
        alignment = OxmlElement('w:jc')
        alignment.set(qn('w:val' if location == 'alignment' else 'w:private'), value)
        paragraph._p.get_or_add_pPr().append(alignment)
    else:
        drawing = etree.SubElement(paragraph._p, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr')
        drawing.set('descr', value)
    with pytest.raises(SafetyError, match='残留'):
        redact(analyze(saved(doc)), PASSWORD)


@pytest.mark.parametrize('owner', ['document', 'package'])
@pytest.mark.parametrize('target_form', ['relative', 'absolute', 'dot_segments'])
@pytest.mark.parametrize('removed_folder', ['docProps', 'customXml'])
def test_scrub_removes_nonstandard_relationships_to_deleted_parts(owner, target_form, removed_folder):
    doc = Document()
    highlighted(doc.add_paragraph(), 'Fictional Person')
    part_name = f'{removed_folder}/review-extra.xml'
    part = Part(PackURI('/' + part_name), 'application/xml', b'<extra>public</extra>', doc.part.package)
    source_owner = doc.part if owner == 'document' else doc.part.package
    source_owner.relate_to(part, CUSTOM_RELATIONSHIP)
    source_owner.rels.get_or_add_ext_rel(CUSTOM_RELATIONSHIP, '/' + part_name)
    package = Package(saved(doc))
    rels_name = 'word/_rels/document.xml.rels' if owner == 'document' else '_rels/.rels'
    relation = next(rel for rel in package.roots[rels_name]
                    if rel.get('Type') == CUSTOM_RELATIONSHIP and rel.get('TargetMode') != 'External')
    if target_form == 'absolute':
        relation.set('Target', '/' + part_name)
    elif target_form == 'dot_segments':
        target = f'{removed_folder}/./review-extra.xml'
        relation.set('Target', '../' + target if owner == 'document' else target)
    source = package.serialize()
    Package(source)
    out, mapping, _ = redact(analyze(source), PASSWORD, acknowledged=True)
    output = Package(out)
    assert part_name not in output.parts
    remaining = [rel for rel in output.roots[rels_name] if rel.get('Type') == CUSTOM_RELATIONSHIP]
    assert len(remaining) == 1 and remaining[0].get('TargetMode') == 'External'
    assert remaining[0].get('Target') == '/' + part_name
    assert text(restore(out, mapping, PASSWORD, acknowledged=True)[0]) == text(source)


def test_group_rename_includes_unmarked_and_retained_members_only_in_that_group():
    plan = analyze(simple([('林知远', True), ('、林知远、林知远、林知远', False)]))
    original_group = plan.occurrences[0].group
    plan.occurrences[-1].group = 'Other entity'
    plan.occurrences[1].excluded = True
    retained_key = plan.occurrences[1].key
    plan.rename_group(original_group, 'Representative')
    plan.rename_group('Representative', 'Representative')
    assert [item.group for item in plan.occurrences] == ['Representative'] * 3 + ['Other entity']
    assert next(item for item in plan.occurrences if item.key == retained_key).excluded
    _, _, report = redact(plan, PASSWORD, confirmed_exclusions=True)
    assert report['entity_count'] == 2


@pytest.mark.parametrize('new_name', ['', '  ', 'Other entity'])
def test_group_rename_validation_does_not_partially_change_groups(new_name):
    plan = analyze(simple([('林知远', True), ('、', False), ('周映禾', True)]))
    plan.occurrences[-1].group = 'Other entity'
    before = [item.group for item in plan.occurrences]
    with pytest.raises(SafetyError):
        plan.rename_group(before[0], new_name)
    assert [item.group for item in plan.occurrences] == before


def test_rename_rejects_a_stale_group_before_changing_any_item():
    plan = analyze(simple([('林知远', True)]))
    before = [item.group for item in plan.occurrences]
    with pytest.raises(SafetyError, match='已不存在'):
        plan.rename_group('missing group', 'Representative')
    assert [item.group for item in plan.occurrences] == before


def test_numeric_word_language_identifier_remains_supported():
    doc = Document()
    paragraph = doc.add_paragraph()
    highlighted(paragraph, '0409')
    language = OxmlElement('w:lang')
    language.set(qn('w:val'), '0409')
    paragraph.runs[0]._r.get_or_add_rPr().append(language)
    source = saved(doc)
    out, mapping, _ = redact(analyze(source), PASSWORD)
    assert text(restore(out, mapping, PASSWORD)[0]) == text(source)
