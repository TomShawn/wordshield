"""Gemini review 对应的虚构边界回归测试。"""
from io import BytesIO
import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pytest

from scripts.generate_fixtures import PASSWORD, highlighted, replace_parts, saved
from tests.test_core import simple, text
from wordshield.core import Package, SafetyError, _malformed, analyze, redact, restore


@pytest.mark.parametrize('term', ['200', '24', '100', '001', '1', '0', 'main', 'id', 'auto', 'Calibri', 'SECRET', '宋体'])
def test_structural_values_do_not_block_redaction(term):
    plan = analyze(simple([(term, True)]))
    out, encrypted, _ = redact(plan, PASSWORD)
    assert text(out) == '【SECRET_001】'
    assert text(restore(out, encrypted, PASSWORD)[0]) == term


@pytest.mark.parametrize('sentence', ['Board Secretary', 'trade secret', 'Top Secret', 'secretary'])
def test_ordinary_secret_words_are_allowed(sentence):
    source = simple([('林知远', True), (' ' + sentence, False)])
    out, encrypted, _ = redact(analyze(source), PASSWORD)
    result, report = restore(out, encrypted, PASSWORD)
    assert text(result) == text(source)
    assert report['status'] == '通过'
    assert not report['malformed_tokens']


@pytest.mark.parametrize('part', ['textbox', 'comment'])
def test_unique_yellow_outside_supported_scope_blocks(part):
    doc = Document()
    highlighted(doc.add_paragraph(), '林知远')
    paragraph = OxmlElement('w:p')
    run = OxmlElement('w:r')
    properties, highlight = OxmlElement('w:rPr'), OxmlElement('w:highlight')
    highlight.set(qn('w:val'), 'yellow')
    properties.append(highlight)
    run.append(properties)
    node = OxmlElement('w:t')
    node.text = '仅在特殊区域出现的虚构文字'
    run.append(node)
    paragraph.append(run)
    if part == 'textbox':
        box = OxmlElement('w:txbxContent')
        box.append(paragraph)
        doc.add_paragraph().add_run()._r.append(box)
        source = saved(doc)
    else:
        comments = OxmlElement('w:comments')
        comment = OxmlElement('w:comment')
        comment.append(paragraph)
        comments.append(comment)
        from lxml import etree
        source = replace_parts(saved(doc), {'word/comments.xml': etree.tostring(comments)})
    plan = analyze(source)
    assert any(i.code == 'UNSUPPORTED_HIGHLIGHT' and i.blocking for i in plan.issues)
    with pytest.raises(SafetyError):
        redact(plan, PASSWORD, acknowledged=True)


def test_split_preserves_unmarked_group_override():
    source = simple([('林知远', True), ('和林知远。', False), ('周映禾沈清岳', True)])
    plan = analyze(source)
    unmarked = next(o for o in plan.occurrences if not o.highlighted)
    unmarked.group = '另一位同名者'
    item = next(o for o in plan.occurrences if o.text == '周映禾沈清岳')
    plan.split(item.key, [3])
    assert next(o for o in plan.occurrences if o.key == unmarked.key).group == '另一位同名者'
    assert all(re.fullmatch(r'E\d{3,}', o.group) for o in plan.occurrences if o.text in {'周映禾', '沈清岳'})


def test_malformed_offset_is_original_position():
    value = '前【SECRET_001】后【SECRET_ 002】'
    assert _malformed(value)[0]['offset'] == value.index('SECRET_ 002')


def test_unmarked_retention_is_location_scoped_and_reported():
    source = simple([('林知远', True), ('另有林知远', False)])
    plan = analyze(source)
    out, encrypted, report = redact(plan, PASSWORD, include_unmarked=False, confirmed_exclusions=True)
    assert text(out) == '【SECRET_001】另有林知远'
    assert report['excluded_count'] == 1 and report['exclusions_confirmed']
    from wordshield.core import report_bytes
    assert '林知远'.encode() not in report_bytes(report)
    restored, result = restore(out, encrypted, PASSWORD)
    assert text(restored) == text(source) and result['status'] == '通过'


def test_explicit_retention_does_not_hide_attribute_residue():
    source = simple([('林知远', True), ('另有林知远', False)])
    package = Package(source)
    package.roots['word/document.xml'].set('private_note', '林知远')
    with pytest.raises(SafetyError, match='残留'):
        redact(analyze(package.serialize()), PASSWORD, include_unmarked=False, confirmed_exclusions=True)


def test_punctuation_can_be_retained_without_spreading_replacements():
    source = simple([('林知远、周映禾', True), ('。公开甲、公开乙', False)])
    plan = analyze(source)
    plan.split(plan.occurrences[0].key, [3, 4])
    punctuation = next(o for o in plan.occurrences if o.highlighted and o.text == '、')
    plan.exclude(punctuation.key, True)
    assert not any(not o.highlighted and o.text == '、' for o in plan.occurrences)
    with pytest.raises(SafetyError, match='单独确认'):
        redact(plan, PASSWORD)
    out, encrypted, report = redact(plan, PASSWORD, confirmed_exclusions=True)
    assert text(out) == '【SECRET_001】、【SECRET_002】。公开甲、公开乙'
    assert report['entity_count'] == 2 and report['excluded_count'] == 1
    assert text(restore(out, encrypted, PASSWORD)[0]) == text(source)


def test_retained_unmarked_group_survives_split():
    source = simple([('林知远', True), ('另有林知远。', False), ('周映禾沈清岳', True)])
    plan = analyze(source)
    old = next(o for o in plan.occurrences if not o.highlighted)
    plan.exclude(old.key, True)
    item = next(o for o in plan.occurrences if o.text == '周映禾沈清岳')
    plan.split(item.key, [3])
    assert next(o for o in plan.occurrences if o.key == old.key).excluded


def test_actual_body_numeric_and_content_attributes_still_checked():
    from wordshield.core import residual_parts
    source = simple([('200', True), ('金额200', False)])
    assert 'word/document.xml' in residual_parts(Package(source), ['200'])
    package = Package(simple([('公开', False)]))
    paragraph = package.roots['word/document.xml'].find('.//' + qn('w:p'))
    paragraph.set('private_amount', '200')
    assert 'word/document.xml' in residual_parts(package, ['200'])


def test_custom_style_name_remains_part_of_residual_scan():
    from docx.enum.style import WD_STYLE_TYPE
    doc = Document()
    highlighted(doc.add_paragraph(), '林知远')
    doc.styles.add_style('林知远专用样式', WD_STYLE_TYPE.PARAGRAPH)
    with pytest.raises(SafetyError, match='残留'):
        redact(analyze(saved(doc)), PASSWORD)


def test_style_inherited_yellow_in_textbox_blocks():
    doc = Document()
    doc.styles['Normal'].font.highlight_color = 7
    box = OxmlElement('w:txbxContent')
    paragraph, run, node = OxmlElement('w:p'), OxmlElement('w:r'), OxmlElement('w:t')
    node.text = '虚构文本框'
    run.append(node)
    paragraph.append(run)
    box.append(paragraph)
    doc.add_paragraph().add_run()._r.append(box)
    assert any(i.code == 'UNSUPPORTED_HIGHLIGHT' for i in analyze(saved(doc)).issues)


def test_structure_warning_separate_from_placeholder_result():
    out, encrypted, _ = redact(analyze(simple([('林知远', True)])), PASSWORD)
    final = replace_parts(out, {'word/media/image1.png': b'fictional image data'})
    _, report = restore(final, encrypted, PASSWORD, acknowledged=True)
    assert report['placeholder_status'] == '通过'
    assert report['document_review_required']
    assert report['status'] == '回填通过，文档需审查'
    assert report['issues']


def test_streamed_read_checks_actual_limit(monkeypatch):
    import wordshield.core as core
    import zipfile
    source = simple([('林知远', True)])
    original_open = zipfile.ZipFile.open
    class OversizedStream(BytesIO):
        pass
    def forged_open(archive, entry, *args, **kwargs):
        if getattr(entry, 'filename', '') == 'word/document.xml':
            return OversizedStream(b'x' * (core.MAX_PART + 1))
        return original_open(archive, entry, *args, **kwargs)
    monkeypatch.setattr(zipfile.ZipFile, 'open', forged_open)
    with pytest.raises(SafetyError, match='实际解压'):
        Package(source)


def test_streamed_read_checks_declared_length(monkeypatch):
    import zipfile
    source = simple([('林知远', True)])
    original_open = zipfile.ZipFile.open
    def forged_open(archive, entry, *args, **kwargs):
        if getattr(entry, 'filename', '') == 'word/document.xml':
            return BytesIO(b'fictional short data')
        return original_open(archive, entry, *args, **kwargs)
    monkeypatch.setattr(zipfile.ZipFile, 'open', forged_open)
    with pytest.raises(SafetyError, match='声明长度'):
        Package(source)
