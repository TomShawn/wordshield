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


@pytest.mark.parametrize('sentence', [
    'Board Secretary', 'trade secret', 'Top Secret', 'secretary',
    'See the clause [See 5] for details.', 'Refer to [Section 12].',
    'Classification: TOP SECRET 2024', 'Keep it secret 10 days.',
])
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


def test_trailing_space_highlight_catches_unmarked_and_attribute_residue():
    doc = Document()
    first = doc.add_paragraph()
    highlighted(first, '林知远 ')
    first.add_run('签署合同。')
    doc.add_paragraph('另一位当事人为林知远，住址略。')
    source = saved(doc)
    plan = analyze(source)
    assert [(o.text, o.highlighted) for o in plan.occurrences] == [('林知远', True), ('林知远', False)]
    out, encrypted, _ = redact(plan, PASSWORD)
    assert '林知远' not in text(out)
    assert text(restore(out, encrypted, PASSWORD)[0]) == text(source)
    package = Package(simple([('林知远 ', True)]))
    package.roots['word/document.xml'].set('private_note', '林知远')
    with pytest.raises(SafetyError, match='残留'):
        redact(analyze(package.serialize()), PASSWORD)


def test_whitespace_only_highlight_does_not_spread_to_all_spaces():
    source = simple([('林知远', True), ('与', False), (' ', True), ('周映禾 在 云岚市 会谈', False)])
    plan = analyze(source)
    assert [o.text for o in plan.occurrences] == ['林知远']


def test_scrub_clears_revision_author_and_printer_settings():
    doc = Document()
    highlighted(doc.add_paragraph(), '林知远')
    insertion = OxmlElement('w:ins')
    insertion.set(qn('w:id'), '1')
    insertion.set(qn('w:author'), 'Fictional Reviewer')
    insertion.set(qn('w:initials'), 'FR')
    insertion.set(qn('w:date'), '2026-11-11T00:00:00Z')
    run, node = OxmlElement('w:r'), OxmlElement('w:t')
    node.text = '公开修订内容'
    run.append(node)
    insertion.append(run)
    doc.add_paragraph()._p.append(insertion)
    source = replace_parts(saved(doc), {'word/printerSettings/printerSettings1.bin': b'\x00' * 32})
    plan = analyze(source)
    assert not any(i.blocking for i in plan.issues)
    out, encrypted, _ = redact(plan, PASSWORD, acknowledged=True)
    exported = Package(out)
    assert 'word/printerSettings/printerSettings1.bin' not in exported.parts
    xml = exported.parts['word/document.xml'].decode('utf-8')
    assert 'Fictional Reviewer' not in xml and '2026-11-11' not in xml
    restored, report = restore(out, encrypted, PASSWORD, acknowledged=True)
    assert report['placeholder_status'] == '通过'
    assert '林知远' in text(restored)


def test_cross_paragraph_text_does_not_false_positive_residue():
    doc = Document()
    highlighted(doc.add_paragraph(), '林知远')
    doc.add_paragraph('第一段末尾是格林')
    doc.add_paragraph('知远是第二段开头')
    source = saved(doc)
    out, encrypted, _ = redact(analyze(source), PASSWORD)
    assert text(restore(out, encrypted, PASSWORD)[0]) == text(source)


def test_assign_group_rejects_conflicting_text_and_empty_name():
    plan = analyze(simple([('林知远', True), ('、', False), ('周映禾', True)]))
    with pytest.raises(SafetyError, match='不能为空'):
        plan.assign_group(plan.occurrences[0].key, '   ')
    with pytest.raises(SafetyError, match='不同文字'):
        plan.assign_group(plan.occurrences[-1].key, plan.occurrences[0].group)
    plan.assign_group(plan.occurrences[0].key, '自定义分组')
    assert plan.occurrences[0].group == '自定义分组'


def test_export_restored_detects_changed_source_file(tmp_path):
    from hashlib import sha256
    from wordshield.core import export_restored
    source_data = simple([('林知远', True)])
    redacted_bytes, encrypted_bytes, _ = redact(analyze(source_data), PASSWORD)
    final_path = tmp_path / 'final.docx'
    map_path = tmp_path / 'final.wsmap'
    out_path = tmp_path / 'restored.docx'
    report_path = tmp_path / 'restored.json'
    final_path.write_bytes(redacted_bytes)
    map_path.write_bytes(encrypted_bytes)
    original_hash = sha256(redacted_bytes).hexdigest()
    final_path.write_bytes(simple([('其他内容', False)]))
    with pytest.raises(SafetyError, match='发生变化'):
        export_restored(
            final_path, map_path, out_path, report_path, PASSWORD,
            expected_source_hash=original_hash,
        )
    assert not out_path.exists()


def test_surrounding_whitespace_does_not_split_one_entity_into_two_placeholders():
    doc = Document()
    for before, value, after in (('签署人 ', 'John Smith ', 'today.'), ('见证人 ', 'John Smith', '.')):
        paragraph = doc.add_paragraph(before)
        highlighted(paragraph, value)
        paragraph.add_run(after)
    doc.add_paragraph('John Smith agreed.')
    source = saved(doc)
    plan = analyze(source)
    assert {o.text for o in plan.occurrences} == {'John Smith'}
    assert len({o.group for o in plan.occurrences}) == 1
    out, encrypted, report = redact(plan, PASSWORD)
    assert report['entity_count'] == 1
    assert '【SECRET_001】 today.' in text(out)
    assert text(restore(out, encrypted, PASSWORD)[0]) == text(source)


def test_split_drops_whitespace_only_parts_and_trims_edges():
    plan = analyze(simple([('林知远 周映禾', True), ('与周映禾', False)]))
    plan.split(plan.occurrences[0].key, [3, 4])
    highlighted_items = [o for o in plan.occurrences if o.highlighted]
    assert [o.text for o in highlighted_items] == ['林知远', '周映禾']
    plan = analyze(simple([('林知远 周映禾', True), ('与周映禾', False)]))
    plan.split(plan.occurrences[0].key, [3])
    second = next(o for o in plan.occurrences if o.highlighted and o.text == '周映禾')
    unmarked = next(o for o in plan.occurrences if not o.highlighted)
    assert second.start == 4 and unmarked.group == second.group
    out, encrypted, _ = redact(plan, PASSWORD)
    assert text(out) == '【SECRET_001】 【SECRET_002】与【SECRET_002】'


def _header_with_positional_tab():
    doc = Document()
    paragraph = doc.sections[0].header.paragraphs[0]
    highlighted(paragraph, '云岚科技')
    tab = OxmlElement('w:ptab')
    for name, value in (('relativeTo', 'margin'), ('alignment', 'right'), ('leader', 'none')):
        tab.set(qn('w:' + name), value)
    run = OxmlElement('w:r')
    run.append(tab)
    paragraph._p.append(run)
    highlighted(paragraph, '林知远')
    doc.add_paragraph('正文')
    return saved(doc)


def test_positional_tab_is_a_boundary_and_layout_survives_restore():
    from lxml import etree
    source = _header_with_positional_tab()
    plan = analyze(source)
    assert [o.text for o in plan.occurrences] == ['云岚科技', '林知远']
    out, encrypted, _ = redact(plan, PASSWORD)
    restored, report = restore(out, encrypted, PASSWORD)
    assert report['status'] == '通过'

    def sequence(data):
        root = Package(data).roots['word/header1.xml']
        return [('PTAB' if etree.QName(n).localname == 'ptab' else n.text)
                for n in root.iter() if etree.QName(n).localname == 'ptab'
                or (etree.QName(n).localname == 't' and n.text)]
    assert sequence(restored) == sequence(source) == ['云岚科技', 'PTAB', '林知远']


def _png(width=48, height=48, seed=7):
    import random, struct, zlib
    rnd = random.Random(seed)
    raw = b''.join(b'\x00' + bytes(rnd.randrange(256) for _ in range(width * 3)) for _ in range(height))

    def chunk(kind, data):
        body = kind + data
        return struct.pack('>I', len(data)) + body + struct.pack('>I', zlib.crc32(body) & 0xffffffff)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(raw, 0)) + chunk(b'IEND', b''))


@pytest.mark.parametrize('term, in_image_bytes', [
    # 2006 来自图片命名空间 URI，2834 来自 3.1 英寸图片的 EMU 尺寸。
    ('2006', False), ('2834', False),
    # 两字节短词附在图片数据尾部，确定性地模拟随机命中。
    ('35', True), ('42', True), ('17', True),
])
def test_picture_layout_and_image_bytes_do_not_block_short_terms(term, in_image_bytes):
    from docx.shared import Inches
    doc = Document()
    paragraph = doc.add_paragraph('数值 ')
    highlighted(paragraph, term)
    image = _png() + (term.encode('ascii') if in_image_bytes else b'')
    doc.add_picture(BytesIO(image), width=Inches(3.1))
    source = saved(doc)
    out, encrypted, _ = redact(analyze(source), PASSWORD, acknowledged=True)
    assert text(restore(out, encrypted, PASSWORD, acknowledged=True)[0]) == text(source)


def test_picture_alt_text_and_long_binary_matches_are_still_checked():
    from docx.shared import Inches
    from wordshield.core import residual_parts
    doc = Document()
    highlighted(doc.add_paragraph(), '林知远')
    doc.add_picture(BytesIO(_png()), width=Inches(1))
    package = Package(saved(doc))
    doc_pr = package.roots['word/document.xml'].find('.//{*}docPr')
    doc_pr.set('descr', '林知远的照片')
    assert 'word/document.xml' in residual_parts(package, ['林知远'])
    media = next(name for name in package.parts if name.startswith('word/media/'))
    package.parts[media] += 'Lin Zhiyuan'.encode('utf-16le')
    assert media in residual_parts(package, ['Lin Zhiyuan'])
