from collections import Counter
from io import BytesIO
from pathlib import Path
import re
import zipfile

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
import pytest

from scripts.generate_fixtures import PASSWORD, highlighted, make_all, replace_parts, saved
from wordshield.core import (Package, SafetyError, TOKEN, analyze, export_redacted,
                             redact, residual_parts, restore)
from wordshield.crypto import MappingError, decrypt_mapping, encrypt_mapping
from wordshield.storage import OutputError, write_outputs


@pytest.fixture(scope="session")
def samples(tmp_path_factory):
    return make_all(tmp_path_factory.mktemp("fictional"))


def text(data):
    return "\n".join(p.text for p in Package(data).paragraphs())


def simple(segments):
    doc = Document()
    p = doc.add_paragraph()
    for value, marked in segments:
        run = p.add_run(value)
        if marked:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return saved(doc)


def test_contiguous_runs_spaces_and_partial():
    data = simple([("公开前", False), ("  林", True), ("知", True), ("远  ", True), ("公开后", False)])
    plan = analyze(data)
    assert [o.text for o in plan.occurrences] == ["林知远"]
    out, encrypted, _ = redact(plan, PASSWORD)
    assert "公开前  【SECRET_001】  公开后" in text(out)
    result, report = restore(out, encrypted, PASSWORD)
    assert text(result) == text(data)
    assert report["status"] == "通过"


def test_repeat_headers_footers_tables_metadata(samples):
    source = samples["01_复杂格式原稿.docx"]
    plan = analyze(source)
    assert {"word/document.xml", "word/header1.xml", "word/footer1.xml"} <= {o.part for o in plan.occurrences}
    assert any(not o.highlighted for o in plan.occurrences)
    out, encrypted, report = redact(plan, PASSWORD)
    package = Package(out)
    assert not any(n.startswith("docProps/") for n in package.parts)
    assert not residual_parts(package, list({o.text for o in plan.occurrences}))
    mapping = decrypt_mapping(encrypted, PASSWORD)
    assert sum(e["count"] for e in mapping["entries"].values()) == report["replacement_count"]
    counts = Counter(m.group() for p in package.paragraphs() for m in TOKEN.finditer(p.text))
    assert counts["【SECRET_001】"] >= 4
    original_doc, out_doc = Document(BytesIO(source)), Document(BytesIO(out))
    assert len(out_doc.tables) == len(original_doc.tables)
    assert out_doc.sections[0].page_width == original_doc.sections[0].page_width
    assert out_doc.paragraphs[2].runs[1].bold
    restored, report = restore(out, encrypted, PASSWORD)
    assert text(restored) == text(source)
    assert report["status"] == "通过"


def test_no_yellow(samples):
    plan = analyze(samples["03_无高亮.docx"])
    assert not plan.occurrences
    with pytest.raises(SafetyError, match="未发现"):
        redact(plan, PASSWORD)


def test_boundaries_and_adjacent_split(samples):
    plan = analyze(samples["02_相邻与边界.docx"])
    item = next(o for o in plan.occurrences if o.text == "林知远周映禾")
    plan.split(item.key, [3])
    assert not any(o.text == "林知远周映禾" for o in plan.occurrences)
    assert sum(o.text == "林知远" and o.highlighted for o in plan.occurrences) == 2
    assert {"澄溪区", "云岚市", "沈清岳"} <= {o.text for o in plan.occurrences}
    out, encrypted, _ = redact(plan, PASSWORD)
    assert text(restore(out, encrypted, PASSWORD)[0]) == text(plan.package.serialize())


def test_overlap_longest_match(samples):
    plan = analyze(samples["04_包含与同名.docx"])
    unmarked = [o.text for o in plan.occurrences if not o.highlighted]
    assert unmarked == ["星澜科技研究院", "星澜科技"]
    out, encrypted, _ = redact(plan, PASSWORD)
    assert text(restore(out, encrypted, PASSWORD)[0]) == text(samples["04_包含与同名.docx"])


def test_homonyms_separate_groups(samples):
    plan = analyze(samples["04_包含与同名.docx"])
    names = [o for o in plan.occurrences if o.text == "林知远"]
    names[-1].group = "另一人"
    out, encrypted, _ = redact(plan, PASSWORD)
    mapping = decrypt_mapping(encrypted, PASSWORD)
    assert sum(e["text"] == "林知远" for e in mapping["entries"].values()) == 2
    assert text(restore(out, encrypted, PASSWORD)[0]) == text(samples["04_包含与同名.docx"])


def test_different_text_cannot_share_group():
    plan = analyze(simple([("林知远", True), ("与", False), ("周映禾", True)]))
    plan.occurrences[-1].group = plan.occurrences[0].group
    with pytest.raises(SafetyError, match="不同文字"):
        redact(plan, PASSWORD)


def test_unmarked_retention_requires_confirmation():
    plan = analyze(simple([("林知远", True), ("另有林知远", False)]))
    with pytest.raises(SafetyError, match="单独确认"):
        redact(plan, PASSWORD, include_unmarked=False)


def test_risks_and_hidden_residue(samples):
    plan = analyze(samples["05_隐藏修订文本框批注.docx"])
    assert {"HIDDEN", "REVISION", "TEXTBOX", "COMMENTS"} <= {i.code for i in plan.issues}
    with pytest.raises(SafetyError, match="明确确认"):
        redact(plan, PASSWORD)
    with pytest.raises(SafetyError, match="残留"):
        redact(plan, PASSWORD, acknowledged=True)


def test_split_run_restore_and_ai_anomalies(samples):
    _, encrypted, _ = redact(analyze(samples["01_复杂格式原稿.docx"]), PASSWORD)
    good, report = restore(samples["07_AI正常拆run定稿.docx"], encrypted, PASSWORD)
    assert report["status"] == "通过"
    assert text(good) == text(samples["01_复杂格式原稿.docx"])
    bad, report = restore(samples["08_AI删除重复改写定稿.docx"], encrypted, PASSWORD)
    assert "【SECRET_999】" in report["unknown_tokens"]
    assert report["count_changes"]
    assert "【SECRET_003】" in report["missing_tokens"]
    assert report["malformed_tokens"]
    assert report["remaining_tokens"]
    assert "【SECRET_ 003】" in text(bad)
    assert report["status"] == "需人工检查"


def test_deleted_placeholder_no_guess():
    data = simple([("林知远", True), ("与", False), ("周映禾", True)])
    out, encrypted, _ = redact(analyze(data), PASSWORD)
    package = Package(out)
    for root in package.roots.values():
        for node in root.iter(qn("w:t")):
            node.text = (node.text or "").replace("【SECRET_001】", "")
    result, report = restore(package.serialize(), encrypted, PASSWORD)
    assert "【SECRET_001】" in report["missing_tokens"]
    assert "林知远" not in text(result)
    assert "周映禾" in text(result)


@pytest.mark.parametrize("variant", ["【secret_001】", "【SECRET_ 001】", "【S E C R E T_001】", "SECRET_001", "【SECRET_01】"])
def test_modified_placeholder_report(variant):
    data = simple([("林知远", True)])
    out, encrypted, _ = redact(analyze(data), PASSWORD)
    package = Package(out)
    for root in package.roots.values():
        for node in root.iter(qn("w:t")):
            node.text = (node.text or "").replace("【SECRET_001】", variant)
    _, report = restore(package.serialize(), encrypted, PASSWORD)
    assert report["malformed_tokens"]
    assert report["missing_tokens"]


def test_restore_keeps_ai_format():
    out, encrypted, _ = redact(analyze(simple([("林知远", True)])), PASSWORD)
    doc = Document(BytesIO(out))
    run = doc.paragraphs[0].runs[0]
    run.bold, run.italic = True, True
    result, _ = restore(saved(doc), encrypted, PASSWORD)
    run = Document(BytesIO(result)).paragraphs[0].runs[0]
    assert run.bold and run.italic and run.text == "林知远"


def test_restore_uses_format_of_run_covering_most_of_placeholder():
    # Word 保存时把【、SECRET_001、】拆成三个 run；双击只选中中间部分加粗。
    out, encrypted, _ = redact(analyze(simple([("签约方", False), ("林知远", True), ("。", False)])), PASSWORD)
    doc = Document(BytesIO(out))
    paragraph = doc.paragraphs[0]
    token_run = next(run for run in paragraph.runs if "【SECRET_001】" in run.text)
    before, after = token_run.text.split("【SECRET_001】")
    token_run.text = before + "【"
    middle = OxmlElement("w:r")
    middle.append(OxmlElement("w:rPr"))
    middle[0].append(OxmlElement("w:b"))
    middle_text = OxmlElement("w:t")
    middle_text.text = "SECRET_001"
    middle.append(middle_text)
    closing = OxmlElement("w:r")
    closing_text = OxmlElement("w:t")
    closing_text.text = "】" + after
    closing.append(closing_text)
    token_run._r.addnext(middle)
    middle.addnext(closing)
    result, report = restore(saved(doc), encrypted, PASSWORD)
    assert report["status"] == "通过"
    runs = [run for run in Document(BytesIO(result)).paragraphs[0].runs if run.text]
    assert "".join(run.text for run in runs) == "签约方林知远。"
    assert [run.text for run in runs if run.bold] == ["林知远"]


def test_password_and_tampering():
    _, encrypted, _ = redact(analyze(simple([("林知远", True)])), PASSWORD)
    assert "林知远".encode() not in encrypted
    assert encrypt_mapping({"version": 1}, PASSWORD) != encrypt_mapping({"version": 1}, PASSWORD)
    with pytest.raises(MappingError):
        decrypt_mapping(encrypted, "wrong")
    with pytest.raises(MappingError):
        decrypt_mapping(encrypted[:-1] + bytes([encrypted[-1] ^ 1]), PASSWORD)
    with pytest.raises(MappingError):
        encrypt_mapping({"version": 1}, "short")


def test_invalid_mapping_schema():
    out, _, _ = redact(analyze(simple([("林知远", True)])), PASSWORD)
    invalid = encrypt_mapping({"version": 1, "entries": {"bad": {"text": "fictional", "count": 1}}}, PASSWORD)
    with pytest.raises(MappingError):
        restore(out, invalid, PASSWORD)


def test_package_corrupt_duplicate_and_xml_entity():
    with pytest.raises(SafetyError):
        Package(b"not a zip")
    good = simple([("林知远", True)])
    source = zipfile.ZipFile(BytesIO(good))
    malformed = replace_parts(good, {"word/document.xml": b"<broken"})
    with pytest.raises(SafetyError, match="XML"):
        Package(malformed)
    malicious = replace_parts(good, {"word/document.xml": b'<!DOCTYPE a [<!ENTITY x SYSTEM "file:///etc/passwd">]><a>&x;</a>'})
    with pytest.raises(SafetyError, match="实体"):
        Package(malicious)
    duplicate = BytesIO()
    with zipfile.ZipFile(duplicate, "w") as archive:
        archive.writestr("word/document.xml", source.read("word/document.xml"))
        with pytest.warns(UserWarning):
            archive.writestr("word/document.xml", source.read("word/document.xml"))
    with pytest.raises(SafetyError, match="重复"):
        Package(duplicate.getvalue())


def test_attributes_comments_and_paths_residual():
    data = simple([("林知远", True)])
    package = Package(data)
    p = package.roots["word/document.xml"].find(".//" + qn("w:p"))
    p.set("test", "林知远")
    plan = analyze(package.serialize())
    with pytest.raises(SafetyError, match="残留"):
        redact(plan, PASSWORD)


def test_no_overwrite_and_permissions(tmp_path):
    target = tmp_path / "output.docx"
    write_outputs({target: b"test"})
    assert target.stat().st_mode & 0o777 == 0o600
    with pytest.raises(OutputError):
        write_outputs({target: b"overwrite"})
    assert target.read_bytes() == b"test"
    with pytest.raises(OutputError):
        write_outputs({tmp_path / "new": b"x"}, (tmp_path / "new",))


def test_source_changed_and_retry(tmp_path):
    data = simple([("林知远", True)])
    source = tmp_path / "source.docx"
    source.write_bytes(data)
    plan = analyze(data)
    first, _, _ = redact(plan, PASSWORD)
    second, _, _ = redact(plan, PASSWORD)
    assert first == second
    source.write_bytes(simple([("其他", True)]))
    with pytest.raises(SafetyError, match="发生变化"):
        export_redacted(source, tmp_path / "o.docx", tmp_path / "m.wsmap", tmp_path / "r.json", plan, PASSWORD)
    assert not (tmp_path / "o.docx").exists()


def test_styles_inheritance_and_direct_none():
    doc = Document()
    doc.styles["Normal"].font.highlight_color = WD_COLOR_INDEX.YELLOW
    p = doc.add_paragraph()
    p.add_run("林知远")
    p.add_run("公开").font.highlight_color = WD_COLOR_INDEX.AUTO
    plan = analyze(saved(doc))
    assert [o.text for o in plan.occurrences] == ["林知远"]
    out, _, _ = redact(plan, PASSWORD)
    assert "公开" in text(out)


def test_footnotes_endnotes():
    data = simple([("林知远", True)])
    extra = ('<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
             '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr>'
             '<w:t>周映禾</w:t></w:r></w:p></w:footnote></w:footnotes>').encode()
    data = replace_parts(data, {"word/footnotes.xml": extra})
    plan = analyze(data)
    assert "周映禾" in [o.text for o in plan.occurrences]
    out, encrypted, _ = redact(plan, PASSWORD)
    assert "周映禾" not in text(out)
    assert "周映禾" in text(restore(out, encrypted, PASSWORD)[0])


def test_reserved_source_placeholder():
    plan = analyze(simple([("林知远", True), ("【SECRET_001】", False)]))
    with pytest.raises(SafetyError, match="必须修复"):
        redact(plan, PASSWORD)


def test_binary_and_external_relationship_acknowledgment():
    data = simple([("林知远", True)])
    data = replace_parts(data, {"word/media/fictional.png": b"fictional-image"})
    plan = analyze(data)
    assert "BINARY" in {i.code for i in plan.issues}
    with pytest.raises(SafetyError):
        redact(plan, PASSWORD)
    out, _, report = redact(plan, PASSWORD, acknowledged=True)
    assert report["risk_acknowledged"] and Package(out)


def test_refuse_active_content():
    data = replace_parts(simple([("林知远", True)]), {"word/embeddings/data.bin": b"fictional"})
    with pytest.raises(SafetyError, match="必须修复"):
        redact(analyze(data), PASSWORD, acknowledged=True)


def test_no_network_imports():
    package_dir = Path(__file__).resolve().parents[1] / "wordshield"
    source = "\n".join(p.read_text() for p in package_dir.glob("*.py"))
    assert not re.search(r"^(?:from|import) (?:requests|httpx|socket|urllib|openai)\b", source, re.M)


def test_restore_placeholder_split_by_tab_and_paragraph():
    out, encrypted, _ = redact(analyze(simple([("林知远", True)])), PASSWORD)
    doc = Document(BytesIO(out))
    doc.paragraphs[0].runs[0].text = "【SEC\tRET_001】"
    _, report = restore(saved(doc), encrypted, PASSWORD)
    assert report["malformed_tokens"] and report["missing_tokens"]
    doc.paragraphs[0].runs[0].text = "【SEC"
    doc.add_paragraph("RET_001】")
    _, report = restore(saved(doc), encrypted, PASSWORD)
    assert any("跨段落" in item["kind"] for item in report["malformed_tokens"])


def test_zero_width_and_typo_placeholder():
    out, encrypted, _ = redact(analyze(simple([("林知远", True)])), PASSWORD)
    for value in ("【SEC\u200bRET_001】", "【SECREX_001】"):
        doc = Document(BytesIO(out))
        doc.paragraphs[0].runs[0].text = value
        _, report = restore(saved(doc), encrypted, PASSWORD)
        assert report["malformed_tokens"]


def test_table_style_yellow_blocks():
    doc = Document()
    doc.styles["Table Grid"].font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.add_table(rows=1, cols=1).style = "Table Grid"
    highlighted(doc.add_paragraph(), "林知远")
    plan = analyze(saved(doc))
    assert any(i.code == "TABLE_HIGHLIGHT" and i.blocking for i in plan.issues)
    with pytest.raises(SafetyError):
        redact(plan, PASSWORD, acknowledged=True)


def test_external_relationship_requires_ack():
    from docx.opc.constants import RELATIONSHIP_TYPE
    doc = Document()
    highlighted(doc.add_paragraph(), "林知远")
    doc.part.relate_to("https://example.invalid/fictional", RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    plan = analyze(saved(doc))
    assert "EXTERNAL" in {i.code for i in plan.issues}
    with pytest.raises(SafetyError):
        redact(plan, PASSWORD)
    redact(plan, PASSWORD, acknowledged=True)


def test_custom_xml_scrub_and_restore_risk():
    data = simple([("林知远", True)])
    out, encrypted, _ = redact(analyze(data), PASSWORD)
    assert not any(n.startswith("customXml/") for n in Package(out).parts)
    extra = '<root>【SECRET_999】</root>'.encode()
    final = replace_parts(out, {"customXml/fictional.xml": extra})
    with pytest.raises(SafetyError):
        restore(final, encrypted, PASSWORD)
    _, report = restore(final, encrypted, PASSWORD, acknowledged=True)
    assert "【SECRET_999】" in report["unknown_tokens"]
    assert report["remaining_tokens"]


def test_output_rollback_on_write_failure(tmp_path, monkeypatch):
    import wordshield.storage as storage
    original_open = storage.os.open
    one, two = tmp_path / "one.docx", tmp_path / "two.wsmap"
    def failing_open(path, flags, mode):
        if path == two:
            raise OSError("fictional disk error")
        return original_open(path, flags, mode)
    monkeypatch.setattr(storage.os, "open", failing_open)
    with pytest.raises(OSError):
        write_outputs({one: b"fictional", two: b"fictional"})
    assert not one.exists() and not two.exists()


def test_duplicate_export_paths_rejected(tmp_path):
    data = simple([("林知远", True)])
    source, output = tmp_path / "source.docx", tmp_path / "output"
    source.write_bytes(data)
    with pytest.raises(OutputError):
        export_redacted(source, output, output, output, analyze(data), PASSWORD)
    assert not output.exists()


def test_empty_runs_do_not_split_entity():
    plan = analyze(simple([("林", True), ("", False), ("知远", True)]))
    assert [o.text for o in plan.occurrences] == ["林知远"]


def test_report_contains_no_sensitive_text(samples):
    from wordshield.core import report_bytes
    plan = analyze(samples["01_复杂格式原稿.docx"])
    out, encrypted, redaction_report = redact(plan, PASSWORD)
    _, restoration_report = restore(out, encrypted, PASSWORD)
    for item in plan.occurrences:
        assert item.text.encode() not in report_bytes(redaction_report)
        assert item.text.encode() not in report_bytes(restoration_report)


def test_cross_paragraph_yellow_stays_separate(samples):
    plan = analyze(samples["12_跨段落.docx"])
    marked = [o.text for o in plan.occurrences if o.highlighted]
    assert marked == ["云岚市", "澄溪区"]


def test_special_hyphen_not_silently_removed():
    doc = Document()
    p = doc.add_paragraph()
    highlighted(p, "林")
    special = p.add_run()
    special._r.append(OxmlElement("w:noBreakHyphen"))
    highlighted(p, "知远")
    plan = analyze(saved(doc))
    assert [o.text for o in plan.occurrences] == ["林", "知远"]
    assert "SPECIAL_TEXT" in {i.code for i in plan.issues}


def test_report_obscures_nonstandard_internal_paths():
    from wordshield.core import Issue, public_part, report_bytes
    part = "word/media/林知远.png"
    assert "林知远" not in public_part(part)
    assert "林知远".encode() not in report_bytes({"issues": [Issue("BINARY", "需要审查", part).public()]})
    assert public_part("word/header1.xml") == "word/header1.xml"
