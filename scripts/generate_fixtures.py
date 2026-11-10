"""生成全部虚构材料，无需也不接受任何真实保密输入。"""
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import sys
import zipfile

sys.dont_write_bytecode = True

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from wordshield.core import Package, analyze, redact, restore, report_bytes
from wordshield.paths import development_directory, require_external

PASSWORD = "Fictional-Only-2026!"


def highlighted(paragraph, text, *, bold=False, italic=False):
    run = paragraph.add_run(text)
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    run.bold = bold
    run.italic = italic
    return run


def base(title):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "PingFang SC"
    normal.font.size = Pt(11)
    normal.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "PingFang SC")
    title_style = doc.styles["Title"]
    title_style.font.color.rgb = RGBColor(0, 0, 0)
    for border in title_style.element.findall(".//" + qn("w:pBdr")):
        border.getparent().remove(border)
    title_run = doc.add_paragraph(style="Title").add_run(title)
    title_run.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph("本材料中的人物、地点、机构及业务内容均为虚构，仅用于离线软件测试。")
    return doc


def saved(doc):
    stream = BytesIO()
    doc.save(stream)
    return stream.getvalue()


def replace_parts(data, changes):
    source, target = BytesIO(data), BytesIO()
    with zipfile.ZipFile(source) as archive, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for item in archive.infolist():
            out.writestr(item.filename, changes.pop(item.filename, archive.read(item)))
        for name, content in changes.items():
            out.writestr(name, content)
    return target.getvalue()


def make_all(destination=None):
    destination = require_external(Path(destination) if destination is not None else development_directory() / "fixtures")
    destination.mkdir(parents=True, exist_ok=True)
    doc = base("星澜项目虚构会议材料")
    p = doc.add_paragraph("参会人员 ")
    highlighted(p, "林", bold=True)
    highlighted(p, "知", italic=True)
    highlighted(p, "远")
    p.add_run("与")
    highlighted(p, "周映禾")
    p.add_run("在")
    highlighted(p, "云岚市")
    p.add_run("会谈。")
    p = doc.add_paragraph("机构 ")
    highlighted(p, "星澜科技")
    p.add_run("与")
    highlighted(p, "星澜科技研究院")
    p.add_run("分别派员。")
    p = doc.add_paragraph("再次出席 ")
    highlighted(p, "林知远")
    p.add_run("。未标记同文位置 林知远、星澜科技研究院。")
    p = doc.add_paragraph("含前后空格 ")
    highlighted(p, "  沈清岳  ")
    p.add_run("；部分高亮 前")
    highlighted(p, "澄溪区", bold=True)
    p.add_run("后。")
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).text, table.cell(0, 1).text = "负责人", "地点"
    highlighted(table.cell(1, 0).paragraphs[0], "周映禾", bold=True)
    highlighted(table.cell(1, 1).paragraphs[0], "云岚市")
    header = doc.sections[0].header.paragraphs[0]
    header.add_run("项目机构 ")
    highlighted(header, "星澜科技")
    footer = doc.sections[0].footer.paragraphs[0]
    footer.add_run("联系人 ")
    highlighted(footer, "林知远")
    doc.core_properties.author = "林知远"
    doc.core_properties.subject = "星澜科技研究院"
    doc.core_properties.comments = "虚构测试元数据"
    complex_data = saved(doc)
    outputs = {"01_复杂格式原稿.docx": complex_data}

    doc = base("连续相邻高亮与结构边界")
    p = doc.add_paragraph("无分隔相邻高亮 ")
    highlighted(p, "林知远")
    highlighted(p, "周映禾")
    p = doc.add_paragraph("带未高亮空格 ")
    highlighted(p, "林知远")
    p.add_run(" ")
    highlighted(p, "周映禾")
    p = doc.add_paragraph("制表符与换行 ")
    highlighted(p, "澄溪区")
    p.add_run("\t")
    highlighted(p, "云岚市")
    p.add_run("\n")
    highlighted(p, "沈清岳")
    outputs["02_相邻与边界.docx"] = saved(doc)

    doc = base("无敏感信息")
    doc.add_paragraph("公开测试说明，全文没有黄色突出显示。")
    outputs["03_无高亮.docx"] = saved(doc)

    doc = base("包含关系与同名实体")
    for text in ("星澜科技", "星澜科技研究院", "林知远", "林知远"):
        p = doc.add_paragraph("条目 ")
        highlighted(p, text)
    doc.add_paragraph("未标记全称 星澜科技研究院；未标记简称 星澜科技。")
    outputs["04_包含与同名.docx"] = saved(doc)

    doc = base("需人工审查的复杂内容")
    p = doc.add_paragraph()
    highlighted(p, "林知远")
    hidden = doc.add_paragraph().add_run("林知远的隐藏测试文字")
    hidden.font.hidden = True
    insertion = OxmlElement("w:ins")
    insertion.set(qn("w:author"), "虚构修订者")
    insertion.set(qn("w:id"), "1")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "虚构修订记录"
    run.append(text)
    insertion.append(run)
    doc.add_paragraph()._p.append(insertion)
    pict = OxmlElement("w:pict")
    txbx = OxmlElement("w:txbxContent")
    inner_p, inner_r, inner_t = OxmlElement("w:p"), OxmlElement("w:r"), OxmlElement("w:t")
    inner_t.text = "林知远的文本框测试"
    inner_r.append(inner_t)
    inner_p.append(inner_r)
    txbx.append(inner_p)
    pict.append(txbx)
    doc.add_paragraph().add_run()._r.append(pict)
    risk = saved(doc)
    comments = (b'<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                + '<w:comment w:id="0" w:author="虚构批注者"><w:p><w:r><w:t>林知远的批注测试</w:t></w:r></w:p></w:comment></w:comments>'.encode() )
    outputs["05_隐藏修订文本框批注.docx"] = replace_parts(risk, {"word/comments.xml": comments})

    redacted, encrypted, redaction_report = redact(analyze(complex_data), PASSWORD)
    outputs["06_脱敏样例.docx"] = redacted
    (destination / "06_脱敏样例.wsmap").write_bytes(encrypted)
    (destination / "06_脱敏检查报告.json").write_bytes(report_bytes(redaction_report))

    # 从脱敏稿构建 AI 定稿，模拟 run 拆分、删除、重复、未知及改写。
    valid = Package(redacted)
    p = next(p for p in valid.paragraphs() if "【SECRET_001】" in p.text)
    span = next(s for s in p.spans if "【SECRET_001】" in (s.node.text or ""))
    span.node.text = (span.node.text or "").replace("【SECRET_001】", "【SECRET_")
    new_run, new_text = OxmlElement("w:r"), OxmlElement("w:t")
    new_text.text = "001】"
    new_run.append(new_text)
    span.node.getparent().addnext(new_run)
    outputs["07_AI正常拆run定稿.docx"] = valid.serialize()
    draft = Package(redacted)
    removed = False
    for root in draft.roots.values():
        for node in root.iter(qn("w:t")):
            text = node.text or ""
            if "【SECRET_001】" in text and not removed:
                text = text.replace("【SECRET_001】", "", 1)
                removed = True
            text = text.replace("【SECRET_002】", "【SECRET_002】【SECRET_002】")
            text = text.replace("【SECRET_003】", "【SECRET_ 003】")
            text = text.replace("【SECRET_004】", "【SECRET_999】")
            node.text = text
    outputs["08_AI删除重复改写定稿.docx"] = draft.serialize()
    restored, report = restore(outputs["07_AI正常拆run定稿.docx"], encrypted, PASSWORD)
    outputs["09_正常回填结果.docx"] = restored
    (destination / "09_正常回填报告.json").write_bytes(report_bytes(report))
    restored_bad, bad_report = restore(outputs["08_AI删除重复改写定稿.docx"], encrypted, PASSWORD)
    outputs["10_异常回填待检查.docx"] = restored_bad
    (destination / "10_异常回填报告.json").write_bytes(report_bytes(bad_report))
    doc = base("样式继承黄色高亮")
    doc.styles["Normal"].font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.add_paragraph("林知远")
    outputs["11_样式继承高亮.docx"] = saved(doc)
    doc = base("跨段落连续高亮")
    highlighted(doc.add_paragraph(), "云岚市")
    highlighted(doc.add_paragraph(), "澄溪区")
    outputs["12_跨段落.docx"] = saved(doc)
    for name, data in outputs.items():
        (destination / name).write_bytes(data)
    (destination / "说明.txt").write_text(
        "全部材料均为虚构。\n测试映射表密码 Fictional-Only-2026!\n"
        "05 含刻意构造的不支持结构，必须告警并阻止残留原文导出。\n"
        "08 和 10 含刻意保留的占位符异常，用于检验检查报告。\n", encoding="utf-8")
    return outputs


if __name__ == "__main__":
    result = make_all()
    print(f"已生成 {len(result)} 个虚构 DOCX 与加密映射表和检查报告。")
