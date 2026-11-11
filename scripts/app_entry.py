"""应用入口。WORDSHIELD_SELF_TEST=1 时只做打包自检：导入依赖，并在内存中用虚构文档完成一次
脱敏与回填，不打开界面，不写文件。"""
import os
import sys


def self_test() -> None:
    from io import BytesIO

    import tkinter  # noqa: F401
    from docx import Document
    from docx.enum.text import WD_COLOR_INDEX

    import wordshield.gui  # noqa: F401
    from wordshield.core import analyze, redact, restore

    document = Document()
    paragraph = document.add_paragraph("客户 ")
    paragraph.add_run("虚构测试人").font.highlight_color = WD_COLOR_INDEX.YELLOW
    paragraph.add_run(" 签署了合同。")
    cell = document.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0]
    cell.add_run("虚构测试人").font.highlight_color = WD_COLOR_INDEX.YELLOW
    source = BytesIO()
    document.save(source)

    password = "self-test-only-password"
    redacted, mapping, _ = redact(analyze(source.getvalue()), password)
    redacted_text = [p.text for p in Document(BytesIO(redacted)).paragraphs]
    if "客户 【SECRET_001】 签署了合同。" not in redacted_text:
        sys.exit("WordShield self-test failed: redaction")
    restored, report = restore(redacted, mapping, password)
    result = Document(BytesIO(restored))
    if (report["status"] != "通过" or "客户 虚构测试人 签署了合同。" not in [p.text for p in result.paragraphs]
            or result.tables[0].cell(0, 0).text != "虚构测试人"):
        sys.exit("WordShield self-test failed: restoration")
    print("WordShield self-test OK")


if os.environ.get("WORDSHIELD_SELF_TEST") == "1":
    self_test()
    sys.exit(0)

from wordshield.gui import main

if __name__ == "__main__":
    main()
