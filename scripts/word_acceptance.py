"""用本机 Microsoft Word 验收虚构样本。只处理 generate_fixtures.py 生成的虚构材料。

检查四件事：
1. Word 能直接打开全部样本（包括脱敏和回填结果），没有修复提示或报错。
2. 原稿经 Word 打开并另存后，WordShield 识别出的黄色高亮与原稿一致。
3. 脱敏结果经 Word 打开并另存后仍能完整回填，结果与未经 Word 另存时相同，Word 也能打开回填结果。
4. 在 Word 中实际编辑脱敏稿：在占位符前后加字、改格式、追加段落后能完整回填并保留编辑；
   误删或多复制占位符时如实报告；开启修订后编辑时要求确认风险。

首次运行时 Word 可能请求访问输出文件夹，请在 Word 中授予访问权限；macOS 也可能询问是否允许控制 Word。
用法：scripts/dev.sh word-acceptance [输出文件夹]
"""
from __future__ import annotations

from pathlib import Path
import subprocess
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from generate_fixtures import PASSWORD, make_all
from io import BytesIO

from docx import Document

from wordshield.core import TOKEN, Package, SafetyError, analyze, redact, restore
from wordshield.crypto import decrypt_mapping
from wordshield.paths import development_directory, require_external

WORD_SCRIPT = """
on run argv
  set source to item 1 of argv
  set fileName to item -1 of my splitPath(source)
  with timeout of 120 seconds
    tell application "Microsoft Word"
      open (POSIX file source)
      -- open 可能在文档就绪前返回，等到同名文档出现。
      set d to missing value
      repeat 120 times
        try
          set d to document fileName
          exit repeat
        end try
        delay 0.5
      end repeat
      if d is missing value then error "Word 在 60 秒内没有打开文档"
      -- 其余参数每三个一组：操作、占位符、值，模拟人在 Word 中编辑。
      set k to 3
      repeat while k + 2 ≤ (count of argv)
        set op to item k of argv
        set tok to item (k + 1) of argv
        set val to item (k + 2) of argv
        set k to k + 3
        if op is "track" then
          set track revisions of d to true
        else if op is "append" then
          set e to (end of content of text object of d) - 1
          set content of (create range d start e end e) to (return & val)
        else
          set i to (my findPos(tok, content of text object of d)) - 1
          if i < 0 then error "找不到 " & tok
          set j to i + (length of tok)
          if content of (create range d start i end j) is not tok then error "定位偏移 " & tok
          if op is "after" then
            set content of (create range d start j end j) to val
          else if op is "before" then
            set content of (create range d start i end i) to val
          else if op is "bold" then
            set bold of font object of (create range d start i end j) to true
          else if op is "bold-inner" then
            -- 双击占位符时 Word 只选中括号内的 SECRET_00x。
            set bold of font object of (create range d start (i + 1) end (j - 1)) to true
          else if op is "italic-tail" then
            set italic of font object of (create range d start (i + (val as integer)) end j) to true
          else if op is "font" then
            set name of font object of (create range d start i end j) to val
          else if op is "color" then
            set color index of font object of (create range d start i end j) to red
          else if op is "delete" then
            set content of (create range d start i end j) to ""
          else if op is "duplicate" then
            set content of (create range d start j end j) to tok
          else
            error "未知操作 " & op
          end if
        end if
      end repeat
      set body to content of text object of d
      if (item 2 of argv) is not "" then
        save as d file name (item 2 of argv) file format format document
        -- 另存后文档改名，需按新文件名关闭。
        set d to document (item -1 of my splitPath(item 2 of argv))
      end if
      close d saving no
    end tell
  end timeout
  return body
end run

on findPos(needle, haystack)
  return offset of needle in haystack
end findPos

on splitPath(posixPath)
  set AppleScript's text item delimiters to "/"
  set parts to text items of posixPath
  set AppleScript's text item delimiters to ""
  return parts
end splitPath
"""


def word(source: Path, save_as: Path | None = None, edits: tuple[tuple[str, str, str], ...] = ()) -> str:
    """用 Word 打开文档、按 edits 编辑并返回正文；给出 save_as 时由 Word 另存为 docx。"""
    arguments = [str(source), str(save_as) if save_as else ""] + [part for edit in edits for part in edit]
    result = subprocess.run(["/usr/bin/osascript", "-", *arguments], input=WORD_SCRIPT,
                            capture_output=True, text=True, timeout=180)
    if result.returncode:
        raise RuntimeError(result.stderr.strip() or "Word 未能打开文档")
    if save_as and not save_as.is_file():
        raise RuntimeError("Word 未生成另存文件")
    return result.stdout


def texts(data: bytes) -> list[tuple[str, str]]:
    # Word 另存时会在表格后补空段落、补脚注尾注分隔符，只比较有文字的段落。
    return [(p.part, p.text) for p in Package(data).paragraphs() if p.text.strip()]


def highlighted(data: bytes) -> list[str]:
    return sorted(o.text for o in analyze(data).occurrences if o.highlighted)


def formatted(data: bytes, text: str, attribute: str) -> bool:
    """正文中第一处 text 覆盖到的全部 run 是否都设置了 bold 或 italic 等属性。"""
    for paragraph in Document(BytesIO(data)).paragraphs:
        start = paragraph.text.find(text)
        if start < 0:
            continue
        position, covered = 0, []
        for run in paragraph.runs:
            if position < start + len(text) and position + len(run.text) > start:
                covered.append(bool(getattr(run, attribute)))
            position += len(run.text)
        return bool(covered) and all(covered)
    return False


def edit_scenarios(fixtures: Path, output: Path, check) -> None:
    """在 Word 中实际编辑脱敏稿，再回填。插入的文字都带【〔〕】以外的少见括号，便于识别。"""
    redacted = fixtures / "06_脱敏样例.docx"
    mapping = (fixtures / "06_脱敏样例.wsmap").read_bytes()
    entities = {token: item["text"] for token, item in decrypt_mapping(mapping, PASSWORD)["entries"].items()}
    expected, expected_report = restore(redacted.read_bytes(), mapping, PASSWORD)
    inserted = ("〔插入甲〕", "〔插入乙〕", "〔新增段落〕")

    # 1. 正常改稿：在占位符前后加字、整体加粗、后半段改斜体、换字体、改颜色、追加段落。
    edited = output / "W6_Word编辑脱敏稿.docx"
    word(redacted, edited, (
        ("after", "【SECRET_001】", inserted[0]), ("before", "【SECRET_004】", inserted[1]),
        ("bold", "【SECRET_002】", ""), ("italic-tail", "【SECRET_003】", "4"),
        ("bold-inner", "【SECRET_004】", ""),
        ("font", "【SECRET_005】", "Times New Roman"), ("color", "【SECRET_006】", ""),
        ("append", "", inserted[2]),
    ))
    restored, report = restore(edited.read_bytes(), mapping, PASSWORD)
    check("Word 编辑后回填状态为通过", report["status"] == "通过")
    check("Word 编辑后全部占位符都已回填", report["restored_count"] == expected_report["restored_count"])

    def without_insertions(data: bytes) -> list[tuple[str, str]]:
        rows = []
        for part, text in texts(data):
            for marker in inserted:
                text = text.replace(marker, "")
            if text.strip():
                rows.append((part, text))
        return rows

    check("Word 编辑后去掉新增文字，回填结果与直接回填一致", without_insertions(restored) == texts(expected))
    body = "".join(text for _, text in texts(restored))
    check("Word 中新增的文字和段落都保留", all(marker in body for marker in inserted))
    check("Word 中加粗的占位符回填后原文仍为粗体", formatted(restored, entities["【SECRET_002】"], "bold"))
    check("只加粗括号内文字的占位符回填后原文为粗体", formatted(restored, entities["【SECRET_004】"], "bold"))
    restored_path = output / "W7_Word编辑后回填.docx"
    restored_path.write_bytes(restored)
    check("Word 打开编辑后回填结果且无残留占位符", not TOKEN.search(word(restored_path)))

    # 2. 误删和误复制占位符：应如实报告，不猜测位置。
    damaged = output / "W8_Word误删误复制.docx"
    word(redacted, damaged, (("delete", "【SECRET_007】", ""), ("duplicate", "【SECRET_002】", "")))
    report = restore(damaged.read_bytes(), mapping, PASSWORD)[1]
    check("Word 中删除的占位符列入缺失", report["missing_tokens"] == ["【SECRET_007】"])
    check("Word 中多复制一份和删除的占位符都列入数量变化",
          {item["token"]: item["actual"] - item["expected"] for item in report["count_changes"]}
          == {"【SECRET_002】": 1, "【SECRET_007】": -1})
    check("误删误复制后回填状态为需人工检查", report["status"] == "需人工检查")

    # 3. 开启修订后编辑：修订属于需审查内容，未确认风险时拒绝回填，确认后回填并提示审查。
    tracked = output / "W9_Word修订模式编辑.docx"
    word(redacted, tracked, (("track", "", ""), ("after", "【SECRET_001】", inserted[0])))
    try:
        restore(tracked.read_bytes(), mapping, PASSWORD)
        refused = False
    except SafetyError:
        refused = True
    check("修订模式编辑后，未确认风险时拒绝回填", refused)
    report = restore(tracked.read_bytes(), mapping, PASSWORD, acknowledged=True)[1]
    check("修订模式编辑后，确认风险可回填且要求审查文档",
          report["placeholder_status"] == "通过" and report["document_review_required"])


def main() -> int:
    folder = require_external(Path(sys.argv[1]) if len(sys.argv) > 1
                              else development_directory() / "word-acceptance")
    # Word 运行在沙盒中，每个文件夹都要单独授权；样本与 Word 另存结果放在同一文件夹，只需授权一次。
    fixtures = word_saved = folder
    for old in folder.glob("W*_*.docx"):
        old.unlink()
    make_all(fixtures)
    failures: list[str] = []

    def check(label: str, condition: bool) -> None:
        print(("通过  " if condition else "失败  ") + label)
        if not condition:
            failures.append(label)

    for document in sorted(fixtures.glob("[0-9]*.docx")):
        try:
            body = word(document)
            check(f"Word 打开 {document.name}", bool(body.strip()))
        except (RuntimeError, subprocess.TimeoutExpired) as error:
            check(f"Word 打开 {document.name}：{error}", False)
    if failures:
        return 1

    original = (fixtures / "01_复杂格式原稿.docx").read_bytes()
    resaved_source = word_saved / "W1_Word另存原稿.docx"
    word(fixtures / "01_复杂格式原稿.docx", resaved_source)
    resaved = resaved_source.read_bytes()
    check("Word 另存后的原稿识别出相同的黄色高亮", highlighted(resaved) == highlighted(original))
    check("Word 另存后的原稿段落文字不变", texts(resaved) == texts(original))

    redacted = (fixtures / "06_脱敏样例.docx").read_bytes()
    mapping = (fixtures / "06_脱敏样例.wsmap").read_bytes()
    expected, expected_report = restore(redacted, mapping, PASSWORD)
    resaved_redacted = word_saved / "W2_Word另存脱敏稿.docx"
    word(fixtures / "06_脱敏样例.docx", resaved_redacted)
    restored, report = restore(resaved_redacted.read_bytes(), mapping, PASSWORD)
    check("Word 另存后的脱敏稿回填状态为通过", report["status"] == expected_report["status"] == "通过")
    check("Word 另存后的脱敏稿全部占位符都已回填",
          report["restored_count"] == expected_report["restored_count"] > 0)
    check("Word 另存后的回填文字与直接回填一致", texts(restored) == texts(expected))
    restored_path = word_saved / "W3_Word另存后回填.docx"
    restored_path.write_bytes(restored)
    check("Word 打开 Word 另存后的回填结果且无残留占位符", not TOKEN.search(word(restored_path)))

    # 用户实际从 Word 保存的原稿开始：脱敏、经 Word 另存、再回填，应回到原稿文字。
    redacted_from_word, mapping_from_word, _ = redact(analyze(resaved), PASSWORD)
    redacted_path = word_saved / "W4_Word原稿脱敏.docx"
    redacted_path.write_bytes(redacted_from_word)
    round_trip = word_saved / "W5_Word原稿脱敏后Word另存.docx"
    word(redacted_path, round_trip)
    final, final_report = restore(round_trip.read_bytes(), mapping_from_word, PASSWORD)
    check("Word 原稿完整往返（脱敏、Word 另存、回填）状态为通过", final_report["status"] == "通过")
    check("Word 原稿完整往返后的段落文字回到原稿", texts(final) == texts(resaved))

    edit_scenarios(fixtures, word_saved, check)

    print(f"\n样本与 Word 另存结果位于 {folder}")
    print("全部通过" if not failures else f"{len(failures)} 项失败")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
