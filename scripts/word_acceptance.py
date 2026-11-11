"""用本机 Microsoft Word 验收虚构样本。只处理 generate_fixtures.py 生成的虚构材料。

检查三件事：
1. Word 能直接打开全部样本（包括脱敏和回填结果），没有修复提示或报错。
2. 原稿经 Word 打开并另存后，WordShield 识别出的黄色高亮与原稿一致。
3. 脱敏结果经 Word 打开并另存后仍能完整回填，结果与未经 Word 另存时相同，Word 也能打开回填结果。

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
from wordshield.core import TOKEN, Package, analyze, redact, restore
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
      set body to content of text object of d
      if (count of argv) > 1 then
        save as d file name (item 2 of argv) file format format document
        -- 另存后文档改名，需按新文件名关闭。
        set d to document (item -1 of my splitPath(item 2 of argv))
      end if
      close d saving no
    end tell
  end timeout
  return body
end run

on splitPath(posixPath)
  set AppleScript's text item delimiters to "/"
  set parts to text items of posixPath
  set AppleScript's text item delimiters to ""
  return parts
end splitPath
"""


def word(source: Path, save_as: Path | None = None) -> str:
    """用 Word 打开文档并返回正文；给出 save_as 时由 Word 另存为 docx。"""
    arguments = [str(source)] + ([str(save_as)] if save_as else [])
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

    print(f"\n样本与 Word 另存结果位于 {folder}")
    print("全部通过" if not failures else f"{len(failures)} 项失败")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
