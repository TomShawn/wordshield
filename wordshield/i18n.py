"""User-interface translations. Document contents never enter this module."""
from __future__ import annotations

from copy import deepcopy
import re


TEXTS = {
    "app_title": ("Word 离线脱敏与回填", "WordShield — Offline Word Redaction & Restoration"),
    "app_heading": ("Word 离线脱敏与回填", "WordShield"),
    "language": ("语言", "Language"),
    "disclaimer": (
        "仅识别 Word 黄色“文本突出显示颜色”，无法自动发现所有未标记敏感信息。\n"
        "请使用本机非同步文件夹。文档、映射表与回填结果均由本机处理。",
        "Only yellow Word Text Highlight Color is detected. The app cannot discover all unmarked sensitive information.\n"
        "Use a local, non-synchronized folder. Documents, mappings, and restored output are processed locally.",
    ),
    "status_choose": ("请选择需要处理的 DOCX 文件。", "Choose a DOCX file to begin."),
    "tab_redact": ("脱敏 Word", "Redact Word"),
    "tab_restore": ("回填 Word", "Restore Word"),
    "choose_analyze": ("选择 DOCX 并识别", "Choose DOCX and analyze"),
    "no_file": ("尚未选择文件", "No file selected"),
    "preview_help": (
        "预览仅在内存显示。相同分组生成同一占位符，不同文字必须使用不同分组。",
        "The preview exists only in memory. Items in one group share a placeholder; different text must use different groups.",
    ),
    "column_text": ("敏感文字（首尾空格不计入）", "Sensitive text (outer spaces excluded)"),
    "column_group": ("实体分组", "Entity group"),
    "column_origin": ("来源 / 处理", "Source / action"),
    "column_location": ("位置", "Location"),
    "edit_group": ("调整所选条目分组", "Assign item to group"),
    "rename_group": ("重命名整个分组", "Rename group"),
    "rename_group_prompt": ("为当前分组的 {count} 个条目输入新名称。", "Enter a new name for all {count} items in this group."),
    "rename_group_failed": ("无法重命名分组", "Cannot rename group"),
    "separate_group": ("同名异实体单独分组", "Separate entity"),
    "split_highlight": ("拆分连续高亮", "Split highlight"),
    "toggle_retain": ("保留／恢复所选原文", "Retain / replace text"),
    "include_unmarked": (
        "同时替换预览中的未标记同文位置（请逐项确认其实体分组）",
        "Also replace unmarked identical text shown in the preview (verify each entity group)",
    ),
    "risk_pending": ("检查结果将在识别后显示。", "Inspection results will appear after analysis."),
    "confirm_preview": ("我已核对全部敏感文字、位置和分组，确认导出", "I reviewed all sensitive text, locations, and groups and approve export"),
    "confirm_risk": (
        "我已人工审查上述无法可靠处理的内容，明确接受剩余风险",
        "I manually reviewed the content that cannot be checked reliably and accept the remaining risk",
    ),
    "export_redacted": ("设置密码并导出脱敏文件与加密映射表", "Set password and export redacted file and encrypted mapping"),
    "restore_intro": (
        "导入含占位符的 AI 定稿和原加密映射表。回填结果属于保密文件，请保存在本地。",
        "Import the final DOCX containing placeholders and its encrypted mapping. Restored output is confidential; save it locally.",
    ),
    "final_docx": ("AI 定稿 DOCX", "Final DOCX"),
    "encrypted_mapping": ("加密映射表", "Encrypted mapping"),
    "choose_file": ("选择文件", "Choose file"),
    "inspect_restore": ("检查定稿并回填", "Inspect and restore"),
    "restore_pending": (
        "回填后将在此处显示报告。已删除的占位符只报告，不推测恢复位置。",
        "The restoration report will appear here. Deleted placeholders are reported without guessing a restoration location.",
    ),
    "open_restored": ("打开最近生成的回填文档", "Open latest restored document"),
    "origin_highlight": ("黄色高亮", "Yellow highlight"),
    "origin_unmarked": ("未标记同文", "Unmarked identical text"),
    "action_retain": ("保留", "Retain"),
    "action_replace": ("替换", "Replace"),
    "must_fix": ("必须修复 ", "Must fix: "),
    "review": ("需审查 ", "Review: "),
    "metadata_notice": ("导出会清除文档属性、自定义 XML 数据、缩略图和文档变量。", "Export removes document properties, custom XML data, thumbnails, and document variables."),
    "retention_notice": ("保留项必须单独确认。其他已知原文残留仍会阻止导出。", "Retained items require separate confirmation. Other known residual original text still blocks export."),
    "location": ("{part} 第 {paragraph} 段，第 {character} 字符", "{part}, paragraph {paragraph}, character {character}"),
    "processing_stopped": ("处理已停止，请根据提示处理后重试。", "Processing stopped. Follow the message and try again."),
    "not_exported": ("未导出", "Not exported"),
    "io_failed": ("文件读写失败。", "File access failed."),
    "io_failed_title": ("文件读写失败", "File access failed"),
    "io_failed_body": ("请检查文件是否存在、文件夹权限和磁盘空间。", "Check that the file exists and that folder permissions and disk space are available."),
    "processing_failed": ("处理失败，已停止。", "Processing failed and stopped."),
    "processing_failed_title": ("处理失败", "Processing failed"),
    "processing_failed_body": ("发生无法处理的内部错误。请使用虚构材料复现后联系开发者。", "An internal error occurred. Reproduce it with fictional material before contacting the developer."),
    "choose_word": ("选择 Word 文档", "Choose Word document"),
    "word_doc": ("Word 文档", "Word document"),
    "analyzing": ("正在本地识别高亮和检查 DOCX 内部结构…", "Analyzing highlights and DOCX structure locally…"),
    "analysis_summary": ("识别到 {marked} 处连续黄色高亮，另有 {unmarked} 处未标记同文位置。", "Detected {marked} contiguous yellow-highlighted region(s) and {unmarked} unmarked identical occurrence(s)."),
    "no_highlight_title": ("未发现黄色高亮", "No yellow highlight found"),
    "no_highlight_body": ("请先用 Word 的黄色文本突出显示颜色标记敏感文字。", "Mark sensitive text with Word's yellow Text Highlight Color first."),
    "select_item_title": ("请选择条目", "Select an item"),
    "select_item_body": ("请先选择预览中的一行。", "Select a row in the preview first."),
    "group_title": ("实体分组", "Entity group"),
    "group_prompt": ("输入所选条目要加入的分组名称。", "Enter the group to assign this selected item to."),
    "split_title": ("拆分连续高亮", "Split contiguous highlight"),
    "split_prompt": ("在第几个字符后拆分？多个位置用英文逗号分隔，例如 2,5。文字中间的空格也计入。", "Split after which character? Separate multiple positions with commas, for example 2,5. Spaces inside the text count."),
    "split_failed": ("无法拆分", "Cannot split"),
    "integer_positions": ("请输入整数位置。", "Enter integer positions."),
    "mapping_password": ("映射表密码", "Mapping password"),
    "password_prompt": ("请输入映射表密码", "Enter the mapping password"),
    "password_requirement": ("（至少 12 个字符，遗失后无法恢复）", " (at least 12 characters; a lost password cannot be recovered)"),
    "confirm_password": ("确认密码", "Confirm password"),
    "confirm_password_prompt": ("请再次输入密码", "Enter the password again"),
    "bad_password_title": ("密码不符合要求", "Password does not meet requirements"),
    "bad_password_body": ("两次密码必须相同，且至少 12 个字符。", "Both passwords must match and contain at least 12 characters."),
    "choose_document_title": ("请选择文档", "Choose a document"),
    "choose_document_body": ("请先选择 DOCX 并核对预览。", "Choose a DOCX and review the preview first."),
    "review_preview_title": ("请核对预览", "Review the preview"),
    "review_preview_body": ("请核对预览后勾选人工确认。", "Review the preview and select the manual confirmation."),
    "confirm_retention_title": ("确认保留原文", "Confirm retained original text"),
    "confirm_retention_body": ("预览中有 {count} 处将保留原文，其中 {highlighted} 处来自黄色高亮。\n这些原文会出现在导出文档中。请核对灰色保留条目，并确认这些位置可公开保留，是否继续？", "The preview retains original text in {count} location(s), including {highlighted} yellow-highlighted location(s).\nThis text will remain in the exported document. Review the gray retained rows and confirm that these locations may remain visible. Continue?"),
    "save_redacted": ("保存脱敏 Word（请选择本地非同步文件夹和新文件名）", "Save redacted Word document (choose a local non-synchronized folder and new name)"),
    "redacted_filename": ("脱敏文档.docx", "redacted_document.docx"),
    "redaction_complete_status": ("已完成脱敏，共替换 {count} 处，生成 DOCX、加密映射表与检查报告。", "Redaction complete. Replaced {count} occurrence(s) and created a DOCX, encrypted mapping, and inspection report."),
    "redaction_complete_title": ("脱敏完成", "Redaction complete"),
    "redaction_complete_body": ("已在所选文件夹生成脱敏 Word、.wsmap 加密映射表和检查报告。\n请将映射表与密码分开离线保存。", "Created the redacted Word document, encrypted .wsmap, and inspection report in the selected folder.\nStore the mapping and password separately offline."),
    "redacting": ("正在本地替换、扫描残留并加密映射表…", "Replacing text, scanning residuals, and encrypting the mapping locally…"),
    "choose_final": ("选择 AI 定稿", "Choose final DOCX"),
    "choose_mapping": ("选择加密映射表", "Choose encrypted mapping"),
    "mapping_file": ("加密映射表", "Encrypted mapping"),
    "choose_files_title": ("请选择文件", "Choose files"),
    "choose_files_body": ("请选择 AI 定稿 DOCX 和加密映射表。", "Choose the final DOCX and encrypted mapping."),
    "must_fix_title": ("需要先修复", "Fix required"),
    "must_fix_body": ("存在必须修复的安全问题，请查看检查列表。", "The document contains a blocking safety issue. Review the inspection list."),
    "confirm_review_title": ("确认人工审查", "Confirm manual review"),
    "confirm_review_body": ("定稿存在无法可靠处理的内容。\n请阅读检查列表并人工审查后确认接受风险，是否继续？", "The final document contains content that cannot be checked reliably.\nRead the inspection list, complete a manual review, and confirm that you accept the risk. Continue?"),
    "inspecting_final": ("正在本地检查定稿结构…", "Inspecting the final DOCX structure locally…"),
    "save_restored": ("保存回填 Word（保密文件，请选择本地非同步文件夹和新文件名）", "Save restored Word document (confidential: choose a local non-synchronized folder and new name)"),
    "restored_filename": ("回填文档.docx", "restored_document.docx"),
    "restore_result": ("回填检查结果 {status}", "Restoration result: {status}"),
    "restored_count": ("已回填 {count} 处", "Restored {count} occurrence(s)"),
    "placeholder_check": ("占位符校验 {status}", "Placeholder check: {status}"),
    "structure_review": ("文档结构审查 {status}", "Document structure review: {status}"),
    "review_required": ("需要人工审查", "Manual review required"),
    "no_structure_warning": ("未发现结构告警", "No structural warning found"),
    "unknown_tokens": ("未知占位符 {value}", "Unknown placeholders: {value}"),
    "missing_tokens": ("定稿中消失的占位符 {value}", "Placeholders missing from final DOCX: {value}"),
    "count_changes": ("出现次数变化", "Occurrence count changes"),
    "count_change": ("  {token}  原来 {expected} 次，定稿 {actual} 次", "  {token}: expected {expected}, final document {actual}"),
    "malformed_count": ("疑似修改或拆开 {count} 处", "Possibly modified or split placeholders: {count}"),
    "remaining_count": ("仍含占位符或疑似标记的内部部件 {count} 个", "Internal parts still containing placeholders or suspicious markers: {count}"),
    "none": ("无", "None"),
    "package_valid": ("DOCX 包、XML 和 python-docx 结构检查通过。", "DOCX package, XML, and python-docx structure checks passed."),
    "report_saved": ("完整检查报告已与回填 DOCX 一起保存。请打开文档检查排版。", "The full inspection report was saved with the restored DOCX. Open the document and inspect its layout."),
    "restore_complete_status": ("已生成回填文档与检查报告，结果为{status}。", "Created the restored document and inspection report. Result: {status}."),
    "restoring": ("正在本地解密、回填并检查异常…", "Decrypting, restoring, and checking anomalies locally…"),
    "no_output_title": ("尚无输出", "No output yet"),
    "no_output_body": ("请先完成回填。", "Complete a restoration first."),
    "ui_failed_title": ("界面操作失败", "Interface operation failed"),
    "ui_failed_body": ("界面操作未完成，请重新选择文件后重试。", "The interface operation did not complete. Select the file again and retry."),
    "busy_title": ("正在处理", "Processing"),
    "busy_body": ("请等待本次本地处理完成后关闭。", "Wait for the current local operation to finish before closing."),
}


ISSUE_EN = {
    "存在文本框，需要在 Word 中人工审查。": "A text box is present and requires manual review in Word.",
    "存在修订记录，请优先接受或拒绝修订后重新导入。": "Tracked changes are present. Accept or reject them before importing again.",
    "存在删除修订，其原文可能仍保存在文件内。": "A deletion revision is present; its original text may remain in the file.",
    "存在移动修订。": "A move revision is present.",
    "存在隐藏文字或隐藏样式。": "Hidden text or a hidden style is present.",
    "存在网页隐藏文字。": "Web-hidden text is present.",
    "存在内容控件，需要人工检查绑定数据。": "A content control is present; review its bound data manually.",
    "存在域，Word 更新域可能改变显示内容。": "A field is present. Updating fields in Word may change displayed content.",
    "存在简单域，需要人工检查。": "A simple field is present and requires manual review.",
    "存在外部导入内容，无法可靠处理。": "Externally imported content is present and cannot be processed reliably.",
    "存在兼容性备用内容，需要人工审查。": "Compatibility fallback content is present and requires manual review.",
    "存在数据绑定，脱敏时将清除绑定数据与关联。": "Data binding is present. Redaction removes bound data and its relationship.",
    "存在公式对象，需要人工检查其中的文字。": "A math object is present; inspect its text manually.",
    "存在绘图或图片，需要人工检查其中的文字与替代文本。": "A drawing or image is present; inspect its text and alternative text manually.",
    "存在旧式绘图或形状，需要人工审查。": "A legacy drawing or shape is present and requires manual review.",
    "存在符号字体字符，需要人工审查。": "A symbol-font character is present and requires manual review.",
    "存在特殊连字符，其两侧高亮分别识别，请人工核对。": "A non-breaking hyphen is present. Highlights on each side are detected separately; review them manually.",
    "存在可选连字符，其两侧高亮分别识别，请人工核对。": "A soft hyphen is present. Highlights on each side are detected separately; review them manually.",
    "未支持的区域含黄色高亮文字，请在 Word 中将其移到正文或清理后重新导入。": "An unsupported area contains yellow-highlighted text. Move it to supported body text or remove it in Word, then import again.",
    "存在批注或批注人员信息，需要人工审查。": "Comments or reviewer information are present and require manual review.",
    "存在外部链接或模板关系，需要人工审查。": "An external link or template relationship is present and requires manual review.",
    "表格样式含黄色高亮，需改为直接文本高亮后重新导入。": "A table style contains yellow highlighting. Convert it to direct text highlighting and import again.",
    "存在非主要文档 XML，需人工审查其中内容。": "A non-primary document XML part is present and requires manual review.",
    "存在图片、嵌入对象或嵌入字体，无法检查其中的文字。": "An image, embedded object, or embedded font is present; its text cannot be inspected.",
    "存在无法可靠检查的二进制内容。": "Binary content that cannot be inspected reliably is present.",
    "存在自定义数据，脱敏时会清除。": "Custom XML data is present and will be removed during redaction.",
    "存在数字签名或活动/二进制对象，请在 Word 中移除后重试。": "A digital signature or active/binary object is present. Remove it in Word and retry.",
    "源文档已含 SECRET 标记，请先处理以避免占位符冲突。": "The source already contains a SECRET marker. Resolve it before continuing to avoid placeholder collisions.",
}


ERROR_EN = {
    "请选择高亮条目，并填写递增且不重复的拆分位置。": "Select a highlighted item and enter increasing, non-duplicate split positions.",
    "拆分位置必须在文字内部。": "Split positions must be inside the text.",
    "DOCX 超过 128 MB 容量限制。": "The DOCX exceeds the 128 MB limit.",
    "DOCX 包含重复条目或条目过多。": "The DOCX contains duplicate or too many package entries.",
    "DOCX 解压后超过容量限制。": "The uncompressed DOCX exceeds the size limit.",
    "DOCX 包含不安全或过大的内部条目。": "The DOCX contains an unsafe or oversized internal entry.",
    "DOCX 使用了不支持的压缩方式。": "The DOCX uses an unsupported compression method.",
    "DOCX 实际解压数据超过容量限制。": "The actual uncompressed DOCX data exceeds the limit.",
    "DOCX 内部条目的声明长度与读取长度不一致。": "A DOCX entry's declared length does not match the data read.",
    "无法读取 DOCX，请检查文件格式与完整性。": "The DOCX cannot be read. Check its format and integrity.",
    "DOCX 包含不允许的 XML 实体声明。": "The DOCX contains a prohibited XML entity declaration.",
    "DOCX 包含不允许的 XML 文档类型声明。": "The DOCX contains a prohibited XML document type declaration.",
    "DOCX 的内部 XML 已损坏。": "The DOCX contains malformed internal XML.",
    "文件不是受支持的 DOCX。": "The file is not a supported DOCX.",
    "仅支持常规 DOCX，Strict OOXML 请先在 Word 中另存为标准 DOCX。": "Only regular DOCX is supported. Save Strict OOXML as a standard DOCX in Word first.",
    "DOCX 的关系或主文档结构无法通过 python-docx 检查。": "The DOCX relationships or main document structure failed validation.",
    "替换跨越了不支持的结构边界，已阻止导出。": "A replacement crosses an unsupported structure boundary, so export was blocked.",
    "存在必须修复的结构或安全问题，请查看检查列表。": "A blocking structural or safety issue is present. Review the inspection list.",
    "存在无法可靠检查的内容，请阅读并明确确认风险后再导出。": "Some content cannot be checked reliably. Review it and explicitly accept the risk before export.",
    "预览内容或分组无效，请重新导入。": "The preview content or grouping is invalid. Import the document again.",
    "不同文字不能共用同一分组，请修改分组。": "Different text values cannot share one group. Change the grouping.",
    "分组名称不能为空。": "The group name cannot be empty.",
    "所选分组已不存在，请重新选择。": "The selected group no longer exists. Select it again.",
    "分组名称已存在，请使用新名称。": "The group name already exists. Choose a new name.",
    "脱敏范围存在重叠，请重新导入。": "Redaction ranges overlap. Import the document again.",
    "预览中有保留原文的位置，必须单独确认这些保留位置后才能导出。": "The preview contains retained original text. Confirm those locations separately before export.",
    "未发现需要替换的黄色文本高亮，请先在 Word 中标记并核对保留条目。": "No replaceable yellow-highlighted text was found. Mark text in Word and review retained items.",
    "占位符计数校验失败，已阻止导出。": "Placeholder count validation failed, so export was blocked.",
    "映射表条目无效。": "The mapping contains invalid entries.",
    "映射表包含无效条目。": "The mapping contains an invalid entry.",
    "三个输出文件路径必须不同。": "The three output paths must be different.",
    "源文件已发生变化，请重新导入并确认预览。": "The source file changed. Import it again and review the new preview.",
    "文档与报告的输出路径必须不同。": "The document and report output paths must be different.",
    "输出文件路径不能重复。": "Output paths cannot be duplicated.",
    "输出路径已存在或与输入文件相同，请选择新文件名。": "An output path already exists or matches an input file. Choose a new filename.",
    "输出文件夹不存在。": "The output folder does not exist.",
    "映射表密码至少需要 12 个字符，请妥善离线保存。": "The mapping password must contain at least 12 characters. Store it safely offline.",
    "映射表超出容量限制。": "The mapping exceeds the size limit.",
    "映射表格式不支持或文件已损坏。": "The mapping format is unsupported or the file is damaged.",
    "密码不正确，或映射表已损坏。": "The password is incorrect or the mapping is damaged.",
    "映射表版本不支持。": "The mapping version is unsupported.",
    "请选择项目目录和应用包之外的文件夹，以保持程序目录干净。": "Choose a folder outside the source project and application bundle to keep them clean.",
}


STATUS_EN = {
    "通过": "Passed",
    "需人工检查": "Manual review required",
    "回填通过，文档需审查": "Restoration passed; document review required",
}


NOTICE_EN = {
    "仅处理人工标记及其相同文字。无法自动发现全部未标记敏感信息。":
        "Only manually marked text and identical occurrences are processed. The app cannot discover all unmarked sensitive information.",
    "消失的占位符仅列入报告，不猜测恢复位置。结构检查通过不等同于 Microsoft Word 实际打开验证。":
        "Missing placeholders are reported without guessing a restoration location. Structural validation is not a substitute for opening the document in Microsoft Word.",
}


def tr(language: str, key: str, **values) -> str:
    pair = TEXTS[key]
    text = pair[0] if language == "zh" else pair[1]
    return text.format(**values)


def issue_text(language: str, message: str) -> str:
    return message if language == "zh" else ISSUE_EN.get(message, "A document structure requires manual review.")


def error_text(language: str, message: str) -> str:
    if language == "zh":
        return message
    residual = re.fullmatch(
        r"检测到已标记文字仍残留在 DOCX 内部，已阻止导出。请处理未标记位置、隐藏内容或内部属性。涉及部件数 (\d+)",
        message,
    )
    if residual:
        return f"Known marked text remains inside the DOCX, so export was blocked. Review unmarked locations, hidden content, and internal attributes. Affected parts: {residual.group(1)}"
    return ERROR_EN.get(message, "Processing stopped because document validation failed.")


def status_text(language: str, status: str) -> str:
    return status if language == "zh" else STATUS_EN.get(status, status)


def notice_text(language: str, notice: str) -> str:
    return notice if language == "zh" else NOTICE_EN.get(notice, notice)


def part_text(language: str, part: str) -> str:
    if language == "en" and re.fullmatch(r"内部部件 [0-9a-f]{12}", part):
        return part.replace("内部部件 ", "Internal part ", 1)
    return part


def localized_report(report: dict, language: str) -> dict:
    """Translate human-readable report values without changing its stable schema."""
    if language == "zh":
        return report
    result = deepcopy(report)
    result["operation"] = {"脱敏": "Redaction", "回填": "Restoration"}.get(
        result.get("operation"), result.get("operation")
    )
    if "status" in result:
        result["status"] = status_text("en", result["status"])
    if "placeholder_status" in result:
        result["placeholder_status"] = status_text("en", result["placeholder_status"])
    if "notice" in result:
        result["notice"] = notice_text("en", result["notice"])
    for issue in result.get("issues", []):
        issue["message"] = issue_text("en", issue["message"])
    for item in result.get("malformed_tokens", []):
        if item.get("kind") == "疑似被改写或拆开的占位符":
            item["kind"] = "Possibly modified or split placeholder"
        elif item.get("kind") == "疑似跨段落拆开的占位符":
            item["kind"] = "Placeholder possibly split across paragraphs"
    for collection in ("issues", "malformed_tokens", "remaining_tokens", "excluded_locations"):
        for item in result.get(collection, []):
            if "part" in item:
                item["part"] = part_text(language, item["part"])
    return result
