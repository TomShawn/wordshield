"""DOCX 包级检查与原位文本节点替换。运行时没有网络、遥测或内容日志。"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from hashlib import sha256
from io import BytesIO
import json
import posixpath
from pathlib import Path, PurePosixPath
import re
import zipfile

from docx import Document
from lxml import etree

from .crypto import decrypt_mapping, encrypt_mapping, MappingError
from .storage import OutputError, write_outputs

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCUMENT_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
# 图片定位、尺寸与编号所在的命名空间。其中的数字和布尔值是版式数据，不是正文。
DRAWING_LAYOUT_NS = frozenset({
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.microsoft.com/office/drawing/2010/main",
})
DRAWING_LAYOUT_TEXT = frozenset({"posOffset", "align", "pctWidth", "pctHeight", "pctPosHOffset", "pctPosVOffset"})
# 二进制部件只按足够长的字节序列匹配，避免两三个字节的短词在图片数据中随机命中。
MIN_BINARY_MATCH = 4
NS = {"w": W}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
TOKEN = re.compile(r"【SECRET_[0-9]{3,6}】")
STORY = re.compile(r"word/(document|header[^/]*|footer[^/]*|footnotes|endnotes)\.xml$")
MAX_PACKAGE = 128 * 1024 * 1024
MAX_PART = 32 * 1024 * 1024
READ_CHUNK = 1024 * 1024


def public_part(name: str) -> str:
    """报告仅公开标准部件名，其他内部文件名使用不可逆位置编号。"""
    standard = (r"word/(document|header\d+|footer\d+|footnotes|endnotes|styles|stylesWithEffects|"
                r"settings|fontTable|numbering|webSettings|comments\d*)\.xml|"
                r"docProps/(core|app|custom)\.xml|\[Content_Types\]\.xml|_rels/\.rels|"
                r"word/_rels/[a-zA-Z]+\d*\.xml\.rels")
    if re.fullmatch(standard, name) or not name:
        return name
    return "内部部件 " + sha256(name.encode("utf-8")).hexdigest()[:12]


class SafetyError(ValueError):
    """错误信息只含类型与位置，不包含文档原文。"""


@dataclass(frozen=True)
class Issue:
    code: str
    message: str
    part: str = ""
    blocking: bool = False

    def public(self) -> dict:
        return {"code": self.code, "message": self.message, "part": public_part(self.part),
                "blocking": self.blocking}


@dataclass
class Span:
    start: int
    end: int
    node: etree._Element
    yellow: bool


@dataclass
class Paragraph:
    part: str
    number: int
    text: str
    spans: list[Span]

    @property
    def location(self) -> str:
        return f"{self.part} 第 {self.number + 1} 段"


@dataclass
class Occurrence:
    part: str
    paragraph: int
    start: int
    end: int
    text: str
    group: str
    highlighted: bool = True
    excluded: bool = False

    @property
    def key(self) -> str:
        return f"{self.part}:{self.paragraph}:{self.start}:{self.end}"

    @property
    def location(self) -> str:
        return f"{self.part} 第 {self.paragraph + 1} 段，第 {self.start + 1} 字符"


@dataclass
class Analysis:
    package: "Package"
    occurrences: list[Occurrence]
    issues: list[Issue]
    source_hash: str

    def new_group(self, reserved: tuple[str, ...] = ()) -> str:
        used = {o.group for o in self.occurrences} | set(reserved)
        number = 1
        while f"E{number:03d}" in used:
            number += 1
        return f"E{number:03d}"

    def exclude(self, key: str, excluded: bool) -> None:
        item = next(o for o in self.occurrences if o.key == key)
        item.excluded = excluded
        self.refresh_unmarked()

    def rename_group(self, group: str, new_name: str) -> None:
        new_name = new_name.strip()
        if not new_name:
            raise SafetyError("分组名称不能为空。")
        members = [item for item in self.occurrences if item.group == group]
        if not members:
            raise SafetyError("所选分组已不存在，请重新选择。")
        if new_name != group and any(item.group == new_name for item in self.occurrences):
            raise SafetyError("分组名称已存在，请使用新名称。")
        for item in members:
            item.group = new_name

    def assign_group(self, key: str, group: str) -> None:
        group = group.strip()
        if not group:
            raise SafetyError("分组名称不能为空。")
        item = next(o for o in self.occurrences if o.key == key)
        if any(o is not item and o.group == group and o.text != item.text for o in self.occurrences):
            raise SafetyError("不同文字不能共用同一分组，请修改分组。")
        item.group = group

    def split(self, key: str, offsets: list[int]) -> None:
        """在用户指定的字符边界拆分连续高亮，不猜测实体边界。"""
        item = next(o for o in self.occurrences if o.key == key)
        if not item.highlighted or not offsets or offsets != sorted(set(offsets)):
            raise SafetyError("请选择高亮条目，并填写递增且不重复的拆分位置。")
        if offsets[0] <= 0 or offsets[-1] >= len(item.text):
            raise SafetyError("拆分位置必须在文字内部。")
        groups = {o.text: o.group for o in self.occurrences
                  if o is not item and o.highlighted and not o.excluded}
        parts = []
        boundaries = [0, *offsets, len(item.text)]
        for start, end in zip(boundaries, boundaries[1:]):
            trimmed = _trim(item.text, start, end)
            if trimmed is None:
                continue  # 仅含空白的片段保留为普通文字
            start, end = trimmed
            text = item.text[start:end]
            if text not in groups:
                # 提前占用已分配编号，避免同一次拆分生成重复分组。
                groups[text] = self.new_group(tuple(groups.values()))
            group = groups[text]
            parts.append(Occurrence(item.part, item.paragraph, item.start + start,
                                    item.start + end, text, group, excluded=item.excluded))
        index = self.occurrences.index(item)
        self.occurrences[index:index + 1] = parts
        self.refresh_unmarked()

    def refresh_unmarked(self) -> None:
        previous = {o.key: o for o in self.occurrences if not o.highlighted}
        self.occurrences = [o for o in self.occurrences if o.highlighted]
        groups: dict[str, str] = {}
        for item in self.occurrences:
            if not item.excluded and item.text.strip():
                groups.setdefault(item.text, item.group)
        terms = sorted(groups, key=len, reverse=True)
        occupied_by_paragraph: dict[tuple[str, int], list[tuple[int, int]]] = {}
        for item in self.occurrences:
            occupied_by_paragraph.setdefault((item.part, item.paragraph), []).append((item.start, item.end))
        for paragraph in self.package.paragraphs():
            occupied = occupied_by_paragraph.setdefault((paragraph.part, paragraph.number), [])
            for start, end, term in term_matches(paragraph.text, terms):
                if any(start < b and end > a for a, b in occupied):
                    continue
                item = Occurrence(paragraph.part, paragraph.number, start, end,
                                  term, groups[term], False)
                old = previous.get(item.key)
                if old is not None and old.text == term:
                    item.group, item.excluded = old.group, old.excluded
                self.occurrences.append(item)
                occupied.append((start, end))


class Package:
    def __init__(self, data: bytes):
        self.parts: dict[str, bytes] = {}
        self.roots: dict[str, etree._Element] = {}
        if len(data) > MAX_PACKAGE:
            raise SafetyError("DOCX 超过 128 MB 容量限制。")
        try:
            with zipfile.ZipFile(BytesIO(data)) as archive:
                entries = archive.infolist()
                names = [entry.filename for entry in entries]
                if len(entries) > 2048 or len(names) != len(set(names)):
                    raise SafetyError("DOCX 包含重复条目或条目过多。")
                if sum(entry.file_size for entry in entries) > MAX_PACKAGE:
                    raise SafetyError("DOCX 解压后超过容量限制。")
                total_read = 0
                for entry in entries:
                    name = entry.filename
                    if (entry.flag_bits & 1 or entry.file_size > MAX_PART or
                            name.startswith("/") or ".." in PurePosixPath(name).parts or
                            "\\" in name):
                        raise SafetyError("DOCX 包含不安全或过大的内部条目。")
                    if entry.compress_type not in {zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED}:
                        raise SafetyError("DOCX 使用了不支持的压缩方式。")
                    if not entry.is_dir():
                        content = bytearray()
                        with archive.open(entry) as stream:
                            while chunk := stream.read(READ_CHUNK):
                                total_read += len(chunk)
                                if len(content) + len(chunk) > MAX_PART or total_read > MAX_PACKAGE:
                                    raise SafetyError("DOCX 实际解压数据超过容量限制。")
                                content.extend(chunk)
                        if len(content) != entry.file_size:
                            raise SafetyError("DOCX 内部条目的声明长度与读取长度不一致。")
                        self.parts[name] = bytes(content)
        except (zipfile.BadZipFile, RuntimeError, OSError, NotImplementedError):
            raise SafetyError("无法读取 DOCX，请检查文件格式与完整性。") from None
        for name, content in self.parts.items():
            if name.endswith((".xml", ".rels")) or name == "[Content_Types].xml":
                if b"<!DOCTYPE" in content.upper() or b"<!ENTITY" in content.upper():
                    raise SafetyError("DOCX 包含不允许的 XML 实体声明。")
                try:
                    root = etree.fromstring(content, etree.XMLParser(
                        resolve_entities=False, no_network=True, load_dtd=False,
                        huge_tree=False, remove_blank_text=False))
                    if root.getroottree().docinfo.doctype:
                        raise SafetyError("DOCX 包含不允许的 XML 文档类型声明。")
                    self.roots[name] = root
                except etree.XMLSyntaxError:
                    raise SafetyError("DOCX 的内部 XML 已损坏。") from None
        if "word/document.xml" not in self.roots or "[Content_Types].xml" not in self.roots:
            raise SafetyError("文件不是受支持的 DOCX。")
        if self.roots["word/document.xml"].tag != f"{{{W}}}document":
            raise SafetyError("仅支持常规 DOCX，Strict OOXML 请先在 Word 中另存为标准 DOCX。")
        try:
            Document(BytesIO(data))
        except Exception:
            raise SafetyError("DOCX 的关系或主文档结构无法通过 python-docx 检查。") from None
        self.styles = {}
        self._default_p_styles: tuple[str, ...] = ()
        self._style_highlight: dict[str | None, str | None] = {}
        self._default_highlight: str | None = None
        root = self.roots.get("word/styles.xml")
        if root is not None:
            self.styles = {s.get(f"{{{W}}}styleId"): s for s in root.findall("w:style", NS)}
            self._default_p_styles = tuple(
                key for key, style in self.styles.items()
                if style.get(f"{{{W}}}default") == "1" and style.get(f"{{{W}}}type") == "paragraph"
            )
            default_rpr = root.find("w:docDefaults/w:rPrDefault/w:rPr", NS)
            if default_rpr is not None:
                hl = default_rpr.find("w:highlight", NS)
                if hl is not None:
                    self._default_highlight = hl.get(f"{{{W}}}val")
            for style_id in self.styles:
                visited = set()
                curr = style_id
                while curr in self.styles and curr not in visited:
                    visited.add(curr)
                    style = self.styles[curr]
                    rpr = style.find("w:rPr", NS)
                    if rpr is not None:
                        hl = rpr.find("w:highlight", NS)
                        if hl is not None:
                            self._style_highlight[style_id] = hl.get(f"{{{W}}}val")
                            break
                    parent = style.find("w:basedOn", NS)
                    curr = parent.get(f"{{{W}}}val") if parent is not None else None

    def yellow(self, node: etree._Element, paragraph: etree._Element) -> bool:
        run = node.getparent()
        rpr = run.find(f"{{{W}}}rPr") if run is not None else None
        if rpr is not None:
            hl = rpr.find(f"{{{W}}}highlight")
            if hl is not None:
                return hl.get(f"{{{W}}}val") == "yellow"
            style_node = rpr.find(f"{{{W}}}rStyle")
            if style_node is not None:
                val = self._style_highlight.get(style_node.get(f"{{{W}}}val"))
                if val is not None:
                    return val == "yellow"
        ppr = paragraph.find(f"{{{W}}}pPr")
        p_style = ppr.find(f"{{{W}}}pStyle") if ppr is not None else None
        p_style_ids = (p_style.get(f"{{{W}}}val"),) if p_style is not None else self._default_p_styles
        for style_id in p_style_ids:
            val = self._style_highlight.get(style_id)
            if val is not None:
                return val == "yellow"
        return self._default_highlight == "yellow"

    def paragraphs(self) -> list[Paragraph]:
        result = []
        p_tag = f"{{{W}}}p"
        r_tag = f"{{{W}}}r"
        t_tag = f"{{{W}}}t"
        txbx_tag = f"{{{W}}}txbxContent"
        boundary_tags = {f"{{{W}}}{tag}" for tag in (
            "tab", "ptab", "br", "cr", "drawing", "pict", "object",
            "fldChar", "instrText", "delText", "sym", "noBreakHyphen",
            "softHyphen", "footnoteReference", "endnoteReference",
        )}
        for name, root in self.roots.items():
            if not STORY.fullmatch(name):
                continue
            for number, paragraph in enumerate(root.iter(p_tag)):
                # 文本框独立审查，修订等仍作为风险提示且参与残留扫描。
                if any(a.tag == txbx_tag for a in paragraph.iterancestors()):
                    continue
                p_iter = paragraph.iter(p_tag)
                next(p_iter, None)
                has_nested_p = next(p_iter, None) is not None
                chunks, spans, position = [], [], 0
                for node in paragraph.iter():
                    if has_nested_p and node is not paragraph:
                        nearest_p = next((a for a in node.iterancestors() if a.tag == p_tag), None)
                        if nearest_p is not paragraph:
                            continue
                    if node.tag == t_tag and node.getparent().tag == r_tag:
                        text = node.text or ""
                        spans.append(Span(position, position + len(text), node,
                                          self.yellow(node, paragraph)))
                    elif node.tag in boundary_tags:
                        text = "\ufffc"  # 结构边界，绝不跨越替换
                    else:
                        continue
                    chunks.append(text)
                    position += len(text)
                result.append(Paragraph(name, number, "".join(chunks), spans))
        return result

    def issues(self) -> list[Issue]:
        issues: list[Issue] = []
        supported_nodes = {span.node for p in self.paragraphs() for span in p.spans}
        tags = {
            "txbxContent": ("TEXTBOX", "存在文本框，需要在 Word 中人工审查。"),
            "ins": ("REVISION", "存在修订记录，请优先接受或拒绝修订后重新导入。"),
            "del": ("REVISION", "存在删除修订，其原文可能仍保存在文件内。"),
            "moveFrom": ("REVISION", "存在移动修订。"),
            "moveTo": ("REVISION", "存在移动修订。"),
            "vanish": ("HIDDEN", "存在隐藏文字或隐藏样式。"),
            "webHidden": ("HIDDEN", "存在网页隐藏文字。"),
            "sdt": ("CONTROL", "存在内容控件，需要人工检查绑定数据。"),
            "fldChar": ("FIELD", "存在域，Word 更新域可能改变显示内容。"),
            "fldSimple": ("FIELD", "存在简单域，需要人工检查。"),
            "altChunk": ("ALTCHUNK", "存在外部导入内容，无法可靠处理。"),
            "AlternateContent": ("ALTERNATE", "存在兼容性备用内容，需要人工审查。"),
            "dataBinding": ("BINDING", "存在数据绑定，脱敏时将清除绑定数据与关联。"),
            "oMath": ("MATH", "存在公式对象，需要人工检查其中的文字。"),
            "drawing": ("DRAWING", "存在绘图或图片，需要人工检查其中的文字与替代文本。"),
            "pict": ("DRAWING", "存在旧式绘图或形状，需要人工审查。"),
            "sym": ("SPECIAL_TEXT", "存在符号字体字符，需要人工审查。"),
            "noBreakHyphen": ("SPECIAL_TEXT", "存在特殊连字符，其两侧高亮分别识别，请人工核对。"),
            "softHyphen": ("SPECIAL_TEXT", "存在可选连字符，其两侧高亮分别识别，请人工核对。"),
        }
        for name, root in self.roots.items():
            for node in root.iter():
                if node.tag not in {f"{{{W}}}t", f"{{{W}}}delText", f"{{{W}}}instrText"}:
                    continue
                paragraph = next((a for a in node.iterancestors() if a.tag == f"{{{W}}}p"), None)
                if (node not in supported_nodes and node.text and paragraph is not None
                        and self.yellow(node, paragraph)):
                    issues.append(Issue("UNSUPPORTED_HIGHLIGHT",
                        "未支持的区域含黄色高亮文字，请在 Word 中将其移到正文或清理后重新导入。",
                        name, True))
                    break
            local_names = {etree.QName(n).localname for n in root.iter() if isinstance(n.tag, str)}
            for tag, (code, message) in tags.items():
                if tag in local_names:
                    issues.append(Issue(code, message, name))
            if "comment" in name.lower():
                issues.append(Issue("COMMENTS", "存在批注或批注人员信息，需要人工审查。", name))
            if name.endswith(".rels"):
                for rel in root:
                    if rel.get("TargetMode") == "External":
                        issues.append(Issue("EXTERNAL", "存在外部链接或模板关系，需要人工审查。", name))
                        break
            if name == "word/styles.xml":
                if root.xpath(".//w:style[@w:type='table']//w:highlight[@w:val='yellow']", namespaces=NS):
                    issues.append(Issue("TABLE_HIGHLIGHT", "表格样式含黄色高亮，需改为直接文本高亮后重新导入。", name, True))
            if not (STORY.fullmatch(name) or name in {
                "[Content_Types].xml", "word/styles.xml", "word/stylesWithEffects.xml",
                "word/settings.xml", "word/fontTable.xml", "word/numbering.xml",
                "word/webSettings.xml", "word/people.xml", "docProps/core.xml",
                "docProps/app.xml", "docProps/custom.xml"
            } or name.endswith(".rels") or name.startswith(("word/theme/", "customXml/"))):
                issues.append(Issue("OTHER_XML", "存在非主要文档 XML，需人工审查其中内容。", name))
        for name in self.parts:
            if name.startswith("word/printerSettings/"):
                continue
            if name.startswith(("word/media/", "word/embeddings/", "word/fonts/")):
                issues.append(Issue("BINARY", "存在图片、嵌入对象或嵌入字体，无法检查其中的文字。", name))
            elif name not in self.roots and not name.startswith("docProps/"):
                issues.append(Issue("OTHER_BINARY", "存在无法可靠检查的二进制内容。", name))
            if name.startswith("customXml/"):
                issues.append(Issue("CUSTOM_XML", "存在自定义数据，脱敏时会清除。", name))
            if name.startswith("_xmlsignatures/") or "vbaproject" in name.lower() or name.endswith(".bin"):
                issues.append(Issue("ACTIVE_CONTENT", "存在数字签名或活动/二进制对象，请在 Word 中移除后重试。", name, True))
        return list(dict.fromkeys(issues))

    def scrub(self) -> None:
        removed = {name for name in self.parts
                   if name.startswith(("docProps/", "customXml/", "word/printerSettings/"))
                   or name == "word/people.xml"}
        for name, root in self.roots.items():
            if name.endswith(".rels"):
                for rel in list(root):
                    kind = rel.get("Type", "")
                    deleted_target = (rel.get("TargetMode") != "External"
                                      and _relationship_target(name, rel.get("Target", "")) in removed)
                    if deleted_target or any(part in kind for part in (
                            "metadata/", "/extended-properties", "/custom-properties",
                            "/customXml", "/thumbnail", "/printerSettings", "/people")):
                        root.remove(rel)
            if name == "[Content_Types].xml":
                for item in list(root):
                    if item.get("PartName", "").lstrip("/") in removed:
                        root.remove(item)
            for element in list(root.iter()):
                if not isinstance(element.tag, str):
                    continue
                local = etree.QName(element).localname
                if local in {"docVars", "rsids", "rsid", "docId", "dataBinding", "printerSettings"}:
                    parent = element.getparent()
                    if parent is not None:
                        parent.remove(element)
                for attr in list(element.attrib):
                    qattr = etree.QName(attr)
                    if qattr.localname.startswith("rsid"):
                        del element.attrib[attr]
                    elif qattr.namespace in {W, None} and qattr.localname == "author":
                        element.attrib[attr] = ""
                    elif qattr.namespace in {W, None} and qattr.localname in {"initials", "date"}:
                        del element.attrib[attr]
        for name in removed:
            self.parts.pop(name, None)
            self.roots.pop(name, None)

    def serialize(self) -> bytes:
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, content in self.parts.items():
                if name in self.roots:
                    content = etree.tostring(self.roots[name], encoding="UTF-8", xml_declaration=True, standalone=True)
                # 创建新的 ZIP 条目，清空原 ZIP 注释、额外字段和文件时间。
                info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                archive.writestr(info, content)
        return buffer.getvalue()


def _trim(text: str, start: int, end: int) -> tuple[int, int] | None:
    """高亮首尾的空白不属于实体，留在原文中，同一实体因此始终使用同一个占位符。"""
    value = text[start:end]
    stripped = value.strip()
    if not stripped:
        return None
    start += len(value) - len(value.lstrip())
    return start, start + len(stripped)


def term_matches(text: str, terms: list[str]):
    """左起、最长优先匹配，避免包含关系引起的重复替换。"""
    if not terms:
        return
    pattern = re.compile("|".join(re.escape(term) for term in terms))
    for match in pattern.finditer(text):
        yield match.start(), match.end(), match.group()


def analyze(data: bytes) -> Analysis:
    package = Package(data)
    paragraphs = package.paragraphs()
    occurrences, groups = [], {}
    for paragraph in paragraphs:
        active: list[Span] = []
        def finish():
            if not active:
                return
            trimmed = _trim(paragraph.text, active[0].start, active[-1].end)
            if trimmed:
                start, end = trimmed
                text = paragraph.text[start:end]
                group = groups.setdefault(text, f"E{len(groups) + 1:03d}")
                occurrences.append(Occurrence(paragraph.part, paragraph.number, start, end, text, group))
            active.clear()
        for span in paragraph.spans:
            if span.start == span.end:
                continue
            if not span.yellow or (active and active[-1].end != span.start):
                finish()
            if span.yellow:
                active.append(span)
        finish()
    result = Analysis(package, occurrences, package.issues(), sha256(data).hexdigest())
    if any(TOKEN.search(p.text) or _malformed(p.text) for p in paragraphs):
        result.issues.append(Issue("TOKEN_COLLISION", "源文档已含 SECRET 标记，请先处理以避免占位符冲突。", blocking=True))
    result.refresh_unmarked()
    return result


def replace(paragraph: Paragraph, start: int, end: int, value: str, clear_yellow: bool = False) -> None:
    covered = [s for s in paragraph.spans if s.start < end and s.end > start]
    if not covered or sum(min(s.end, end) - max(s.start, start) for s in covered) != end - start:
        raise SafetyError("替换跨越了不支持的结构边界，已阻止导出。")
    # 新文字放进覆盖字符最多的 run，沿用其格式；并列时取第一个。Word 会把【、SECRET_001、】
    # 拆成三个 run，用户双击只选中中间部分加粗时，回填原文应随之加粗，而不是沿用括号的格式。
    target = max(range(len(covered)),
                 key=lambda i: (min(covered[i].end, end) - max(covered[i].start, start), -i))
    for index, span in enumerate(covered):
        text = span.node.text or ""
        left, right = max(start - span.start, 0), min(end - span.start, len(text))
        span.node.text = text[:left] + (value if index == target else "") + text[right:]
        span.node.set(XML_SPACE, "preserve")
        if clear_yellow:
            run = span.node.getparent()
            properties = run.find("w:rPr", NS)
            if properties is None:
                properties = etree.Element(f"{{{W}}}rPr")
                run.insert(0, properties)
            highlight = properties.find("w:highlight", NS)
            if highlight is None:
                highlight = etree.SubElement(properties, f"{{{W}}}highlight")
            highlight.set(f"{{{W}}}val", "none")


def check_issues(issues: list[Issue], acknowledged: bool) -> None:
    if any(i.blocking for i in issues):
        raise SafetyError("存在必须修复的结构或安全问题，请查看检查列表。")
    if issues and not acknowledged:
        raise SafetyError("存在无法可靠检查的内容，请阅读并明确确认风险后再导出。")


FORMATTING_ELEMENTS = frozenset("""
    sz szCs spacing ind position kern w tcW tblW tblInd trHeight gridCol pgSz pgMar
    cols col docGrid tabs tab tblLook tblLayout tblCellMar top bottom left right start end
    b i bCs iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof
    snapToGrid vanish webHidden color highlight u vertAlign shd textEffect textScale
    fitText lang rtl cs rFonts font altName charset family pitch sig panose1
    basedOn next link pStyle rStyle tblStyle numStyleLink styleLink uiPriority qFormat
    semiHidden unhideWhenUsed locked autoRedefine hidden contextualSpacing
    keepNext keepLines pageBreakBefore widowControl suppressAutoHyphens
    suppressLineNumbers suppressOverlap adjustRightInd mirrorIndents jc textDirection
    textAlignment outlineLvl ilvl numId abstractNumId lvl lvlOverride startOverride
    numFmt lvlJc suff multiLevelType nsid tmpl isLgl lvlRestart
    tblBorders tcBorders pBdr bdr insideH insideV tblCellSpacing cellSpacing tcFitText tcNoWrap
    tblStyleRowBandSize tblStyleColBandSize abstractNum num
    vAlign vMerge hMerge cantSplit tblHeader tblPrChange pPrChange rPrChange
    footnotePr endnotePr footnote footnoteReference endnote endnoteReference
    separator continuationSeparator comment commentRangeStart commentRangeEnd commentReference
    gridSpan gridBefore gridAfter pgNumType hyphenationZone drawingGridHorizontalSpacing
    drawingGridVerticalSpacing displayHorizontalDrawingGridEvery displayVerticalDrawingGridEvery
    bookmarkStart bookmarkEnd proofErr lastRenderedPageBreak
    defaultTabStop zoom characterSpacingControl view compatSetting latentStyles lsdException
""".split())


WORD_IDENTIFIER_ATTRIBUTES = {
    "http://schemas.microsoft.com/office/word/2010/wordml": {"paraId", "textId"},
    "http://schemas.microsoft.com/office/word/2012/wordml": {"paraId", "paraIdParent", "textId"},
    "http://schemas.microsoft.com/office/word/2016/wordml/cid": {"paraId", "durableId"},
}

FORMATTING_ENUM_ATTRIBUTES = {
    ("headerReference", "type"): {"default", "first", "even"},
    ("footerReference", "type"): {"default", "first", "even"},
    ("jc", "val"): {"left", "right", "center", "both", "start", "end", "distribute",
                    "mediumKashida", "highKashida", "lowKashida", "numTab", "thaiDistribute"},
    ("lvlJc", "val"): {"left", "right", "center", "both", "start", "end"},
    ("spacing", "lineRule"): {"auto", "atLeast", "exact"},
    ("tab", "val"): {"clear", "left", "right", "center", "decimal", "bar", "num", "start", "end"},
    ("multiLevelType", "val"): {"singleLevel", "multilevel", "hybridMultilevel"},
    ("u", "val"): {"single", "words", "double", "thick", "dotted", "dottedHeavy",
                   "dash", "dashedHeavy", "dashLong", "dashLongHeavy", "dotDash", "dashDotHeavy",
                   "dotDotDash", "dashDotDotHeavy", "wave", "wavyHeavy", "wavyDouble", "none"},
    ("docGrid", "type"): {"default", "lines", "linesAndChars", "snapToChars"},
    ("tblLayout", "type"): {"autofit", "fixed"},
    ("vAlign", "val"): {"top", "center", "both", "bottom"},
    ("vertAlign", "val"): {"baseline", "superscript", "subscript"},
    ("textDirection", "val"): {"lrTb", "tbRl", "btLr", "lrTbV", "tbRlV", "tbLrV"},
    ("br", "type"): {"page", "column", "textWrapping"},
    ("br", "clear"): {"none", "left", "right", "all"},
    ("highlight", "val"): {"black", "blue", "cyan", "green", "magenta", "red", "yellow",
                           "white", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed",
                           "darkYellow", "darkGray", "lightGray", "none"},
}
for _element in ("tcW", "tblW", "tblInd", "tblCellSpacing", "top", "bottom", "left", "right", "start", "end"):
    FORMATTING_ENUM_ATTRIBUTES[(_element, "type")] = {"auto", "dxa", "nil", "pct"}
for _element in ("top", "bottom", "left", "right", "start", "end", "insideH", "insideV", "bdr", "tl2br", "tr2bl"):
    FORMATTING_ENUM_ATTRIBUTES[(_element, "val")] = {
        "nil", "none", "single", "thick", "double", "dotted", "dashed", "dotDash", "dotDotDash",
        "triple", "thinThickSmallGap", "thickThinSmallGap", "thinThickThinSmallGap",
        "thinThickMediumGap", "thickThinMediumGap", "thinThickThinMediumGap",
        "thinThickLargeGap", "thickThinLargeGap", "thinThickThinLargeGap", "wave",
        "doubleWave", "dashSmallGap", "dashDotStroked", "threeDEmboss", "threeDEngrave", "outset", "inset",
    }


def _relationship_target(rels_part: str, target: str) -> str:
    directory = posixpath.dirname(posixpath.dirname(rels_part))
    return posixpath.normpath(posixpath.join(directory, target)).lstrip("/")


def _standard_part(name: str) -> bool:
    return bool(re.fullmatch(
        r"\[Content_Types\]\.xml|_rels/\.rels|"
        r"word/(document|styles|stylesWithEffects|settings|fontTable|numbering|webSettings|"
        r"header\d+|footer\d+|footnotes|endnotes|comments\d*)\.xml|"
        r"word/_rels/(document|header\d+|footer\d+|footnotes|endnotes|numbering|fontTable)\.xml\.rels|"
        r"word/theme/theme\d+\.xml|word/media/image\d+\.[a-zA-Z0-9]+", name))


def _structural_attribute(part: str, node: etree._Element, attr: str, value: str) -> bool:
    """仅排除已知格式/结构值，正文、书签名称、说明和自定义属性仍检查。"""
    element, attribute = etree.QName(node), etree.QName(attr)
    local = element.localname
    if attribute.localname == "uri" and (
            value.startswith(("http://schemas.openxmlformats.org/", "http://schemas.microsoft.com/"))
            or re.fullmatch(r"\{[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}", value)):
        return True
    if element.namespace in DRAWING_LAYOUT_NS:
        if attribute.localname in {"name", "descr", "title"}:
            return False
        if attribute.localname == "relativeFrom":
            return bool(re.fullmatch(r"[A-Za-z]+", value))
        return bool(re.fullmatch(r"[-+]?\d+(?:\.\d+)?|[0-9A-Fa-f]{8}|true|false|on|off", value))
    if attribute.namespace == DOCUMENT_RELATIONSHIPS_NS and attribute.localname in {"id", "embed", "link"}:
        return bool(re.fullmatch(r"rId\d+", value))
    if attribute.localname in WORD_IDENTIFIER_ATTRIBUTES.get(attribute.namespace, ()):
        return bool(re.fullmatch(r"[0-9A-Fa-f]{8}", value))
    if attribute.namespace == "http://schemas.openxmlformats.org/markup-compatibility/2006":
        return attribute.localname in {"Ignorable", "MustUnderstand", "ProcessContent"}
    if element.namespace == W and attribute.namespace == W:
        if value in FORMATTING_ENUM_ATTRIBUTES.get((local, attribute.localname), ()):
            return True
        if (local in {"lang", "themeFontLang"} and attribute.localname in {"val", "eastAsia", "bidi"}
                and re.fullmatch(r"[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*|x-none", value)):
            return True
        if local == "rFonts" and attribute.localname == "hint":
            return value in {"default", "eastAsia", "cs"}
        if local in {"rFonts", "font", "altName"}:
            return attribute.localname in {"ascii", "hAnsi", "eastAsia", "cs", "name", "val",
                                           "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme"}
        if local == "lsdException" or local == "compatSetting":
            return attribute.localname in {"name", "uri", "val", "locked", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat"}
        if local == "style" and node.get(f"{{{W}}}customStyle") not in {"1", "true", "on"}:
            return attribute.localname in {"styleId", "default", "customStyle"}
        if local == "style" and attribute.localname in {"default", "customStyle"}:
            return value in {"0", "1", "true", "false", "on", "off"}
        if local == "style" and attribute.localname == "styleId" and re.fullmatch(r"Heading[1-9]Char", value):
            return True
        if local == "name" and node.getparent() is not None and node.getparent().tag == f"{{{W}}}style":
            return (node.getparent().get(f"{{{W}}}customStyle") not in {"1", "true", "on"}
                    or bool(re.fullmatch(r"Heading [1-9] Char", value)))
        if local == "clrSchemeMapping":
            return bool(re.fullmatch(r"light[12]|dark[12]|accent[1-6]|hyperlink|followedHyperlink", value))
        if local == "tblStylePr" and attribute.localname == "type":
            return value in {"wholeTable", "firstRow", "lastRow", "firstCol", "lastCol",
                             "band1Vert", "band2Vert", "band1Horz", "band2Horz", "neCell", "nwCell", "seCell", "swCell"}
        if local == "lvlText" and attribute.localname == "val":
            return bool(re.fullmatch(r"(?:%[1-9]|[\s.()、\-])+", value))
        if local in {"color", "shd", "top", "bottom", "left", "right", "insideH", "insideV"}:
            if attribute.localname in {"themeColor", "themeFill"}:
                return bool(re.fullmatch(r"accent[1-6]|text[12]|background[12]|dark[12]|light[12]|hyperlink|followedHyperlink", value))
        if local in FORMATTING_ELEMENTS:
            # 字体、样式引用为技术标识。其余仅跳过数字、颜色值及格式枚举。
            if local in {"basedOn", "next", "link", "pStyle", "rStyle", "tblStyle", "numStyleLink", "styleLink"}:
                return attribute.localname == "val"
            if attribute.localname in {"name", "instr", "author", "date", "descr", "title"}:
                return False
            return bool(re.fullmatch(r"[0-9A-Fa-f]+|[-+]?\d+(?:\.\d+)?|auto|none|true|false", value))
    if element.namespace == DRAWING_NS:
        if local in {"latin", "ea", "cs", "font"} and attribute.localname in {"typeface", "script", "pitchFamily", "charset"}:
            return True
        if local == "prstDash" and attribute.localname == "val":
            return True
        if local == "schemeClr" and attribute.localname == "val":
            return bool(re.fullmatch(r"accent[1-6]|tx[12]|bg[12]|dk[12]|lt[12]|hlink|folHlink|phClr", value))
        return local not in {"t", "fld"} and bool(re.fullmatch(r"[-+]?\d+(?:\.\d+)?|[0-9A-Fa-f]{6}", value))
    if element.namespace == "http://schemas.openxmlformats.org/officeDocument/2006/math":
        if attribute.namespace == element.namespace and local == "defJc" and attribute.localname == "val":
            return value in {"left", "right", "center", "centerGroup"}
        return local in {"smallFrac", "lMargin", "rMargin", "wrapIndent"} and value.isdecimal()
    if element.namespace == "urn:schemas-microsoft-com:office:office":
        return (local, attribute.localname) in {("shapedefaults", "spidmax"), ("idmap", "data")} and value.isdecimal()
    if element.namespace == "http://schemas.microsoft.com/office/word/2010/wordml":
        return local == "defaultImageDpi" and attribute.localname == "val" and value.isdecimal()
    if element.namespace == R and local == "Relationship":
        if attribute.localname in {"Id", "TargetMode"}:
            return True
        if attribute.localname == "Type":
            return value.startswith(("http://schemas.openxmlformats.org/", "http://schemas.microsoft.com/"))
        if attribute.localname == "Target" and node.get("TargetMode") != "External":
            return _standard_part(_relationship_target(part, value))
    if element.namespace == "http://schemas.openxmlformats.org/package/2006/content-types":
        if attribute.localname == "PartName":
            return _standard_part(value.lstrip("/"))
        if attribute.localname == "Extension":
            return bool(re.fullmatch(r"[A-Za-z0-9]{1,10}", value))
        if attribute.localname == "ContentType":
            return value.startswith(("application/vnd.openxmlformats-", "application/xml", "image/"))
    return False


def _xml_text_blocks(root: etree._Element) -> str:
    chunks: list[str] = []
    boundary_tags = {f"{{{W}}}{tag}" for tag in ("p", "tc", "tr", "br", "cr", "tab", "ptab")}
    for event, node in etree.iterwalk(root, events=("start", "end")):
        if event == "start":
            if node.tag in boundary_tags:
                chunks.append("\ufffc")
            qname = etree.QName(node) if isinstance(node.tag, str) else None
            layout = (qname is not None and qname.namespace in DRAWING_LAYOUT_NS
                      and qname.localname in DRAWING_LAYOUT_TEXT)
            if node.text and not layout:
                chunks.append(node.text)
        else:
            if node.tag in boundary_tags:
                chunks.append("\ufffc")
            if node.tail:
                chunks.append(node.tail)
    return "".join(chunks)


def residual_parts(package: Package, terms: list[str]) -> list[str]:
    """按内容语义检查 XML，避免把命名空间、标签名和格式数字误当原文。"""
    found = set()
    for name, data in package.parts.items():
        haystacks = [] if _standard_part(name) else [name]
        root = package.roots.get(name)
        if root is not None:
            haystacks.append(_xml_text_blocks(root))
            haystacks.extend(value for node in root.iter() if isinstance(node.tag, str)
                             for attr, value in node.attrib.items()
                             if not _structural_attribute(name, node, attr, value))
            haystacks.extend(str(node) for node in root.xpath("//comment() | //processing-instruction()"))
            # XML 的内容已解码，原始字节扫描仅用于非 XML 部件。
            data = b""
        for term in terms:
            if any(term in text for text in haystacks) or any(
                    len(encoded) >= MIN_BINARY_MATCH and encoded in data
                    for encoded in (term.encode(e) for e in ("utf-8", "utf-16le", "utf-16be"))):
                found.add(name)
                break
    return sorted(found)


def _validate_plan(analysis: Analysis) -> None:
    paragraphs = {(p.part, p.number): p for p in analysis.package.paragraphs()}
    assigned, intervals = {}, {}
    for item in analysis.occurrences:
        p = paragraphs.get((item.part, item.paragraph))
        if (p is None or not item.group.strip() or not item.text or item.start < 0 or
                item.end <= item.start or p.text[item.start:item.end] != item.text):
            raise SafetyError("预览内容或分组无效，请重新导入。")
        if item.group in assigned and assigned[item.group] != item.text:
            raise SafetyError("不同文字不能共用同一分组，请修改分组。")
        assigned[item.group] = item.text
        key = (item.part, item.paragraph)
        for start, end in intervals.setdefault(key, []):
            if item.start < end and item.end > start:
                raise SafetyError("脱敏范围存在重叠，请重新导入。")
        intervals[key].append((item.start, item.end))


def redact(analysis: Analysis, password: str, *, acknowledged: bool = False,
           include_unmarked: bool = True, confirmed_exclusions: bool = False) -> tuple[bytes, bytes, dict]:
    # 自定义 XML 将完整移除，其存在本身只作清理通知。
    check_issues([i for i in analysis.issues if i.code != "CUSTOM_XML"], acknowledged)
    _validate_plan(analysis)
    # 每次从原始内存包克隆，失败与重试不会修改预览状态。
    package = Package(analysis.package.serialize())
    paragraphs = {(p.part, p.number): p for p in package.paragraphs()}
    selected = [o for o in analysis.occurrences
                if not o.excluded and (o.highlighted or include_unmarked)]
    excluded = [o for o in analysis.occurrences
                if o.excluded or (not o.highlighted and not include_unmarked)]
    if excluded and not confirmed_exclusions:
        raise SafetyError("预览中有保留原文的位置，必须单独确认这些保留位置后才能导出。")
    if not selected:
        raise SafetyError("未发现需要替换的黄色文本高亮，请先在 Word 中标记并核对保留条目。")
    # 扫描副本只遮蔽本次已替换位置和人工明确确认的保留位置。
    # 输出副本照常写入占位符，保留项仍保存原文。
    scan_package = Package(analysis.package.serialize())
    scan_paragraphs = {(p.part, p.number): p for p in scan_package.paragraphs()}
    for item in sorted(selected + excluded, key=lambda o: (o.part, o.paragraph, o.start), reverse=True):
        replace(scan_paragraphs[(item.part, item.paragraph)], item.start, item.end, "\ufffc")
    group_tokens, entries = {}, {}
    for item in selected:
        token = group_tokens.setdefault(item.group, f"【SECRET_{len(group_tokens) + 1:03d}】")
        entry = entries.setdefault(token, {"text": item.text, "count": 0})
        entry["count"] += 1
    for item in sorted(selected, key=lambda o: (o.part, o.paragraph, o.start), reverse=True):
        replace(paragraphs[(item.part, item.paragraph)], item.start, item.end,
                group_tokens[item.group], clear_yellow=True)
    package.scrub()
    scan_package.scrub()
    # 除用户逐位置确认的保留项外，已知正文和非结构属性残留仍为硬性阻止。
    terms = sorted({o.text for o in selected if o.text.strip()}, key=len, reverse=True)
    residue = residual_parts(scan_package, terms)
    if residue:
        raise SafetyError("检测到已标记文字仍残留在 DOCX 内部，已阻止导出。请处理未标记位置、隐藏内容或内部属性。涉及部件数 " + str(len(residue)))
    output = package.serialize()
    checked = Package(output)
    counts = Counter(m.group() for p in checked.paragraphs() for m in TOKEN.finditer(p.text))
    if counts != Counter({token: item["count"] for token, item in entries.items()}):
        raise SafetyError("占位符计数校验失败，已阻止导出。")
    mapping = {"version": 1, "source_sha256": analysis.source_hash,
               "redacted_sha256": sha256(output).hexdigest(), "entries": entries}
    encrypted = encrypt_mapping(mapping, password)
    report = {"operation": "脱敏", "entity_count": len(entries), "replacement_count": sum(counts.values()),
              "unmarked_count": sum(not o.highlighted for o in selected),
              "excluded_count": len(excluded), "exclusions_confirmed": bool(excluded) and confirmed_exclusions,
              "excluded_locations": [{"part": public_part(o.part), "paragraph": o.paragraph + 1,
                                      "start": o.start, "end": o.end} for o in excluded],
              "metadata_cleared": True, "package_valid": True,
              "risk_acknowledged": acknowledged, "issues": [i.public() for i in analysis.issues],
              "notice": "仅处理人工标记及其相同文字。无法自动发现全部未标记敏感信息。"}
    return output, encrypted, report


def _mapping_entries(mapping: dict) -> dict:
    entries = mapping.get("entries")
    if not isinstance(entries, dict) or not entries or len(entries) > 100000:
        raise MappingError("映射表条目无效。")
    for token, item in entries.items():
        if (not isinstance(token, str) or not TOKEN.fullmatch(token) or not isinstance(item, dict)
                or not isinstance(item.get("text"), str) or not item["text"]
                or not isinstance(item.get("count"), int) or isinstance(item["count"], bool)
                or item["count"] < 1 or TOKEN.search(item["text"])):
            raise MappingError("映射表包含无效条目。")
    return entries


def _malformed(text: str) -> list[dict]:
    # 等长遮蔽保留原始字符偏移。普通英文 SECRET/secretary 不属于占位符。
    remainder = TOKEN.sub(lambda m: " " * len(m.group()), text)
    gap = r"[\s\ufffc\u200b-\u200f\ufeff]*"
    tight_gap = r"[\ufffc\u200b-\u200f\ufeff]*"
    separated_secret = gap.join("SECRET")
    identifier = (
        separated_secret
        + r"(?:" + gap + r"[_＿]" + gap + r"[0-9０-９]+"
        + r"|" + tight_gap + r"[0-9０-９]{2,}"
        + r"|" + gap + r"[0-9０-９]{2,}" + gap + r"[】\]])"
    )
    typo = r"[【\[]\s*SE(?!CRET)[A-Z\s]{2,10}[_＿]\s*[0-9０-９]+\s*[】\]]?"
    no_number = r"[【\[]\s*" + separated_secret + gap + r"[_＿]?" + gap + r"[】\]]"
    pattern = re.compile(identifier + "|" + typo + "|" + no_number, re.I)
    return [{"offset": m.start(), "kind": "疑似被改写或拆开的占位符"}
            for m in pattern.finditer(remainder)]


def restore(data: bytes, encrypted: bytes, password: str, *, acknowledged: bool = False) -> tuple[bytes, dict]:
    entries = _mapping_entries(decrypt_mapping(encrypted, password))
    package = Package(data)
    issues = package.issues()
    check_issues(issues, acknowledged)
    paragraphs = package.paragraphs()
    counts = Counter(m.group() for p in paragraphs for m in TOKEN.finditer(p.text))
    all_tokens = set(counts)
    for root in package.roots.values():
        internal_text = "".join(root.itertext())
        all_tokens.update(TOKEN.findall(internal_text))
        for node in root.iter():
            if isinstance(node.tag, str):
                for value in node.attrib.values():
                    all_tokens.update(TOKEN.findall(value))
    unknown = sorted(all_tokens - set(entries))
    missing = sorted(set(entries) - set(counts))
    changed = [{"token": t, "expected": e["count"], "actual": counts[t]}
               for t, e in entries.items() if counts[t] != e["count"]]
    malformed = [{"part": public_part(p.part), "paragraph": p.number + 1, **item}
                 for p in paragraphs for item in _malformed(p.text)]
    # 跨段落断开通常无法自动修复，但仍纳入异常检测。
    for name in {p.part for p in paragraphs}:
        combined = "\ufffc".join(p.text for p in paragraphs if p.part == name)
        local_count = sum(len(_malformed(p.text)) for p in paragraphs if p.part == name)
        if len(_malformed(combined)) > local_count:
            malformed.append({"part": public_part(name), "kind": "疑似跨段落拆开的占位符"})
    restored = Counter()
    for paragraph in paragraphs:
        for match in reversed(list(TOKEN.finditer(paragraph.text))):
            token = match.group()
            if token in entries:
                replace(paragraph, match.start(), match.end(), entries[token]["text"])
                restored[token] += 1
    # 非主文档区域的标记保留并报告，不静默替换批注/修订/属性。
    remaining = []
    for name, root in package.roots.items():
        text = "".join(root.itertext()) + " ".join(v for n in root.iter()
                if isinstance(n.tag, str) for v in n.attrib.values())
        found = sorted(set(TOKEN.findall(text)))
        suspicious = bool(_malformed(text))
        if found or suspicious:
            remaining.append({"part": public_part(name), "tokens": found, "suspicious": suspicious})
    output = package.serialize()
    Package(output)
    placeholder_status = "需人工检查" if unknown or missing or changed or malformed or remaining else "通过"
    status = placeholder_status
    if placeholder_status == "通过" and issues:
        status = "回填通过，文档需审查"
    report = {"operation": "回填", "package_valid": True,
              "restored_count": sum(restored.values()), "unknown_tokens": unknown,
              "missing_tokens": missing, "count_changes": changed,
              "malformed_tokens": malformed, "remaining_tokens": remaining,
              "issues": [i.public() for i in issues], "risk_acknowledged": acknowledged,
              "status": status, "placeholder_status": placeholder_status,
              "document_review_required": bool(issues),
              "notice": "消失的占位符仅列入报告，不猜测恢复位置。结构检查通过不等同于 Microsoft Word 实际打开验证。"}
    return output, report


def report_bytes(report: dict) -> bytes:
    return json.dumps(report, ensure_ascii=False, indent=2).encode("utf-8")


def export_redacted(source: Path, output: Path, mapping: Path, report_path: Path,
                    analysis: Analysis, password: str, report_language: str = "zh", **options) -> dict:
    if len({p.resolve() for p in (output, mapping, report_path)}) != 3:
        raise OutputError("三个输出文件路径必须不同。")
    if sha256(source.read_bytes()).hexdigest() != analysis.source_hash:
        raise SafetyError("源文件已发生变化，请重新导入并确认预览。")
    result, encrypted, report = redact(analysis, password, **options)
    from .i18n import localized_report
    write_outputs({output: result, mapping: encrypted,
                   report_path: report_bytes(localized_report(report, report_language))}, (source,))
    return report


def export_restored(source: Path, mapping: Path, output: Path, report_path: Path,
                    password: str, report_language: str = "zh",
                    expected_source_hash: str | None = None, **options) -> dict:
    if output.resolve() == report_path.resolve():
        raise OutputError("文档与报告的输出路径必须不同。")
    source_bytes = source.read_bytes()
    if expected_source_hash is not None and sha256(source_bytes).hexdigest() != expected_source_hash:
        raise SafetyError("源文件已发生变化，请重新导入并确认预览。")
    result, report = restore(source_bytes, mapping.read_bytes(), password, **options)
    from .i18n import localized_report
    write_outputs({output: result,
                   report_path: report_bytes(localized_report(report, report_language))}, (source, mapping))
    return report
