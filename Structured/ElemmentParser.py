import re
import os
import time
from .Elements import (
    Toc,
    Title,
    TableTitle,
    ImageTitle,
    DynamicTitle,
    Paragraph,
    BlankPara,
    Table,
    Element
)
from typing import Any, Dict, List, Optional, Iterator
from .Metadata import FileMetadata, ElemMetadata
from dataclasses import field

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table as _DocxTable
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .utils import (
    convert_doc_to_docx,
    table_to_markdown,
)


def parse_docx(
        file_path: str,
        parse_table_as_markdown=True
):
    """
    解析 Word 文档（支持 .doc 和 .docx），提取结构化元素。

    支持将表格内容转换为 markdown 格式（可选）。

    参数:
    ----------
    file_path : str
        Word 文件的完整路径。支持 .doc 或 .docx 文件。
    parse_table_as_markdown : bool, default=True
        是否将表格转换为 markdown 格式。如果为 False，表格以纯文本字符串返回。

    返回:
    ----------
    List[Element]
        返回一个由结构化元素组成的列表，元素类型包括 Paragraph、Title、Table 等。

    异常:
    ----------
    FileNotFoundError:
        当文件路径不存在时抛出。
    IsADirectoryError:
        当路径是文件夹而不是文件时抛出。
    TypeError:
        当文件类型既不是 .doc 也不是 .docx 时抛出。
    """
    # 路径存在性和合法性校验
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在：{file_path}")
    if not os.path.isfile(file_path):
        raise IsADirectoryError(f"该路径是文件夹，不是文件: {file_path}")

     # 检查文件类型并转换 .doc -> .docx
    _, file_extension = os.path.splitext(file_path)
    if file_extension == ".doc":
        print(1)
        file_path = convert_doc_to_docx(file_path)
    elif file_extension == ".docx":
        pass
    else:
        raise TypeError(f"不支持处理 {file_extension} 类型文件")

    # 提取元信息
    dir_name = os.path.dirname(file_path)
    file_name = os.path.basename(file_path)

    timestamp = os.path.getmtime(file_path)
    last_modified_time = time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime(timestamp))

    # 构建文档级元数据
    file_metadata = FileMetadata(
        filename=file_name,
        file_directory=dir_name,
        filetype=".docx",
        last_modified=last_modified_time,
    )

    # 解析文档元素，返回迭代器并转换为列表
    elements = _DocxParser.iter_document_elements(
        file_path,
        file_metadata,
        table_as_markdown=parse_table_as_markdown
    )
    return list(elements)


class _DocxParser:
    def __init__(
            self,
            docx_file: Document,
            file_metadata: FileMetadata,
            table_as_markdown: bool = True
    ):
        """
        初始化 _DocxParser 实例，用于解析 Word 文档结构。

        参数:
        ----------
        docx_file : Document
            由 `python-docx` 加载后的 Word 文档对象，表示一个 .docx 文件。
        file_metadata : FileMetadata
            当前文档的文件级元数据（如文件名、路径、修改时间等）。
        table_as_markdown : bool, default=True
            是否将表格内容解析为 Markdown 字符串。如果为 False，保留结构化表格元素。

        成员变量:
        ----------
        self.docx_file : Document
            存储传入的 Word 文档对象，供后续遍历处理。
        self.styles_map : Dict[str, Style]
            样式字典，key 是样式 ID，value 是对应的样式对象，用于辅助判断元素类型。
        self.file_metadata : FileMetadata
            存储文档级元信息，供各元素关联。
        self.table_as_markdown : bool
            控制表格解析格式的开关，在解析过程中判断是否需要转为 Markdown。
        """
        self.docx_file = docx_file
        self.styles_map = {s.style_id: s for s in self.docx_file.styles}
        self.file_metadata = file_metadata
        self.table_as_markdown = table_as_markdown

    @classmethod
    def iter_document_elements(
            cls,
            file_path: Optional[str] = None,
            file_metadata: FileMetadata = field(default_factory=FileMetadata),
            table_as_markdown=True
    ) -> Iterator[Element]:
        """
        以类方法形式初始化 _DocxParser，并迭代文档中的结构化元素。

        参数:
        ----------
        file_path : Optional[str]
            Word 文件的路径，必须是 .docx 格式。
        file_metadata : FileMetadata
            文件的元信息（如文件名、目录、最后修改时间等）。
        table_as_markdown : bool
            是否将表格转换为 Markdown 格式。

        返回:
        ----------
        Iterator[Element]
            逐个返回文档中的结构化元素，包括段落、表格、控件等。
        """
        instance = cls(
            docx_file=Document(file_path),
            file_metadata=file_metadata,
            table_as_markdown=table_as_markdown
        )

        return instance._iter_document_elements()

    def _iter_document_elements(self) -> Iterator[Element]:
        """
        遍历 docx 文件的 <w:body>，逐个处理并返回结构化元素。

        包含处理：
        - 普通段落（<w:p>）
        - 表格（<w:tbl>）
        - 内容控件（<w:sdt>），包括其嵌套的段落和表格

        返回:
        ----------
        Iterator[Element]
            通过 yield 方式逐个返回结构化元素对象。
        """
        body = self.docx_file.element.body

        for element in body:
            if isinstance(element, CT_P):
                # 普通段落
                yield from self._process_paragraph_block(element)
            elif isinstance(element, CT_Tbl):
                # 表格
                yield from self._process_table_block(element)
            elif element.tag == qn("w:sdt"):
                # 内容控件（Smart Document Tag）
                sdt_content = element.find(qn("w:sdtContent"))
                if sdt_content is None:
                    continue

                # 遍历控件内部元素
                for child in sdt_content:
                    if isinstance(child, CT_P):
                        yield from self._process_paragraph_block(child)
                    elif isinstance(child, CT_Tbl):
                        yield from self._process_table_block(child)
                    # 控件内嵌套控件，递归解析
                    elif child.tag == qn("w:sdt"):
                        inner = child.find(qn("w:sdtContent"))
                        if inner is None:
                            continue
                        for grand in inner:
                            if isinstance(grand, CT_P):
                                yield from self._process_paragraph_block(grand)
                            elif isinstance(grand, CT_Tbl):
                                yield from self._process_table_block(grand)

    def _process_paragraph_block(self, element):
        """
        处理 <w:p> 段落元素，将其解析为对应的结构化段落类型。

        根据段落样式内容判断该段落是空段、标题、目录、表格标题、图片标题还是普通段落。
        每种类型将返回对应的 Element 子类对象（通过 yield 逐个输出）。

        参数:
        ----------
        element : CT_P
            Word 文档中的段落对象，对应 XML 的 <w:p> 标签。

        产出:
        ----------
        Iterator[Element]
            根据段落类型返回 BlankPara、Toc、TableTitle、ImageTitle、Title 或 Paragraph 实例。
        """
        # 获取样式名，默认为"Normal"
        name = "Normal"

        # 提取段落的样式 ID
        pPr = element.find(qn("w:pPr"))
        if pPr is not None:
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                style_id = pStyle.get(qn("w:val"), "")
                style = self.styles_map.get(style_id)
                if style:
                    # 样式名或基样式名
                    name = style.name if style.name else "Normal"

        # 空段落检测
        if self._is_blank_para_element(element):
            yield BlankPara(
                text=element.text,
                file_metadata=self.file_metadata,
                elem_metadata=ElemMetadata(
                    style_name=name
                )
            )
        # 目录（TOC）段落检测
        elif self._is_toc_element(element):
            yield Toc(
                text=element.text,
                file_metadata=self.file_metadata,
                elem_metadata=ElemMetadata(
                    style_name=name
                )
            )
        # 表格标题检测
        elif self._is_table_title_element(element):
            yield TableTitle(
                text=element.text,
                file_metadata=self.file_metadata,
                elem_metadata=ElemMetadata(
                    style_name=name
                )
            )
        # 图片标题检测
        elif self._is_image_title_element(element):
            yield ImageTitle(
                text=element.text,
                file_metadata=self.file_metadata,
                elem_metadata=ElemMetadata(
                    style_name=name
                )
            )
        # 动态标题检测
        elif self._is_dynamic_title_element(element):
            yield DynamicTitle(
                text=element.text.strip(),
                file_metadata=self.file_metadata,
                elem_metadata=ElemMetadata(
                    style_name=name
                )
            )
        # 普通标题检测（带编号的标题行）
        elif self._is_title_element(element):
            yield Title(
                text=element.text,
                level=self._get_title_level(element),
                file_metadata=self.file_metadata,
                elem_metadata=ElemMetadata(
                    style_name=name
                )
            )
        
        # 默认剩余段落
        else:
            yield Paragraph(
                text=element.text,
                file_metadata=self.file_metadata,
                elem_metadata=ElemMetadata(
                    style_name=name
                )
            )

    def _process_table_block(self, element):
        """
        处理 <w:tbl> 表格元素，提取其文本内容并返回结构化的 Table 元素。

        参数:
        ----------
        element : CT_Tbl
            Word 文档中的表格对象，对应 XML 的 <w:tbl> 标签。

        产出:
        ----------
        Iterator[Table]
            返回提取后的 Table 元素对象，包含表格的纯文本和（可选）Markdown 表达。
        """
        docx_table = _DocxTable(element, self.docx_file)

        # 提取每行、每个单元格的文本
        rows = []
        for row in docx_table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            rows.append(cells)

        # 拼成一个字符串表示，可用制表符和换行分隔
        table_str = "\n".join("\t".join(cells) for cells in rows)

        yield Table(
            # text=table_str,
            text=table_str,
            file_metadata=self.file_metadata,
            text_as_markdown=table_to_markdown(docx_table) if self.table_as_markdown else ""
        )

    def _is_table_title_element(self, element):
        """
        判断段落是否为“表格标题”。

        表格标题形式示例：
        - 表1-2 表名
        - 表 1.2.3 表格说明
        - 续表1.3 表格名称

        匹配逻辑：
        使用正则表达式匹配是否以“表”或“续表”开头，后跟编号，再后面跟标题文本。

        参数:
        ----------
        element : CT_P
            Word 段落对象。

        返回:
        ----------
        bool : 若为表格标题，返回 True；否则返回 False。
        """
        text = element.text
        re.sub(r"\s+", " ", text.strip())

        # 定义表序号正则表达式
        table_num_pattern = r"(?:续)?表\s*\d+(?:[\-\—－\.．]\d+)*"

        if re.match(rf"^\s*({table_num_pattern})\s*(.*)", text):
            return True
        return False

    def _is_image_title_element(self, element):
        """
        判断段落是否为“图片标题”。

        图片标题形式示例：
        - 图1-1 图片说明
        - 图1.2.3-4 名称

        匹配逻辑：
        使用正则表达式匹配是否以“图”开头，后跟编号，再后面跟说明文字。

        参数:
        ----------
        element : CT_P
            Word 段落对象。

        返回:
        ----------
        bool : 若为图片标题，返回 True；否则返回 False。
        """
        text = element.text
        re.sub(r"\s+", " ", text.strip())

        # 定义表序号正则表达式
        image_num_pattern = r"图\s*\d+(?:[\-\—－\.．]\d+)*"

        if re.match(rf"^\s*({image_num_pattern})\s*(.*)", text):
            return True
        return False

    def _is_blank_para_element(self, element):
        """
        判断段落是否为空段落。

        空段落定义如下：
        - element 为 None
        - element.text 为 None
        - 去除所有空白字符后仍为空字符串

        参数:
        ----------
        element : CT_P
            Word 段落对象。

        返回:
        ----------
        bool : 若为空段落，返回 True；否则返回 False。
        """
        if (
                element is None or
                element.text is None or
                element.text.strip() == ""
        ):
            return True
        return False

    def _is_toc_element(self, element):
        """
        判断段落是否为目录项（Table of Contents）。

        检测逻辑：
        1. 样式名检查：段落样式名或基样式名中包含 "toc"（不区分大小写）；
        2. 字段代码检查：段落中包含 instrText 元素，其内容包含 "TOC"。

        参数:
        ----------
        element : CT_P
            Word 段落对象。

        返回:
        ----------
        bool : 若为目录段落，返回 True；否则返回 False。
        """
        # 样式名检查
        pPr = element.find(qn("w:pPr"))
        if pPr is not None:
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                style_id = pStyle.get(qn("w:val"), "")
                style = self.styles_map.get(style_id)
                if style:
                   # 样式 ID 本身
                    if style_id.upper().startswith("TOC"):
                        return True
                    # 样式名/基样式名包含 "toc"
                    name = (style.name or "").lower()
                    base = getattr(style, "base_style", None)
                    base_name = (base.name or "").lower() if base else ""
                    if "toc" in name or "toc" in base_name:
                        return True

        # # ---- instrText 检查 ----
        # for instr_txt in element.findall(f".//{qn('w:instrText')}"):
        #     txt = (instr_txt.text or "").strip().lower()
        #     if txt.startswith("toc"):
        #         print(f"instrText: {txt}")  # 目录定义
        #         return True
        #     if txt.startswith("hyperlink \\l \"_toc"):  # 目录条目
        #         print(f"instrText: {txt}")
        #         return True

        # # ---- fldSimple 检查 ----
        # for fld in element.findall(f".//{qn('w:fldSimple')}"):
        #     instr = fld.get(qn("w:instr"), "").strip().lower()
        #     if instr.startswith("toc"):
        #         print(f"fldSimple: {txt}")
        #         return True
        #     if instr.startswith("hyperlink \\l \"_toc"):
        #         print(f"fldSimple: {txt}")
        #         return True

        return False

    def _is_dynamic_title_element(self, element) -> bool:
        """
        判断段落是否为“动态标题”。

        动态标题要求：
        1. 文本长度不超过30字符（去除前后空白后计算）。
        2. 符合以下任一编号格式（不区分全/半角括号，容忍空格）：
           - 圆圈数字①–⑳ 开头
           - 数字后缀括号，如 2) 或 2）
           - 括号内阿拉伯数字，如 (2) 或 （3）
           - 括号内中文大写数字，如 (二) 或 （三）
           - 字母后缀括号，如 a) 或 b）
        3. 句尾不能是任何标点符号，必须以中文字符、英文字母或数字结尾。

        参数:
        ----------
        element : CT_P
            Word 段落对象。

        返回:
        ----------
        bool : 若为动态标题，返回 True；否则返回 False。
        """
        if element is None or element.text is None:
            return False

        # 规范化空白
        text = re.sub(r"\s+", " ", element.text.strip())

        # 长度限制
        if len(text) > 30:
            return False

        # 各种编号格式的子模式
        p_circle       = r"[\u2460-\u2473]\s*.+[\u4e00-\u9fffA-Za-z0-9]"
        p_suffix_num   = r"\d+\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]"
        p_suffix_letter= r"[A-Za-z]\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]"
        p_paren_ar     = r"[\(\uFF08]\s*\d+\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]"
        p_paren_ch     = r"[\(\uFF08]\s*[一二三四五六七八九十]+\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]"

        # 组合所有子模式
        dynamic_pattern = re.compile(
            rf"^(?:{p_circle}|{p_suffix_num}|{p_suffix_letter}|{p_paren_ar}|{p_paren_ch})$"
        )

        return bool(dynamic_pattern.match(text))

    def _is_title_element(self, element):
        """
        判断段落是否为章节标题（如“1 总则”、“2.1 范围”这类结构）。

        检测逻辑：
        ---------------------
        1. 段落文本必须非空，且长度小于一定阈值（<35）；
        2. 样式名或其基样式名中包含 "heading"（不区分大小写）；
        3. 段落属性中包含 <w:outlineLvl> 元素，表示其为大纲级别；
        4. 正则匹配“编号+空格+内容”的结构，例如：
            - 1 总则
            - 2.1 范围
            - 3.2.5 特殊情况说明

        参数:
        ----------
        element : CT_P
            Word 段落对象。

        返回:
        ----------
        bool : 若为标题段落，返回 True；否则返回 False。
        """
        # 文本及长度判断
        text = element.text.strip()
        if not text or len(text) >= 35:
            return False
        
        if "《" in text or "》" in text:
            return False

        # 读取 <w:pPr>/<w:pStyle w:val="..."/>
        style_id = None
        pPr = element.find(qn("w:pPr"))
        if pPr is not None:
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                style_id = pStyle.get(qn("w:val"))

        # 样式名和基样式名判断
        style_name = ""
        based_on_name = ""
        if style_id and style_id in self.styles_map:
            style = self.styles_map[style_id]
            style_name = (style.name or "").lower()
            base = getattr(style, "base_style", None)
            if base is not None:
                based_on_name = (base.name or "").lower()

        if any(x in style_name for x in ("heading", "Heading")):
            return True
        if any(x in based_on_name for x in ("heading", "Heading")):
            return True

        # 检查 <w:outlineLvl>
        if pPr is not None and pPr.find(qn("w:outlineLvl")) is not None:
            return True

        # 简单正则匹配“1.1 概述”、“2 总则”这类章节标题
        # if re.match(r"^\d+(\.\d+)*\s+[\u4e00-\u9fa5\w]+$", text):
        if re.match(r"^\d+(\.\d+)*\s+.+$", text):
            return True

        return False

    def _get_title_level(self, element):
        """
        获取标题级别（H1、H2 等）。

        通过段落属性中的 <w:outlineLvl> 元素获取标题级别，
        如果不存在则返回 0（表示 H0）。

        参数:
        ----------
        element : CT_P
            Word 段落对象。

        返回:
        ----------
        int : 标题级别，0 表示 H0，1 表示 H1，以此类推。
        """
        # 1. 检查 <w:outlineLvl>
        pPr = element.find(qn("w:pPr"))
        if pPr is not None:
            outline = pPr.find(qn("w:outlineLvl"))
            if outline is not None:
                val = outline.get(qn("w:val"))
                # val 从 "0" 开始，转成 int 后 +1
                return int(val) + 1

        # 2. 检查样式名及其继承
        style_level = None
        style_id = None

        def get_builtin_style_level(style_name: str):
            """
            只匹配 Word 内置标题样式，返回级别 1-9
            """
            if not style_name:
                return None
            name = style_name.lower().strip()
            m_en = re.match(r"^heading\s*([1-9])$", name)
            if m_en:
                return int(m_en.group(1))
            return None

        if pPr is not None:
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                style_id = pStyle.get(qn("w:val"))

        if style_id and style_id in self.styles_map:
            style = self.styles_map[style_id]
            # 尝试匹配样式名（只匹配内置标题）
            style_level = get_builtin_style_level(style.name)
            # 如果没匹配成功，再看基样式
            if style_level is None:
                base = getattr(style, "base_style", None)
                if base and base.name:
                    style_level = get_builtin_style_level(base.name)

        if style_level:
            return style_level

        # 3. 用文本前缀正则匹配数字编号
        text = element.text.strip()
        # 匹配 1.  或 1.2  或 1.2.3 之后跟空格
        m = re.match(r"^(\d+(?:\.\d+)*)\s*", text)
        if m:
            nums = m.group(1).split(".")
            return len(nums)

        # 4. 如果都失败，返回 -1
        return -1
