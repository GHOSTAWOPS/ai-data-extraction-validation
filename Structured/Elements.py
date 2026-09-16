from abc import ABC, abstractmethod
from dataclasses import dataclass, field
import uuid
import hashlib
from typing import Any, Dict, List, Pattern, Optional
from .Metadata import FileMetadata, ElemMetadata
import re


DYNAMIC_TITLE_PATTERNS: Dict[str, Pattern] = {
    "circle_num": re.compile(r'^[\u2460-\u2473]\s*.+[\u4e00-\u9fffA-Za-z0-9]$'),
    "suffix_bracket_num": re.compile(r'^\d+\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]$'),
    "suffix_letter": re.compile(r'^[A-Za-z]\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]$'),
    "parenthesized_arabic": re.compile(r'^[\(\uFF08]\s*\d+\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]$'),
    "parenthesized_chinese": re.compile(r'^[\(\uFF08]\s*[一二三四五六七八九十]+\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]$'),
}


@dataclass
class Element(ABC):
    """
    文档元素基类，所有具体元素继承自此类。
    """
    element_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    file_metadata: FileMetadata = field(default_factory=FileMetadata)
    elem_metadata: ElemMetadata = field(default_factory=ElemMetadata)

    @abstractmethod
    def to_dict(self) -> Dict[str, Any]:
        """
        序列化元素为字典，包含类型、ID、内容、metadata。
        """
        ...

    @classmethod
    def deterministic_id(cls, text: str, page: Optional[int] = None, filename: Optional[str] = None) -> str:
        """
        根据内容、页码和文件名生成可重现的 SHA-256 ID。
        """
        sha = hashlib.sha256()
        sha.update(text.encode('utf-8'))
        if page is not None:
            sha.update(str(page).encode('utf-8'))
        if filename:
            sha.update(filename.encode('utf-8'))
        return sha.hexdigest()


@dataclass
class TextElement(Element):
    """
    文本元素基类，包含纯文本内容。
    """
    text: str = ""

    def to_dict(self) -> Dict[str, Any]:
        data = {
            "element_type": self.__class__.__name__,
            "element_id": self.element_id,
            "text": self.text,
            "file_metadata": self.file_metadata.to_dict(),
            "elem_metadata": self.elem_metadata.to_dict(),
        }
        return data


@dataclass
class Toc(TextElement):
    """目录元素"""
    category = "Toc"


@dataclass
class Title(TextElement):
    """文档标题（如 H1、H2 等）"""
    category = "Title"
    level: int = -1  # 标题级别，默认为 -1


@dataclass
class TableTitle(TextElement):
    """表格标题或描述"""
    category = "TableTitle"


@dataclass
class ImageTitle(TextElement):
    """图片标题或描述"""
    category = "ImageTitle"


@dataclass
class Paragraph(TextElement):
    """普通段落文本"""
    category = "Paragraph"


@dataclass
class BlankPara(TextElement):
    """空段落"""
    category = "BlankPara"


@dataclass
class Table(Element):
    """
    表格元素，text 存储表格纯文本信息
    
    text_as_html 可存 HTML 表示。
    text_as_markdown 可存 MD 表示。
    """
    category = "Table"
    text: str = field(default_factory=str)

    # ---- 表格特有信息 ----
    text_as_html: Optional[str] = None  # 表格对应的 HTML 表示
    text_as_markdown: Optional[str] = None  # 表格对应的 Markdown 表示

    def to_dict(self) -> Dict[str, Any]:
        data = {
            "element_type": "Table",
            "element_id": self.element_id,
            "data": self.text,
            "file_metadata": self.file_metadata.to_dict(),
            "elem_metadata": self.elem_metadata.to_dict(),
        }
        return data


@dataclass
class DynamicTitle(TextElement):
    """
    动态标题：标题层级在初始化时未知，只能通过后续处理赋值。
    支持多种编号格式的匹配。
    """
    category: str = "DynamicTitle"
    level: Optional[int] = None
    matched_style: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        # 扩展父类的字典，加入 matched_style 与 level
        data = super().to_dict()
        data.update({
            "matched_style": self.matched_style,
            "level": self.level,
        })
        return data

    def detect_pattern(self) -> Optional[str]:
        """
        检查 self.text 匹配哪种编号样式，设置 matched_style 并返回样式名；
        若都不匹配，返回 None。
        """
        for style_name, pattern in DYNAMIC_TITLE_PATTERNS.items():
            if pattern.match(self.text):
                self.matched_style = style_name
                return style_name
        return None

    def set_level(self, level: int) -> None:
        """
        后续处理中调用，为该标题赋予具体层级。
        """
        self.level = level
