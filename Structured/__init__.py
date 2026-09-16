"""
Structured 包

该包用于解析 Word 文档（.docx），提取段落、标题、表格、图片标题等结构化内容，并提供元数据管理与结构化组织功能。

模块内容包括：
- parse_docx: 解析 Word 文档为原始元素列表。
- clean_and_structure_elements: 对原始元素进行清洗与结构化处理。

- FileMetadata / ElemMetadata: 提供文档级与元素级的元数据信息。
- 各类元素对象（Paragraph, Title, Table 等），用于表示 Word 中的各种内容。
- TitleHierarchyManager: 管理标题结构的辅助类，用于计算标题的 parent_path 与层级。
"""
from .ElemmentParser import parse_docx
from .StructureParser import clean_and_structure_elements
from .Metadata import (
    FileMetadata,
    ElemMetadata
)
from .Elements import (
    Paragraph,
    BlankPara,
    Toc,
    Title,
    TableTitle,
    ImageTitle,
    DynamicTitle,
    Table
)

__all__ = [
    "parse_docx",
    "clean_and_structure_elements",
    "FileMetadata",
    "ElemMetadata",
    "Paragraph",
    "BlankPara",
    "Toc",
    "Title",
    "TableTitle",
    "ImageTitle",
    "DynamicTitle",
    "Table",
]