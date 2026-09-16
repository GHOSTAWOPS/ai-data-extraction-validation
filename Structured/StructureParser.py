import re
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
from typing import Any, Tuple, Dict, List, Optional, Iterator
from .Metadata import FileMetadata, ElemMetadata
from .DataStructure import TitleHierarchyManager
from .utils import build_toc_tree


def clean_and_structure_elements(
        elements,
        clean_image_title=True
):
    """
    对解析后的元素列表进行结构化整理与清洗。

    功能说明：
    ---------------------
    - 封装调用 _StructureParser 的静态方法，对元素列表进行统一的结构清洗；
    - 可选项：是否对图片标题进行清理（默认开启）；
    - 返回处理后的元素结果列表，结构更规整，便于后续分析或转换。

    参数:
    ----------
    elements : List[Element]
        原始解析得到的元素列表（段落、表格、标题等）；
    
    clean_image_title : bool, optional
        是否清除误识别的图片标题（默认: True）。

    返回:
    ----------
    List[Element] :
        经过结构化和清洗处理后的元素列表。
    """
    result = _StructureParser.clean_and_structure_elements(
        elements,
        clean_img_title=clean_image_title
    )
    return result


class _StructureParser:
    def __init__(
            self,
            elements: List[Any],
            clean_img_title=True
    ):
        """
        初始化结构化解析器。

        参数:
        ----------
        elements : List[Any]
            文档中提取的所有元素（如段落、标题、表格等）。

        clean_img_title : bool, optional
            是否在清理时去除图片标题（默认: True）。
        """
        self.elements = elements
        self.clean_img_title = clean_img_title

    @classmethod
    def clean_and_structure_elements(
            cls,
            elements: List[Any],
            clean_img_title=True
    ) -> Tuple[Any, List[Any]]:
        """
        对元素进行清洗和结构化处理（类方法封装）。

        参数:
        ----------
        elements : List[Any]
            初始提取的元素列表；

        clean_img_title : bool, optional
            是否清除误识别的图片标题（默认: True）。

        返回:
        ----------
        Tuple[Any, List[Any]]
            - toc_tree: 构建的目录树（如果有）；
            - elements: 清洗和结构化后的元素列表。
        """
        instance = cls(
            elements=elements,
            clean_img_title=clean_img_title
        )
        return instance._clean_and_structure_elements()

    def _clean_and_structure_elements(self) -> Tuple[Any, List[Any]]:
        """
        实际执行结构化和清洗工作的私有方法。

        步骤说明：
        ---------------------
        1. 尝试生成目录树（TOC Tree）；
        2. 为每个元素注入其层次结构中的父节点信息；
        3. 清除无用元素（如空段落 BlankPara，以及图片标题 ImageTitle）；

        返回:
        ----------
        Tuple[Any, List[Any]]
            - toc_tree: 构建的目录树结构（如有）；
            - elements: 经过清洗后的有效元素列表。
        """
        toc_tree = self._get_toc_tree()
        self._inject_title_hierarchy()
        self._annotate_element_hierarchy_by_level()
        self._clean_elements()
        return toc_tree, self.elements

    def _get_toc_tree(self) -> str:
        """
        提取并构建目录树（TOC Tree），并从元素列表中移除目录部分。

        处理逻辑：
        -------------------------
        1. 查找所有 Toc 类型元素的位置；
        2. 提取其文本内容，用于构造目录树；
        3. 删除目录段前及目录段内所有元素；
        4. 返回构造好的目录树字符串。

        返回:
        ----------
        toc_tree : str
            构造的目录树结构（字符串形式，若无目录则为空）。
        """
        toc_indices = [i for i, e in enumerate(self.elements) if isinstance(e, Toc)]
        if not toc_indices:
            # 没有 Toc，不做删除，返回空目录
            return ""

        # 提取所有 Toc 文本
        toc_texts = [self.elements[i].text.strip() for i in toc_indices]

        # 构造目录树字符串
        toc_tree = build_toc_tree(toc_texts)

        # 删除目录段及之前的元素（保留目录后内容）
        first = toc_indices[-1]
        self.elements = self.elements[first + 1:]

        return toc_tree


    def _inject_title_hierarchy(self):
        """
        为 Title 和 DynamicTitle 元素注入层级 level 与 parent_path。

        逻辑说明：
        - 利用 标题栈 与 模板层级字典 维护结构层次；
        - Title 根据显式 level 入栈并构建路径；
        - DynamicTitle 根据编号样式推断层级并入栈；
        - 每次入栈前弹出高层级标题，确保路径正确。
        """
        manager = TitleHierarchyManager()

        for elem in self.elements:
            # Title 处理
            if isinstance(elem, Title):
                if elem.level == -1:
                    continue
                manager.pop_until_smaller_level(elem.level)
                manager.push_title(elem)
                manager.reset_dynamic_levels()

            # DynamicTitle 处理
            elif isinstance(elem, DynamicTitle):
                style = elem.detect_pattern()
                if style is None:
                    continue

                current_level = manager.title_level_dict.get(style, -1)
                if current_level == -1:
                    current_level = manager.get_current_level() + 1
                    manager.title_level_dict[style] = current_level

                elem.set_level(current_level)
                manager.pop_until_smaller_level(current_level)
                manager.push_title(elem)


    def _annotate_element_hierarchy_by_level(self):
        """
        根据已注入的标题层级信息，为其它元素设置 parent_id 和 parent_path。

        处理方式：
        - 标题元素更新当前层级参考；
        - 表格标题元素用于标记后续表格；
        - 表格元素优先关联最近的表格标题，否则继承最近标题的层级；
        - 其它元素继承最近标题的层级路径。
        """
        last_heading_elem = None  # 最近的有意义标题（Title 或 DynamicTitle）
        last_table_title_elem = None  # 最近的 TableTitle

        for el in self.elements:
            # 更新最近的标题（Title / DynamicTitle），要求其 parent_path 非空
            if isinstance(el, (Title, DynamicTitle)):
                if getattr(el, "elem_metadata", None) and el.elem_metadata.parent_path:
                    last_heading_elem = el
                continue

            # 更新最近的 TableTitle
            if isinstance(el, TableTitle):
                last_table_title_elem = el
                el.elem_metadata.parent_id = last_heading_elem.element_id
                el.elem_metadata.parent_path = last_heading_elem.elem_metadata.parent_path
                continue


            if isinstance(el, Table):
                if hasattr(el, "elem_metadata"):
                    # if last_table_title_elem is not None:
                    #     el.elem_metadata.parent_id = last_table_title_elem.element_id
                    #     el.elem_metadata.parent_path = last_table_title_elem.text.strip()
                    if last_heading_elem is not None:
                        el.elem_metadata.parent_id = last_heading_elem.element_id
                        el.elem_metadata.parent_path = last_heading_elem.elem_metadata.parent_path
                continue

            # 其它元素（例如段落、图片标题等）
            if hasattr(el, "elem_metadata") and last_heading_elem is not None:
                el.elem_metadata.parent_id = last_heading_elem.element_id
                el.elem_metadata.parent_path = last_heading_elem.elem_metadata.parent_path

    def _clean_elements(self):
        """
        清理无效元素。

        默认删除：
        -------------------
        - BlankPara：空段落；
        - ImageTitle：图片标题（可选，视 `clean_img_title` 参数而定）。
        """
        clean_elements = [BlankPara]
        if self.clean_img_title:
            clean_elements.append(ImageTitle)

        self.elements = [el for el in self.elements if not isinstance(el, tuple(clean_elements))]


