import re
from typing import List, Dict, Optional

class TitleHierarchyManager:
    """
    标题层级管理器，用于维护文档中 Title 与 DynamicTitle 的结构关系。

    功能：
    ----------
    - 通过 标题栈 维护当前层级上下文；
    - 动态标题样式与对应层级的映射维护；
    - 提供层级判断、路径拼接与栈操作接口。
    """
    def __init__(self):
        """
        初始化正则模板与层级字典，并创建空标题栈。
        """
        self.patterns: Dict[str, re.Pattern] = {
            "circle_num": re.compile(r'^[\u2460-\u2473]\s*.+[\u4e00-\u9fffA-Za-z0-9]$'),
            "suffix_bracket_num": re.compile(r'^\d+\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]$'),
            "suffix_letter": re.compile(r'^[A-Za-z]\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]$'),
            "parenthesized_arabic": re.compile(r'^[\(\uFF08]\s*\d+\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]$'),
            "parenthesized_chinese": re.compile(r'^[\(\uFF08]\s*[一二三四五六七八九十]+\s*[)\uFF09]\s*.+[\u4e00-\u9fffA-Za-z0-9]$'),
        }

        self.title_level_dict: Dict[str, int] = {k: -1 for k in self.patterns}
        self.title_stack: List = []

    def reset_dynamic_levels(self):
        """
        将所有动态标题模板对应的层级重置为 -1。
        """
        for k in self.title_level_dict:
            self.title_level_dict[k] = -1

    def get_current_level(self) -> int:
        """
        获取当前标题栈顶元素的层级（若栈为空返回 -1）。
        """
        return self.title_stack[-1].level if self.title_stack else -1

    def get_current_path(self) -> str:
        """
        获取当前标题栈顶元素的 parent_path（若栈为空返回空字符串）。
        """
        return self.title_stack[-1].elem_metadata.parent_path if self.title_stack else ""

    def push_title(self, elem):
        """
        将标题元素压入栈，并构建其 parent_path。
        """
        # 设置 parent_path
        parent_path = self.get_current_path()
        if parent_path:
            new_path = f"{parent_path}|{elem.text}"
        else:
            new_path = elem.text
        elem.elem_metadata.parent_path = new_path
        self.title_stack.append(elem)

    def pop_until_smaller_level(self, current_level: int):
        """
        弹出标题栈中所有层级大于等于 current_level 的元素。
        """
        while self.title_stack and self.title_stack[-1].level >= current_level:
            self.title_stack.pop()