from dataclasses import dataclass, asdict
from typing import Any, Dict, Optional


@dataclass
class FileMetadata:
    """
    存储关于源文件的所有元信息。

    filename: Optional[str] = None           # 文件名
    file_directory: Optional[str] = None     # 文件路径
    filetype: Optional[str] = None           # 文件类型
    last_modified: Optional[str] = None      # 最后修改时间
    """
    # ---- 文件级信息 ----
    filename: Optional[str] = None           # 文件名
    file_directory: Optional[str] = None     # 文件路径
    filetype: Optional[str] = None           # 文件类型
    last_modified: Optional[str] = None      # 最后修改时间

    def to_dict(self) -> Dict[str, Any]:
        """
        将 Metadata 转为字典，结合 asdict 递归展开 dataclass 字段。
        """
        return asdict(self)


@dataclass
class ElemMetadata:
    """
    存储关于 Element 的所有元信息。

    page_number: Optional[int] = None        # 元素所在页码
    style_name: Optional[str] = None         # 应用的样式名
    parent_id: Optional[str] = None          # 父元素 ID
    category_depth: Optional[int] = None     # 同类型元素中的层级深度
    """
    # ---- 版面位置信息 ----
    page_number: Optional[int] = None        # 元素所在页码

    # ---- 文本样式信息 ----
    style_name: Optional[str] = None         # 应用的样式名

    # ---- 结构层级信息 ----
    parent_id: Optional[str] = ""            # 父元素 ID
    parent_path: Optional[str] = ""          # 层次路径

    def to_dict(self) -> Dict[str, Any]:
        """
        将 Metadata 转为字典，结合 asdict 递归展开 dataclass 字段。
        """
        return asdict(self)

