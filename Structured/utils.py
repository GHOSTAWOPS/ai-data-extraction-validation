import os
import re
import platform
import subprocess
from typing import List, Tuple
from docx.table import Table as _DocxTable

from docx import Document
from docx.oxml.ns import qn


def convert_doc_to_docx(doc_path: str):
    """
    使用 LibreOffice 将 .doc 文件转换为 .docx 文件。

    参数:
        doc_path (str): 需要转换的 .doc 文件路径

    返回:
        str: 转换后生成的 .docx 文件路径

    异常:
        subprocess.CalledProcessError: 如果转换失败，会抛出该异常，提示手动转换
    """
    # 检测当前操作系统类型，决定调用哪个 LibreOffice 可执行程序
    system = platform.system().lower()
    if system.startswith("windows"):
        cmd_bin = "soffice"
    else:
        cmd_bin = "libreoffice"

    # 调用 LibreOffice 命令行工具进行转换
    output_dir = os.path.dirname(doc_path)
    command = [
        cmd_bin,
        "--headless",
        "--convert-to",
        "docx",
        doc_path,
        "--outdir",
        output_dir,
    ]
    try:
        subprocess.run(command, check=True)
        print(f"转换成功，文件已保存为: {doc_path.replace('.doc', '.docx')}")
        return doc_path.replace(".doc", ".docx")
    except subprocess.CalledProcessError as e:
        raise subprocess.CalledProcessError(f"将 {doc_path} 转换为 .docx 文件失败，请手动转换")


def table_to_markdown(table: _DocxTable) -> str:
    """
    将 Word 文档中的表格（_DocxTable）转换为 Markdown 格式的字符串。
    
    规则：
    - 第一行作为 Markdown 表头
    - 第二行自动生成分隔符（---）
    - 其余行为数据行
    - 表格内的竖线 | 会被转义成 \|，避免干扰 Markdown 格式
    
    参数：
        table (_DocxTable): 传入的 Word 表格对象
    
    返回：
        str: 转换后的 Markdown 格式表格字符串
    """
    rows: List[List[str]] = []
    for row in table.rows:
        cells = [cell.text.strip().replace("\n", "") for cell in row.cells]
        rows.append(cells)

    # 最多列数
    n_cols = max(len(r) for r in rows)
    # 确保每行都有足够单元格
    norm_rows = [r + [""]*(n_cols - len(r)) for r in rows]

    # 转义竖线
    def esc(cell: str) -> str:
        return cell.replace("|", "\\|")

    # 拼表头
    header = "| " + " | ".join(esc(c) for c in norm_rows[0]) + " |"
    # 拼分隔符
    separator = "| " + " | ".join("---" for _ in range(n_cols)) + " |"
    # 拼数据行
    body = "\n".join(
        "| " + " | ".join(esc(c) for c in row) + " |"
        for row in norm_rows[1:]
    )

    return "\n".join([header, separator, body])


def build_toc_tree(toc_texts: List[str]) -> str:
    """
    根据一组目录条目文本（如 "1 概述", "1.1 背景" 等）构造带缩进的目录树字符串。
    
    规则：
    - 通过正则提取章节编号和标题
    - 按章节号的层级生成缩进（每一级缩进两个空格）
    - 输出格式为 "1.1.2 标题"，缩进表示层级关系
    
    参数：
        toc_texts (List[str]): 目录文本列表，每项格式类似 "1.2 标题"
    
    返回：
        str: 带层级缩进的目录字符串，方便阅读
    """
    tree: List[Tuple[List[int], str]] = []
    for text in toc_texts:
        # 用正则提取章节号和标题
        m = re.match(r"^(\d+(?:\.\d+)*)\s*(.*)", text)
        if not m:
            continue
        num, title = m.group(1), m.group(2)
        # 把 "1.2.3" 切成 [1,2,3]
        level = [int(x) for x in num.split(".")]
        tree.append((level, title))

    # 构建缩进字符串
    lines = []
    for level, title in tree:
        indent = "  " * (len(level) - 1)
        lines.append(f"{indent}{'.'.join(str(i) for i in level)} {title}")
    return "\n".join(lines)


if __name__ == "__main__":
    path = "/home/ljc/project/Demo2.9/data/test_data/合并文档测试/08-理塘-预可-工程布置及建筑物-20250618.doc"
    convert_doc_to_docx(path)