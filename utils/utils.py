from docx.table import Table
from langchain_community.vectorstores import Chroma, FAISS
from langchain_community import embeddings
from langchain_community.chat_models import ChatOllama
from langchain_core.runnables import RunnablePassthrough
from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
from langchain.text_splitter import (
    CharacterTextSplitter,
    RecursiveCharacterTextSplitter,
)
from langchain_core.prompts.prompt import PromptTemplate
from langchain.schema import Document

import numpy as np
import matplotlib.pyplot as plt
from scipy.spatial.distance import cosine
from sklearn.decomposition import PCA
import matplotlib as mpl

# from data_processor import init_query_info
# import win32com.client
# import pythoncom
import os
import tempfile
import io
import re
import time
from datetime import timedelta


# from PyPDF2 import PdfReader
from langchain_community.document_loaders import PyPDFLoader
from datetime import datetime
from docx import Document as DocxDocument
from collections import defaultdict, Counter
from typing import List, Dict
import json
import subprocess
from config import CONFIG
import math
import platform
import jieba
import pandas as pd

# from data_processor import doc_to_pdf_and_load_pages

from utils.logger_util import logger
import json


# 停用
def get_appendix(source_doc):
    """
    从源文件中读取附表，提取每张表第一列除第一行的元素
    作为prompt的关键词
    """
    paras = []
    keyword = "附表"
    query_info = []

    # 找到包含 "附表" 的段落位置
    appendix_idx = None
    for idx, para in enumerate(source_doc.paragraphs):
        if keyword in para.text.strip():
            paras.append(idx)

    # 如果找到了 "附表" 关键字，提取从该段落开始之后的表格
    # 最后一个“附表”才是真正附表的开始，正文可能会出现“附表”二字
    if paras:
        appendix_idx = paras[-1]

        tbl_idx = 0
        for element in source_doc.element.body[appendix_idx + 1 :]:
            if element.tag.endswith("tbl"):  # 这是表格元素
                table = Table(element, source_doc)

                # !!!
                if tbl_idx > 2:
                    break
                tbl_idx += 1

                # query_info
                keywords_result = get_keyword_list(table)

                assert len(keywords_result[0]) == len(keywords_result[1])
                query_info.extend(
                    [
                        keywords_result[1][i]
                        for i in range(len(keywords_result[0]))
                        if keywords_result[0][i]
                    ]
                )

    return query_info


# 停用
def get_keyword_list(table):
    """
    从一个表格中 生成 当前表格所需要的 keywords_list 以及 query_list
    """
    keywords = [i.cells[0].text.strip() for i in table.rows]
    levl_list = []
    word_list = []
    flag_list = []

    for _, keyword in enumerate(keywords[1:], start=1):

        # 使用正则表达式匹配前面的数字部分
        match = re.match(r"^([\d\.]+)\s+(.*)", keyword)
        if match:
            index = match.group(1)  # 提取数字部分
            text = match.group(2)  # 提取后面的词语部分

        kw = keyword.split(" ")
        levl_list.append(kw[0])
        word_list.append(kw[1] + kw[2] if len(kw) > 2 else kw[1])

    for i in range(0, len(levl_list)):
        flag = True
        for j in range(i + 1, len(levl_list)):
            # 将索引按照 '.' 分割成列表
            parts1 = levl_list[i].split(".")
            parts2 = levl_list[j].split(".")
            # 首先判断第一个索引的长度是否比第二个少1，且第二个以第一个为开头
            if len(parts2) == len(parts1) + 1 and parts2[: len(parts1)] == parts1:
                # 第 i 个关键字是第 j 个关键字的直接上级
                word_list[j] = word_list[i] + "|" + word_list[j]
                flag = False
        flag_list.append(flag)

    return flag_list, word_list


# def filter_gt(relevant_docs, answer):
#     """
#     过滤掉出现answer答案的chunk

#     Args:
#         relevant_docs (list): 待过滤的文档列表
#         answer (str/int): 需要过滤的答案

#     Returns:
#         list: 过滤后的文档列表
#     """
#     answer_str = str(answer)
#     ans_list = answer_str.split("|")

#     # 判断是否为有效数字（整数或小数，且小数点两边都有数字）
#     def is_valid_number(s):
#         return re.match(r"^-?(\d+\.\d+|\d+)$", s) is not None

#     def simplify_unit(unit):
#         unit = unit.strip()
#         match = re.match(r"[^\d\s/]+", unit)  # 匹配以字母或特殊符号开头的连续单位片段
#         return match.group() if match else ""

#     filtered_docs = []
#     true_docs = []

#     # 如果answer是有效数字
#     if len(ans_list) == 2 and is_valid_number(ans_list[1]):
#         # 提取并简化单位
#         unit = simplify_unit(ans_list[0])
#         num_str = ans_list[1]

#         # 构建正则表达式，确保精确匹配数值
#         pattern = r"(?<![\d.])" + re.escape(num_str) + r"(?!\d|\.\d)"

#         # 根据单位构建新的正则表达式
#         if unit:
#             unit_pattern = re.escape(unit)
#             pattern = r"(?<![\d.])" + re.escape(num_str) + r"\s*" + unit_pattern

#         for doc in relevant_docs:
#             if not re.search(pattern, doc.page_content):
#                 filtered_docs.append(doc)

#             else:
#                 true_docs.append(doc)

#     else:
#         # 对于文本，直接使用普通的包含匹配
#         for doc in relevant_docs:
#             if answer_str not in doc.page_content:
#                 filtered_docs.append(doc)
#             else:
#                 true_docs.append(doc)

#     return filtered_docs, true_docs


def filter_gt(relevant_docs, answer):
    """
    过滤掉出现answer答案的chunk，支持乘法形式的单位和值匹配

    Args:
        relevant_docs (list): 待过滤的文档列表，每个元素应有 page_content 属性
        answer (str/int): 需要过滤的答案，例如 "m×cm×dm|7.0×8.5×5.0"

    Returns:
        tuple: (不包含答案的文档列表, 包含答案的文档列表)
    """
    # print(answer)

    def is_valid_number(s):
        return re.match(r"^-?(\d+\.\d+|\d+)$", s) is not None

    def simplify_unit(unit):
        unit = unit.strip()
        match = re.match(r"[^\d\s/]+", unit)  # 匹配非数字、空格、斜杠的单位前缀
        return match.group() if match else ""

    answer_str = str(answer)
    if "|" not in answer_str:
        # 普通文本处理
        true_docs = []
        rele_docs = []

        for doc in relevant_docs:
            content = doc.page_content
            if answer_str in content:
                true_docs.append(doc)
            else:
                rele_docs.append(doc)
        return rele_docs, true_docs

    else:
        unit_part, value_part = answer_str.split("|")
        answer = value_part + unit_part
        # print(answer)
        true_docs = []
        rele_docs = []

        for doc in relevant_docs:
            content = doc.page_content
            if answer in content:
                true_docs.append(doc)
            else:
                rele_docs.append(doc)
        return rele_docs, true_docs

    # # 支持 × 或 * 作为乘法符号分隔
    # unit_part, value_part = answer_str.split("|")
    # unit_parts = re.split(r"[×*]", unit_part)
    # value_parts = re.split(r"[×*]", value_part)

    # # 广播单位：如果只有一个单位但多个数值，则自动广播
    # if len(unit_parts) == 1 and len(value_parts) > 1:
    #     unit_parts = unit_parts * len(value_parts)

    # # 如果单位和数值个数仍不一致，无法配对，跳过过滤
    # if len(unit_parts) != len(value_parts):
    #     return relevant_docs, []

    # # 构造正则模式，每个元素是 "数值 + 单位" 的模式
    # def build_pattern(val, unit):
    #     val = re.escape(val.strip())
    #     unit = re.escape(simplify_unit(unit.strip()))
    #     #zty 2025/8/29修改 保证后面也不是小数点
    #     # return rf"(?<![\d.]){val}\s*{unit}(?!\w)"
    #     return rf"(?<![\d.]){val}(?![\d.])"
    #     #return rf"(?<![\d.]){val}(?!\d)"

    # patterns = []
    # for val, unit in zip(value_parts, unit_parts):
    #     flag = is_valid_number(val.strip())
    #     if flag:
    #         pattern = re.compile(build_pattern(val, unit))
    #         patterns.append(pattern)

    # # patterns = [re.compile(build_pattern(val, unit)) 
    # #             for val, unit in zip(value_parts, unit_parts)
    # #             if is_valid_number(val.strip())]

    # filtered_docs = []
    # true_docs = []

    # if len(patterns) != 0: 
    #     for doc in relevant_docs:
    #         content = doc.page_content
    #         if all(p.search(content) for p in patterns):
    #             true_docs.append(doc)
    #         else:
    #             filtered_docs.append(doc)
    # else:
    #     return ([doc for doc in relevant_docs if answer_str not in doc.page_content],
    #             [doc for doc in relevant_docs if answer_str in doc.page_content])

    # print(f"answer_str: {answer_str}, ")
    # print("------------------------")
    # for i in true_docs:
    #     print(i.page_content)
    # print()

    # return filtered_docs, true_docs

# 暂时废弃
def pdf_to_pages(doc_file_path):
    pass
    # text_splitter = RecursiveCharacterTextSplitter.from_tiktoken_encoder(
    #     separators=CONFIG.CHUNK_SEPARATORS,
    #     chunk_size=CONFIG.CHUNK_SIZE,
    #     chunk_overlap=math.ceil(CONFIG.CHUNK_SIZE * CONFIG.CHUNK_OVERLAP_ratio),
    # )

    # system = platform.system()
    # if system == "Windows":
    #     chunks = doc_to_pdf_and_load_pages(doc_file_path, text_splitter)
    # elif system == "Linux":
    #     loader = PyPDFLoader("./data/LuShanReport-OnlyText.pdf")
    #     chunks = loader.load_and_split(text_splitter=text_splitter)
    # else:
    #     raise RuntimeError(f"当前系统是 {system}，需要 Windows 或 Linux")

    # # 确保每个chunk都包含页面信息
    # for chunk in chunks:
    #     if "page" not in chunk.metadata:
    #         chunk.metadata["page"] = chunk.metadata.get("source", "").split("_")[-1]

    # return chunks


def error_handler(error_message):
    """
    装饰器：捕获异常并抛出指定的错误信息
    """

    def decorator(func):
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)  # 调用原函数
            except Exception:
                raise RuntimeError(error_message)  # 捕获异常并抛出指定的 RuntimeError

        return wrapper

    return decorator


@error_handler("文档处理出错，请联系技术人员")
def text_table_to_chunks(
    elements_text,
    elements_table_only_text,
    elements_table_idx,
    elements_table_summary,
    markdown_table_all,
):
    text_splitter = RecursiveCharacterTextSplitter.from_tiktoken_encoder(
        separators=CONFIG.CHUNK_SEPARATORS,
        chunk_size=CONFIG.CHUNK_SIZE,
        chunk_overlap=math.ceil(CONFIG.CHUNK_SIZE * CONFIG.CHUNK_OVERLAP_ratio),
    )
    # !!!!!!!表格用更小的分块器
    text_splitter_table = RecursiveCharacterTextSplitter.from_tiktoken_encoder(
        separators=["\n\n\n", "\n\n", "\n", "。", "！", "？"],
        chunk_size=CONFIG.CHUNK_SIZE / 3,
        chunk_overlap=math.ceil((CONFIG.CHUNK_SIZE / 3) * CONFIG.CHUNK_OVERLAP_ratio),
    )

    chunks = text_splitter.split_text(elements_text)
    chunks = [Document(page_content=chunk) for chunk in chunks]
    for chunk in chunks:
        chunk.metadata["category"] = "text"
    if CONFIG.EXTRACT_MODEL_SUMMARY_EMBEDDING_TABLE_SWITCH:
        table_pages = [Document(page_content=chunk) for chunk in elements_table_summary]
    else:
        table_pages = [
            Document(page_content=chunk) for chunk in elements_table_only_text
        ]
    for (
        chunk,
        table_doc_id,
        elements_table_summary_i,
        markdown_table,
        elements_table_only_text_i,
    ) in zip(
        table_pages,
        elements_table_idx,
        elements_table_summary,
        markdown_table_all,
        elements_table_only_text,
    ):
        table_chunk = CONFIG.TABLE_CHUNK_SWITCH
        if table_chunk:
            #!!!!!! 这里修改了分块器
            if CONFIG.TABLE_SUMMARY_SPLITTER == 1:
                chunks_table = text_splitter_table.split_text(chunk.page_content)
            else:
                chunks_table = text_splitter.split_text(chunk.page_content)
            chunks_table = [
                Document(page_content=chunk_table) for chunk_table in chunks_table
            ]
            for chunk_table in chunks_table:
                chunk_table.metadata["category"] = "table"
                chunk_table.metadata["table_doc_id"] = table_doc_id
                chunk_table.metadata["elements_table_summary"] = (
                    chunk_table.page_content
                )
                chunk_table.metadata["markdown_table"] = markdown_table
                chunk_table.metadata["elements_table_only_text"] = (
                    elements_table_only_text_i
                )
            chunks += chunks_table
        else:
            chunk.metadata["category"] = "table"
            chunk.metadata["table_doc_id"] = table_doc_id
            chunk.metadata["elements_table_summary"] = chunk.page_content
            chunk.metadata["markdown_table"] = markdown_table
            chunk.metadata["elements_table_only_text"] = elements_table_only_text_i
            # chunks+=chunk
            chunks.append(chunk)
    # chunks += table_pages
    for idx, chunk in enumerate(chunks):
        chunk.metadata["chunkid"] = idx

    return chunks

@error_handler("文档处理出错，请联系技术人员")
def zbb_text_table_to_chunks(
    elements_text,
    elements_table,
):
    text_splitter = RecursiveCharacterTextSplitter.from_tiktoken_encoder(
        separators=CONFIG.CHUNK_SEPARATORS,
        chunk_size=CONFIG.CHUNK_SIZE,
        chunk_overlap=math.ceil(CONFIG.CHUNK_SIZE * CONFIG.CHUNK_OVERLAP_ratio),
    )
    # !!!!!!!表格用更小的分块器
    text_splitter_table = RecursiveCharacterTextSplitter.from_tiktoken_encoder(
        separators=["\n\n\n", "\n\n", "\n", "。", "！", "？"],
        chunk_size=CONFIG.CHUNK_SIZE / 3,
        chunk_overlap=math.ceil((CONFIG.CHUNK_SIZE / 3) * CONFIG.CHUNK_OVERLAP_ratio),
    )
    chunks = []
    for key, value in elements_text.items():
        chunks_i = text_splitter.split_text(value)
        chunks_i = [Document(page_content=chunk) for chunk in chunks_i]
        # zhy修改2025/6/4
        # clean_key = key.replace("|", " ")
        # chunks_i = [
        #     Document(page_content=f"{clean_key}\n{chunk}") for chunk in chunks_i
        # ]
        for chunk in chunks_i:
            chunk.metadata["category"] = "text"
            chunk.metadata["path"] = key
            # zhy8/10,embedding 加 path
            # clean_path = key.replace("|", " ")
            # chunk.page_content = f"{clean_path}\n{chunk.page_content}" 

        chunks += chunks_i

    if CONFIG.EXTRACT_MODEL_SUMMARY_EMBEDDING_TABLE_SWITCH:
        table_pages = [Document(page_content=chunk) for chunk in elements_table["summary"]]
    else:
        table_pages = [
            Document(page_content=chunk) for chunk in elements_table["text"]
        ]
    for (
        chunk,
        table_doc_id,
        elements_table_summary_i,
        markdown_table,
        elements_table_only_text_i,
        elements_table_path
    ) in zip(
        table_pages,
        elements_table['idx'],
        elements_table['summary'],
        elements_table['markdown'],
        elements_table['text'],
        elements_table['path']
    ):
        table_chunk = CONFIG.TABLE_CHUNK_SWITCH
        if table_chunk:
            #!!!!!! 这里修改了分块器
            if CONFIG.TABLE_SUMMARY_SPLITTER == 1:
                chunks_table = text_splitter_table.split_text(chunk.page_content)
            else:
                chunks_table = text_splitter.split_text(chunk.page_content)
            chunks_table = [
                Document(page_content=chunk_table) for chunk_table in chunks_table
            ]
            for chunk_table in chunks_table:
                chunk_table.metadata["category"] = "table"
                chunk_table.metadata["table_doc_id"] = table_doc_id
                chunk_table.metadata["elements_table_summary"] = (
                    chunk_table.page_content
                )
                chunk_table.metadata["markdown_table"] = markdown_table
                chunk_table.metadata["elements_table_only_text"] = (
                    elements_table_only_text_i
                )
                chunk_table.metadata["path"] = (
                    elements_table_path
                )
            chunks += chunks_table
        else:
            chunk.metadata["category"] = "table"
            chunk.metadata["table_doc_id"] = table_doc_id
            chunk.metadata["elements_table_summary"] = chunk.page_content
            chunk.metadata["markdown_table"] = markdown_table
            chunk.metadata["elements_table_only_text"] = elements_table_only_text_i
            chunk.metadata["path"] = (
                    elements_table_path
                )
            # chunks+=chunk
            chunks.append(chunk)
    # chunks += table_pages
    for idx, chunk in enumerate(chunks):
        chunk.metadata["chunkid"] = idx
        # zhy8/10,embedding 加 path
        clean_path = chunk.metadata["path"].replace("|", " ")
        chunk.page_content = f"{clean_path}\n{chunk.page_content}" 

    return chunks

def extract_table_chunks_to_json(
    chunks,
    docx_file_name,
    json_path,
):
    """
    提取表格相关的 chunks 并保存为 JSON 文件
    """
    # 生成包含表格信息的字典
    # table_data = []
    table_data = {}
    for chunk in chunks:
        if chunk.metadata.get("category") == "table":
            # table_entry = {
            #     "chunkid": chunk.metadata.get("chunkid"),
            #     "table_doc_id": chunk.metadata.get("table_doc_id"),
            #     "content": chunk.page_content,
            #     "summary": chunk.metadata.get("elements_table_summary"),
            #     "markdown": chunk.metadata.get("markdown_table"),
            #     "plain_text": chunk.metadata.get("elements_table_only_text"),
            # }
            # table_data.append(table_entry)
            table_entry = {
                chunk.metadata.get("table_doc_id"): chunk.metadata.get(
                    "elements_table_summary"
                ),
            }
            table_data.update(table_entry)

    # 保存为 JSON 文件
    json_file = f"{json_path}/{docx_file_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.json"

    with open(json_file, "w", encoding="utf-8") as f:
        json.dump(table_data, f, ensure_ascii=False, indent=2)

    # 删除目录中存储的旧的 JSON 文件
    files_to_delete = []
    for filename in os.listdir(json_path):
        file_path = os.path.join(json_path, filename)
        if (
            os.path.isfile(file_path)
            and filename.startswith(docx_file_name)
            and file_path != json_file
        ):
            files_to_delete.append(file_path)

    for file_path in files_to_delete:
        try:
            os.remove(file_path)
        except Exception as e:
            raise RuntimeError(f"删除文件 {file_path} 失败: {e}")

    return table_data


def timing_decorator():
    """
    装饰器：用于计算函数执行时间
    """

    def decorator(func):
        def wrapper(*args, **kwargs):
            start_time = time.time()  # 记录开始时间
            result = func(*args, **kwargs)  # 调用原函数
            execution_time = time.time() - start_time  # 计算时间差（秒）
            execution_time_str = str(
                timedelta(seconds=execution_time)
            )  # 格式化为timedelta对象
            logger.info(f"函数 {func.__name__} 运行时间: {execution_time_str}")
            return result  # 返回函数的执行结果

        return wrapper

    return decorator

    # def wrapper(*args, **kwargs):
    #     start_time = time.time()  # 记录开始时间
    #     result = func(*args, **kwargs)  # 调用原函数
    #     execution_time = timedelta(seconds=int(time.time() - start_time))  # 计算时间
    #     logger.info(f"函数 {func.__name__} 运行时间: {execution_time}")
    #     return result  # 返回函数的执行结果

    # return wrapper


def expand_relevant_docs_text(relevant_docs, chunks):
    for doc in relevant_docs:
        # original_length = len(doc.page_content)
        doc.page_content = expand_chunks_text(doc.metadata["chunkid"], chunks)
        # expanded_length = len(doc.page_content)
    return relevant_docs


def expand_chunks_text(chunk_id, chunks):
    text = chunks[chunk_id].page_content
    if chunk_id > 0:
        # 将前一个chunk的文本补充到前面，直到遇见回车符或者chunk的开始
        prev_text = chunks[chunk_id - 1].page_content
        newline_pos = prev_text.rfind("\n")
        if newline_pos != -1:
            text = prev_text[newline_pos + 1 :] + text
        else:
            text = text
    if chunk_id < len(chunks) - 1:
        # 将后一个chunk的文本补充到后面，直到遇见回车符或者chunk的结束
        next_text = chunks[chunk_id + 1].page_content
        newline_pos = next_text.find("\n")
        if newline_pos != -1:
            text = text + next_text[:newline_pos]
        else:
            text = text
    return text


def cosine_similarity_manual(text1: str, text2: str):
    """
    计算两个字符串的 余弦相似度，结果越接近 1 表示两个字符串越相近。
    """
    # 统计词频生成变量
    vec1 = Counter(jieba.lcut(text1))
    vec2 = Counter(jieba.lcut(text2))

    # 计算点积
    intersection = set(vec1.keys()) & set(vec2.keys())
    dot_product = sum([vec1[x] * vec2[x] for x in intersection])

    # 计算模长
    norm1 = math.sqrt(sum([val**2 for val in vec1.values()]))
    norm2 = math.sqrt(sum([val**2 for val in vec2.values()]))

    # 计算余弦相似度
    return dot_product / (norm1 * norm2) if norm1 and norm2 else 0.0


def chunk_filter(input_list: List[Document], output_size=-1):
    """
    对 input_list 去重：
       如果 output_size 为默认值，那么需要对 input_list 完整去重
       如果 output_size 不为默认值，那么只需要从 input_list 中取出不重复的前 output_size 个元素
    """
    output_size = len(input_list) if output_size == -1 else output_size
    output_list = []

    for chunk1 in input_list:
        ## 将表格总结后的内容分块后需要将page_content转成markdown格式去重
        if "category" in chunk1.metadata and chunk1.metadata["category"] == "table":
            # 将page_content转成markdown格式
            chunk1.page_content = chunk1.metadata["markdown_table"]
        text1 = chunk1.page_content

        not_similar = True

        # 遍历 output_list 查看是否有与 chunk1 相似的 chunk
        for chunk2 in output_list:
            # 将搜索到的表格内容转成markdown格式
            if "category" in chunk2.metadata and chunk2.metadata["category"] == "table":
                # 将page_content转成markdown格式
                chunk2.page_content = chunk2.metadata["markdown_table"]
            text2 = chunk2.page_content
            similarity = cosine_similarity_manual(text1, text2)
            if similarity >= 0.96:
                not_similar = False
                break

        # 如果没有，将其加入 output_list
        if not_similar:
            output_list.append(chunk1)
            if len(output_list) >= output_size:
                # 控制加入 output_list 的 chunk 数
                break

    return output_list


def parse_directory(dir_name):
    # 查找 _去特性表.doc 文件
    doc_files = [
        f
        for f in os.listdir(dir_name)
        if f.endswith(("_去特性表.doc", "_去特性表.docx"))
    ]

    doc_file_path = None
    for f in doc_files:
        if f.endswith("_去特性表.doc") or f.endswith("_去特性表.docx"):
            doc_file_path = os.path.join(dir_name, f)
            break

    # 如果未找到“_去特性表”文件，则将唯一的doc文件视作信息来源
    if not doc_file_path and len(doc_files) == 1:
        doc_file_path = os.path.join(dir_name, doc_files[0])

    if not doc_file_path:
        raise FileNotFoundError(
            f"警告: 找不到 '_去特性表.doc' 或 '_去特性表.docx' 文件!"
        )

    # 查找 所有样本_提取.xlsx 以及 所有样本.xlsx 文件
    extract_file_path = os.path.join(dir_name, "所有样本_提取.xlsx")
    gt_excel_file_path = os.path.join(dir_name, "所有样本.xlsx")

    if not os.path.exists(extract_file_path):
        raise FileNotFoundError(f"警告: 找不到 '所有样本_提取.xlsx' 文件!")

    if not os.path.exists(gt_excel_file_path):
        raise FileNotFoundError(f"警告: 找不到 '所有样本.xlsx' 文件!")

    # 如果都存在，返回文件路径的列表
    return doc_file_path, extract_file_path, gt_excel_file_path


# 中间过程数据处理，用于中间结果输出
def extract_mid_process(ans_list):
    ans = []
    ans_list = list(dict.fromkeys(ans_list))
    for ans_true in ans_list:
        # zhy修改2025/3/4
        if not ans_true.strip():  # 过滤空字符串
            continue
        parts = ans_true.split("|")  # 分隔字符串

        if len(parts) == 1:
            ans.append([parts[0], "", "", ""])
        elif len(parts) == 2:
            if "###" in parts[1]:
                parts[1] = parts[1].strip("#")
            ans.append([parts[0], "", "", parts[1]])
        elif len(parts) == 3:
            if "###" in parts[2]:
                parts[2] = parts[2].strip("#")
            ans.append([parts[0], "", parts[1], parts[2]])
        # zhy添加2025/3/4/16：58
        else:
            if "###" in parts[3] and "###" in parts[-1]:
                parts[3] = parts[3].strip("#")
                parts[-1] = parts[-1].strip("#")
            ans.append([parts[0], parts[1], parts[2], " ".join(parts[3:])])
        ans = [row[:4] for row in ans]
        # parts = ans_true.split("|")
        # if len(parts) == 2:
        # ans.append([parts[0], "", parts[1]])
        # else:
        #     ans.append(parts)
        # ans.append(ans_true.split("|"))
    df = pd.DataFrame(ans, columns=["关键词", "单位", "数量", "备注"])
    return df


def split_answer2(asr: str):
    """
    对 asr 字符串做分隔符匹配以及字符分割。
    如果所有分隔符都失败，则返回原本的字符串
    """
    split_char = ["|", " ", "\t", "/", "\n", ",", "-"]

    for char in split_char:
        if char in asr:
            tokens = asr.split(char)
            asr = [token for token in tokens if token]
            logger.info(asr)
            return asr

    # 如果没有匹配到分隔符，直接返回原字符串作为单一元素
    return [asr]


def check_mid_process(ans_list, ref_ans_list, true_docs_list, topk_docs_list, topk):
    """
    让“页面预览备注”与填入 Excel 的备注保持一致的中间结果构造：
      S1: 有参考且无冲突
          在参考文段中找到关键词“{kw}”以及参考答案“{ground}”的相关内容。
      S2: 无参考且无冲突
          未在原文中找到关键词“{kw}”以及参考答案“{ground}”的相关内容。
      S3: 有参考且有冲突
          关键词“{kw}”在原文中存在不止一处可能答案。
          右侧列出了包含参考答案“{ground}”的文段以及与参考答案冲突的文段。
      S4: 无参考且有冲突
          在原文中找到的可能答案与参考答案“{ground}”皆不相同。
          右侧列出了与参考答案相关的冲突文段。
    """
    # 去重但保持顺序
    seen, ordered_ans = set(), []
    for a in ans_list:
        if a not in seen:
            ordered_ans.append(a)
            seen.add(a)

    rows = []
    for ans_true, ref_ans, true_docs, topk_docs in zip(
        ordered_ans, ref_ans_list, true_docs_list, topk_docs_list
    ):
        # 关键词
        kw = ans_true.split("|")[0].strip()

        # -------- 与 fill_check_answer_to_xlsx 完全一致的 flags / is_topk_relative 计算 ----------
        # split_answer2(ans_true) -> ["关键词","正确/错误","正确/错误",...]
        parts = split_answer2(ans_true)
        if len(parts) == topk + 1:
            flags = [s.strip() == "正确" for s in parts[1:]]  # True=一致(非冲突)
        else:
            # 与 Excel 写入函数保持一致：异常时按“全正确”回退（即无冲突）
            flags = [True] * topk

        is_topk_relative = True
        for v in flags:
            is_topk_relative = is_topk_relative and v
        is_topk_relative = not is_topk_relative  # True 表示“存在冲突文段”

        has_true = isinstance(true_docs, (list, tuple)) and len(true_docs) > 0

        # 参考答案文本直接用模板里的 ground truth（保持与 Excel 写入一致）
        ground = ref_ans.strip()

        # -------- 四种情形的备注文案（逐字对齐你的填表函数） ----------
        if has_true and (not is_topk_relative):
            note = f"在参考文段中找到关键词“{kw}”以及参考答案“{ground}”的相关内容。"
        elif (not has_true) and (not is_topk_relative):
            note = f"未在原文中找到关键词“{kw}”以及参考答案“{ground}”的相关内容。"
        elif has_true and is_topk_relative:
            note = (
                f"关键词“{kw}”在原文中存在不止一处可能答案。"
            )
        else:  # (not has_true) and is_topk_relative
            note = (
                f"在原文中找到的可能答案与参考答案“{ground}”皆不相同。"
            )

        # —— 单位/数量列（与模板保持一致）：支持 “单位|数量” 或仅“描述” —— #
        gt_parts = [x.strip() for x in ground.split("|") if x.strip() != ""]
        if len(gt_parts) >= 2:
            unit, qty = gt_parts[0], gt_parts[1]
        elif len(gt_parts) == 1:
            unit, qty = "", gt_parts[0]  # 描述型：放到“数量”列展示
        else:
            unit = qty = ""

        rows.append([kw, unit, qty, note])

    return pd.DataFrame(rows, columns=["关键词", "单位", "数量", "备注"])



if __name__ == "__main__":
    pass
