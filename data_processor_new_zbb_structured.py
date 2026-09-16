import json
import subprocess
import tempfile

# import pythoncom
# import win32com
from docx2pdf import convert
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import pandas as pd
import re
import os
import shutil
from utils.logger_util import logger
from unstructured.partition.docx import partition_docx
from unstructured.partition.doc import partition_doc
from langchain_community.document_loaders import PyPDFLoader

from bs4 import BeautifulSoup
from langchain_core.prompts.prompt import PromptTemplate
from langchain_core.output_parsers import StrOutputParser
from models.llm_model import get_llm_model
from unstructured.documents.elements import Table
from unstructured.cleaners.core import clean, clean_non_ascii_chars
from models.prompt_engineering import table_summary_prompt
from utils.utils import error_handler
from config import CONFIG
import psutil
import signal
from datetime import datetime
import os
from Structured.ElemmentParser import parse_docx
from Structured.StructureParser import clean_and_structure_elements

def read_docx_file(file_path):
    # 创建一个Document对象
    pydoc = Document(file_path)

    return None, pydoc


def doc_to_pdf_and_load_pages(doc_path, text_splitter):
    word = None
    doc = None
    pdf_content = None
    temp_pdf_path = None

    try:
        # 初始化 COM 库
        pythoncom.CoInitialize()

        # 创建 Word 应用程序对象
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        # 打开文档
        doc = word.Documents.Open(doc_path)

        # 创建临时文件路径
        temp_pdf_path = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name

        # 将文档保存为 PDF
        doc.SaveAs(temp_pdf_path, FileFormat=17)  # 17 表示 PDF 格式

        # 读取 PDF 文件内容
        loader = PyPDFLoader(temp_pdf_path)
        pages = loader.load_and_split(text_splitter=text_splitter)

        return pages

    except Exception as e:
        print(f"发生错误: {str(e)}")
        return None

    finally:
        # 清理资源
        if doc:
            doc.Close()
        if word:
            word.Quit()
        if temp_pdf_path and os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        # 确保 COM 库被正确关闭
        pythoncom.CoUninitialize()


def ensure_persistent_copy(file_path: str) -> str:
    """
    把 gradio 临时文件复制到固定目录，避免被系统清理
    """
    UPLOAD_DIR = "/home/ljc/project/Demo2.9/tmp"

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    # 避免中文路径问题，可以改名为固定文件名
    target_path = os.path.join(UPLOAD_DIR, "input.docx")
    shutil.copy(file_path, target_path)
    return target_path


@error_handler("文档解析失败，请检查Docx文档是否正确")
def zbb_structure_parser_process_docx(docx_file_path,json_file_path=None):
    docx_file_name = os.path.splitext(os.path.basename(docx_file_path))[0]
    name_word = docx_file_name.split("_")[0]

    #  # 先复制文件到固定目录
    # safe_path = ensure_persistent_copy(docx_file_path)
    
    elements = parse_docx(file_path=docx_file_path, parse_table_as_markdown=True)
    _, elements_docx = clean_and_structure_elements(elements)
    # 解析文本 & 清洗文本
    elements_text = zbb_extract_text_from_docx(elements_docx)
    # 解析表格
    elements_table = zbb_extract_table_from_docx(elements_docx,name_word,json_file_path)

    return elements_docx, elements_text, elements_table

def zbb_extract_text_from_docx(elements_docx):
    '''
    文本提取
    '''
    elements_text_dict = {}  # 将列表改为字典
    for idx, chunk in enumerate(elements_docx):
        if chunk.category == 'Paragraph':
            parent_path = elements_docx[idx].elem_metadata.parent_path
            if parent_path not in elements_text_dict:
                elements_text_dict[parent_path] = ""  # 初始化字典中的字符串
            else:
                elements_text_dict[parent_path] += "\n"  # 如果已有字符串，先换行
            elements_text_dict[parent_path] += chunk.text.replace(" ", "")  # 将文本添加到对应的字符串中

    return elements_text_dict

def zbb_extract_table_from_docx(elements_docx,docx_file_name,json_file_path=None):
    elements_table_summary = []#总结后的文本
    elements_table_idx = [] #table在elements_docx中的id
    elements_table_only_text=[]#纯文本
    markdown_table_all = []#markdown格式表格
    skip_table_idx=[]#跳过的table的id
    elements_table_path = []#table的路径
    # 输出每个表格的Markdown格式
    logger.info("提取表格中")
    table_data = {}
    table_id = 0
    for idx, element in enumerate(elements_docx):
        ##step1.提取表格的前一段，后一段，并判断是否为表格标题或者注释。
        if is_valid_table_candidate(element.category, element.text):  # 先判断是否为表格
            # 判断是否为续表
            if idx in skip_table_idx:
                continue 
            precontext_table = ""
            postcontext_table = ""
            continue_table = ""
            continue_table_context = ""#续表的纯文本
            context_continue_table = ""#续表间的文本
            if idx - 1 >= 0 and idx - 1 < len(elements_docx):
                if is_table_caption(
                    elements_docx[idx - 1], elements_docx, idx - 1
                ):  # 再判断前一块是否为表格标题
                    precontext_table = elements_docx[idx - 1].text.replace(
                        " ", ""
                    )  # 只有标题
            if idx - 2 >= 0 and idx - 1 < len(elements_docx):
                if is_table_caption(
                    elements_docx[idx - 2], elements_docx, idx - 1
                ):  # 再判断前一块是否为表格标题
                    precontext_table = (
                        elements_docx[idx - 2].text.replace(" ", "")
                        + "\n"
                        + elements_docx[idx - 1].text.replace(" ", "")
                    )  # 标题和注释（单位相关）
            next_idx = idx + 1
            while next_idx < len(elements_docx) and next_idx >= 0:
                if is_valid_contiue_table_candidate(elements_docx[next_idx],element):#判断是否为续表
                    continue_table+= elements_docx[next_idx].text_as_markdown
                    continue_table_context+=elements_docx[next_idx].text.replace(" ", "")
                    skip_table_idx.append(next_idx)
                if is_valid_contiue_table_candidate(elements_docx[next_idx+1],element):#判断是否为续表
                    if not is_valid_contiue_table_candidate(elements_docx[next_idx],element):
                        context_continue_table = elements_docx[next_idx].text # table后一段字段，相关注释
                        # if '续表' not in context_continue_table and '续上表' not in context_continue_table:
                        #     break
                    continue_table= continue_table+'\n'+context_continue_table+'\n'+elements_docx[next_idx+1].text_as_markdown
                    continue_table_context=continue_table_context+'\n'+context_continue_table+'\n'+elements_docx[next_idx+1].text.replace(" ", "")
                    skip_table_idx.append(next_idx+1)
                else:
                    break
                next_idx += 2
            if next_idx < len(elements_docx) and next_idx >= 0:
                if not is_valid_contiue_table_candidate(elements_docx[next_idx],element):#判断是否为续表
                    postcontext_table = elements_docx[next_idx].text.strip().split('\n')[0]  # table后一段字段，相关注释
                else:
                    postcontext_table = elements_docx[next_idx+1].text.strip().split('\n')[0]  # table后一段字段，相关注释
            # zhy_6_9
            # if next_idx < len(elements_docx) and next_idx >= 0:
            #     if not is_valid_contiue_table_candidate(elements_docx[next_idx], element):  # 不是续表
            #         post_candidate = elements_docx[next_idx]
            #     else:
            #         if next_idx + 1 < len(elements_docx):
            #             post_candidate = elements_docx[next_idx + 1]
            #         else:
            #             post_candidate = None

            #     # 判断是否是表格标题，如果是则设为空；否则提取前几行内容作为后文注释
            #     if post_candidate:
            #         if is_table_caption(post_candidate, elements_docx, elements_docx.index(post_candidate)):
            #             postcontext_table = ""
            #         else:
            #             postcontext_table = post_candidate.text.strip().split('\n')[0]  # table后一段字段，相关注释


            # #===================总结表格内容并转成markdown格式========================#
            markdown_table = element.text_as_markdown
            llm_table_summary = (
                precontext_table 
                + "\n" + markdown_table.replace(" ", "") 
                + "\n"+continue_table
                + "\n" + postcontext_table
            )
            # logger.info(f'文档第{idx}个表格的总结前：{llm_table_summary}')
            # 更新会话
            # print("llm_table_summary:", len(llm_table_summary))
            table_summary_prompt_table=table_summary_prompt(llm_table_summary)
            model = get_llm_model(model_type = "ollama_table")
            # model = get_llm_model(model_type="tongyi")
            rag_chain = (
            {
                "table": lambda _: llm_table_summary,
            }
            | table_summary_prompt_table
            | model
            | StrOutputParser()
            )
            if CONFIG.EXTRACT_MODEL_SUMMARY_EMBEDDING_TABLE_SWITCH or CONFIG.EXTRACT_MODEL_SUMMARY_RERANK_TABLE_SWITCH:
                # 是否含有json文件
                if json_file_path and os.path.exists(json_file_path):
                    # 加载json文件
                    with open(json_file_path, "r", encoding="utf-8") as f:
                        json_data = json.load(f)
                    # 获取表格对应id的值
                    llm_table_summary_after = json_data.get(f"{table_id}")
                    # 如果没有值，则重新生成总结
                    if llm_table_summary_after is None:
                        llm_table_summary_after = rag_chain.invoke({})
                        # llm_table_summary_after = rag_chain.invoke({"table": llm_table_summary})
                        table_entry = {
                                    table_id:llm_table_summary_after,
                                    }
                        json_data.update(table_entry)
                        with open(json_file_path, "w", encoding="utf-8") as f:
                            json.dump(json_data, f, ensure_ascii=False, indent=2)
                    table_id += 1
                else:        
                    llm_table_summary_after= rag_chain.invoke({})#是否总结表格
                    # llm_table_summary_after = rag_chain.invoke({"table": llm_table_summary})
                    if "deepseek" in CONFIG.TABLE_SUMMARY_MODEL:
                        logger.info(f'文档第{table_id}个表格的总结后,删除<think>标签之前的内容：{llm_table_summary_after}')
                        # 使用正则表达式删除 <think>...</think> 之间的内容
                        llm_table_summary_after = re.sub(r'<think>.*?</think>\s*', '', llm_table_summary_after, flags=re.DOTALL)
                        logger.info(f'文档第{table_id}个表格的总结后,删除<think>标签之后的内容：{llm_table_summary_after}')
                    table_entry = {
                    table_id:llm_table_summary_after,
                    }
                    table_data.update(table_entry)
                    table_id += 1 
                # print("llm_table_summary_after:", len(llm_table_summary_after))
            else:
                llm_table_summary_after= llm_table_summary#没有总结就是markdown格式表格（总结处）
            logger.info(f'文档第{table_id-1}个表格的总结后：{llm_table_summary_after}')
            # llm_table_summary+=lllm_table_summary_after#合并表格后的总结内容
            elements_table_summary.append(llm_table_summary_after)
            elements_table_idx.append(idx)  # table,在elements_docx中的id
            markdown_table_all.append(llm_table_summary)
            # #====================end==================================#

            # #===================没有总结表格,没有转成markdown格式========================#
            table_text = (
                precontext_table
                + "\n"
                + element.text.replace(" ", "")
                + "\n"
                +continue_table_context
                + "\n"
                + postcontext_table
            )
            elements_table_only_text.append(table_text)#纯文本
            elements_table_path.append(element.elem_metadata.parent_path)
            # #===================end==============================================#

    # 是否含有json文件 并将表格总结写入json文件
    if json_file_path is None:
        logger.info(f"将表格总结写入json文件")
        dir_part = CONFIG.json_output_folder_path
        # 保存为 JSON 文件
        json_file = f"{dir_part}/{docx_file_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.json"

        with open(json_file, "w", encoding="utf-8") as f:
            json.dump(table_data, f, ensure_ascii=False, indent=2)

        # 删除目录中存储的旧的 JSON 文件
        files_to_delete = []
        for filename in os.listdir(dir_part):
            file_path = os.path.join(dir_part, filename)
            if os.path.isfile(file_path) and filename.startswith(docx_file_name) and file_path != json_file:
                files_to_delete.append(file_path)
        
        for file_path in files_to_delete:
            try:
                os.remove(file_path)
            except Exception as e:
                raise RuntimeError(f"删除文件 {file_path} 失败: {e}")
    ##如果总结了表格需要清理大模型缓存
    if CONFIG.TABLE_SUMMARY_MODEL not in "tongyiqianwen72b":
        if CONFIG.EXTRACT_MODEL_SUMMARY_EMBEDDING_TABLE_SWITCH or CONFIG.EXTRACT_MODEL_SUMMARY_RERANK_TABLE_SWITCH:
            logger.info(f"Terminated process {CONFIG.TABLE_SUMMARY_MODEL}")
            subprocess.run(['ollama', 'stop',CONFIG.TABLE_SUMMARY_MODEL], check=True)    

    return {
        "text": elements_table_only_text,
        "idx": elements_table_idx,
        "summary": elements_table_summary,
        "markdown": markdown_table_all,
        "path": elements_table_path
    }        
                


def extract_text_from_docx(elements_docx):
    excluded_categories = {
        "FigureCaption",
        "PageBreak",
        "Formula",
        "Header",
        "Footer",
        "Image",
        "PageNumber",
        "Table",
    }
    elements_text_list = []
    for idx, chunk in enumerate(elements_docx):
        if (
            chunk.category in excluded_categories
            or is_table_caption(chunk, elements_docx, idx)
            or is_fig_caption(chunk)
            or is_annotation(chunk)
        ):
            continue

        elements_text_list.append(chunk.text.replace(" ", ""))

    return "\n".join(elements_text_list)


def is_valid_caption(text):
    """判断文本是否符合表格标题的基本条件"""
    return (
        len(text) > 0
        and len(text) < 100
        and not text.endswith("。")
        and not text.endswith("：")
    )

# 定义函数：提取表格第一行的数据
def get_first_row_data(soup):
    # 找到表格中的第一个 <tr> 元素
    first_row = soup.find('tr')
    if first_row:
        # 提取第一行中所有 <td> 或 <th> 元素的文本内容
        cells = first_row.find_all(['td', 'th'])
        return [cell.get_text(strip=True).replace(" ", "") for cell in cells]
    return []  # 如果没有找到行，返回空列表

def is_valid_table_candidate(category, text):
    """判断文本是否符合作为表格的基本条件"""
    return category == "Table" and len(text) > 10

def is_valid_contiue_table_candidate(contiue_table,table_frist):
    """判断文本是否符合作为表格的基本条件
        判断第一行是否相同
    """
    if is_valid_table_candidate(contiue_table.category,contiue_table.text):# 再判断后一块是否为表格的续表
            # 提取表头及分隔线
        def extract_header(text):
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            header = lines[0] if len(lines) > 0 else ''
            return header
        
        h1 = extract_header(contiue_table.text_as_markdown)
        h2 = extract_header(contiue_table.text_as_markdown)
        
        
        # 列内容比对
        cols1 = [c.strip() for c in h1.split('|')[1:-1]]
        cols2 = [c.strip() for c in h2.split('|')[1:-1]]
        return contiue_table.category == "Table" and len(contiue_table.text) > 10 and set(cols1).issubset(set(cols2)) 
    else:
        return False

def is_table_caption(chunk, elements_docx, idx):
    """
    判断当前 chunk 是否为表格标题。

    :param chunk: 当前文本块
    :param elements_docx: 所有文档元素的列表
    :param idx: 当前文本块在列表中的索引
    :return: 如果是表格标题，返回 True；否则返回 False
    """

    # 清理当前文本和后续两个文本的空格
    text = chunk.text.replace(" ", "")
    next_chunk = None
    next_next_chunk = None
    if idx + 1 < len(elements_docx):
        next_chunk = elements_docx[idx + 1]
        text_1 = next_chunk.text.replace(" ", "")
    else:
        return False

    if idx + 2 < len(elements_docx):
        next_next_chunk = elements_docx[idx + 2]
        text_2 = next_next_chunk.text.replace(" ", "")
    else:
        text_2 = ""

    # 检查后续的两个块是否为符合条件的表格标题
    if (
        next_chunk
        and is_valid_table_candidate(next_chunk.category, text_1)
        and is_valid_caption(text)
    ):
        return True

    if (
        next_next_chunk
        and is_valid_table_candidate(next_next_chunk.category, text_2)
        and is_valid_caption(text_1)
        and is_valid_caption(text)
    ):
        return True

    return False


def is_fig_caption(chunk):
    text = chunk.text.replace(" ", "") if chunk.text else ""
    if len(text) < 100 and text.startswith("图"):
        if len(text) > 1 and text[1].isdigit():
            return True
    return False


def is_annotation(chunk):
    text = chunk.text.replace(" ", "") if chunk.text else ""
    if len(text) >= 2 and len(text) <= 100 and text.startswith("注"):
        return True
    return False


# 将每个表格转化为Markdown格式
def html_table_to_markdown(html_table):
    # 解析HTML表格
    soup = BeautifulSoup(html_table, "html.parser")
    rows = soup.find_all("tr")
    # 提取表格内容
    markdown_lines = []
    for i, row in enumerate(rows):
        cols = [col.get_text(strip=True) or " " for col in row.find_all(["td", "th"])]
        line = "| " + " | ".join(cols) + " |"
        markdown_lines.append(line)
        # 添加分隔线
        if i == 0:
            header_line = "| " + " | ".join(["---"] * len(cols)) + " |"
            markdown_lines.append(header_line)

    # 转换为Markdown格式
    return "\n".join(markdown_lines)


if __name__ == "__main__":
    safe_path = "/home/ljc/project/Demo2.9/tmp/input.docx"

    elements = parse_docx(file_path=safe_path, parse_table_as_markdown=True)
    _, elements_docx = clean_and_structure_elements(elements)
    # 解析文本 & 清洗文本
    elements_text = zbb_extract_text_from_docx(elements_docx)
    # 解析表格
    elements_table = zbb_extract_table_from_docx(elements_docx,name_word,json_file_path)


