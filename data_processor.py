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


@error_handler("excel表格解析错误，请检查excel表格样式是否合法然后重传")
def unstructured_process_docx(docx_file_path,json_file_path=None):
    _, ext = os.path.splitext(docx_file_path)
    if ext.lower() == ".doc":
        elements_docx = partition_doc(filename=docx_file_path)
    else:
        elements_docx = partition_docx(filename=docx_file_path)
    docx_file_name = os.path.splitext(os.path.basename(docx_file_path))[0]
    name_word = docx_file_name.split("_")[0]
    # 解析文本 & 清洗文本
    elements_text = extract_text_from_docx(elements_docx)
    # 解析表格
    elements_table_only_text, elements_table_idx,elements_table_summary,markdown_table_all = extract_table_from_docx(elements_docx,name_word,json_file_path)

    return elements_docx, elements_text, elements_table_only_text, elements_table_idx,elements_table_summary,markdown_table_all

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
        # 解析HTML
        soup1 = BeautifulSoup(contiue_table.metadata.text_as_html, 'html.parser')
        soup2 = BeautifulSoup(table_frist.metadata.text_as_html, 'html.parser')

        # 获取两个表格第一行的数据
        first_row_data1 = get_first_row_data(soup1)
        first_row_data2 = get_first_row_data(soup2)
        # # 获取表头
        # headers1 = [header.text for header in soup1.find_all('th')]
        # headers2 = [header.text for header in soup2.find_all('th')]
        # print(f"first_row_data1:{set(first_row_data1)},first_row_data2:{set(first_row_data2)}")
        # logger.info(f"{set(first_row_data1).issubset(set(first_row_data2))}")
        return contiue_table.category == "Table" and len(contiue_table.text) > 10 and set(first_row_data1).issubset(set(first_row_data2)) 
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


def extract_table_from_docx(elements_docx,docx_file_name,json_file_path=None):
    elements_table_summary = []#总结后的文本
    elements_table_idx = [] #table在elements_docx中的id
    elements_table_only_text=[]#纯文本
    markdown_table_all = []#markdown格式表格
    skip_table_idx=[]#跳过的table的id
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
                    continue_table+= html_table_to_markdown(elements_docx[next_idx].metadata.text_as_html)
                    continue_table_context+=elements_docx[next_idx].text.replace(" ", "")
                    skip_table_idx.append(next_idx)
                if is_valid_contiue_table_candidate(elements_docx[next_idx+1],element):#判断是否为续表
                    if not is_valid_contiue_table_candidate(elements_docx[next_idx],element):
                        context_continue_table = elements_docx[next_idx].text # table后一段字段，相关注释
                        # if '续表' not in context_continue_table and '续上表' not in context_continue_table:
                        #     break
                    continue_table= continue_table+'\n'+context_continue_table+'\n'+html_table_to_markdown(elements_docx[next_idx+1].metadata.text_as_html)
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

            #===================每一行作为一个chunk========================#
            # markdown_table = html_table_to_markdown(element.metadata.text_as_html)
            # # 按行分割表格
            # lines = markdown_table.strip().split('\n')
            # # 过滤掉表头和分隔线，提取数据行
            # data_rows = [line for line in lines if line.startswith('|') and not line.startswith('|-----')]
            # # 合并表头和数据行
            # for i, row in enumerate(data_rows):
            #     # 跳过第一行，表头
            #     if i==0:
            #         table_frist= row
            #         continue
            #     markdown_table_row=precontext_table+'\n'+table_frist+'\n'+row+'\n'+postcontext_table
            #     elements_table.append(markdown_table_row)
            #     elements_table_idx.append(idx)
            #====================end==================================#
            
            #===================每一行作为一个chunk并大模型总结========================#
            # markdown_table = html_table_to_markdown(element.metadata.text_as_html)
            # logger.info(f'文档第{idx}个表格，原始表格内容：{markdown_table}') 
            # # 按行分割表格
            # lines = markdown_table.strip().split('\n')
            # for i, row in enumerate(lines):
            #     # 跳过第一行，表头
            #     if i==0:
            #         table_frist= row
            #         continue
            #     # 跳过'|---|'行
            #     if '---' in row:
            #         continue
            #     # 合并表头
            #     if table_frist.strip().split('|')[1:-1][0] == row.strip().split('|')[1:-1][0]:
            #         table_frist=table_frist+ '\n' + row
            #         continue
            #     # 合并表格前后文
            #     markdown_table_row=precontext_table+'\n'+table_frist+'\n'+row+'\n'+postcontext_table
            #     logger.info(f'文档第{idx}个表格，第{i}行表格的总结前：{markdown_table_row}')
            #     # # 更新会话
            #     table_summary_prompt_table=table_summary_prompt()
            #     model = get_llm_model(model_type = "ollama_table")
            #     rag_chain = (
            #     {
            #         "table": lambda _: markdown_table_row,
            #     }
            #     | table_summary_prompt_table
            #     | model
            #     | StrOutputParser()
            #     )
            #     llm_table_summary_before= rag_chain.invoke({})#没有合并表格的总结内容
            #     logger.info(f'文档第{idx}个表格，第{i}行表格的总结后：{llm_table_summary_before}')
            #     elements_table.append(llm_table_summary_before)
            #     elements_table_idx.append(idx)
            #===================end==========================#

            # #===================总结表格内容并转成markdown格式========================#
            markdown_table = html_table_to_markdown(element.metadata.text_as_html)
            llm_table_summary = (
                precontext_table 
                + "\n" + markdown_table.replace(" ", "") 
                + "\n"+continue_table
                + "\n" + postcontext_table
            )
            # logger.info(f'文档第{idx}个表格的总结前：{llm_table_summary}')
            # 更新会话
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
                        llm_table_summary_after= rag_chain.invoke({})
                        table_entry = {
                                    table_id:llm_table_summary_after,
                                    }
                        json_data.update(table_entry)
                        with open(json_file_path, "w", encoding="utf-8") as f:
                            json.dump(json_data, f, ensure_ascii=False, indent=2)
                    table_id += 1
                else:        
                    llm_table_summary_after= rag_chain.invoke({})#是否总结表格
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

            
                
    return elements_table_only_text, elements_table_idx,elements_table_summary,markdown_table_all


def unstruct_text_cleaner(text):
    text = clean(
        text,
        extra_whitespace=True,
        bullets=True,
    )

    return text


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
    print(len("图1.1描述的是奥斯卡假大空建行的卡十大科技沙克斯的"))
    assert (
        is_fig_caption("图1.1描述的是奥斯卡假大空建行的卡十大科技沙克斯的") == False
    )  # 带小数的图号

    print("所有测试用例通过！")


