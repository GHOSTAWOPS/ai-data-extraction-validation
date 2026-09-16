from typing import List, Dict
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt
import pandas as pd
import subprocess
import re
import os
import json
import random
import zipfile
from datetime import datetime
from docx2pdf import convert
import xml.dom.minidom



class DocxManager:
    """
    该类用于管理一个 Docx 文件及其相关操作，在实例化时接受两个地址：
    1. file_path 需要进行管理的文件地址，必须以 .docx 或者 .doc 结尾，否则会报错
    2. save_path 文件发生修改后将要被保存到的目录地址，默认为当前工作目录
    """

    process_file_path: str
    process_file_name: str

    timestamp: str

    save_file_path: str
    save_file_name: str

    docx_file: Document
    sentences_info: List[dict]
    bookmarks_info: Dict

    def __init__(self, file_path: str, save_path: str = None):
        file_name = os.path.basename(file_path)
        self.process_file_name, file_extension = os.path.splitext(file_name)

        if file_extension.lower() == ".doc":
            self.process_file_path = convert_doc_to_docx(file_path)
        elif file_extension.lower() == ".docx":
            self.process_file_path = file_path
        else:
            raise ValueError("不支持的文件格式，请提供 .doc 或 .docx 文件。")

        if self.process_file_path is None:
            raise ValueError("该 .doc 文件转换为 .docx 文件失败")

        self.timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        self.docx_file = Document(self.process_file_path)
        self.sentences_info = self.__generate_sentences_info()
        self.bookmarks_info = self.__init_bookmarks_info()

        # 生成新的文件名
        timestamp_pattern = r"\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}"
        match = re.search(f"({timestamp_pattern})$", self.process_file_name)
        if match:
            new_base = re.sub(f"_{timestamp_pattern}$", "", self.process_file_name)
        else:
            new_base = self.process_file_name
    
        self.save_file_name = f"{new_base}_{self.timestamp}"

        if save_path is None:
            # self.save_path = os.path.join(os.getcwd(), f"{self.file_name}.docx")
            self.save_file_path = os.getcwd()
        else:
            # self.save_path = os.path.join(save_path, f"{self.file_name}.docx")
            self.save_file_path = save_path

    def __generate_sentences_info(self):
        """
        从 self.docx 文件中分析、记录所有句子的信息，每一句的相关信息存储为一部字典。sentences_info 是字典的列表
        字典结构如下:
        {
            "para_idx": 句子所在文段序号
            "context": 句子清洗后的文本内容
            "start_char_idx": 句子在其段落中的起始字符的位置
            "end_char_idx": 句子在其段落中的结束字符的位置 
        }
        """
        sentence_info_list = []
        delimiters = r"[。！？\n：；，、]"  # 句子分隔符
        for para_idx, paragraph in enumerate(self.docx_file.paragraphs):
            paragraph_text = ""
            for run in paragraph.runs:
                paragraph_text += run.text
            sentences = re.split(delimiters, paragraph_text)

            start_idx = 0
            for sentence in sentences:
                sentence = sentence.strip()
                if sentence:
                    end_idx = start_idx + len(sentence)
                    sentence_info = {
                        "para_idx": para_idx,
                        "context": re.sub(r"\s+", "", sentence),
                        "start_char_idx": start_idx,
                        "end_char_idx": end_idx,
                    }
                    sentence_info_list.append(sentence_info)
                    start_idx = end_idx + 1  # 跳过分隔符
        return sentence_info_list

    def __init_bookmarks_info(self):
        """
        初始化书签信息，返回一个字典，包含下列信息：
        1. next_id: 下一个书签的 ID
        2. bookmark_list: 书签列表，每个书签包含下列信息：
            - idx: 书签 ID
            - pos: 书签在文档中的位置，包含起始段落、起始字符、结束段落、结束字符
            - name: 书签名称
            - text: 书签对应的文本
        """
        json_filename = f"output_jsons/bookmark_jsons/{self.process_file_name}.json"

        # 检查文件是否存在
        if os.path.isfile(json_filename):
            try:
                # 加载现有文件
                with open(json_filename, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, FileNotFoundError) as e:
                print(f"警告：文件 {json_filename} 存在但无法解析，返回空字典。错误信息：{e}")
                return {}
        else:
            return {"next_id": 502, "bookmark_list": []}

        # bookmarks = []
        # # 遍历所有段落和表格，按文档顺序提取书签
        # max_id = 502
        # for para_idx, para in enumerate(self.docx_file.paragraphs):
        #     element = para._p
        #     bookmark = element.find(qn("w:bookmarkStart"))
        #     if bookmark is not None:
        #         bookmark_name = bookmark.get(qn("w:name"))
        #         bookmark_id = int(bookmark.get(qn("w:id")))
        #         bookmarks.append((para_idx, bookmark_id, bookmark_name))
        #         if bookmark_id >= max_id:
        #             max_id = bookmark_id + 1

        # # 按 bookmark_id 排序
        # bookmarks.sort(key=lambda x: x[1])

        # return bookmarks, max_id

    def find_text_in_docx(self, structure_path: str, target_text: str):
        """
        structure_path: 文档结构路径，例如："工程概况|政策背景|生态流量保障工作发展"

        通过在 self.sentence_info 中匹配 target_text 的起始句与结束句，找到 target_text 的起始段落与起始字符，结束段落与结束字符
        """
        # target_text预处理
        if not target_text.strip():
            raise KeyError("警告：待匹配文段为空！")

        # target_text预处理
        sentences = extract_sentences_info(target_text)

        # 待返回的匹配结果
        match = {
            "start_para": -1,
            "start_char_idx": -1,
            "end_para": -1,
            "end_char_idx": -1,
        }
        result = {
            "not_found": True,
            "chunk_pos": match
        }

        # 根据structure_path找到匹配的起始位置
        if not structure_path:
            start_pos = 0
        else:
            nodes = structure_path.replace(" ", "").split("|")
            nodes = [node for node in nodes if node]  # 去除空节点
            start_pos = 0
            for node in nodes:
                for idx, sentence in enumerate(self.sentences_info[start_pos:]):
                    if node in sentence["context"]:
                        start_pos += idx
                        break
        print(f"开始匹配文本：{self.sentences_info[start_pos]['context']}")

        # 优先进行全句匹配尝试（直接在句子窗口中比对 target_text）
        full_text = "".join(sentences)  # 要匹配的完整文本
        window_size = 25

        for i in range(start_pos, len(self.sentences_info) - window_size + 1):
            window = self.sentences_info[i: i + window_size]
            context_list = [w["context"] for w in window]

            # 拼接字符串，同时记录字符的累积位置
            joined_text = ""
            offsets = []  # offsets[i] 表示 window[i] 在 joined_text 中的起始位置
            for ctx in context_list:
                offsets.append(len(joined_text))
                joined_text += ctx

            pos = joined_text.find(full_text)
            if pos != -1:
                # 找到开始和结束句索引
                start_idx = end_idx = None
                for j, offset in enumerate(offsets):
                    if offset <= pos < offset + len(context_list[j]):
                        start_idx = j
                    if offset <= pos + len(full_text) <= offset + len(context_list[j]):
                        end_idx = j
                        break

                if start_idx is not None and end_idx is not None:
                    start_info = window[start_idx]
                    end_info = window[end_idx]
                    match["start_para"] = start_info["para_idx"]
                    match["start_char_idx"] = start_info["start_char_idx"]
                    match["end_para"] = end_info["para_idx"]
                    match["end_char_idx"] = end_info["end_char_idx"]
                    result["not_found"] = False
                    return result

        # 仅对开头句以及结尾句（以及随机中间句）进行匹配
        for i in range(0, 3):
            if i == 0:
                start_sen = sentences[0]
                end_sen = sentences[-1]
            elif i == 1:
                start_sen = sentences[0]
                end_sen = sentences[-2]
            elif i == 2:
                start_sen = sentences[1]
                end_sen = sentences[-1]

            is_start_found = False
            is_end_found = False

            # 随机挑两句中间句(排除首尾)
            if len(sentences)>= 4:
                mid_sens=random.sample(sentences[1:-1],2)
            else:
                mid_sens =sentences[0:-1]
            is_mids_found ={m: False for m in mid_sens}

            for sen_start_idx, sen_start_info in enumerate(self.sentences_info[start_pos:]):
                # if sen_start_info["context"] == start_sen:
                if sen_start_info["context"].endswith(start_sen):
                    is_start_found = True

                    # 寻找中间句
                    window = self.sentences_info[sen_start_idx + start_pos: sen_start_idx + start_pos + 25]
                    for w in window:
                        for m in mid_sens:
                            if w["context"] == m:
                                is_mids_found[m] = True
                    # 如果任意中间句未找到，则跳过该起始句
                    if not all(is_mids_found.values()):
                        is_start_found = False
                        continue

                    # 寻找结束句
                    for sen_end_idx, sen_end_info in enumerate(
                        self.sentences_info[start_pos:]
                    ):
                        if (
                            # sen_end_info["context"] == end_sen
                            sen_end_info["context"].startswith(end_sen)
                            and sen_end_idx - sen_start_idx <= 25
                        ):
                            match["start_para"] = sen_start_info["para_idx"]
                            match["start_char_idx"] = sen_start_info["start_char_idx"]
                            match["end_para"] = sen_end_info["para_idx"]
                            match["end_char_idx"] = sen_end_info["end_char_idx"]
                            is_end_found = True
                            break
                    if is_start_found and is_end_found:
                        break
            
            if match["start_para"] != -1 and match["end_para"] != -1:
                result["not_found"] = False

            if result["not_found"] is False:
                break

        return result


    def add_bookmark_by_pos(self, chunk_pos: dict, target_text: str):
        """
        在 self.docx_file 文件中匹配指定文本并添加书签。
        """
        # 如果该标签已经存在，则直接返回该标签名称
        for bookmark in self.bookmarks_info["bookmark_list"]:
            if chunk_pos == bookmark["pos"]:
                return bookmark["name"]

        # 声明新加入标签的标签名以及标签id
        bookmark_idx = self.bookmarks_info["next_id"]
        bookmark_name = f"参考文段_{bookmark_idx}"

        start_para = self.docx_file.paragraphs[chunk_pos["start_para"]]
        start_idx = chunk_pos["start_char_idx"]
        end_para = self.docx_file.paragraphs[chunk_pos["end_para"]]
        end_idx = chunk_pos["end_char_idx"]

        # 创建书签的起始标记以及结尾标记
        bookmark_start = OxmlElement("w:bookmarkStart")
        bookmark_start.set(qn("w:name"), bookmark_name)
        bookmark_start.set(qn("w:id"), f"{bookmark_idx}")
        bookmark_end = OxmlElement("w:bookmarkEnd")
        bookmark_end.set(qn("w:id"), f"{bookmark_idx}")
        bookmark_end.set(qn("w:name"), bookmark_name)

        # 在起始段落中添加 bookmark_start
        len_sum = 0
        for run_idx, run in enumerate(start_para.runs):
            if len_sum == start_idx:
                start_para._p.insert(start_para._p.index(run._r), bookmark_start)
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW  # 标记为黄色
                break
            elif len_sum < start_idx < len(run.text) + len_sum:
                split_idx = start_idx - len_sum
                new_run = start_para.add_run(run.text[split_idx : len(run.text)])
                run.text = run.text[:split_idx]

                start_para._p.insert(start_para._p.index(run._r) + 1, new_run._r)
                start_para._p.insert(start_para._p.index(run._r) + 1, bookmark_start)
                new_run.font.highlight_color = WD_COLOR_INDEX.YELLOW  # 标记为黄色
                break
            else:
                len_sum += len(run.text)

        # 在结束段落中添加 bookmark_end
        len_sum = 0
        for run_idx, run in enumerate(end_para.runs):

            if len_sum + len(run.text) == end_idx:
                end_para._p.insert(end_para._p.index(run._r) + 1, bookmark_end)
                run.font.highlight_color = WD_COLOR_INDEX.TURQUOISE  # 标记为蓝色
                break
            elif len_sum < end_idx < len(run.text) + len_sum:
                split_idx = end_idx - len_sum + 1
                new_run = end_para.add_run(run.text[split_idx : len(run.text)])
                run.text = run.text[:split_idx]
                end_para._p.insert(end_para._p.index(run._r) + 1, new_run._r)
                end_para._p.insert(end_para._p.index(run._r) + 1, bookmark_end)
                run.font.highlight_color = WD_COLOR_INDEX.TURQUOISE  # 标记为蓝色
                break
            else:
                len_sum += len(run.text)

        # 更新bookmark_info
        self.bookmarks_info["next_id"] += 1
        bookmark_info = {"idx": bookmark_idx, "pos": chunk_pos, "name": bookmark_name, "text": target_text}
        self.bookmarks_info["bookmark_list"].append(bookmark_info)

        return bookmark_name

    def save_docx_file(self):
        """
        将 self.docx_file 保存到 self.save_path
        将 self.bookmarks_info 保存到 output_jsons/bookmark_jsons
        """
        # 保存 docx 文件
        docx_path = os.path.join(self.save_file_path, f"{self.save_file_name}.docx")
        self.docx_file.save(docx_path)
        print(f"Docx 文件已保存到 {docx_path}")

        # 保存书签信息
        if not os.path.exists("output_jsons/bookmark_jsons"):
            os.makedirs("output_jsons/bookmark_jsons")
        json_path = f"output_jsons/bookmark_jsons/{self.save_file_name}.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(self.bookmarks_info, f, ensure_ascii=False, indent=4)
        print(f"书签信息已保存到 {json_path}")


def read_docx_file(doc_path: str, get_pdf=False):
    """
    将 doc_path 地址所指向的 docx 文件读取为 Document 类。
    同时，将该 docx 文件转换为 pdf 格式并储存在当前工作目录下。
    """
    if get_pdf:
        pdf_path = os.path.join(os.getcwd(), doc_path[:-5] + ".pdf")
        convert(doc_path, pdf_path)
    return Document(doc_path)


def convert_doc_to_docx(doc_path: str):
    """
    使用 LibreOffice 将 .doc 文件转换为 .docx 文件
    """
    # 检查文件是否存在
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"文件不存在: {doc_path}")

    # 调用 LibreOffice 命令行工具进行转换
    output_dir = os.path.dirname(doc_path)
    command = [
        "libreoffice",
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
        print(f"转换失败: {e}")
        return None


def extract_table_info(para: str):
    """
    分析 para 是否表示一个 markdown 型式的表格，返回字典如下：
    result = {
        "is_table": 表示是否为表格的标识位
        "tbl_info": 表格的序号 + 表格的名称
    }
    """
    # 清洗空格并提取前两行
    lines = [re.sub(r"\s+", " ", line.strip()) for line in para.split("\n")[:2]]
    line1 = lines[0] if len(lines) > 0 else ""
    line2 = lines[1] if len(lines) > 1 else ""

    # 定义表序号正则表达式
    table_num_pattern = r"(?:续)?表\s*\d+(?:[\-\—－\.．]\d+)*"

    # 定义返回结果字典
    result = {
        "is_table": False,
        "tbl_info": para,
    }

    # 情况1：表信息在一行
    if re.match(rf"^\s*({table_num_pattern})\s*(.*)", line1):
        match = re.match(rf"^\s*({table_num_pattern})\s*(.*)", line1)
        if match:
            table_num = match.group(1)
            table_name = match.group(2)
            result["is_table"] = True
            result["tbl_info"] = f"{table_num} {table_name}"
            return result

    # 情况2：表信息在两行
    elif line2:
        # 检查第二行是否包含表序号
        if re.match(rf"^\s*({table_num_pattern})\s*(.*)", line2):
            match = re.match(rf"^\s*({table_num_pattern})\s*(.*)", line2)
            if match:
                table_num = match.group(1)
                table_name = line1
                result["is_table"] = True
                result["tbl_info"] = f"{table_num} {table_name}"
                return result

    return result


def extract_sentences_info(para: str):
    """
    用正则表达式按照 “句号、感叹号、问号、回车、冒号、分号” 对一个段落进行切分，去掉空白并过滤空字符串
    """
    para = para.strip().strip('“”"\'')

    idx = para.find("\n")
    if idx != -1:
        para = para[idx+1:]
    else:
        para = ""  # 如果没有换行符，整段删掉

    sentences = re.split(r"[。！？\n：；，、]", para)
    sentences = [
        re.sub(r"\s+", "", sentence.strip())
        for sentence in sentences
        if sentence.strip()
    ]

    return sentences


def extract_docx_to_txt(docx_path, output_dir):
    """
    将 docx 文件解压成 xml，并将 xml 内容格式化后写入 txt 文件
    :param docx_path: 输入的 word 文件路径 (.docx)
    :param output_dir: 输出目录路径
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with zipfile.ZipFile(docx_path, "r") as docx:
        for file_name in docx.namelist():
            if file_name.endswith(".xml"):
                xml_bytes = docx.read(file_name)
                xml_content = xml_bytes.decode("utf-8")

                try:
                    # 格式化 XML，带缩进
                    dom = xml.dom.minidom.parseString(xml_content)
                    pretty_xml = dom.toprettyxml(indent="  ", encoding="utf-8")
                    xml_text = pretty_xml.decode("utf-8")
                except Exception:
                    # 如果解析失败，就保持原始内容
                    xml_text = xml_content

                # 保存成 txt 文件
                txt_file_path = os.path.join(
                    output_dir, file_name.replace("/", "_") + ".txt"
                )
                with open(txt_file_path, "w", encoding="utf-8") as f:
                    f.write(xml_text)

                print(f"已导出: {txt_file_path}")


def remove_toc_and_before(docx_path, output_path):
    """
    删除最后一个 TOC（含 SDT 内 TOC）及其之前的内容
    """
    ext = os.path.splitext(docx_path)[1].lower()
    if ext == ".doc":
        docx_path = convert_doc_to_docx(docx_path)
        output_path+='x'
        if docx_path is None:
            raise RuntimeError("DOC 转换为 DOCX 失败")

    elif ext != ".docx":
        raise ValueError(f"不支持的文件类型: {ext}，只支持 .doc 或 .docx")


    print(f"删除临时文件：{docx_path}")
    doc = Document(docx_path)
    body_elements = list(doc.element.body)

    last_toc_idx = None
    for idx, el in enumerate(body_elements):
        if contains_toc(el, doc):
            last_toc_idx = idx

    if last_toc_idx is not None:
        # 删除最后一个 TOC 及其之前的所有块元素
        for i in range(last_toc_idx, -1, -1):
            body_elements[i].getparent().remove(body_elements[i])

    doc.save(output_path)


def is_toc_paragraph(p):
    """
    判断段落是否为目录段落
    """
    # 样式名检测
    pPr = p._element.find(qn("w:pPr"))
    if pPr is not None:
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is not None:
            style_val = pStyle.get(qn("w:val"), "")
            if "toc" in style_val.lower():
                return True
    # 字段检测
    for instr in p._element.findall(".//" + qn("w:instrText")):
        if instr.text and "toc" in instr.text.lower():
            return True
    return False


def contains_toc(el, doc):
    """
    判断元素中是否包含 TOC
    el 可以是段落、表格或 sdt
    """
    tag = el.tag
    if tag == qn("w:p"):
        return is_toc_paragraph(Paragraph(el, doc))
    elif tag == qn("w:sdt"):
        # SDT 内部内容递归检查
        sdt_content = el.find(qn("w:sdtContent"))
        if sdt_content is not None:
            for child in sdt_content:
                if contains_toc(child, doc):
                    return True
    # 表格可以忽略 TOC 检测
    return False


