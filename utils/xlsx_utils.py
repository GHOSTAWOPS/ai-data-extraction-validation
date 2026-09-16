from typing import List, Optional
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.text import InlineFont
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.rich_text import TextBlock, CellRichText
from utils.docx_utils import DocxManager
from utils.docx_utils import extract_table_info
from utils.logger_util import logger
from utils.heading_mapper import get_heading_level_and_number
from copy import deepcopy, copy
import pandas as pd
import re
import math

# print("math module imported successfully:", math.ceil(2.5))

from utils.translate_unit import find_matching_unit
from utils.utils import error_handler


@error_handler("excel表格解析错误，请检查excel表格样式是否合法然后重传")
def read_xlsx_file(xlsx_path: str):
    """
    将 xlsx_path 地址所指向的 .xlsx 文件读取为 Workbook 类。
    """
    wb = load_workbook(filename=xlsx_path, rich_text=True)
    return wb


def read_xlsx_sheet(wb: Workbook, verify=False):
    """
    读取 xlsx 文件中的工作表。

    参数：
        wb (Workbook): openpyxl 的工作簿对象。
        verify (bool): 是否进行核验流程并复制最后一个工作表。

    返回：
        Tuple[Worksheet, int]：返回选定的工作表对象和原工作表数量减 1。
    """
    if verify:
        sheets = wb.worksheets
        sheet_count = len(sheets)
        last_sheet = sheets[-1]
        last_sheet.title = "Sheet" + str(sheet_count)

        if sheet_count > 1:
            # 复制最后一个工作表并返回新工作表及原数量+1
            ws = wb.copy_worksheet(last_sheet)
            check_count = sheet_count + 1
        else:
            v1 = last_sheet["F1"].value
            v2 = last_sheet["G1"].value
            if v1 == "参考文段" and v2 == "冲突文段":
                ws = wb.copy_worksheet(last_sheet)
                check_count = sheet_count + 1
            else:
                ws = last_sheet
                check_count = sheet_count

        ws.title = "NEW! Sheet" + str(check_count)
    else:
        ws = wb.active
        check_count = 0
    return ws, check_count - 1


def get_merged_info(ws: Worksheet):
    """
    检查工作表中从第二行开始的每一行，“数量”及“单位”两列的单元格是否被合并。

    参数：
        ws (Worksheet): openpyxl 的工作表对象。

    返回：
        List[bool]: 长度为 ws.max_row-1 的布尔列表。列表中第 i 项对应第 i+2 行，
                    True 表示该行在指定列范围内所有单元格均属于某个合并区域，False 则表示至少存在未合并的单元格。
    """
    merged_ranges = ws.merged_cells.ranges
    merged_rows = []

    for row_idx in range(2, ws.max_row + 1):  # 从第2行开始
        is_merged = True
        for col in [3, 4]:  # 固定检测C列和D列（列索引从1开始）
            cell = ws.cell(row=row_idx, column=col)
            in_merge = any(
                merged.bounds[1] <= cell.row <= merged.bounds[3] and
                merged.bounds[0] <= cell.column <= merged.bounds[2]
                for merged in merged_ranges
            )
            if not in_merge:
                is_merged = False
                break
        merged_rows.append(is_merged)
    # merged_rows = []
    # for row_idx in range(2, ws.max_row + 1):
    #     is_merged = True
    #     for col in [3, 4]:  # C列和D列
    #         cell = ws.cell(row=row_idx, column=col)
    #         # 直接检查单元格是否属于合并单元格
    #         if not isinstance(cell, MergedCell):
    #             is_merged = False
    #             break
    #     merged_rows.append(is_merged)


    return merged_rows


def analysis_keyword_list_from_ws(ws: Worksheet):
    """
    从工作表中提取用于查询的关键字列表和标识列表。

    参数：
        ws (Worksheet): openpyxl 的工作表对象，数据从第2行开始，第一列为标题/层级，
                        第二列为关键词。

    返回：
        Tuple:
            flag_list (List[bool]): 与行数相同长度的布尔列表。True 表示该关键词需要作为独立查询，
                                    False 表示无需查询。
            word_list (List[str]): 清洗并根据层级关系拼接后的关键词列表。
    """    
    # 1. 构建层级信息列表 levl_list：
    #    每项为 [level, number_str]，level 表示标题层级，number_str 为提取的编号或 "10000" 。
    levl_list = []
    temp_parent_level = 0 # 当前的上级标题级别，初始化为0
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        raw_levl = row[0].value if hasattr(row[0], 'value') else None
        if raw_levl is None or raw_levl.strip == "":
            # 如果标题为空，级别设置为 temp_parent_level + 1
            levl_list.append([temp_parent_level + 1, str(10000)])  # 如果标题为空，设置为10000
        else:
            # 处理标题的级别和编号
            level, number = get_heading_level_and_number(raw_levl)
            if level == 10000:
                # 如果标题不符合任何格式，级别设置为 temp_parent_level + 1
                levl_list.append([temp_parent_level + 1, str(10000)])
            else:
                levl_list.append([level, str(number)])
                temp_parent_level = level

    # 2. 构建初始关键词列表 word_list：去除空格
    word_list: List[str] = [
        (row[1].value or "").replace(" ", "")
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row)
    ]

    # 3. 根据层级关系生成 flag_list：
    #    遍历每个关键词，若其下方存在直接下级（level+1），则自己不需独立查询
    flag_list = []
    for i in range(0, len(levl_list)):
        flag = True
        heading_level_i = levl_list[i][0]
        for j in range(i + 1, len(levl_list)):
            heading_level_j = levl_list[j][0]

            # 如果索引 i 与 索引 j 满足：heading_level_j == heading_level_i + 1 
            # 则 j 是 i 的直接下级
            if heading_level_j == heading_level_i + 1:
                word_list[j] = word_list[i] + "|" + word_list[j]
                flag = False

            # 如果索引 j 的级别小于等于 i 的级别，
            # 则说明 j 不是 i 的下级
            elif heading_level_j <= heading_level_i:
                break

            # 如果索引 j 的级别小于 i 的级别且差距较大
            # 则说明 j 不是 i 的直接下级，不做处理
            else:
                pass
        
        flag_list.append(flag)

    return flag_list, word_list


def get_answers_from_xlsx(
    ws: Worksheet, 
    need_query: List[bool], 
    is_special: List[bool]
):
    """
    根据 need_query 和 is_special 两个布尔列表，从工作表中提取答案和答案路径。

    参数：
        ws (Worksheet): openpyxl 的工作表对象。
        need_query (List[bool]): 与数据行对应，True 表示需要查询该行答案。
        is_special (List[bool]): 与数据行对应，True 表示该行为特殊行，仅提取某一列。

    返回：
        Tuple:
            asr_list (List[Optional[str]]): 提取的答案列表，按行顺序。
            asr_path_list (List[Optional[str]]): 提取的答案路径列表，按行顺序。
    """
    asr_list = []
    asr_path_list = []

    for row_id, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row)):
        if need_query[row_id]:
            # 提取答案
            if is_special[row_id]:
                value = row[2].value or "无"
                asr_list.append(value)
            else:
                val2 = row[2].value or "无"
                val3 = row[3].value or "无"
                asr_list.append(flatten_cell_value(val2) + "|" + str(val3))

            # 提取“备注”中的答案路径部分
            asr_source = str(row[4].value) or " "
            pattern = r"答案路径：(.*?)\s*答案证据："
            match = re.search(pattern, asr_source, re.DOTALL)
            if match:
                asr_path_list.append(match.group(1).strip())
            else:
                asr_path_list.append("无")
        else:
            # 不需要查询的行，填 None
            asr_list.append(None)
            asr_path_list.append(None)

    return asr_list, asr_path_list


@error_handler("从excel中提取表格信息出错，请检查excel表格样式是否合法然后重传")
def init_query_info_from_xlsx(ws: Worksheet, verify=False):
    """
    初始化功能特性表的查询信息，返回一个包含各类分析结果的字典。

    参数：
        ws (Worksheet): openpyxl 表格对象
        verify (bool): 是否提取答案及路径

    返回：
        Dict[str, Any]: 初始化结果字典，结构如下：
        {
            "sheet_name": # 该工作表的名称
            "rows": # 该表总行数
            "need_query": # 代表每一行是否需要查询，True为需要，False则忽略跳过(忽略第一行表头)
            "is_special": # 如果为True,代表这一项只需要查一个和“方式”“形式”有关的结果
            "keywords": # 每一行的第一列内容，表示要查询的内容
            "answers": # 需要查询的关键词的答案内容，仅在 verify=True 时存在；
                        对于一个 keyword :
                        当 need_query=False 时，answer 为 None；
                        当 is_special=True 时，answer 为描述性的短句；
                        当 is_special=False 时，answer为 "单位|数量" 的字符串
            "answers_path": # 答案内容所在段落的文章层次结构信息，仅在 verify=True 时存在；
        }
    """
    # 1. 删除多余行（前3列都为空的行）
    if verify:
        unmerge_all_cells(ws)
    for row in range(ws.max_row, 1, -1):
        values = [ws.cell(row=row, column=col).value for col in range(1, 4)]
        if all(v is None for v in values):
            ws.delete_rows(row)
    if verify:
        merge_special_cells(ws)

    # 2. 提取查询结构信息
    need_query, keywords = analysis_keyword_list_from_ws(ws)
    is_special = get_merged_info(ws)
    unmerge_all_cells(ws)

    # 3. 核验模式下提取答案内容
    if verify:
        asr_list, asr_path_list = get_answers_from_xlsx(ws, need_query, is_special)
    else:
        asr_list, asr_path_list = (None,None)        

    # 4. 删除多余列
    max_column = ws.max_column
    if max_column >= 5:
        ws.delete_cols(6, max_column - 5)

    # 5. 设置所有单元格背景色为白色
    for row in ws.iter_rows():
        for cell in row:
            cell.fill = PatternFill(fill_type=None)

    # 6. 构建返回字典
    tbl_info = {
        "sheet_name": ws.title,
        "rows": ws.max_row,
        "need_query": need_query,
        "is_special": is_special,
        "keywords": keywords,
        "answers": asr_list,
        "answers_path": asr_path_list,
    }

    return tbl_info


def extract_index_column_from_keywords(xlsx_path: str):
    """
    函数作用：
        - 提取关键词前的层级编号作为索引（如“1.2.3 名词” → “索引: 1.2.3”，“关键词: 名词”）；
        - 在原关键词列前插入索引列；
        - 添加“索引”和“关键词”表头；

    参数:
        xlsx_path (str): 要处理的 Excel 文件路径
    """
    wb = load_workbook(filename=xlsx_path)
    ws = wb.active
    
    # 删除多余行
    for row in range(ws.max_row, 1, -1):  # 从最后一行开始
        cell_value = ws.cell(row=row, column=1).value  # 获取第一列的值
        if cell_value is None:  # 检查第一列的值是否为空
            ws.delete_rows(row)  # 删除该行

    # 获取原文档中关键词一列并分析索引
    keywords = [row[0].value for row in ws.iter_rows(min_row=2, max_row=ws.max_row)]
    levl_list = []
    word_list = []

    dflag = 1
    temp_parent = "0"
    for keyword in keywords:
        # 使用正则表达式匹配前面的数字部分
        match = re.match(r"^([\d\.]+)\s*(.*)", keyword)
        if match:
            levl = match.group(1)  # 提取数字部分
            text = match.group(2)  # 提取后面的词语部分
            dflag = 1
            temp_parent = levl
        else:
            levl = temp_parent + "." + str(dflag)
            text = keyword
            dflag += 1
        levl_list.append(levl)
        word_list.append(text.strip())

    # 备份合并区域 并 解除原合并
    original_merged_ranges = ws.merged_cells.ranges.copy()
    for merged_range in original_merged_ranges:
        ws.unmerge_cells(str(merged_range))

    # 插入第一列
    ws.insert_cols(1)

    # 重新计算合并区域 并 重新合并
    new_merged_ranges = []
    for merged_range in original_merged_ranges:
        min_row = merged_range.min_row
        min_col = merged_range.min_col + 1
        max_row = merged_range.max_row
        max_col = merged_range.max_col + 1
        new_range = (
            f"{get_column_letter(min_col)}{min_row}"
            f":{get_column_letter(max_col)}{max_row}"
        )
        new_merged_ranges.append(new_range)
    for new_range in new_merged_ranges:
        ws.merge_cells(new_range)

    # 填写表头
    ws["A1"] = "索 引"
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B1"] = "关键词"
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")

    # 填写索引
    for i in range(0, len(levl_list)):
        ws.cell(row=i + 2, column=1).value = levl_list[i]
        ws.cell(row=i + 2, column=2).value = word_list[i]
        ws.cell(row=i + 2, column=1).alignment = Alignment(horizontal="left", vertical="center")

    save_path = xlsx_path.split(".")[0] + "_new.xlsx"
    wb.save(save_path)


def reformat_index_column_style(input_path: str, output_path: str):
    """
    函数逻辑：
        - 将Excel表格中的首列编号按层级进行格式化（支持中文数字、编号样式），
        - 并保留原始单元格样式、合并单元格结构，输出为一个新的Excel文件。

    参数:
        input_path (str): 原始Excel文件路径
        output_path (str): 处理后的Excel文件保存路径
    """
    # 数字→中文映射
    def generate_chinese_numerals_upto_100():
        digits = {
            1: '一', 2: '二', 3: '三', 4: '四', 5: '五',
            6: '六', 7: '七', 8: '八', 9: '九', 10: '十'
            }
        chinese_numerals = {}
        for i in range(1, 101):
            if i == 10:
                chinese_numerals[i] = '十'
            elif i < 10:
                chinese_numerals[i] = digits[i]
            elif i < 20:
                chinese_numerals[i] = '十' + digits[i % 10]
            elif i % 10 == 0:
                chinese_numerals[i] = digits[i // 10] + '十'
            else:
                chinese_numerals[i] = digits[i // 10] + '十' + digits[i % 10]
        return chinese_numerals

    chinese_numerals = generate_chinese_numerals_upto_100()

    def map_index(value):
        if pd.isna(value):
            return value
        s = str(value).strip()
        parts = s.split('.')
        depth = len(parts) - 1
        try:
            idx = int(parts[-1])
        except ValueError:
            return ''
        if depth == 0:
            cn = chinese_numerals.get(idx, str(idx))
            return f"{cn}、"
        elif depth == 1:
            return f"{idx}."
        elif depth == 2:
            return f"（{idx}）"
        elif depth == 3:
            return f"{idx}）"
        else:
            return ''

    # 加载数据
    df = pd.read_excel(input_path)
    first_col = df.columns[0]
    df[first_col] = df[first_col].apply(map_index)

    # 读取旧工作簿和新建工作簿
    wb_old = load_workbook(input_path)
    ws_old = wb_old.active
    wb_new = Workbook()
    ws_new = wb_new.active

    # 复制表头及样式
    for j, col_name in enumerate(df.columns):
        cell_old = ws_old.cell(row=1, column=j + 1)
        cell_new = ws_new.cell(row=1, column=j + 1, value=col_name)

        cell_new.font = copy(cell_old.font)
        cell_new.fill = copy(cell_old.fill)
        cell_new.border = copy(cell_old.border)
        cell_new.alignment = copy(cell_old.alignment)
        cell_new.number_format = cell_old.number_format
        cell_new.protection = copy(cell_old.protection)

    # 写入数据，并复制样式
    for i, row in df.iterrows():
        for j, val in enumerate(row):
            row_idx = i + 2  # 数据从第二行开始
            col_idx = j + 1

            cell_old = ws_old.cell(row=row_idx, column=col_idx)
            cell_new = ws_new.cell(row=row_idx, column=col_idx, value=val)

            cell_new.font = copy(cell_old.font)
            cell_new.fill = copy(cell_old.fill)
            cell_new.border = copy(cell_old.border)
            cell_new.alignment = copy(cell_old.alignment)
            cell_new.number_format = cell_old.number_format
            cell_new.protection = copy(cell_old.protection)

    # 复制合并单元格
    for merged_range in ws_old.merged_cells.ranges:
        ws_new.merge_cells(str(merged_range))


    for col_idx in range(1, ws_new.max_column + 1):
        col_letter = get_column_letter(col_idx)
        if col_idx == 2:
            ws_new.column_dimensions[col_letter].width = 40
        else:
            ws_new.column_dimensions[col_letter].width = 20

    # 保存文件
    wb_new.save(output_path)
    print(f"处理完成，结果已保存到 {output_path}")


def adjust_cell_styles(ws: Worksheet, query_info: dict, verify=False):
    """
    动态调整 excel 表格的列宽与行高，并规范所有单元格的格式
    调整列宽列高时会考虑 query_info 中所包含的信息
    """
    # 设置单元格边框样式
    border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
        diagonal=None,
    )

    # 动态调整列宽，并设置单元格格式
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)  # 获取列字母
        for cell in col:
            # 设置单元格的边框样式
            cell.border = border
            # 设置单元格文字的排布方式
            original_alignment = cell.alignment
            cell.alignment = Alignment(
                vertical="center",  # 垂直居中
                horizontal=original_alignment.horizontal,  # 保留水平对齐
                shrink_to_fit=original_alignment.shrink_to_fit,  # 保留单元格内容自动缩放信息
                indent=original_alignment.indent,  # 保留缩进信息
                wrap_text=True,  # 设置自动换行
            )
            if col_letter in ["A", "B", "E"]:  # A列（索引列） B列（关键词列） E列（备注列） 动态调整列宽
                try:
                    if cell.value:
                        cell_value = str(cell.value).split("\n")
                        max_token = ""
                        for i in cell_value:
                            if len(i) > len(max_token):
                                max_token = i
                        # 中文字符宽度视为2，英文字符宽度视为1.3
                        adjusted_length = (
                            sum(
                                2.2 if "\u4e00" <= char <= "\u9fff" else 1.3
                                for char in max_token
                            )
                            + 4
                        )
                        max_length = max(max_length, adjusted_length)
                except:
                    pass
            elif col_letter in ["C", "D"]:  # C列（单位列） D列（数量列）固定列宽为15
                max_length = 15
            else:  # 在核验模式下，文段列固定列宽为50
                max_length = 50
        ws.column_dimensions[col_letter].width = max_length

    # 动态调整行高
    for row_id, row in enumerate(ws.iter_rows(min_row=2)):
        if not verify and query_info["need_query"][row_id]:
            max_height = 15
            for cell in row:
                col_letter = get_column_letter(cell.column)

                if col_letter in ["A", "B", "D"]:  # A列（索引列）B列（关键词列）D列（数量列）一般不会出现height大于15的情况
                    pass
                elif col_letter == "C":  # C列（单位列），假如发生了单元格合并，需要额外处理
                    if query_info["is_special"][row_id] and cell.value:
                        cell_height = math.ceil(len(str(cell.value)) / 14) * 15
                        max_height = max(cell_height, max_height)
                elif col_letter == "E":  # E列（备注列），输出内容较为固定，按照换行数调整行高
                    if cell.value:
                        max_height = max(
                            len(str(cell.value).split("\n")) * 15, max_height
                        )
            ws.row_dimensions[row[0].row].height = max_height
        elif verify:
            max_height = 15
            for cell in row:
                col_letter = get_column_letter(cell.column)

                if col_letter in ["A", "B", "D"]:  # A列（索引列）B列（关键词列）D列（数量列）一般不会出现height大于15的情况
                    pass
                elif col_letter == "E":  # E列（备注列），输出内容较为固定，按照换行数调整行高
                    if cell.value:
                        max_height = max(
                            len(str(cell.value).split("\n")) * 15, max_height
                        )
                else:  # 粘贴文段的列表稍复杂
                    if cell.value:
                        text_list = str(cell.value).split("\n")
                        cell_height = sum(
                            math.ceil(len(text) / 25) * 14 for text in text_list
                        )
                        max_height = max(cell_height, max_height)
            ws.row_dimensions[row[0].row].height = max_height

    merge_special_cells(ws)


def split_answer(asr: str):
    """
    分割一个带有“答案+证据+路径”的字符串，提取结构化结果。

    参数：
        asr: 原始答案字符串，格式为 "答案###证据###路径"（路径可选）

    返回：
        List[str]: 形如 ["答案", "单位", ..., "答案证据：xxx\n答案路径：xxx"]
    """
    parts = asr.split("###")

    # 初始化分块内容，避免 IndexError
    main_answer = parts[0].strip() if len(parts) > 0 else ""
    evidence = parts[1].strip() if len(parts) > 1 else ""
    path = parts[2].strip() if len(parts) > 2 else ""

    # 构造 answer_source
    if path and evidence:
        if path.startswith("|"):
            path = path[1:]
        answer_source = f"答案路径：{path}\n答案证据：{evidence}"
    elif not path and evidence:
        answer_source = f"答案证据：{evidence}"
    else:
        answer_source = ""

    # 常见分隔符，尝试分割答案
    split_chars = ["|", " ", "\t", "/", "\n", ",", "-"]

    # 复用原始分割逻辑
    for char in split_chars:
        if char in main_answer:
            tokens = [token.strip() for token in main_answer.split(char) if token.strip()]
            # 如果有提取到答案来源，追加到结果末尾
            if answer_source:
                tokens.append(answer_source)
            if tokens:
                return tokens

    # 处理没有分割符的情况
    result = []
    if main_answer:
        result.append(main_answer)
    if answer_source:
        result.append(answer_source)

    return result if result else [asr]


@error_handler("表格填写出错，请联系技术人员")
def fill_extract_answer_to_xlsx(
    asr_list: List[str],
    query_info: dict,
    tg_ws: Worksheet,
    # 数据提取功能相关参数
    extract_with_proof=False,
):
    """
    true_chunks: 通过正则表达式匹配，存在正确答案的被筛掉的文段列表
    topk_chunks: 向 AI 询问的是否与关键词相关的文档列表

    填写模式。根据查询表 query_info 的信息，将 asr_list 中所有答案填写到 tg_wb 的对应表的对应行中去；
    asr 格式如下： “关键词|单位|数量” 或 “关键词|描述短句” 或 “关键词|无”

    """
    if tg_ws.max_column < 5:
        tg_ws.insert_cols(5)
    tg_ws["E1"] = "备 注"
    tg_ws["E1"].alignment = Alignment(horizontal="center", vertical="center")

    p = 0
    # 读取标准单位格式的文件
    with open("./utils/unit.txt", "r", encoding="utf-8") as file:
        units = [line.strip() for line in file]

    # 填写答案完整流程
    for row_id, row in enumerate(tg_ws.iter_rows(min_row=2, min_col=3)):
        if query_info["need_query"][row_id] and p < len(asr_list):
            asr = split_answer(asr_list[p])

            kw = query_info["keywords"][row_id]

            # asr 异常处理
            if extract_with_proof:
                if query_info["is_special"][row_id]:
                    asr = asr if len(asr) == 3 else ["无", "无"]
                    kw = asr.pop(0) if len(asr) == 3 else kw
                else:
                    asr = asr if len(asr) == 4 else ["无", "无", "无"]
                    kw = asr.pop(0) if len(asr) == 4 else kw
            else:
                if query_info["is_special"][row_id]:
                    asr = asr if len(asr) == 2 else ["无"]
                    kw = asr.pop(0) if len(asr) == 2 else kw
                else:
                    asr = asr if len(asr) == 3 else ["无", "无"]
                    kw = asr.pop(0) if len(asr) == 3 else kw

            # 设置错误行背景色
            orange_fill = PatternFill(
                start_color="FFC000", end_color="FFC000", fill_type="solid"
            )

            # 填表模式下填写关键词的数量单位 或 描述 或 备注
            if asr[0] == "无":  # 当没有查到答案时就只有一个“无”字
                for cell in tg_ws[row_id + 2]:
                    cell.fill = orange_fill
                row[-1].value = f"未在原文中找到关键词“{kw}”的相关内容。"
            elif query_info["is_special"][
                    row_id
            ]:  
                # 填写描述性短句
                row[0].value = asr[0]
                # 填写备注
                if extract_with_proof and len(asr) == 2:
                    row[2].value = asr[1]
            else:  # 填写一般答案（单位 + 数量）
                # 填写单位
                text = []
                # 填写单位时修改其格式成标准单位格式
                tranlsate_unit_value = find_matching_unit(asr[0], units)
                if tranlsate_unit_value:  # 是否转换成功
                    asr[0] = tranlsate_unit_value
                tk_list = re.split(r"([a-zA-Z]+\d+)", asr[0])
                for tk in tk_list:
                    mt = re.match(r"([a-zA-Z]+)(\d+)", tk)
                    if mt is not None:
                        text.append(mt.group(1))
                        text.append(
                            TextBlock(
                                InlineFont(vertAlign="superscript"), mt.group(2)
                            )
                        )
                    else:
                        text.append(tk)
                row[0].value = CellRichText(text)
                # 填写数量
                row[1].value = asr[1]
                # 填写备注
                if extract_with_proof and len(asr) == 3:
                    row[2].value = asr[2]
            p += 1
        else:
            pass

    # 调整 xlsx 表格格式
    adjust_cell_styles(tg_ws, query_info)


@error_handler("表格填写出错，请联系技术人员")
def fill_check_answer_to_xlsx(
    asr_list: List[str],
    query_info: dict,
    tg_ws: Worksheet,
    # 数据核验功能相关参数
    topk=3,
    true_chunks=[],
    topk_chunks=[],
    proof_list=[],
    # verify_turn=1, 
    check_with_proof=True,
    docx_manager: Optional[DocxManager] = None,
):
    # 1. 检查是否传入 Docx 文档管理器
    if not isinstance(docx_manager, DocxManager):
        raise TypeError(f"{docx_manager} 不是一个 DocxManager 类")

    # 2. 定义在核验模式下将 chunk 填入指定单元格的方法
    def fill_chunk_in_cell(cell, structure_path, chunk, proof="参考答案来源"):
        parse_result = extract_table_info(chunk)
        if parse_result["is_table"]:
            cell.value = parse_result["tbl_info"]
        else:
            find_result = docx_manager.find_text_in_docx(structure_path, chunk)
            if find_result["not_found"]:
                tp_text = [
                    TextBlock(InlineFont(color="FF0000", b=True), "对以下文段添加超链接失败，请自行检查：\n"),
                    chunk
                ]
                cell.value = CellRichText(tp_text)
            else:
                bookmark_name = docx_manager.add_bookmark_by_pos(find_result["chunk_pos"], chunk)
                hyperlink = f"{docx_manager.save_file_name}.docx#{bookmark_name}"
                if check_with_proof:
                    cell.value = f"...{proof}...(点击跳转原文)"
                else:
                    cell.value = f"点击跳转到参考文段"
                cell.hyperlink = hyperlink
                cell.font = Font(color="0000FF", underline="single")  # 蓝色字体+下划线

    # 3. 添加 备注列 & 文段展示列
    if tg_ws.max_column < 5:
        tg_ws.insert_cols(5)
    tg_ws["E1"] = "备 注"
    tg_ws["E1"].alignment = Alignment(horizontal="center", vertical="center")

    tg_ws.insert_cols(6, amount=2)
    # tg_ws["F1"] = "参考文段"
    # tg_ws["G1"] = "冲突文段"

    tg_ws["F1"] = "冲突文段"
    tg_ws["G1"] = "参考文段"
    tg_ws["F1"].alignment = Alignment(horizontal="center", vertical="center")
    tg_ws["G1"].alignment = Alignment(horizontal="center", vertical="center")

        
    # 4. 根据 query_info 、将 asr_list 的答案填入 tg_ws 中
    asr_ls_ptr = 0
    qry_info_ptr = 0
    row_ptr = 2

    while row_ptr <= tg_ws.max_row:
        if query_info["need_query"][qry_info_ptr] and asr_ls_ptr < len(asr_list):
            asr = split_answer(asr_list[asr_ls_ptr])
            
            # 取得当前 关键词 以及 参考答案
            kw = query_info["keywords"][qry_info_ptr]
            ground_truth = query_info["answers"][qry_info_ptr]

            # asr 异常处理
            if len(asr) == topk + 1:
                asr.pop(0)
            else:
                asr = ["正确"]*topk
           
            # 设置错误答案所在单元格的格式
            orange_fill = PatternFill(
                start_color="FFC000", end_color="FFC000", fill_type="solid"
            )

            # 统计 asr 结果 
            bool_list = [True if i.strip() == "正确" else False for i in asr]
            is_topk_relative = True
            for value in bool_list:
                is_topk_relative = is_topk_relative and value
            is_topk_relative = not is_topk_relative

            # 情形一：过滤出文段，并且 topk_chunks 都与关键词不相关，认为该关键词在原文中只有一个答案且填写正确
            if len(true_chunks[asr_ls_ptr]) != 0 and not is_topk_relative:
                tg_ws.cell(row=row_ptr, column=5).value = f"在参考文段中找到关键词“{kw}”以及参考答案“{ground_truth}”的相关内容。"
                tg_ws.cell(row=row_ptr, column=7).value = f"参考答案“{ground_truth}”对应段落"

                # tg_ws.cell(row=row_ptr, column=7).value = f"下方展示包含参考答案“{ground_truth}”的文段"

                # 根据参考文段数量插入行数
                true_chunk_num = len(true_chunks[asr_ls_ptr])
                insert_row_nums = true_chunk_num
                tg_ws.insert_rows(idx=row_ptr + 1, amount=insert_row_nums)                        

                # 插入所有参考文段
                offset = 1
                for chunk in true_chunks[asr_ls_ptr]:
                    fill_chunk_in_cell(
                        tg_ws.cell(row=row_ptr + offset, column=7),
                        chunk.metadata["path"], 
                        chunk.page_content,
                    )
                    offset += 1
                    
                # row_ptr 移动
                row_ptr += insert_row_nums

            # 情形二：没过滤出任何文段，并且 topk_chunks 都与关键词不相关，认为该关键词在原文中没有相关文段        
            elif len(true_chunks[asr_ls_ptr]) == 0 and not is_topk_relative:
                # 抓取当前处理行
                row = list(tg_ws.rows)[row_ptr - 1]

                # 设置背景色
                for cell in row:
                    cell.fill = orange_fill

                # 填写备注信息
                tg_ws.cell(row=row_ptr, column=5).value = f"未在原文中找到关键词“{kw}”以及参考答案“{ground_truth}”的相关内容。"

            # 情形三：过滤出文段，但是 topk_chunks 中部分文段与关键词相关，认为该关键词的答案在原文中存在冲突
            elif len(true_chunks[asr_ls_ptr]) != 0 and is_topk_relative:
                # 抓取当前处理行
                row = list(tg_ws.rows)[row_ptr - 1]
                
                # 设置背景色
                for cell in row:
                    cell.fill = orange_fill

                # 填写备注信息
                tg_ws.cell(row=row_ptr, column=5).value = (
                    f"关键词“{kw}”在原文中存在不止一处可能答案。\n"
                    f"右侧列出了包含参考答案“{ground_truth}”的文段以及与参考答案冲突的文段。"
                )
                tg_ws.cell(row=row_ptr, column=7).value = f"参考答案“{ground_truth}”对应段落"
                tg_ws.cell(row=row_ptr, column=6).value = f"冲突文段对应段落"

                # tg_ws.cell(row=row_ptr, column=6).value = f"下方展示包含参考答案“{ground_truth}”的文段"
                # tg_ws.cell(row=row_ptr, column=7).value = f"下方展示冲突文段"

                # 根据参考文段以及冲突文段数量插入行数
                conflict_chunks_num = sum(not i for i in bool_list)
                true_chunks_num = len(true_chunks[asr_ls_ptr])
                insert_row_nums = conflict_chunks_num if conflict_chunks_num > true_chunks_num else true_chunks_num

                tg_ws.insert_rows(idx=row_ptr + 1, amount=insert_row_nums)                        

                # 插入所有参考文段
                offset = 1
                for chunk in true_chunks[asr_ls_ptr]:
                    fill_chunk_in_cell(
                        tg_ws.cell(row=row_ptr + offset, column=7),
                        chunk.metadata["path"], 
                        chunk.page_content,
                    )
                    offset += 1

                # 插入所有冲突文段
                offset = 1
                for idx, i in enumerate(bool_list):
                    if not i:
                        # 添加指向文段的超链接
                        fill_chunk_in_cell(
                            tg_ws.cell(row=row_ptr + offset, column=6), 
                            topk_chunks[asr_ls_ptr][idx].metadata["path"],
                            topk_chunks[asr_ls_ptr][idx].page_content, 
                            proof_list[asr_ls_ptr][idx] if check_with_proof else None
                        )
                        offset += 1

                # row_ptr 移动
                row_ptr += insert_row_nums

            # 情形四：没过滤出任何文段，并且 topk_chunks 中部分文段与关键词相关，认为该关键词的答案在原文中存在冲突
            elif len(true_chunks[asr_ls_ptr]) == 0 and is_topk_relative:
                # 抓取当前处理行
                row = list(tg_ws.rows)[row_ptr - 1]
                
                # 设置背景色
                for cell in row:
                    cell.fill = orange_fill

                # 填写备注信息  
                tg_ws.cell(row=row_ptr, column=5).value = (
                    f"在原文中找到的可能答案与参考答案“{ground_truth}”皆不相同。\n"
                    f"右侧列出了与参考答案相冲突的文段。"
                )
                tg_ws.cell(row=row_ptr, column=6).value = f"冲突答案对应段落"

                # 根据冲突文段数量插入行数
                conflict_chunks_num = sum(not i for i in bool_list)
                insert_row_nums = conflict_chunks_num
                tg_ws.insert_rows(idx=row_ptr + 1, amount=insert_row_nums)                        

                # 插入所有冲突文段
                offset = 1
                for idx, i in enumerate(bool_list):
                    if not i:
                        # 添加指向文段的超链接
                        fill_chunk_in_cell(
                            tg_ws.cell(row=row_ptr + offset, column=6), 
                            topk_chunks[asr_ls_ptr][idx].metadata["path"],
                            topk_chunks[asr_ls_ptr][idx].page_content, 
                            proof_list[asr_ls_ptr][idx] if check_with_proof else None
                        )
                        offset += 1
                
                # row_ptr 移动
                row_ptr += insert_row_nums

            # 当前答案填写完成，移动 asr_ls_ptr
            asr_ls_ptr += 1
        else:
            pass
        
        # 当前行处理完成：移动 row_ptr 以及 qry_info_ptr
        row_ptr += 1
        qry_info_ptr += 1

    # 调整 xlsx 表格格式
    adjust_cell_styles(tg_ws, query_info, verify=True)
    
    # 导出添加标签后的 docx 文档
    docx_manager.save_docx_file()


def replace_hyperlink_timestamp(wb:Workbook, docx_manager: DocxManager):
    """
    补丁：替换超链接中的时间戳
    """
    new_timestamp = docx_manager.timestamp
    timestamp_pattern = re.compile(r'\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}')
    
    # 遍历所有工作表
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                # 检查单元格是否有超链接
                if cell.hyperlink:
                    original_target = cell.hyperlink.target
                    # 替换时间戳
                    new_target = timestamp_pattern.sub(new_timestamp, original_target)
                    if new_target != original_target:
                        cell.hyperlink.target = new_target


def unmerge_all_cells(ws):
    """
    解除 worksheet 中的所有合并单元格
    """
    # 注意：要先复制列表，否则边遍历边修改会出错
    for merge_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merge_range))

def merge_special_cells(ws):
    """
    检查 ws 中的 C、D 列：
    如果某一行 C 列不为空 且 D 列为空，
    则合并该行的 C、D 两个单元格
    """
    for row in range(1, ws.max_row + 1):  # 遍历所有行
        c_cell = ws.cell(row=row, column=3)  # C列
        d_cell = ws.cell(row=row, column=4)  # D列

        if c_cell.value is not None and d_cell.value is None:
            ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)


def flatten_cell_value(value):
    """
    将 Excel 单元格的内容统一转换为字符串。

    功能：
        - 如果值是 `CellRichText`（可能包含带格式的文本块），则提取并拼接所有文本块内容；
        - 如果值是普通字符串，直接返回；
        - 其他类型（如数字、日期等），转换为字符串；
        - 如果值为 None，返回空字符串。

    """
    if isinstance(value, CellRichText):
        return ''.join(
            block.text if isinstance(block, TextBlock) else str(block)
            for block in value
        )
    elif isinstance(value, str):
        return value
    else:
        return str(value) if value is not None else ''


if __name__ == "__main__":
    xlsx_path = "/home/ljc/project/Demo2.9/data/test_data/芦山抽水蓄能电站/所有样本_提取_new_新标号.xlsx"

    wb = read_xlsx_file(xlsx_path=xlsx_path)
    ws, _ = read_xlsx_sheet(wb=wb)

    init_query_info_from_xlsx(ws)
    print()