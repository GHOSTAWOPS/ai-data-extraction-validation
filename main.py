from data_processor_new_zbb_structured import zbb_structure_parser_process_docx
from utils.utils import (
    pdf_to_pages,
    text_table_to_chunks,
    timing_decorator,
    extract_mid_process,
    check_mid_process,
    zbb_text_table_to_chunks,
)

from utils.xlsx_utils import *
from utils.docx_utils import *
from utils.utils import extract_table_chunks_to_json

from RAGAS.ragas_eval_correct import init_ragas

from models.extract_model import call_extract_model
from models.vectorstore import random_collection_name, add_doc_to_kb
from models.check_model import call_check_model
from models.new_check_model import call_check_new_model
from config import CONFIG
import time
import gradio as gr
from data_processor import unstructured_process_docx
from datetime import datetime
import os
from utils.file_cleanup import continuous_cleanup
import threading
from utils.logger_util import logger
from datetime import timedelta
from langchain.schema import Document
import shutil
import asyncio
import chromadb
from docx import Document
from utils.MergeDoc import merge_docx_in_folder

def check_files(doc_file, excel_file):
    # 加入 None 判断
    if doc_file and excel_file:
        return gr.update(interactive=True)
    else:
        return gr.update(interactive=False)


def check_files_pre(doc_file, state_files=None):
    """决定预处理按钮是否可点。
    优先依据 state_files（累积的已上传路径）判断；若无 state 信息，再回退到单次 input 的 doc_file 判断。
    这样可以避免在 add_file_and_clear 清空上传控件（返回 None）后被其他 change 事件误判为无文件。
    """
    # state_files 可能是 gr.State([]) 传来的列表
    if state_files:
        try:
            if isinstance(state_files, (list, tuple)) and len(state_files) > 0:
                return gr.update(interactive=True)
        except Exception:
            pass

    # 退化判断：如果上传控件当前有文件（未被清空），也允许启用
    if doc_file is not None:
        return gr.update(interactive=True)
    return gr.update(interactive=False)

def get_all_templates(template_level=None):
    if not template_level:
        return []

    # 决定路径
    level_path = (
        CONFIG.template_user_path if template_level == "客户级"
        else CONFIG.template_system_path
    )

    if not os.path.exists(level_path) or not os.path.isdir(level_path):
        return []

    # 显示为：客户级:test.xlsx 或 系统级:xxx.xlsx
    return [
        f"{template_level}:{f}"
        for f in os.listdir(level_path)
        if os.path.isfile(os.path.join(level_path, f))
    ]


def upload_template(excel_file, template_level):
    if not excel_file:
        return None, "❌ 请上传模板文件", gr.update()
    if not template_level:
        return None, "❌ 请先选择模板级别（客户级 / 系统级）", gr.update()

    # ✅ 构造存储目录
    target_dir = CONFIG.template_user_path if template_level == "客户级" else CONFIG.template_system_path
    os.makedirs(target_dir, exist_ok=True)

    # ✅ 生成目标文件路径
    filename = os.path.basename(excel_file)
    dest_path = os.path.join(target_dir, filename)

    # ✅ 检查是否已存在
    if os.path.exists(dest_path):
        return (
            None,
            f"⚠️ 模板文件已存在: {filename}",
            gr.update(choices=get_all_templates(template_level)),
        )

    # ✅ 模板格式检查
    try:
        wb = read_xlsx_file(excel_file)
        ws, _ = read_xlsx_sheet(wb, verify=False)
        _ = init_query_info_from_xlsx(ws)
    except Exception as e:
        return (
            None,
            f"❌ 模板格式错误：{str(e)}",
            gr.update(choices=get_all_templates(template_level)),
        )

    # ✅ 执行保存
    shutil.copy(excel_file, dest_path)
    print(f"✅ 模板已保存至：{dest_path}")

    return (
        None,
        f"✅ 模板上传成功: {filename}",
        gr.update(choices=get_all_templates(template_level)),
    )


def delete_template(template_name, template_level):
    # ✅ 参数检查
    if not template_level:
        return (
            "❌ 模板级别未指定（客户级 / 系统级）",
            None,
            gr.update(choices=[]),
        )
    if not template_name:
        return (
            "❌ 未选择要删除的模板",
            None,
            gr.update(choices=get_all_templates(template_level)),
        )

    # ✅ 构造目标路径
    target_dir = CONFIG.template_user_path if template_level == "客户级" else CONFIG.template_system_path
    file_path = os.path.join(target_dir, template_name)

    # ✅ 检查文件是否存在
    if not os.path.exists(file_path):
        return (
            f"⚠️ 模板文件不存在: {template_name}",
            None,
            gr.update(choices=get_all_templates(template_level)),
        )

    # ✅ 删除文件
    try:
        os.remove(file_path)
        print(f"🗑️ 删除成功: {file_path}")
        return (
            f"✅ 模板已删除: {template_name}",
            None,
            gr.update(choices=get_all_templates(template_level)),
             gr.update(value=None),  # ✅ 清空预览 DataFrame
        )
    except Exception as e:
        return (
            f"✅ 模板已删除: {template_name}",
            None,  # 清空 File
            gr.update(choices=get_all_templates(template_level)),
            gr.update(value=None),  # 清空 DataFrame
        )


def dropdown_change(display_name, template_level):
    if not display_name or not template_level:
        return None, None

    # 从 display_name 中拆出文件名（如 '客户级:test.xlsx' -> 'test.xlsx'）
    if ":" in display_name:
        _, template_name = display_name.split(":", 1)
    else:
        template_name = display_name

    base_dir = CONFIG.template_user_path if template_level == "客户级" else CONFIG.template_system_path
    file_path = os.path.join(base_dir, template_name)

    try:
        wb = read_xlsx_file(file_path)
        ws, _ = read_xlsx_sheet(wb, verify=False)

        ans = []
        for row in ws.iter_rows(values_only=True):
            if len(row) >= 2:
                ans.append([row[0], row[1]])
            elif len(row) == 1:
                ans.append([row[0], None])
        df = pd.DataFrame(ans, columns=["索引", "关键词"])

        return df, file_path
    except Exception as e:
        return f"❌ 读取失败：{str(e)}", None


def extract_dropdown_change(display_name, template_level):
    if not display_name or not template_level:
        return None

    if ":" in display_name:
        _, template_name = display_name.split(":", 1)
    else:
        template_name = display_name

    template_dir = (
        CONFIG.template_user_path if template_level == "客户级"
        else CONFIG.template_system_path
    )

    return os.path.join(template_dir, template_name)


def refresh_dropdown():
    return gr.update(choices=get_all_templates())


def load_markdown_file(page):
    if page == "extract":
        with open(
            "project_instructions/markdown/extract.md", "r", encoding="utf-8"
        ) as f:
            return f.read()
    elif page == "verify":
        with open(
            "project_instructions/markdown/verify.md", "r", encoding="utf-8"
        ) as f:
            return f.read()
    elif page == "preprocess":
        with open(
            "project_instructions/markdown/preprocess.md", "r", encoding="utf-8"
        ) as f:
            return f.read()
    else:
        pass


def back_to_home():
    return (
        gr.update(visible=True),
        gr.update(visible=False),
        gr.update(visible=False),
        gr.update(visible=False),
        gr.update(visible=False),
        gr.update(visible=False),
        gr.update(visible=False),
    )


def create_tab_switch_handler(visible_index):
    def handler():
        updates = [gr.update(visible=False)] * 5   # 创建7个False的更新
        updates[visible_index] = gr.update(visible=True)  # 设置指定索引为True
        return tuple(updates)

    return handler


# docx文件预处理
def preprocess_file(
    doc_file_path,
    progress=None,
):
    try:
        yield "文档解析中"
        # 提取输入文件文件名
        docx_file_name = os.path.splitext(os.path.basename(doc_file_path))[0]
        name_word = docx_file_name.split("_")[0]
        logger.info(f"文件名：{name_word}")

        # unstructured解析docx文档
        start_time = time.time()  # 记录开始时间
        logger.info("unstructured解析docx文档")
        yield "正在加载文档中表格信息……"
        (elements_docx, elements_text, elements_table) = (
            zbb_structure_parser_process_docx(doc_file_path, json_file_path=None)
        )

        yield "处理完成，系统已加载该文档表格信息"

    except Exception as e:
        logger.error(f"处理过程中出现错误：{str(e)}", exc_info=True)
        # 返回具体错误信息供前端展示
        yield f"处理失败：{str(e)}"


# 工程数据提取async
def process_extract(
    doc_file_path=None,
    excel_file_path=None,
    gt_excel_file_path=None,
    progress=None,
):  # 主函数
    """
    Demo1.0:
        logger.info("pdf文档解析切块")
        pages = pdf_to_pages(doc_file_path)
    """
    try:
        yield None, "文档解析中", []
        start_time = time.time()  # 记录开始时间

        # # 通过 thread_id 以及 task_id 检测并发是否有效
        # thread_id = threading.get_ident()
        # task_id = f"ID-{hash(start_time)}"

        # # 打印当前进程id
        # print(f"进程 PID: {os.getpid()}, 正在运行...")

        ## gt值输入的区别：原gt需要自己输入到getgroundtruth里面
        # 现在输入文档，直接用原文档的底色判断正、负、零样本（标准文档格式）
        # ===========gt输入方式的选择(用一种取消一种)========#
        # print(f"Thread-{thread_id} | Task {task_id} STARTED | 提取文档的gt值 开始")

        ## 原文档gt输入方式（只有芦山报告的测试）
        # gt = getGroundTruth(doc_file_path)

        ## 现文档输入方式（多个报告的测试，包含芦山报告），
        # 需要在main函数输入gt_excel_file_path（test_data下的每个文件夹下对应一个所有样本.xlsx）
        # gt值获取方式，看GroundTruth代码read_xlsx_and_convert函数
        gt = []
        logger.info("提取文档的gt值")
        if gt_excel_file_path is not None:
            gt = read_xlsx_and_convert(gt_excel_file_path)

        # print(f"Thread-{thread_id} | Task {task_id} STARTED | 提取文档的gt值 结束")
        # ====================end======================#

        ## 提取输入文件文件名，水电站的名字（rerank的query使用）
        docx_file_name = os.path.splitext(os.path.basename(doc_file_path))[0]
        name_word = docx_file_name.split("_")[0]
        logger.info(f"文件名：{name_word}")

        ## 寻找docx文件对应的json文件
        json_files = []
        for filename in os.listdir(CONFIG.json_output_folder_path):
            file_path = os.path.join(CONFIG.json_output_folder_path, filename)
            if os.path.isfile(file_path) and filename.startswith(name_word):
                json_files.append(file_path)

        json_file = None if len(json_files) == 0 else json_files[0]
        if json_file is not None:
            logger.info(f"找到对应的json文件：{json_file}")
            yield None, f"找到 {name_word}.docx 所对应的表格信息", []
        else:
            logger.info(f"没有找到对应的json文件")
            yield None, f"系统首次处理 {name_word}.docx ，请耐心等待文件预处理", []

        # ==============================unstructured文档分析器===========================#
        # print(f"Thread-{thread_id} | Task {task_id} STARTED | 解析Docx文档 开始")

        logger.info("unstructured解析docx文档")
        (elements_docx, elements_text, elements_table) = (
            zbb_structure_parser_process_docx(doc_file_path, json_file_path=json_file)
        )
        yield None, "正在生成向量数据库……", []
        logger.info("将文本切分为chunk")
        chunks = zbb_text_table_to_chunks(elements_text, elements_table)
        logger.info(f"文档被分割成 {len(chunks)} 块")
        logger.info("|".join(str(len(chunk.page_content)) for chunk in chunks))

        # print(f"Thread-{thread_id} | Task {task_id} STARTED | 解析Docx文档 结束")
        # =================================end==================================#

        # ==============================Demo1.0文档分析器================================#
        ## 记得修改extract_model.py的context赋值块，选择没有转成markdown的那一块
        # logger.info("pdf文档解析切块")
        # chunks = pdf_to_pages(doc_file_path)
        # elements_docx=None
        # =================================end==================================#

        # 查看整个文档切块情况#
        file_name_without_extension = os.path.splitext(os.path.basename(doc_file_path))[
            0
        ]
        output_file = f"./output_excels/chunk_txt/{file_name_without_extension}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
        # 将 context 写入文件,将检索到的上下文写入context_output.txt
        with open(output_file, "w", encoding="utf-8") as f:
            for i, doc in enumerate(chunks):
                f.write(
                    f"第{i}块：{doc.page_content}" + "\n"
                )  # 换行区分不同qurey写入上下文
        logger.info(f" 全文档切割情况：{output_file}")

        # print(f"Thread-{thread_id} | Task {task_id} STARTED | chunk嵌入向量数据库 开始")
        logger.info("将chunck送入embedding")
        random_id = random_collection_name()
        retriever = add_doc_to_kb(chunks, random_id)

        logger.info("从excel中提取表格信息，包括keywords")
        wb = read_xlsx_file(excel_file_path)
        ws, _ = read_xlsx_sheet(wb, verify=False)
        tbl_info = init_query_info_from_xlsx(ws)
        # print(f"Thread-{thread_id} | Task {task_id} STARTED | chunk嵌入向量数据库 结束")

        # ============================提取结果评估器=============================#
        kwargs = {}
        if CONFIG.RAGAS_SWITCH:
            # 如果开关打开，启动提取评估器
            API_KEY = os.environ["QWEN_API_KEY"]
            evaluator = init_ragas(api_key=API_KEY)
            kwargs["evaluator"] = evaluator
            logger.info("启动提取评估器")

            # 如果开关打开，生成中间结果存储文件
            folder_path = "output_jsons/extract_asr_jsons"
            os.makedirs(folder_path, exist_ok=True)

            timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            filename = f"data_{timestamp}.json"
            asr_file_path = os.path.join(folder_path, filename)
    
            # 创建一个空的 JSON 文件，作为后续写入中间数据的容器
            with open(asr_file_path, 'w', encoding='utf-8') as f:
                json.dump({}, f, ensure_ascii=False, indent=4)
            
            logger.info(f"中间结果将保存到: {file_path}")
        # =================================end==================================#

        yield None, "大模型问答启动中", []
        # print(f"Thread-{thread_id} | Task {task_id} STARTED | 查询大模型 开始")
        logger.info("调用大模型，生成答案")
        ans_list = []
        asr_info_list = []
        middle_output = []
        # zhy 8/21
        preprocess_time = time.time()
        with open("zhy_test/pre_timing_results.jsonl", "a", encoding="utf-8") as f:
            record = {"abscissa": 'preprocess', "time": preprocess_time - start_time}
            f.write(json.dumps(record, ensure_ascii=False) + "\n")

        for ans_true, asr_info in call_extract_model(
            tbl_info, retriever, chunks, name_word, elements_docx, progress, gt=gt, **kwargs
        ):
            ans_list.append(ans_true)
            asr_info_list.append(asr_info)
            middle_output = extract_mid_process(ans_list)
            yield None, "正在处理，请稍后", middle_output
        # print(f"Thread-{thread_id} | Task {task_id} STARTED | 查询大模型 结束")

        # print(f"Thread-{thread_id} | Task {task_id} STARTED | 输出文档生成 开始")
        logger.info("将答案填入表格，生成处理后的excel文件")
        # tg_wb = fill_answer_to_xlsx(
        #     ans_list,
        #     tbl_info,
        #     wb,
        #     extract_with_proof=CONFIG.EXTRACT_MODEL_SOURCE_SWITCH,
        # )

        if CONFIG.RAGAS_SWITCH:
            with open(asr_file_path, 'w', encoding='utf-8') as f:
                json.dump(asr_info_list, f, ensure_ascii=False, indent=4)

        fill_extract_answer_to_xlsx(
            ans_list,
            tbl_info,
            ws,
            extract_with_proof=CONFIG.EXTRACT_MODEL_SOURCE_SWITCH,
        )

        logger.info("保存处理后的excel文件")
        excel_output_path = (
            CONFIG.extract_output_folder_path
            + f"/output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        wb.save(excel_output_path)
        # print(f"Thread-{thread_id} | Task {task_id} STARTED | 输出文档生成 结束")

        execution_time = time.time() - start_time

        execution_time_str = str(
            timedelta(seconds=execution_time)
        )  # 格式化为timedelta对象
        logger.info(f"提取函数运行时间: {execution_time_str}")
        logger.info(f"向量数据库清理缓存：")
        chromadb.api.client.SharedSystemClient.clear_system_cache()
        yield excel_output_path, "处理完成，请在右侧点击文件下载", middle_output

    except Exception as e:
        logger.error(f"处理过程中出现错误：{str(e)}", exc_info=True)
        # 返回具体错误信息供前端展示
        yield None, f"处理失败：{str(e)}", []


# 工程数据复核
def process_check(
    doc_file_path, excel_file_path, gt_excel_file_path=None, progress=None, json_file=None
):
    """
    logger.info("pdf文档解析切块")
    chunks = pdf_to_pages(doc_file_path)
    """
    try:
        yield None, "文档解析中", []
        start_time = time.time()  # 记录开始时间
        # gt = getGroundTruth(doc_file_path)
        gt = []
        logger.info("提取文档的gt值")
        if gt_excel_file_path is not None:
            gt = read_xlsx_and_convert_check_model(gt_excel_file_path)
        logger.info("从excel中提取表格信息，包括keywords")

        # !!! 彬彬交给你了复核次数（time_check）
        wb = read_xlsx_file(excel_file_path)
        ws, time_check = read_xlsx_sheet(wb, verify=True)
        tbl_info = init_query_info_from_xlsx(ws, verify=True)

        logger.info(f"第{time_check+1}次复核")
        ## 提取输入文件文件名，水电站的名字（rerank的query使用）
        docx_file_name = os.path.splitext(os.path.basename(doc_file_path))[0]
        name_word = docx_file_name.split("_")[0]
        logger.info(f"文件名：{name_word}")

        ## 寻找docx文件对应的json文件
        json_files = []
        for filename in os.listdir(CONFIG.json_output_folder_path):
            file_path = os.path.join(CONFIG.json_output_folder_path, filename)
            if os.path.isfile(file_path) and filename.startswith(name_word):
                json_files.append(file_path)

        json_file = None if len(json_files) == 0 else json_files[0]
        if json_file is not None:
            logger.info(f"找到对应的json文件：{json_file}")
            yield None, f"找到 {name_word}.docx 所对应的表格信息", []
        else:
            logger.info(f"没有找到对应的json文件")
            yield None, f"系统首次处理 {name_word}.docx ，请耐心等待文件预处理", []

        # logger.info("unstructured解析docx文档")
        # (
        #     elements_docx,
        #     elements_text,
        #     elements_table_only_text,
        #     elements_table_idx,
        #     elements_table_summary,
        #     markdown_table_all,
        # ) = unstructured_process_docx(doc_file_path, json_file_path=json)

        # yield None, "正在生成向量数据库……", []
        # logger.info("将文本切分为chunk")
        # chunks = text_table_to_chunks(
        #     elements_text,
        #     elements_table_only_text,
        #     elements_table_idx,
        #     elements_table_summary,
        #     markdown_table_all,
        # )
        logger.info("zbb_structured解析docx文档")
        (elements_docx, elements_text, elements_table) = (
            zbb_structure_parser_process_docx(doc_file_path, json_file_path=json_file)
        )
        yield None, "正在生成向量数据库……", []
        logger.info("将文本切分为chunk")
        chunks = zbb_text_table_to_chunks(elements_text, elements_table)
        logger.info(f"文档被分割成 {len(chunks)} 块")
        logger.info("|".join(str(len(chunk.page_content)) for chunk in chunks))

        logger.info("将chunck送入embedding")
        logger.info(f"embedding模型路径：{CONFIG.EMBED_MODEL}")
        random_id = random_collection_name()
        retriever = add_doc_to_kb(chunks, random_id)

        output_dir_path = (
            CONFIG.check_output_folder_path
            + f"/output_{datetime.now().strftime('%Y%m%d_%H%M%S')}_folder"
        )
        excel_output_path = os.path.join(output_dir_path, "excel_output.xlsx")

        logger.info(f"文件将会保存到以下目录：{output_dir_path}")
        # 创建文件夹（如果不存在）
        os.makedirs(output_dir_path, exist_ok=True)

        logger.info(f"生成docx文档管理器")
        docx_manager = DocxManager(file_path=doc_file_path, save_path=output_dir_path)

        yield None, "大模型问答启动中", []
        logger.info("调用大模型，生成答案")

        ans_list = []
        middle_output = []
        ref_ans_list = []
        for (
            ans,
            ref_ans,
            true_docs_list,
            topk_docs_list,
            all_proof_list,
        ) in call_check_new_model(
            tbl_info,
            retriever,
            chunks,
            name_word,
            elements_docx,
            progress,
            gt=gt,
            time_check=time_check,
        ):
            ans_list.append(ans)
            ref_ans_list.append(ref_ans)
            middle_output = check_mid_process(
                ans_list,
                ref_ans_list,
                true_docs_list,
                topk_docs_list,
                topk=CONFIG.CHECK_RERANK_TOP_K,
            )
            yield None, "大模型正在处理，请稍等", middle_output

        logger.info(f"生成的答案：")
        logger.info("\n".join(ans_list))

        logger.info("将答案填入表格，生成处理后的excel文件")
        fill_check_answer_to_xlsx(
            ans_list,
            tbl_info,
            ws,
            topk=CONFIG.CHECK_RERANK_TOP_K,
            true_chunks=true_docs_list,
            topk_chunks=topk_docs_list,
            proof_list=all_proof_list,
            check_with_proof=True,
            docx_manager=docx_manager,
        )
        replace_hyperlink_timestamp(wb, docx_manager)

        logger.info("保存处理后的excel文件")
        # excel_output_path = (
        #     CONFIG.check_output_folder_path
        #     + f"/output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        # )
        wb.save(excel_output_path)

        execution_time = time.time() - start_time  # 计算时间差（秒）
        execution_time_str = str(
            timedelta(seconds=execution_time)
        )  # 格式化为timedelta对象
        logger.info(f"校验函数运行时间: {execution_time_str}")

        zip_path = output_dir_path + ".zip"
        # 压缩文件夹
        shutil.make_archive(output_dir_path, "zip", output_dir_path)
        logger.info(f"向量数据库清理缓存：")
        chromadb.api.client.SharedSystemClient.clear_system_cache()
        yield zip_path, "处理完成，请在右侧点击文件下载", middle_output

    except Exception as e:
        logger.error(f"处理过程中出现错误：{str(e)}", exc_info=True)
        # 返回具体错误信息供前端展示
        yield None, f"处理失败：{str(e)}", []


# ===== 以下均为新增工具函数 ========================
# 解析 "客户级:foo.xlsx" => ("客户级", "foo.xlsx")
def _split_item_text(txt: str, fallback_level: str) -> tuple[str, str]:
    if not txt:
        return fallback_level, ""
    if ":" in txt:
        lv, name = txt.split(":", 1)
        return lv, name
    # 没有前缀时，退化为用右边radio的级别
    return fallback_level, txt


# 取UploadButton传来的临时文件路径（不同gradio版本对象结构不同）
def _extract_temp_path(file_obj):
    if file_obj is None:
        return None
    if isinstance(file_obj, (list, tuple)):
        file_obj = file_obj[0]
    if hasattr(file_obj, "name"):
        return file_obj.name
    if hasattr(file_obj, "path"):
        return file_obj.path
    if isinstance(file_obj, dict):
        return file_obj.get("name") or file_obj.get("path")
    return str(file_obj)


def _level_dir(level: str) -> str:
    return CONFIG.template_user_path if level == "客户级" else CONFIG.template_system_path


def preview_only(selected_name: str, level: str):
    """
    调用你原有的 dropdown_change，但只返回“可被 Dataframe 接收”的结果。
    - dropdown_change 可能返回 (df, path) / str / None
    - 这里统一返回 pandas.DataFrame
    """
    try:
        out = dropdown_change(selected_name, level)
    except Exception as e:
        return pd.DataFrame([{"提示": f"读取失败：{e}"}])

    # 兼容 (df, path) 二元组
    if isinstance(out, tuple):
        df = out[0]
    else:
        df = out

    # 统一成 DataFrame
    if df is None:
        return pd.DataFrame(columns=["索引", "关键词"])
    if isinstance(df, str):
        # 把错误/提示信息也放进表格，避免报错
        return pd.DataFrame([{"提示": df}])
    if not isinstance(df, pd.DataFrame):
        # 万一是 list/dict，尽量转一下；转不动就给空表
        try:
            return pd.DataFrame(df)
        except Exception:
            return pd.DataFrame(columns=["索引", "关键词"])
    return df


# ---------- 列表&表格刷新 ----------
# 表格刷新：返回“模板名/更新时间”的二维数组 + meta列表（给select用）+ 清空选中
def refresh_table(level: str, keyword: str):
    items = get_all_templates(level)  # e.g. ["客户级:a.xlsx", "客户级:b.xlsx"]
    rows, meta = [], []
    keyword = (keyword or "").strip()

    for it in items:
        lv, name = _split_item_text(it, level)
        if keyword and keyword not in name:
            continue
        path = os.path.join(_level_dir(lv), name)
        try:
            ts = os.path.getmtime(path)
            mtime = time.strftime("%Y-%m-%d", time.localtime(ts))
        except Exception:
            mtime = "-"
        rows.append([name, mtime])           # 表格展示只放“模板名 / 时间”
        meta.append({"level": lv, "name": name})  # 另存一份结构化数据，供选中行解析

    # 返回：DataFrame值、meta列表（放到一个State里）、清空当前选中
    return rows, meta, ""


# Dataframe选中行 => 得到 "客户级:xxx.xlsx" 这样的选择值
def on_table_select(evt: gr.SelectData, meta: list[dict]):
    if not meta:
        return ""
    # evt.index 可能是 (row, col) 或者 row
    idx = evt.index[0] if isinstance(evt.index, (list, tuple)) else evt.index
    if idx is None or idx < 0 or idx >= len(meta):
        return ""
    m = meta[idx]
    return f"{m['level']}:{m['name']}"


# 查看（复用你现有 dropdown_change，只取预览DF即可）
def view_preview_only(selected: str, level_radio: str):
    lv, name = _split_item_text(selected, level_radio)
    if not name:
        return gr.update()  # 不动
    # 你的旧逻辑要求传入 "客户级:xxx.xlsx" + 级别
    df, _ = dropdown_change(f"{lv}:{name}", lv)
    return df


# 下载（直接把文件路径赋给 DownloadButton）
def prepare_download_template(selected: str, level_radio: str):
    lv, name = _split_item_text(selected, level_radio)
    if not name:
        return gr.update(value=None)
    path = os.path.join(_level_dir(lv), name)
    if not os.path.exists(path):
        return gr.update(value=None)
    return path  # DownloadButton会用这个路径作为下载目标


# 删除后刷新
def delete_and_refresh(selected: str, level: str, keyword: str):
    lv, name = _split_item_text(selected, level)
    if name:
        # 你的删除函数签名是 delete_template(template_dropdown, template_level)
        # 其中第一个参数之前就是 "客户级:文件名.xlsx"
        _msg, _x, _y = delete_template(f"{lv}:{name}", lv)
    rows, meta, _ = refresh_table(level, keyword)
    return rows, meta, ""   # 表格、meta、清空选中


# 上传按钮（UploadButton）接收文件对象 -> 拷贝到模板目录 -> 刷新
def save_uploaded_template(file_obj, level: str, keyword: str):
    src = _extract_temp_path(file_obj)
    if not src or not os.path.exists(src):
        rows, meta, _ = refresh_table(level, keyword)
        return rows, meta, ""
    os.makedirs(_level_dir(level), exist_ok=True)
    fname = os.path.basename(src)
    dest = os.path.join(_level_dir(level), fname)
    shutil.copy2(src, dest)  # 如需去重/校验，在此加逻辑

    rows, meta, _ = refresh_table(level, keyword)
    return rows, meta, ""  # 刷新表格 + meta + 清空选中


# ---------- 进入页面时初始化（客户级 & 清空状态） ----------
def _open_template_page_default():
    rows, meta, _ = refresh_table("客户级", "")
    return (
        rows,                     # tpl_table
        meta,                     # names_state
        "",                       # selected_name
        gr.update(value="客户级"), # template_level
        gr.update(value=""),      # template_search
        gr.update(interactive=False),  # view_btn 关
        gr.update(interactive=False),  # down_btn 关
        gr.update(interactive=False),  # del_btn 关
    )


def _split_display(display: str, fallback_level: str):
    """
    将 '客户级:foo.xlsx' 拆成 ('客户级', 'foo.xlsx')。
    如果没有冒号，则用 fallback_level 作为级别。
    """
    display = (display or "").strip()
    if ":" in display:
        lv, name = display.split(":", 1)
        return lv.strip(), name.strip()
    return (fallback_level or "").strip(), display


def _rows_to_df(rows: list[list]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=["模板名", "更新时间"])
    return pd.DataFrame(rows, columns=["模板名", "更新时间"])


def _style_df(df: pd.DataFrame, sel_idx: int):
    """返回带行高亮的 Styler；sel_idx=-1 表示不高亮"""
    def _row_style(row):
        return ['background-color: #EAF3FF' if row.name == sel_idx else '' for _ in row]
    return df.style.apply(_row_style, axis=1)


def _refresh_and_style(level: str, keyword: str, sel_idx: int):
    """按当前级别+关键词拉数据，并用 sel_idx 还原高亮"""
    rows, meta, _ = refresh_table(level, keyword)      # 你已有的函数
    df = _rows_to_df(rows)
    # 选中行如果越界则清空
    if sel_idx < 0 or sel_idx >= len(df):
        sel_idx = -1
    return _style_df(df, sel_idx), meta, sel_idx

def _normalize_to_paths(file_input):#用来支持多文档预处理
    """把 gr.File 的输入统一转成 List[str] 路径列表"""
    if not file_input:
        return []
    # 单文件（旧逻辑可能直接给 str）
    if isinstance(file_input, (str, os.PathLike)):
        return [str(file_input)]
    # 多文件或对象
    items = file_input if isinstance(file_input, (list, tuple)) else [file_input]
    paths = []
    for it in items:
        p = _extract_temp_path(it)  # 你文件里已有的工具
        if p:
            paths.append(p)
    return paths


def preprocess_ui_wrapper(doc_files):
    # ✅ 统一成路径列表
    paths = _normalize_to_paths(doc_files)
    # 合并文件出现的错误信息
    errorinfo=''
    # 初始提示
    if not paths:
        yield _pre_steps_markup(0), "等待上传/校验中…", None
        return

    # 单文件：保持你原有流程
    if len(paths) == 1:
        file_to_process = paths[0]
        yield _pre_steps_markup(0), "等待上传/校验中…", file_to_process
        for msg in preprocess_file(file_to_process):
            txt = str(msg or "")
            if "加载文档中表格信息" in txt or "解析" in txt:
                step = 1
            elif "完成" in txt or "保存" in txt or "缓存" in txt or "系统已加载该文档" in txt:
                step = 2
            else:
                step = 1
            yield _pre_steps_markup(step), txt, file_to_process
        yield _pre_steps_markup(3), "处理完成", file_to_process
        return

    # 多文件：先给出识别到的数量提示
    yield _pre_steps_markup(1), f"检测到 {len(paths)} 个文件，正在准备合并…", None

    merged_ok = False
    try:
        os.makedirs("data/tmp_merged", exist_ok=True)

        # merged_path 需要对应的文件名字：data/tmp_merged/水电站名字.docx
        docx_name = os.path.basename(paths[0]).split("-")[1]
        logger.info(f"合并文件名：{docx_name}")
        merged_path = f"data/tmp_merged/{docx_name}.docx"

        # 为清理后的文件创建专属临时目录
        timestamp = int(time.time() * 1000)
        tmp_dir = os.path.join("data", "tmp_cleaned", f"{docx_name}_{timestamp}")
        os.makedirs(tmp_dir, exist_ok=True)

        # 把每个文件清理后写入临时目录（保留原始文件名）
        for p in paths:
            base = os.path.basename(p)
            cleaned_path = os.path.join(tmp_dir, base)
            remove_toc_and_before(p, cleaned_path)   # 调用你的函数

        # 传入临时目录给 merge_docx_in_folder（函数会读取目录下的所有 docx 并合并）
        errorinfo=merge_docx_in_folder(tmp_dir, output_path=merged_path)
        merged_ok = True

    except Exception as e:
        logger.error(f"合并 docx 失败: {e}")
    finally:
        # 清理临时目录
        try:
            if tmp_dir and os.path.exists(tmp_dir):
                shutil.rmtree(tmp_dir)
        except Exception as cleanup_err:
            logger.warning(f"清理临时目录失败：{cleanup_err}")

    # 继续走单文件解析逻辑
    target_file = merged_path if merged_ok else paths[0]
    download_file = merged_path if merged_ok else None

    for msg in preprocess_file(target_file):
        txt = str(msg or "")
        if "加载文档中表格信息" in txt or "解析" in txt:
            step = 1
        elif "完成" in txt or "保存" in txt or "缓存" in txt or "系统已加载该文档" in txt:
            step = 2
        else:
            step = 1
        yield _pre_steps_markup(step), txt, download_file

    yield _pre_steps_markup(3), f"处理完成！注意：{errorinfo}", download_file

def _pre_steps_markup(active_idx: int) -> str:
    """
    active_idx: 0=就绪/等待, 1=文档/表格解析, 2=入库/缓存完成
    """
    def cls(i):
        if i < active_idx:
            return "step done"
        elif i == active_idx:
            return "step active"
        return "step"
    
    # 连接线状态类名 - 修复逻辑
    def connector_cls(i):
        # 只有当 i < active_idx 时，连接线才发光（表示前一个步骤已完成）
        if i < active_idx:
            return "connector-done"
        else:
            return "connector-waiting"
    
    # 判断连接线是否应该显示动画
    def connector_animation(i):
        # 如果当前步骤正在进行中，且这是它右边的连接线，则显示动画
        if i == active_idx:
            return "connector-progressing"
        return ""
    
    return f"""
<div class="pre-steps">
  <div class="steps">
    <div class="{cls(0)}">
      <div class="badge">{ '进行中' if active_idx==0 else ('已完成' if active_idx>0 else '就绪') }</div>
      <div class="name">上传与就绪</div>
      <div class="state">等待上传 / 校验文件</div>
    </div>
    
    <div class="step-connector">
      <div class="connector-line {connector_cls(0)} {connector_animation(0)}"></div>
    </div>
    
    <div class="{cls(1)}">
      <div class="badge">{ '进行中' if active_idx==1 else ('已完成' if active_idx>1 else '等待') }</div>
      <div class="name">文档解析</div>
      <div class="state">结构化解析、表格信息加载</div>
    </div>
    
    <div class="step-connector">
      <div class="connector-line {connector_cls(1)} {connector_animation(1)}"></div>
    </div>
    
    <div class="{cls(2)}">
      <div class="badge">{ '进行中' if active_idx==2 else ('已完成' if active_idx>2 else '等待') }</div>
      <div class="name">写入缓存</div>
      <div class="state">向量入库 / 结果缓存</div>
    </div>
  </div>
</div>
"""


def process_extract_wrapper(doc_file_path=None, excel_file_path=None):
    # 处理前禁用处理按钮，启用停止按钮
    yield gr.update(interactive=False), gr.update(interactive=True), None, "", []

    # 执行实际处理逻辑
    for result in process_extract(
        doc_file_path=doc_file_path, excel_file_path=excel_file_path
    ):
        yield gr.update(interactive=False), gr.update(interactive=True), *result

    # 处理完成后恢复按钮状态
    yield gr.update(interactive=True), gr.update(interactive=False), *result


def process_check_wrapper(doc_file_path=None, excel_file_path=None):
    """
    仅做透传：把 process_check 的 (zip_path, status_text, middle_df) 原样 yield 出去
    不在这里控制按钮状态（按钮状态用单独轻量事件处理，避免与长任务冲突）
    """
    try:
        # 可选：刚开始给一点提示（不想提示可删）
        yield None, "开始处理…", []
        for zip_or_none, status_text, middle in process_check(doc_file_path, excel_file_path):
            yield zip_or_none, status_text, middle
        # 结束时不额外多 yield，避免输出个数不一致
    except Exception as e:
        yield None, f"处理失败：{e}", []


def process_extract_wrapper(doc_file_path=None, excel_file_path=None):

    try:
        # 可选：一上来给个提示（不想要可删）
        yield None, "开始处理…", []
        for dwn, status, mid in process_extract(
            doc_file_path=doc_file_path, excel_file_path=excel_file_path
        ):
            yield dwn, status, mid
        # 不额外再 yield，保证输出个数稳定
    except Exception as e:
        yield None, f"处理失败：{e}", []


def prepare_index_zip():
    # 源 zip（你已有的）
    src_path = CONFIG.download_path
    if not os.path.exists(src_path):
        # 没文件就不更新按钮
        return gr.update(value=None)

    # 每次点击生成一个唯一名，避免浏览器缓存导致“没反应”
    ts = int(time.time() * 1000)
    out_dir = "downloads"
    os.makedirs(out_dir, exist_ok=True)
    # 如果担心中文名兼容，可换英文名；这里只示范中文 + 时间戳
    dst_path = os.path.join(out_dir, f"索引排序和自动更新_{ts}.zip")

    shutil.copy2(src_path, dst_path)
    # 两种写法都可以：直接返回路径 或 gr.update(value=路径)
    return gr.update(value=dst_path)

def beautify_label(file):
    if file is not None:
        # 上传后显示绿色图标+文字
        return gr.update(label="✅ 已上传")
    else:
        # 恢复原始状态
        return gr.update(label="上传DOCX文件")


if __name__ == "__main__":
    gradio_flag = True
    if gradio_flag:
        with open("css/style2.css", "r", encoding="utf-8") as f:
            custom_css = f.read()
        with gr.Blocks(css=custom_css) as demo:

            with gr.Row(elem_classes="app-shell"):

                # ===== 左侧：固定侧栏 =====
                with gr.Column(elem_id="sidebar", elem_classes="sidebar"):
                    # 品牌与副标题（可换成你的 LOGO）
                    gr.Markdown(
                        """
                        <div class="brand">
                        <div class="brand-icon">⚙️</div>
                        <div class="brand-title">
                            <div class="title">AI数据提取与辅助校验</div>
                            <div class="subtitle">智能提效 · 提升准确</div>
                        </div>
                        </div>
                        """,
                        elem_classes="brand-wrap"
                    )
                    # 导航按钮
                    choose_instructions_btn = gr.Button("📘 工具使用说明", elem_classes=["nav-btn"])
                    choose_preprocess_btn   = gr.Button("📄 工程文档预处理", elem_classes=["nav-btn"])
                    choose_extract_btn      = gr.Button("⚙️ 工程数据提取",   elem_classes=["nav-btn"])
                    choose_review_btn       = gr.Button("🧪 工程数据复核",   elem_classes=["nav-btn"])
                    choose_template_btn     = gr.Button("🗂️ 数据提取模板管理", elem_classes=["nav-btn"])

                # ===== 右侧：内容区=====
                with gr.Column(elem_id="content", elem_classes="content-area"):

                    with gr.Column(visible=False) as instructions_page:
                        with gr.Row():
                            with gr.Column(scale=7):
                                pass

                        with gr.Row():
                            with gr.Column(scale=1, min_width=200):
                                preprocess_instructions_btn = gr.Button("工程文档预处理")
                                template_instructions_btn = gr.Button(
                                    "数据提取模板管理", visible=False
                                )
                                extract_instructions_btn = gr.Button("工程数据提取")
                                review_instructions_btn = gr.Button("工程数据复核")
                            with gr.Column(scale=7):
                                preprocess_markdown = gr.Markdown(
                                    load_markdown_file("preprocess"),
                                    visible=True,
                                )
                                template_markdown = gr.Markdown(
                                    load_markdown_file("template"),
                                    visible=False,
                                )
                                extract_markdown = gr.Markdown(
                                    load_markdown_file("extract"),
                                    visible=False,
                                )
                                review_markdown = gr.Markdown(
                                    load_markdown_file("verify"),
                                    visible=False,
                                )
                        preprocess_instructions_btn.click(
                            lambda: (
                                gr.update(visible=True),
                                gr.update(visible=False),
                                gr.update(visible=False),
                                gr.update(visible=False),
                            ),
                            outputs=[
                                preprocess_markdown,
                                template_markdown,
                                extract_markdown,
                                review_markdown,
                            ],
                        )
                        template_instructions_btn.click(
                            lambda: (
                                gr.update(visible=False),
                                gr.update(visible=True),
                                gr.update(visible=False),
                                gr.update(visible=False),
                            ),
                            outputs=[
                                preprocess_markdown,
                                template_markdown,
                                extract_markdown,
                                review_markdown,
                            ],
                        )
                        extract_instructions_btn.click(
                            lambda: (
                                gr.update(visible=False),
                                gr.update(visible=False),
                                gr.update(visible=True),
                                gr.update(visible=False),
                            ),
                            outputs=[
                                preprocess_markdown,
                                template_markdown,
                                extract_markdown,
                                review_markdown,
                            ],
                        )
                        review_instructions_btn.click(
                            lambda: (
                                gr.update(visible=False),
                                gr.update(visible=False),
                                gr.update(visible=False),
                                gr.update(visible=True),
                            ),
                            outputs=[
                                preprocess_markdown,
                                template_markdown,
                                extract_markdown,
                                review_markdown,
                            ],
                        )

                    with gr.Column(visible=False) as template_page:

                        # 顶部黄条说明（接近参考图“操作指南”banner）
                        with gr.Row():
                            gr.Markdown(
                                """
                    <div class="notice-banner">
                    <div class="notice-title">操作指南</div>
                    <ul class="notice-list">
                        <li>该页面提供模板的统一入口，可以上传、删除、下载模板文件，上传过程将自动进行格式校验。</li>
                        <li>在数据提取页面可以从全部模板文件中选择指定模板进行处理。</li>
                        <li>如需在本地批量增删改，请点击下方“下载工具”获取“索引排序更新工具”。</li>
                    </ul>
                    </div>
                                """,
                                elem_id="template-notice"
                            )

                        # ===== 在构建“模板管理”这页的布局里，找到“模板预览”块之后、‘操作说明’之前，插入以下代码 =====
                        with gr.Column(elem_id="index-download-box"):
                            gr.Markdown("### 下载文件", elem_classes=["tpl-title"])

                            with gr.Row():
                                with gr.Column(scale=6):
                                    gr.Markdown(
                                        """
                                        📂 **辅助工具**  
                                        包括用于本地的数据表索引排序和快速获取下载工具，以及多文档合并工具及使用说明
                                        """,
                                        elem_classes=["index-desc"]
                                    )
                                with gr.Column(elem_classes="right-align"):
                                    idx_zip_btn = gr.DownloadButton(
                                        "辅助工具下载",
                                        elem_classes="btn-index btn-index-sm",   # 绿色按钮
                                        value=None,                 # 初始无文件
                                        interactive=True,
                                    )
                            # 事件：点击后把路径“赋值给下载按钮自身”
                        idx_zip_btn.click(
                            fn=prepare_index_zip,
                            inputs=None,
                            outputs=[idx_zip_btn],
                        )

                        # 三列卡片：上传/删除/下载 说明
                        with gr.Row(elem_classes="three-col-grid"):
                            with gr.Column(elem_classes="guide-card"):
                                gr.Markdown(
                                    """
                    <div class="section-card">
                    <div class="card-title">📤 上传新模板</div>
                    <ul class="card-list">
                        <li><b>步骤1</b>：点击下方<b>上传模板</b>按钮</li>
                        <li><b>步骤2</b>：在本地文件中选择模板上传</li>
                    </ul>
                    <div class="card-tip error">注意：模板文件名不能相同；若需覆盖旧模板，请先删除原模板</div>
                    </div>
                                    """,
                                )
                            with gr.Column(elem_classes="guide-card"):
                                gr.Markdown(
                                    """
                    <div class="section-card">
                    <div class="card-title">🗑️ 删除已有模板</div>
                    <ul class="card-list">
                        <li><b>步骤1</b>：下方选择模板级别和文件</li>
                        <li><b>步骤2</b>：点击下方<b>删除</b>按钮</li>
                    </ul>
                    <div class="card-tip warning">注意：删除后无法恢复，建议定期备份重要模板</div>
                    </div>
                                    """,
                                )
                            with gr.Column(elem_classes="guide-card"):
                                gr.Markdown(
                                    """
                    <div class="section-card">
                    <div class="card-title">⬇️ 下载已有模板</div>
                    <ul class="card-list">
                        <li><b>步骤1</b>：在下方列表中选择要下载的模板</li>
                        <li><b>步骤2</b>：点击链接下载</li>
                    </ul>
                    </div>
                                    """,
                                )

                        # 控件区：级别 / 下拉 / 拖拽上传框 / 按钮与状态
                        # ========== 功能区（级别+搜索+清单+操作） ==========


                        # 统一的“名字列表 state”，不要使用临时 gr.State()
                        names_state = gr.State([])
                        # === 模板预览模块（一个卡片里包含：标题 + 级别单选 + 搜索 + 表格）=== #
                        with gr.Row():
                                gr.Markdown("### 模板预览", elem_id="tpl-title")

                        with gr.Column(elem_classes="tpl-box"):

                            with gr.Row(elem_classes="right-align"):
                                upload_btn = gr.UploadButton(
                                    "上传模板",
                                    file_types=[".xls", ".xlsx"],
                                    elem_classes="btn-process btn-process-sm",  # 继续用你的小号深蓝按钮
                                    scale=0,                                     # 不拉伸
                                    min_width=220                                # 可按需改宽度
                                )
 
                            # 控制区：左边用户级，右边搜索
                            with gr.Row(elem_classes="tpl-controls"):
                                with gr.Column(scale=1, elem_classes=["col-left"]):
                                    template_level = gr.Radio(
                                        ["客户级", "系统级"],
                                        value="客户级",
                                        show_label=False,              # ✅ 隐藏“用户级”行
                                        elem_classes="radio-compact",
                                    )
                                with gr.Column(scale=1, elem_classes=["col-right"]):
                                    template_search = gr.Textbox(
                                        placeholder="输入关键词，回车过滤…",
                                        show_label=False,              # ✅ 这里改为 show_label=False
                                        lines=1,
                                        elem_classes="search-compact",
                                    )


                            # 表格
                        
                            tpl_table = gr.Dataframe(
                                value=[["",""]],
                                datatype=["str", "str"],
                                headers=["模板名", "更新时间"],
                                interactive=False,
                                row_count=(0, "dynamic"),
                                col_count=2,
                                wrap=False,
                                elem_classes="df-compact",
                                
                            )
                            names_state   = gr.State([])          # 存 meta 列表（on_table_select 用）
                            selected_idx  = gr.State(-1)          # 选中行号（保留你现在的 State 用法）
                            selected_name = gr.Textbox(value="", visible=False)  # "客户级:xxx.xlsx"




                        with gr.Row():
                            # 右侧操作：查看 / 下载 / 删除（链接风格），以及上传
                        
                                with gr.Row(elem_classes="table-ops-inline"):
                                    selected_name = gr.Textbox(value="", visible=False)
                                    with gr.Row(elem_classes="ops-left"):
                                        view_btn   = gr.Button("预览", elem_classes="action-btn btn-process")
                                        down_btn = gr.DownloadButton(
                                            "下载",
                                            elem_classes="action-btn btn-process", 
                                            interactive=True)
                                        del_btn    = gr.Button("删除", elem_classes="action-btn btn-stop")
                        
                        
                        # 结构预览（继续用DataFrame）
                        with gr.Column(scale=6):
                            preview_df = gr.Dataframe(
                                value=[["", ""]], 
                                headers=["索引", "关键词"], 
                                row_count=5, 
                                col_count=2, 
                                wrap=True, 
                                elem_classes="shadow-box"
                            )


                    

                        # --- 初始化：禁用按钮 ---
                        view_btn.interactive = False
                        down_btn.interactive = False
                        del_btn.interactive  = False
                        selected_name.value  = ""

                        # --- 小工具：根据是否选中启/禁 三个按钮 ---
                        def _btns_enabled(name: str):
                            on = bool(name)
                            return (
                                gr.update(interactive=on),  # view_btn
                                gr.update(interactive=on),  # down_btn（值还没给，先允许点；或保持 False 直到准备好路径也可）
                                gr.update(interactive=on),  # del_btn
                            )

                        # --- 级别/搜索 -> 刷新表并禁用按钮 ---
                        def _on_filter_change(level: str, keyword: str, cur_idx: int):
                            rows, meta, _ = refresh_table(level, keyword)
                            # 清空选中，按钮禁用
                            return (
                                rows,               # tpl_table
                                meta,               # names_state
                                -1,                 # selected_idx
                                "",                 # selected_name
                                gr.update(interactive=False),                      # view_btn
                                gr.update(value=None, interactive=False),          # down_btn
                                gr.update(interactive=False),                      # del_btn
                            )

                        template_level.change(
                            _on_filter_change,
                            inputs=[template_level, template_search, selected_idx],
                            outputs=[tpl_table, names_state, selected_idx, selected_name, view_btn, down_btn, del_btn],
                        )
                        template_search.submit(
                            _on_filter_change,
                            inputs=[template_level, template_search, selected_idx],
                            outputs=[tpl_table, names_state, selected_idx, selected_name, view_btn, down_btn, del_btn],
                        )


                        # 表格选中 → 解析出 "客户级:xxx.xlsx"
                        def _on_table_select(evt: gr.SelectData, meta: list[dict], level_radio: str):
                            # 用你自己的解析函数把选择转成 "客户级:xxx.xlsx"
                            selected = on_table_select(evt, meta)   # 若 meta/evt 无效则返回 ""
                            # 行号
                            idx = evt.index[0] if isinstance(evt.index, (tuple, list)) else evt.index
                            try:
                                idx = int(idx)
                            except Exception:
                                idx = -1
                            # 返回：selected_name、selected_idx、以及三个按钮的启用状态
                            return (
                                selected,
                                idx,
                                *(
                                    _btns_enabled(selected)  # 这会返回 (view_btn, down_btn, del_btn) 的 gr.update
                                ),
                            )

                        tpl_table.select(
                            fn=_on_table_select,
                            inputs=[names_state, template_level],
                            outputs=[selected_name, selected_idx, view_btn, down_btn, del_btn],
                        )



                        # --- 预览：仅更新 DataFrame；顺手准备下载路径（给 DownloadButton 的 value）---
                        def _restyle_after_action(level: str, keyword: str, cur_idx: int):
                            styled_df, _, saved_idx = _refresh_and_style(level, keyword, cur_idx)
                            return styled_df, saved_idx

                        view_btn.click(
                            fn=preview_only,  # 你的预览函数：inputs=[selected_name, template_level] -> outputs=[preview_df]
                            inputs=[selected_name, template_level],
                            outputs=[preview_df],
                        ).then(
                            fn=_restyle_after_action,
                            inputs=[template_level, template_search, selected_idx],
                            outputs=[tpl_table, selected_idx],
                        )

                        # --- 直接点“下载”：重新准备一次路径（保证最新）---
                        down_btn.click(
                            fn=prepare_download_template,
                            inputs=[selected_name, template_level],
                            outputs=[down_btn],
                        ).then(
                            fn=_restyle_after_action,
                            inputs=[template_level, template_search, selected_idx],
                            outputs=[tpl_table, selected_idx],
                        )

                        # --- 删除：刷新表 + 清空选中 + 禁用按钮 ---
                        def _delete_and_refresh(selected_display: str, current_level: str, keyword: str):
                            # 先把 '客户级:foo.xlsx' 拆成 (level, name)
                            lv, name = _split_display(selected_display, current_level)
                            if name:
                                # 这里一定只传“纯文件名”给 delete_template，级别单独传
                                _msg, *_ = delete_template(name, lv)

                            # 刷新当前级别 + 关键词的表格
                            rows, meta, _ = refresh_table(current_level, keyword)

                            # 返回：表格、meta、清空选中 & 禁用 3 个按钮（查看/下载/删除）
                            return (
                                rows,               # tpl_table
                                meta,               # names_state
                                "",                 # selected_name 清空
                                gr.update(interactive=False),
                                gr.update(value=None, interactive=False),
                                gr.update(interactive=False),
                            )

                        del_btn.click(
                            fn=_delete_and_refresh,
                            inputs=[selected_name, template_level, template_search],
                            outputs=[tpl_table, names_state, selected_name, view_btn, down_btn, del_btn],
                        ).then(
                            fn=_restyle_after_action,
                            inputs=[template_level, template_search, selected_idx],
                            outputs=[tpl_table, selected_idx],
                        )

                        # --- 上传：保存后刷新 + 清空选中 + 禁用按钮 ---
                        upload_btn.upload(
                            fn=save_uploaded_template,
                            inputs=[upload_btn, template_level, template_search],  # 传 UploadButton 自身 + 级别 + 搜索词
                            outputs=[tpl_table, names_state, selected_name],       # 刷新表、刷新 meta、清空选中
                        ).then(
                            # 上传完成后：统一禁用“预览/下载/删除”，等待用户重新选中表格行再启用
                            fn=lambda: (
                                gr.update(interactive=False),  # view_btn
                                gr.update(interactive=False),  # down_btn
                                gr.update(interactive=False),  # del_btn
                            ),
                            inputs=None,
                            outputs=[view_btn, down_btn, del_btn],
                        )


                        


                    with gr.Column(visible=False) as preprocess_page:
                    # 顶部步骤标题 + 说明卡片
                        with gr.Row():
                            with gr.Column(scale=12):
                                
                                gr.Markdown(
                                    """
                                <div class="pre-tip">
                                <div class="title">使用说明</div>
                                <ol style="margin:0 0 6px 20px;">
                                    <li>上传 <b>DOCX</b> 格式的工程文档。</li>
                                    <li>系统会自动完成文档结构化解析与表格提取。</li>
                                    <li>处理完成后，后续提取/复核可复用缓存以加速。</li>
                                </ol>
                                <div class="warn">注意：仅支持 <b>DOCX</b> 文件；若文件较大，解析时间会稍长。</div>
                                </div>
                                    """,
                                    elem_classes=["no-margin"]
                                )

                        with gr.Row():
                            gr.Markdown("📂 第一步：上传 DOCX 文档", elem_classes="step-title")

                        
                        def add_file_and_clear(new_file, state_files):
                            state_files = state_files or []
                            paths = _normalize_to_paths(new_file)

                            # 追加去重
                            for p in paths:
                                if p not in state_files:
                                    state_files.append(p)

                            # 展示列表
                            html = "<ul style='line-height:1.8em;margin:8px 0;'>"
                            for p in state_files:
                                fname = os.path.basename(p)
                                try:
                                    size_k = os.path.getsize(p) / 1024
                                    size_str = f"{size_k:.1f} KB"
                                except Exception:
                                    size_str = "--"
                                html += f"<li>📄 {fname}（{size_str}）</li>"
                            html += "</ul>"

                            # 关键：立刻清空上传框 & 按 state 判定按钮是否可点
                            return state_files, html, None, gr.update(interactive=bool(state_files))
                        def preprocess_from_state(state_files):
                            print(">>> 点击按钮时 state =", state_files)
                            if not state_files:
                                # 三个输出：步骤HTML / 状态文本 / 下载路径
                                yield _pre_steps_markup(0), "请先上传至少一个文件", None
                                return
                            # 逐条透传 generator（务必三元组）
                            for step, msg, path in preprocess_ui_wrapper(state_files):
                                yield step, msg, path

                        with gr.Row():
                            with gr.Column(scale=6):
                                input_doc_file = gr.File(
                                    label="单文档上传或批量上传（支持多选）",
                                    type="filepath",
                                    file_types=[".docx", ".doc"],  # 允许doc/docx
                                    file_count="multiple",
                                    elem_classes="upload-box",
                                )
                                file_state = gr.State([])
                                file_list_html = gr.HTML(label="当前已上传文件")


                        # 第二步：执行预处理
                        with gr.Row():
                            gr.Markdown("📌 **第二步：点击文档预处理**", elem_classes="step-title")

                        with gr.Row():
                            with gr.Column(scale=9):
                                process_button = gr.Button(
                                    "文档预处理",
                                    elem_classes="btn-process",  # 橙色主按钮（与你提取页一致）
                                    interactive=False
                                )
                                input_doc_file.change(
                                    fn=add_file_and_clear,
                                    inputs=[input_doc_file, file_state],
                                    outputs=[file_state, file_list_html, input_doc_file, process_button],
                                )

                        # 处理信息（状态条）
                       # 状态条 + 状态文本
                        with gr.Row():
                            gr.Markdown(" **状态处理**", elem_classes="step-title")
                        pre_steps_html = gr.HTML(value=_pre_steps_markup(0))
                        pre_status = gr.Markdown("")  # 显示“文档解析中/处理完成…”等
                        pre_download_btn = gr.DownloadButton(label="下载处理后的DOCX", value=None)
                        
                        process_event_preprocess = process_button.click(
                            fn=preprocess_from_state,
                            inputs=[file_state],
                            outputs=[pre_steps_html, pre_status, pre_download_btn],
                            concurrency_limit=2,  # 限制最多X个并发用户
                        )

                        # 监听文件上传，动态更新按钮状态
                        input_doc_file.change(
                            fn=check_files_pre,
                            inputs=[input_doc_file, file_state],
                            outputs=[process_button],  # 同时更新按钮
                        )

                    with gr.Column(visible=False) as extract_page:

                        with gr.Row():
                            with gr.Column(scale=12):
                                gr.Markdown("📌 第一步：上传DOCX文档和XLSX模板文件", elem_classes="step-title")
                                gr.Markdown(
                                """
                                <div class="hint-box">
                                <span class="hint-emoji">🔑</span>
                                <span class="hint-title">提示：</span>
                                <span> XLSX模板文件可以在下拉框中选择已有模板，或者直接上传本地模板</span>
                                </div>
                                """
                            )

                        with gr.Row(elem_classes="two-col-grid"):
                            with gr.Column(scale=6, elem_classes="left-span-2"):
                                input_doc_file = gr.File(
                                label="上传DOCX文件",
                                file_types=[".doc", ".docx"],
                                elem_classes="upload-box"
                            )
                                input_doc_file.change(
                                    fn=beautify_label,
                                    inputs=[input_doc_file],
                                    outputs=[input_doc_file]
                                )
                            with gr.Column(scale=6, elem_classes="right-top"):
                                    template_level_extract = gr.Radio(
                                        choices=["客户级", "系统级"],
                                        value="客户级",
                                        label="模板级别",
                                    )

                                    select_template_dropdown = gr.Dropdown(
                                        choices=get_all_templates(),
                                        label="模板列表",
                                        info="选择需要的模板文件（非必选，可自己上传）",
                                    )
                            with gr.Column(scale=6, elem_classes="right-bottom"):
                                    input_excel_file = gr.File(
                                        label="上传EXCEL文件",
                                        file_types=[".xls", ".xlsx"],
                                        elem_classes="upload-box"
                                    )
                                    input_excel_file.change(
                                        fn=beautify_label,
                                        inputs=[input_excel_file],
                                        outputs=[input_excel_file]
                                    )
                        with gr.Row():
                            gr.Markdown("📌 第二步：点击工程数据提取", elem_classes="step-title")

                        with gr.Row():
                            with gr.Column(scale=9):
                                process_button = gr.Button(
                                "工程数据提取", 
                                interactive=False,
                                elem_classes="btn-process" 
                                )
                            with gr.Column(scale=1, min_width=80):
                                reset_button_extract = gr.Button(
                                "停止",
                                interactive=False,
                                elem_classes="btn-stop"
                            )
                
                        with gr.Row():
                            progress_bar = gr.Textbox(
                                info="处理信息",
                                max_lines=1,
                                interactive=False,
                                show_label=False,
                                elem_classes="status-bar",
                            )

                        with gr.Row():
                            gr.Markdown("📥 第三步：数据结果提取", elem_classes="step-title")
                            #output_file = gr.File(label="下载处理后的EXCEL文件", elem_classes="output-box")

                        with gr.Row(elem_classes="result-card"):
                            text_bar = gr.DataFrame(
                            label="处理结果",
                            headers=["关键词", "单位", "数量"],
                            row_count=2,
                            col_count=3,
                            wrap=True,
                            interactive=False,
                            elem_classes="styled-table",
                        )
                        with gr.Row():
                            download_btn = gr.DownloadButton(
                                "⬇️ 下载结果",
                                elem_classes=["btn-success", "download-float"],
                                visible=True,
                                interactive=False
                            )


                        # 监听文件上传，动态更新按钮状态
                        input_doc_file.change(
                            fn=check_files,
                            inputs=[input_doc_file, input_excel_file],
                            outputs=[process_button],  # 同时更新两个按钮
                        )

                        input_excel_file.change(
                            fn=check_files,
                            inputs=[input_doc_file, input_excel_file],
                            outputs=[process_button],  # 同时更新两个按钮
                        )

                        process_button.click(
                            fn=lambda: (
                                gr.update(interactive=False),      # 开始禁用
                                gr.update(interactive=True),       # 停止启用
                                gr.update(value=None, interactive=False)  # 清空并禁用下载按钮（防止点到旧文件）
                            ),
                            inputs=None,
                            outputs=[process_button, reset_button_extract, download_btn],
                            queue=False   # 轻量事件，不进队列
                        )

                        process_event_extract = process_button.click(
                            fn=process_extract_wrapper,            # 用 wrapper 跑长任务
                            inputs=[input_doc_file, input_excel_file],
                            outputs=[download_btn, progress_bar, text_bar],  # 3个固定输出
                            concurrency_limit=2,                   # 你原来的并发限制
                        )

                        # 3.3 长任务完成 -> 复原按钮；下载按钮启用
                        process_event_extract.then(
                            fn=lambda: (
                                gr.update(interactive=True),   # 开始可点
                                gr.update(interactive=False),  # 停止禁用
                                gr.update(interactive=True),   # 下载可点
                            ),
                            inputs=None,
                            outputs=[process_button, reset_button_extract, download_btn],
                            queue=False
                        )
                        #模版选择
                        select_template_dropdown.select(
                            fn=extract_dropdown_change,
                            inputs=[select_template_dropdown, template_level_extract],
                            outputs=[input_excel_file],
                        ).then(
                            fn=check_files,
                            inputs=[input_doc_file, input_excel_file],
                            outputs=[process_button],  # 启用按钮
                        )
                        
                        template_level_extract.change(
                            fn=lambda level: gr.update(choices=get_all_templates(level)),
                            inputs=[template_level_extract],
                            outputs=[select_template_dropdown],
                        )

                        reset_button_extract.click(
                            fn=lambda: (
                                "已停止",                        # progress_bar 提示
                                gr.update(interactive=True),     # 开始可点
                                gr.update(interactive=False),    # 停止禁用
                                # 不返回 text_bar / download_btn，保持现状（你要求“停止后不抹除处理结果”）
                            ),
                            inputs=None,
                            outputs=[progress_bar, process_button, reset_button_extract],
                            cancels=[process_event_extract],      # 只取消这一个长任务
                            queue=False
                        )

                    with gr.Column(visible=False) as review_page:
                        # —— 步骤 1：说明 + 上传 —— 
                        with gr.Column(scale=12):
                            gr.Markdown("📥 **第一步：上传 DOCX 和 XLSX 文件**", elem_classes="step-title")

                        with gr.Row(elem_classes="two-col-grid"):
                            with gr.Column(scale=6, elem_classes="left-span-2"):
                                input_doc_file2 = gr.File(label="上传DOCX文件", file_types=[".doc", ".docx"])
                                input_doc_file2.change(
                                    fn=beautify_label,
                                    inputs=[input_doc_file2],
                                    outputs=[input_doc_file2]
                                )
                            with gr.Column(scale=6, elem_classes="right-top"):
                                input_excel_file2 = gr.File(label="上传EXCEL文件", file_types=[".xls", ".xlsx"])
                                input_excel_file2.change(
                                        fn=beautify_label,
                                        inputs=[input_excel_file2],
                                        outputs=[input_excel_file2]
                                    )

                        # —— 步骤 2：开始/停止（样式与提取页一致）——
                        with gr.Row():
                            gr.Markdown("🖱️ **第二步：点击工程数据复核**", elem_classes="step-title")

                        # 拉满的开始按钮 + 右上角停止按钮
                        with gr.Row():
                            with gr.Column(scale=9):
                                process_button2 = gr.Button(
                                    "工程数据复核",
                                    interactive=False,          # 初始禁用，等文件选好
                                    elem_classes="btn-process"  # 你已定义好的深蓝按钮
                                )
                            with gr.Column(scale=1, min_width=80):
                                reset_button_check = gr.Button(
                                    "停止",
                                    interactive=False,          # 运行后才可用
                                    elem_classes="btn-stop"     # 你已定义好的红色按钮
                                )
                        with gr.Row():
                            progress_bar2 = gr.Textbox(
                                info="处理信息",
                                max_lines=1,
                                interactive=False,
                                show_label=False,
                                elem_classes="status-bar",
                                )

                        # —— 步骤 3：下载 —— 
                        with gr.Row():
                            gr.Markdown("📥 **第三步：下载生成文件**", elem_classes="step-title")

                            #output_file2 = gr.File(label="下载处理后的文件")


                        with gr.Row(elem_classes="result-card"):
                            text_bar2 = gr.DataFrame(
                                label="处理结果",
                                headers=["关键词", "单位", "数量", "备注"],  # 4 个表头
                                row_count=2,
                                col_count=4,                                # ← 改成 4
                                value=[["", "", "", ""]]*1,                 # 可选：初始化 5 行空数据，列数要匹配
                                wrap=True,
                                interactive=False,
                                elem_classes="styled-table",
                            )

                        with gr.Row():
                            download_btn = gr.DownloadButton(
                                "⬇️ 下载结果",
                                elem_classes=["btn-success", "download-float"],
                                visible=True,
                                interactive=False,
                                value=None           # 初始无文件
                            )


                        # ============== 交互 ==============

                        # 1) 文件就绪 -> 启用“工程数据复核”按钮（保持你的逻辑）
                        input_doc_file2.change(
                            fn=check_files,
                            inputs=[input_doc_file2, input_excel_file2],
                            outputs=[process_button2],
                        )
                        input_excel_file2.change(
                            fn=check_files,
                            inputs=[input_doc_file2, input_excel_file2],
                            outputs=[process_button2],
                        )

                        # 2) 点击“开始”：跑长任务（queue=False 关键），同时立刻把“开始禁用/停止启用/下载禁用”
                        process_event_check = process_button2.click(
                            fn=process_check_wrapper,                        # 只产出 zip路径/状态文本/中间表
                            inputs=[input_doc_file2, input_excel_file2],
                            outputs=[download_btn, progress_bar2, text_bar2],
                                                               # 关键：禁用队列，规避取消后 KeyError
                            concurrency_limit=2,
                        )

                        # 立刻切 UI（轻量事件，queue=False，避免排队）：
                        process_button2.click(
                            fn=lambda: (gr.update(interactive=False), gr.update(interactive=True), gr.update(interactive=False), ""),
                            inputs=None,
                            outputs=[process_button2, reset_button_check, download_btn, progress_bar2],
                           
                        )

                        # 长任务结束后 -> “开始”可点，“停止”禁用，“下载”可点
                        process_event_check.then(
                            fn=lambda: (gr.update(interactive=True), gr.update(interactive=False), gr.update(interactive=True)),
                            inputs=None,
                            outputs=[process_button2, reset_button_check, download_btn],
                            queue=False
                        )

                        # 3) 点击“停止”：仅做极短 UI 更新 + 取消长任务（queue=False）
                        def _stop_and_reset():
                            # 仅负责即时 UI：给提示、开始可点、停止禁用、下载禁用
                            return (
                                "已停止",
                                gr.update(interactive=True),
                                gr.update(interactive=False),
                                gr.update(interactive=False),
                                None,  # 清空下载文件
                            )

                        reset_button_check.click(
                            fn=_stop_and_reset,
                            inputs=None,
                            outputs=[progress_bar2, process_button2, reset_button_check, download_btn, download_btn],
                            cancels=[process_event_check],   # 只取消这一个长任务
                            queue=False
                        )

                    
                    
                    # 页面切换逻辑
                    pages = [
                       
                        instructions_page,
                        preprocess_page,
                        template_page,
                        extract_page,
                        review_page,
                    ]
                    choose_instructions_btn.click(create_tab_switch_handler(0), outputs=pages)
                    choose_preprocess_btn.click(create_tab_switch_handler(1), outputs=pages)
                    # 进入“模板管理”页
                    choose_template_btn.click(
                        create_tab_switch_handler(2),   # 保持你原来的索引
                        outputs=pages
                    ).then(
                        fn=_open_template_page_default,
                        inputs=None,
                        outputs=[
                            tpl_table,        # Dataframe
                            names_state,      # gr.State([]) 统一放模板名列表
                            selected_name,    # 当前选中模板名（隐藏框）
                            template_level,   # 单选：客户级/系统级
                            template_search,  # 搜索框
                            view_btn,
                            down_btn,
                            del_btn,
                        ]
                    )

                    choose_extract_btn.click(create_tab_switch_handler(3), outputs=pages).then(
                        fn=lambda: gr.update(choices=get_all_templates("客户级")), inputs=None, outputs=[select_template_dropdown]
                    )
                    choose_review_btn.click(create_tab_switch_handler(4), outputs=pages)

        demo.queue(
            max_size=20            # 等待队列长度，按需
        )


        # demo.launch(server_name="0.0.0.0")  # 启动服务
        demo.queue().launch(
            server_name="0.0.0.0", allowed_paths=["project_instructions/markdown/media"], server_port=7860
            ,
        )

        # demo.launch(share=True)  # 外网链接

    #demo.launch(
    #     auth=("admin", "2024"),  # 登录凭证
         #auth_message="欢迎使用AI数据提取与辅助校验工具",  # 自定义登录提示信息
         #server_port=7860,
        # show_api=False,  # 是否显示API文档
         # favicon_path="./assets/images/login_logo.png",  # 如果提供了文件的路径（.png、.gif 或 .ico），则该路径将用作网页的网站图标。
#     # max_file_size = ""  # 最大上传文件大小，单位为字节。默认情况下，没有限制。
    # )  # 启动服务



