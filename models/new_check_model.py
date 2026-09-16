import json
import re
from docx import Document as pyDocument
from docx.table import Table
from langchain_community.vectorstores import Chroma, FAISS, Weaviate
from utils.utils import filter_gt

# from langchain_milvus import Milvus
from langchain_community import embeddings
from langchain_ollama import OllamaEmbeddings
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
from utils.utils import get_appendix, expand_relevant_docs_text
from config import CONFIG
from data_processor import read_docx_file
import ollama
from datetime import datetime
import time
from tqdm import tqdm
from langchain.schema import BaseOutputParser

import random
import pandas as pd

# from sentence_transformers import SentenceTransformer
from langchain.embeddings import HuggingFaceEmbeddings

# from rerank_model import OllamaReranker
# conan 选择rerank_model 其余的模型选择new_rerank_model
from models.new_rerank_model import rerank_docs, set_seed, rrf_docs  # dcx 修改

# from models.rerank_model import rerank_docs  #dcx 修改于2024/10/31

from models.prompt_engineering import (
    check_embed_prompt,
    check_prompt,
    extract_prompt,
    check_llm_query_keyword,
    check_rerank_prompt,
    splited_check_prompt,
)
from langchain_community.llms.tongyi import Tongyi
from utils.logger_util import logger

from models.llm_model import get_llm_model
import os
from data_processor import html_table_to_markdown, is_table_caption

from models.embedding_model import MultiOutputParser, get_embedding_model

# from models.new_embed_model import MultiOutputParser, get_embedding_model  别开，这个embed效果不好
from utils.utils import error_handler
from chromadb.config import Settings
import shutil
import os
import torch
from models.hybrid_search import hybrid_search
from utils.utils import chunk_filter

import asyncio
from concurrent.futures import ThreadPoolExecutor
from functools import partial


@error_handler("大模型问答处理出错，请联系技术人员")
def call_check_new_model(
    tbl_info,
    retriever,
    chunks,
    name_word,
    elements_docx=None,
    progress=None,
    gt=None,
    time_check=0,
):
    ans_list = []
    results = []
    j = 0  # gt的值遍历
    all_rerank_scores = []
    true_docs_list = []
    topk_docs_list = []
    all_proof_list = []
    txt_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    progress_bar = tqdm if progress is None else progress.tqdm
    rerank_query_data = {}
    ## 寻找docx文件对应的rerank_json文件
    json_files = []
    for filename in os.listdir(CONFIG.json_rerank_folder_path):
        file_path = os.path.join(CONFIG.json_rerank_folder_path, filename)
        if os.path.isfile(file_path) and filename.startswith(name_word):
            json_files.append(file_path)

    rerank_json_path = None if len(json_files) == 0 else json_files[0]

    if time_check > 0 and rerank_json_path:
        logger.info(f"找到rerank的json文件：{rerank_json_path}")
        with open(rerank_json_path, "r", encoding="utf-8") as f:
            rerank_query_data = json.load(f)
    for i in progress_bar(range(len(tbl_info["need_query"]))):
        logger.info(f"表格总长度：{len(tbl_info['need_query'])}")
        if not tbl_info["need_query"][i]:
            continue
        logger.info(f"第{i}个:{tbl_info['keywords'][i]}")
        row_info = {
            "need_query": tbl_info["need_query"][i],
            "is_special": tbl_info["is_special"][i],
            "keywords": tbl_info["keywords"][i],
            "answers": tbl_info["answers"][i],  # 单位|数量  or  描述性的短句
        }
        query_keyword = row_info["keywords"]
        answer = row_info["answers"]

        # 更新会话 ollama/tongyi
        model = get_llm_model(model_type="ollama", query=query_keyword)

        # 需要选择,需要询问的目标:gt[j]['flag']==1,0,-1
        # 1表示正文中出现了正确答案且容易识别真值（正样本）
        # -1表示正文没有出现答案（负样本）
        # 0表示出现了正确答案但不容易识别（如答案带有/、*等字符，）
        if gt and (
            gt[j]["flag"] == 1
        ):  # or gt[j]["flag"] == -1 or gt[j]["flag"]) == 0:
            (
                ans_true,
                rerank_context,
                rerank_scores_topk,
                ans_score,
                context_relevant_docs,
                llm_context,
                true_docs,
                topk_docs,
                proof_list,
                rerank_query_data,
            ) = query_llm(
                query_keyword,
                retriever,
                model,
                txt_time,
                chunks,
                elements_docx,
                name_word,
                row_info["is_special"],
                answer,
                time_check=time_check,
                gt_i=gt[j],
                rerank_query_data=rerank_query_data,
                rerank_json_path=rerank_json_path,
            )
            all_rerank_scores += rerank_scores_topk
            ans_list.append(ans_true)
            true_docs_list.append(true_docs)
            topk_docs_list.append(topk_docs)
            logger_keyword = query_keyword.replace("|", "")  # 相似度搜索，用“”连接
            results.append(
                collect_results(
                    logger_keyword,
                    rerank_context,
                    context_relevant_docs,
                    ans_true,
                    gt[j],
                    row_info["is_special"],
                    ans_score,
                    llm_context,
                )
            )
            all_proof_list.append(proof_list)
        else:
            if gt:
                ans_true = "无| 无 | 无"
            else:
                (
                    ans_true,
                    rerank_context,
                    rerank_scores_topk,
                    ans_score,
                    context_relevant_docs,
                    llm_context,
                    true_docs,
                    topk_docs,
                    proof_list,
                    rerank_query_data,
                ) = query_llm(
                    query_keyword,
                    retriever,
                    model,
                    txt_time,
                    chunks,
                    elements_docx,
                    name_word,
                    row_info["is_special"],
                    answer=answer,
                    time_check=time_check,
                    gt_i=None,
                    rerank_query_data=rerank_query_data,
                    rerank_json_path=rerank_json_path,
                )
            ans_list.append(ans_true)
            true_docs_list.append(true_docs)
            topk_docs_list.append(topk_docs)
            all_proof_list.append(proof_list)

        yield ans_true, answer, true_docs_list, topk_docs_list, all_proof_list
        logger.info(f"填表答案:{ans_true}")
        j += 1

        # # !!!!!! 仅作为寻找合适的超参数使用,只取前X作为统计
        # if len(ans_list) % 20 == 0:
        #     break

    ### ---------存储rerank 的json文件 ---------###
    if time_check == 0:
        dir_part = CONFIG.json_rerank_folder_path
        # 保存为 JSON 文件
        rerank_json_path = f"{dir_part}/{name_word}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.json"
        with open(rerank_json_path, "w", encoding="utf-8") as f:
            json.dump(rerank_query_data, f, ensure_ascii=False, indent=2)
            # 删除目录中存储的旧的 JSON 文件
        files_to_delete = []
        for filename in os.listdir(dir_part):
            file_path = os.path.join(dir_part, filename)
            if (
                os.path.isfile(file_path)
                and filename.startswith(name_word)
                and file_path != rerank_json_path
            ):
                files_to_delete.append(file_path)

        for file_path in files_to_delete:
            try:
                os.remove(file_path)
            except Exception as e:
                raise RuntimeError(f"删除文件 {file_path} 失败: {e}")
    if gt:
        # 在需要的时候打印分析结果
        print_analysis(results)

    # logger.info(f"全部正样本的topk的相似度：total:{len(all_rerank_scores)},{all_rerank_scores}")
    # filtered_scores = [score for score in all_rerank_scores if score is not None]
    # logger.info(f"正样本最小的相似度分数：{min(filtered_scores)}") #未归一化正样本最小的相似度分数：-3.64453125,归一化：0.02546808298636756

    yield ans_true, answer, true_docs_list, topk_docs_list, all_proof_list


# 检查 ans.split("|")[2] 是否为一个可以转换为浮点数的数字，如果是，再进行转换；如果不是，跳过或者执行其他逻辑。
def try_convert_to_float(value):
    try:
        return float(value)
    except ValueError:
        return value  # 或者返回其他合适的默认值


# 获取关键词的上下文和llm的回答
@error_handler("大模型问答处理出错，请联系技术人员")
def query_llm(
    query_keyword,
    retriever,
    model,
    txt_time,
    chunks,
    elements_docx,
    name_word,
    row_info_is_special,
    answer,
    time_check=0,
    gt_i=None,
    rerank_query_data=None,
    rerank_json_path=None,
):
    """
    retriever文段提取关键词设置->rerank文段提取关键词设置->llm询问关键词设置
    Args:
        query_keyword (str): 表格中的关键词
        retriever (VectorStoreRetriever): 文件作为知识库
        rag_prompt(str): llm提示词
        model:llm模型
        gt_i(dict):gt值
        txt_time(str):txt写入文档时间
        name_word(str):文件名（水电站名字）
        time_check(int):复核次数
        row_info_is_special(bool):表格是否为特殊表格
    Returns:
        ans_true(str):最后填表的回答答案
        rerank_context(str): rerank_上下文
        rerank_scores_topk(list[float]):rerank前三得分
        ans_score(str):llm回答自评估得分
        context_relevant_docs:embedding和混合检索后的文档内容
        llm_context:llm的上下文

    """
    # 注：1）query_keyword直接传字符串会报错
    # 需要将query_keyword包装在一个lambda函数中。这样可以确保它被视为一个可调用的对象,而不是直接的字符串值。
    # 2）层级query_keyword，拼接后使用

    ## RAG流程：
    # 1.使用 retriever 获取相关文档（query设置）
    # ret_query_keyword_retriever=embed_prompt(query_keyword)
    ret_query_keyword_retriever = check_embed_prompt(
        query_keyword, row_info_is_special, answer
    )
    logger.info(f"Embedding检索提示词：{ret_query_keyword_retriever}")
    relevant_docs = retriever.get_relevant_documents(ret_query_keyword_retriever)
    # print(f"embedding一号：{relevant_docs[0].page_content}")

    # 混合检索：
    if CONFIG.EXTRACT_MODEL_HYBRID_SEARCH_SWITCH:
        relevant_docs = hybrid_search(
            relevant_docs, ret_query_keyword_retriever, top_k=CONFIG.EMBEDDING_TOP_k
        )

    # 过滤重复的chunk时，向量库检索的topk可以适当调大，去重后再取恰当的top—k是比较合理的，如config参数设置
    # SEARCH_KWARGS = {"k": 80, }
    # EMBEDDING_TOP_k = 60  # 真实embedding环节会输出的topk
    else:
        # relevant_docs = chunk_filter(relevant_docs)  # 过滤重复的chunk
        if len(relevant_docs) < CONFIG.EMBEDDING_TOP_k:
            logger.warning("混合检索并去重后的chucks总数少于topK，debug检查！！！")
        else:
            # zhy_7_13, 测完记得删除
            # test_docs = relevant_docs[: 60]
            # with open("zhy_test/test_docs.txt", "a", encoding="utf-8") as f:
            #     f.write(f"QUERY： {ret_query_keyword_retriever}\n")
            #     f.write(f"EMBEDDING_ANSWER {answer}\n")
            #     for i, doc in enumerate(test_docs):
            #         f.write(f"第{i}块： {doc.page_content.strip()}\n")
            relevant_docs = relevant_docs[: CONFIG.EMBEDDING_TOP_k]  #

    # embedding表格没有转成markdown格式：
    # context_relevant_docs = "\n".join(
    #     [f'第{i}块："{doc.page_content}"' for i, doc in enumerate(relevant_docs)]
    # )  # embedding和混合检索后的文档内容
    # embedding表格转成markdown格式：
    context_relevant_docs = []
    for i, doc in enumerate(relevant_docs):
        # 将搜索到的表格内容转成markdown格式
        if "category" in doc.metadata and doc.metadata["category"] == "table":
            # 将page_content转成markdown格式
            doc.page_content = doc.metadata["markdown_table"]


        embedding_context_content = f"第{i}块：{doc.page_content}"
        context_relevant_docs.append(embedding_context_content)
    # embedding_test_docs = relevant_docs[: 60]
    # with open("zhy_test/test_docs.txt", "a", encoding="utf-8") as f:
    #     f.write(f"QUERY： {ret_query_keyword_retriever}\n")
    #     f.write(f"EMBEDDING_ANSWER {answer}\n")
    #     for i, doc in enumerate(embedding_test_docs):
    #         f.write(f"第{i}块： {doc.page_content.strip()}\n")
    # 将所有文档内容合并成一个字符串
    context_relevant_docs = "\n".join(context_relevant_docs)
    filtered_docs, true_docs = filter_gt(relevant_docs, answer)  # RAG后处理，初步筛选
    logger.info(f"过滤掉的文档：{len(true_docs)}/{len(relevant_docs)}")
    logger.info(f"过滤后进入rerank的文档：{len(filtered_docs)}/{len(relevant_docs)}")
    # 扩充chunk的上下文
    # relevant_docs = expand_relevant_docs_text(relevant_docs, chunks)

    # 2.rerank询问字段：最后三级标题（query设置）以及doc中表格的page_content换成总结后的内容
    ## doc中表格的page_content换成总结后的内容(保证rerank公平)
    logger.info(f"Rerank模型名称：{CONFIG.RERANK_MODEL_PATH}")
    if CONFIG.EXTRACT_MODEL_SUMMARY_RERANK_TABLE_SWITCH:
        for i, doc in enumerate(filtered_docs):
            if doc.metadata.get("category") == "table":
                doc.page_content = doc.metadata["elements_table_summary"]

                # clean_path = doc.metadata['path'].replace("|", " ")[1:]
                # doc.page_content = f'索引：{clean_path}\n{doc.page_content}'
    else:
        for i, doc in enumerate(filtered_docs):
            if doc.metadata.get("category") == "table":
                doc.page_content = doc.metadata["elements_table_only_text"]

                # clean_path = doc.metadata['path'].replace("|", " ")[1:]
                # doc.page_content = f'索引：{clean_path}\n{doc.page_content}'
    rerank_top_k = CONFIG.CHECK_RERANK_TOP_K
    ## rerank的query设置：
    ret_query_keyword_rerank = check_rerank_prompt(
        query_keyword, row_info_is_special, name_word, answer
    )
    if CONFIG.EXTRACT_MODEL_RERANK_SWITCH:  # rerank打开

        # relevant_docs, rerank_scores_topk = rerank_docs(
        #     filtered_docs, ret_query_keyword_rerank, top_k=CONFIG.CHECK_RERANK_TOP_K,time_check=time_check
        # )
        # logger.info(f"Rerank检索提示词：{ret_query_keyword_rerank}")
        if time_check > 0 and rerank_query_data:
            relevant_docs_json = rerank_query_data.get(ret_query_keyword_rerank)
            if relevant_docs_json is None:
                if CONFIG.RRF_SWITCH:
                    relevant_docs, _ = rrf_docs(
                        filtered_docs,
                        ret_query_keyword_rerank,
                        top_k=CONFIG.CHECK_RERANK_TOP_K,
                        time_check=time_check,
                        k = CONFIG.RRF_K,
                    )
                else:
                    relevant_docs, _ = rerank_docs(
                        filtered_docs,
                        ret_query_keyword_rerank,
                        top_k=CONFIG.CHECK_RERANK_TOP_K,
                        time_check=time_check,
                    )
                
                rerank_data = []
                # 表格转成markdown格式：
                for _, doc in enumerate(relevant_docs):
                    if (
                        "category" in doc.metadata
                        and doc.metadata["category"] == "table"
                    ):
                        # 将page_content转成markdown格式
                        if "markdown_table" in doc.metadata:
                            doc.page_content = doc.metadata["markdown_table"]

                    elif "category" in doc.metadata and doc.metadata["category"] == "text":
                        clean_path = doc.metadata['path'].replace("|", " ")
                        # doc.page_content = f'{doc.page_content}\n以上文段的标题是：{clean_path}'
                        doc.page_content = f'索引：{clean_path}\n{doc.page_content}'

                    # 将Document对象转换为字典
                    rerank_data.append(
                        {"page_content": doc.page_content, "metadata": doc.metadata}
                    )
                rerank_query_data_i = {ret_query_keyword_rerank: rerank_data}
                rerank_query_data.update(rerank_query_data_i)
                with open(rerank_json_path, "w", encoding="utf-8") as f:
                    json.dump(rerank_query_data, f, ensure_ascii=False, indent=2)
            else:
                relevant_docs = [
                    Document(
                        page_content=data["page_content"], metadata=data["metadata"]
                    )
                    for data in relevant_docs_json
                ]
        else:
            if CONFIG.RRF_SWITCH:
                print("RRFING!!")
                relevant_docs, _ = rrf_docs(
                    filtered_docs,
                    ret_query_keyword_rerank,
                    top_k=CONFIG.CHECK_RERANK_TOP_K,
                    time_check=time_check,
                    k = CONFIG.RRF_K,
                )
            else:
                relevant_docs, _ = rerank_docs(
                    filtered_docs,
                    ret_query_keyword_rerank,
                    top_k=CONFIG.CHECK_RERANK_TOP_K,
                    time_check=time_check,
                )
            logger.info(f"Rerank检索提示词：{ret_query_keyword_rerank}")
            """
                context内容是否转成markdown格式
                见data_processor.py的extract_table_from_docx函数
            """

            rerank_data = []
            # 表格转成markdown格式：
            for _, doc in enumerate(relevant_docs):
                if "category" in doc.metadata and doc.metadata["category"] == "table":
                    # 将page_content转成markdown格式
                    if "markdown_table" in doc.metadata:
                        doc.page_content = doc.metadata["markdown_table"]
                # elif "category" in doc.metadata and doc.metadata["category"] == "text":
                    # clean_path = doc.metadata['path'].replace("|", "")
                    # doc.page_content = f'{doc.page_content}\n以上文段的标题是：{clean_path}'
                # 将Document对象转换为字典
                rerank_data.append(
                    {"page_content": doc.page_content, "metadata": doc.metadata}
                )
            rerank_query_data_i = {ret_query_keyword_rerank: rerank_data}
            rerank_query_data.update(rerank_query_data_i)
    # zhy
        test_docs = relevant_docs[: 60]
        with open("zhy_test/test_docs.txt", "a", encoding="utf-8") as f:
            f.write(f"QUERY： {ret_query_keyword_retriever}\n")
            f.write(f"RERANK_ANSWER {answer}\n")
            for i, doc in enumerate(test_docs):
                f.write(f"第{i}块： {doc.page_content.strip()}\n")

    else:  # rerank关闭
        rerank_scores_topk = []  # 跳过rerank需要加上

    if time_check * rerank_top_k < len(
        relevant_docs
    ) and time_check * rerank_top_k + rerank_top_k < len(relevant_docs):
        relevant_docs = relevant_docs[
            time_check * rerank_top_k : time_check * rerank_top_k + rerank_top_k
        ]
    elif time_check * rerank_top_k < len(relevant_docs):
        relevant_docs = relevant_docs[time_check * rerank_top_k :]
    else:
        relevant_docs = relevant_docs[:rerank_top_k]
    # # 表格没有转成markdown格式：
    rerank_scores_topk = []
    context = []
    tile = [] #标题
    # 删除“标题：xxx\n”
    # for i, doc in enumerate(relevant_docs):
    #     doc.page_content=re.sub(r'^标题：.*?\n', '', doc.page_content)

    for i, doc in enumerate(relevant_docs):

        rerank_docs_call = f"“{doc.page_content}”"
        tile_ = doc.metadata['path'].replace("|", " ")[0:]
        tile.append(tile_)
        context.append(rerank_docs_call)
        if "rerank_score" in doc.metadata:
            rerank_scores_topk.append(doc.metadata["rerank_score"])
        else:
            # 处理 rerank_score 不存在的情况，例如记录日志或使用默认值
            rerank_scores_topk.append(0)  # 假设使用默认值 0
    # 将所有文档内容合并成一个字符串
    # llm_context = "\n".join(context)
    llm_context = "\n".join([f'第{i}块："{doc}"' for i, doc in enumerate(context)])
    # for doc in context:
    #     if doc.metadata.get("category") == "text":
    #         clean_path = doc.metadata["path"].replace("|", " ")
    #         doc.page_content = f"{clean_path}\n{doc.page_content}"

    # llm_context = "\n".join(
    #     f'第{i}块："{doc.page_content}"'
    #     for i, doc in enumerate(context)
    # )

    rerank_context = llm_context
    # 3.大模型询问关键词：在“一级标题”中“二级标题”的“三级标题”
    ret_query_keyword_llm = check_llm_query_keyword(
        query_keyword, row_info_is_special, answer
    )
    proof_list = []
    # 跳过模型问答，用于测试rag
    if CONFIG.EXTRACT_MODEL_LLM_SWITCH == 0:
        ans_true = f"{ret_query_keyword_llm}|无|无"
        # ans_true = f"{ret_query_keyword_llm}" + random.choice(debug_ans)
        ans_score = 0
    elif CONFIG.EXTRACT_MODEL_LLM_SWITCH == 1:
        if CONFIG.SPLIT_LLM_INVOKE != 1:
            # 提示词工程
            rag_prompt = check_prompt(context,tile, ret_query_keyword_llm, answer)
            # 没有自评估得分,需要改prompt
            rag_chain = (
                {
                    "context": lambda _: context,
                    "tile":lambda _: tile,
                    "query_keyword": lambda _: ret_query_keyword_llm,
                }
                | rag_prompt
                | model
                | StrOutputParser()  # 分析器:用以分割"调表答案"和"llm回答得分"
            )
            ans = rag_chain.invoke({})
            if "deepseek" in CONFIG.LLM_MODEL:
                logger.info(f"删除<think>标签之前的内容：{ans}")
                # 使用正则表达式删除 <think>...</think> 之间的内容
                ans = re.sub(r"<think>.*?</think>\s*", "", ans, flags=re.DOTALL)
                logger.info(f"删除<think>标签之后的内容：{ans}")
            else:
                logger.info(f"大模型回答的内容：{ans}")
            ans_true = ans
            ans_score = 0
        else:
            # !topK分开询问
            ans = f"{ret_query_keyword_llm}"
            for one_context,one_tile in zip(context,tile):
                # 更新会话 ollama
                model = get_llm_model(model_type="ollama", query=query_keyword)
                rag_prompt = splited_check_prompt(
                    one_context,
                    one_tile,
                    ret_query_keyword_llm,
                    answer,
                    self_consistency=CONFIG.PROMPT_SELF_CONSISTENCY,
                )
                rag_chain = (
                    {
                        "context": lambda _: context,
                        "tile": lambda _:one_tile,
                        "query_keyword": lambda _: ret_query_keyword_llm,
                    }
                    | rag_prompt
                    | model
                    | StrOutputParser()
                )
                one_ans = rag_chain.invoke({})
                if CONFIG.PROMPT_SELF_CONSISTENCY:
                    match = re.search(r"\[最终判定\]\s*(正确|错误)", one_ans)
                    one_ans = match.group(1) if match else "正确"
                logger.info(f"大模型回答的内容：{one_ans}")
                # if int(one_ans.split("|")[1]) < 70:
                #     one_ans = "正确"
                # else:
                #     one_ans = one_ans.split("|")[0]
                if "正确" in one_ans:
                    ans += "|正确"
                else:
                    ans += "|错误"
                if CONFIG.CHECK_MODEL_SOURCE_SWITCH == 1:
                    if "正确" in one_ans:
                        proof_list.append("")
                    else:
                        if len(one_ans.split("|")) > 1:
                            proof_list.append(one_ans.split("|")[1])
                        else:
                            proof_list.append("")

            logger.info(f"大模型回答的内容-拼接后：{ans}")
            logger.info(f"大模型回答的内容-proof：{proof_list}")
            ans_true = ans
            ans_score = 0
    else:
        # 提示词工程
        rag_prompt = check_prompt(context, ret_query_keyword_llm, answer)
        # 自评估得分,需要改prompt，高于5，4分的认为回答正确，低于4分的认为回答错误(输出格式:稻谷库存|吨|10000###1)
        # 输出格式中:稻谷库存|吨|10000:str(填表数据),1:str(llm大模型回答得分),"###"用于分割调表数据和llm回答得分
        rag_chain = (
            {
                "context": lambda _: context,
                "query_keyword": lambda _: ret_query_keyword_llm,
            }
            | rag_prompt
            | model
            | MultiOutputParser()  # 分析器:用以分割"调表答案"和"llm回答得分"
        )
        ans = rag_chain.invoke({})
        if "5" in ans["score"] or "4" in ans["score"]:
            ans_true = ans["result"]
        else:
            ans_true = f"{query_keyword}|无|无"
        ans_score = ans["score"]

    ### 测试系统正确率的时间使用
    # 定义保存文件的路径
    if gt_i:
        output_file = (
            f"./output_excels/test_txt/{name_word}_rerank_context_output_{txt_time}.txt"
        )
        # 将 context 写入文件,将检索到的上下文写入context_output.txt
        with open(output_file, "a", encoding="utf-8") as f:
            f.write(f" retriever的询问的字段:{ret_query_keyword_retriever}" + "\n")
            f.write(f" rerank的询问的字段:{ret_query_keyword_rerank}" + "\n")
            f.write(f" LLM的询问的字段:{ret_query_keyword_llm}" + "\n")
            f.write(
                f"ground truth value:{gt_i['value']}"
                + "\t\t"
                + f"ground truth flag:{gt_i['flag']}"
                + "\n"
                + f"ground truth modified answer:{gt_i['modify_value']}"
            )
            f.write(
                "\n".join(
                    f"第{i}块的rerank得分：{score}"
                    for i, score in enumerate(rerank_scores_topk)
                )
                + "\n"
            )
            f.write((f"处理前回答：{ans_true};得分：{ans_score}") + "\n")
            f.write((f"自评估筛选后的回答：{ans_true}") + "\n")
            f.write((f"llm_context(rerank后表格转成markdown格式):") + "\n")
            f.write(llm_context + "\n\n")  # rerank_context的内容
            f.write((f"rerank_context(rerank的表格块是大模型总结格式):") + "\n")
            f.write(rerank_context + "\n\n")  # rerank_context的内容
            f.write((f"embedding_context:") + "\n")
            f.write(context_relevant_docs + "\n\n\n")  # 换行区分不同qurey写入上下文
        logger.info(
            f" retriever,rerank,llm询问关键词以及llm回答以及文本已经写入：{output_file}"
        )

    # --- Extracting query and pos ---
    # # Step 1: 处理 query，去掉其中的 '|'
    # ret_query_keyword_retriever = query_keyword.replace("|", "")  # 相似度搜索，用“”连接

    # # # Step 2: 使用 retriever 获取相关文档
    # relevant_docs = retriever.get_relevant_documents(ret_query_keyword_retriever)

    # # # Step 3: 重排序 rerank 文档
    # ret_query_keyword_rerank = query_keyword.split("|")[-3:]
    # ret_query_keyword_rerank = "的".join(ret_query_keyword_rerank)

    # # # 对文档进行重排序，获取重新排序的文档和得分
    # relevant_docs, rerank_scores_topk = rerank_docs(relevant_docs, ret_query_keyword_rerank)

    # # # Step 4: 生成 context (pos) 为相关文档内容的拼接
    # context = "\n".join([f'第{i}块："{doc.page_content}"' for i, doc in enumerate(relevant_docs)])

    # # # Step 5: 创建一个列表来存储 query 和 pos 样本
    # samples = []
    # for doc in relevant_docs:
    #     query = ret_query_keyword_retriever  # query 为传入的关键词
    #     pos = doc.page_content  # pos 为相关文档的内容
    #     samples.append({"query": query, "pos": pos})

    # # # Step 6: 将样本数据保存到 Pandas DataFrame
    # df_samples = pd.DataFrame(samples)

    # # # Step 7: 将生成的 DataFrame 输出为 CSV 文件（如果需要保存）
    # print("------------------------------------------------")
    # print("write file")
    # output_file_samples = f"./output_excels/test_txt/query_pos_samples_txt_time.csv"
    # file_exists = os.path.exists(output_file_samples)
    # first_sample = df_samples.iloc[:1]
    # first_sample.to_csv(output_file_samples, mode='a', header=not file_exists, index=False, encoding="utf-8")
    # print(df_samples)
    return (
        ans_true,
        rerank_context,
        rerank_scores_topk,
        ans_score,
        context_relevant_docs,
        llm_context,
        true_docs,
        relevant_docs,
        proof_list,
        rerank_query_data,
    )


# 统计正负样本的正确率
def collect_results(
    keyword,
    rerank_context,
    context_relevant_docs,
    ans,
    gt,
    is_special,
    ans_score,
    llm_context,
):
    """
    收集单条数据的匹配结果

    Args:
        keyword:查询关键词
        rerank_context (str): rerank后的上下文文本
        context_relevant_docs:embeding后的上下文
        ans (str): 当前答案
        gt (dict): GroundTruth中的gt值
        is_special(bool):文段还是单位和数量
        ans_score(str):llm回答的得分
        llm_context:llm的上下文
    Returns:
        dict: 当前记录的匹配结果
    """
    current_result = {
        "keyword": keyword,
        "ground_truth": gt["value"],
        "modify_value": gt["modify_value"],
        "ground_truth_flag": gt["flag"],
        "embedding_context_match": False,  # embedding的上下文是否匹配
        "rerank_context_match": False,  # rerank的上下文是否匹配
        "modify_value_embedding_context_match": False,  # modify_value_embedding的上下文是否匹配
        "modify_value_rerank_context_match": False,  # modify_value_rerank的上下文是否匹配
        "positive_context_match": False,  # 正样本llm的上下文是否匹配
        "positive_answer_match": False,  # 正样本llm的回答是否匹配
        "zero_context_match": False,
        "zero_answer_match": False,
        "negative_answer_match": False,
        "reank_context": rerank_context,
        "context_relevant_docs": context_relevant_docs,
        "llm_context": llm_context,
        "answer": (
            (ans.split("|")[1] if is_special else ans.split("|")[2])
            if len(ans.split("|")) >= 2
            else "N/A"
        ),
        "ans_score": ans_score,
    }
    if len(ans.split("|")) >= 2:
        # 正样本检查context匹配 和 检查answer匹配
        if current_result["ground_truth_flag"] == 1:
            if gt["value"] in context_relevant_docs:
                current_result["embedding_context_match"] = True
            if gt["value"] in rerank_context:
                current_result["rerank_context_match"] = True
            if gt["value"] in llm_context:
                current_result["positive_context_match"] = True
            if gt["modify_value"] in context_relevant_docs:
                current_result["modify_value_embedding_context_match"] = True
            if gt["modify_value"] in llm_context:
                current_result["modify_value_rerank_context_match"] = True
            # if is_special:
            #     if gt["value"] in ans.split("|")[1]:
            #         current_result["positive_answer_match"] = True
            # else:
            #     if gt["value"] in ans.split("|")[2] or try_convert_to_float(
            #         gt["value"]
            #     ) == try_convert_to_float(ans.split("|")[2]):
            #         current_result["positive_answer_match"] = True

        # # 零样本检查context匹配 和 检查answer匹配
        # elif current_result["ground_truth_flag"] == 0:
        #     if gt["value"] in llm_context:
        #         current_result["zero_context_match"] = True
        #     if is_special:
        #         if gt["value"] in ans.split("|")[1]:
        #             current_result["zero_answer_match"] = True
        # else:
        #     if gt["value"] in ans.split("|")[2] or try_convert_to_float(
        #         gt["value"]
        #     ) == try_convert_to_float(ans.split("|")[2]):
        #         current_result["zero_answer_match"] = True
        # # 负样本检查answer是否为无
        # else:
        #     if is_special:
        #         if "无" in ans.split("|")[1]:
        #             current_result["negative_answer_match"] = True
        #     else:
        #         if "无" in ans.split("|")[2] and "无" in ans.split("|")[1]:
        #             current_result["negative_answer_match"] = True

    return current_result


def print_analysis(results):
    """
    打印分析结果

    Args:
        results (list): 包含多条collect_results返回结果的列表
    """
    # 计算正样本准确率
    positive_total_count = sum(1 for r in results if r["ground_truth_flag"] == 1)
    positive_context_acc_count = sum(1 for r in results if r["positive_context_match"])
    positive_embedding_context_acc_count = sum(
        1 for r in results if r["embedding_context_match"]
    )
    positive_rerank_context_acc_count = sum(
        1 for r in results if r["rerank_context_match"]
    )
    modify_value_positive_embedding_context_acc_count = sum(
        1 for r in results if r["modify_value_embedding_context_match"]
    )
    modify_value_positive_rerank_context_acc_count = sum(
        1 for r in results if r["modify_value_rerank_context_match"]
    )

    # positive_ans_acc_count = sum(1 for r in results if r["positive_answer_match"])
    positive_context_acc = (
        positive_context_acc_count / positive_total_count
        if positive_total_count != 0
        else 0
    )
    positive_embedding_context_acc = (
        positive_embedding_context_acc_count / positive_total_count
        if positive_total_count != 0
        else 0
    )
    positive_rerank_context_acc = (
        positive_rerank_context_acc_count / positive_total_count
        if positive_total_count != 0
        else 0
    )
    modify_value_positive_embedding_context_acc = (
        modify_value_positive_embedding_context_acc_count / positive_total_count
        if positive_total_count != 0
        else 0
    )
    modify_value_positive_rerank_context_acc = (
        modify_value_positive_rerank_context_acc_count / positive_total_count
        if positive_total_count != 0
        else 0
    )

    # positive_ans_acc = (
    #     positive_ans_acc_count / positive_total_count
    #     if positive_total_count != 0
    #     else 0
    # )
    # 计算零样本准确率
    # zero_total_count = sum(1 for r in results if r["ground_truth_flag"] == 0)
    # zero_context_acc_count = sum(1 for r in results if r["zero_context_match"])
    # zero_ans_acc_count = sum(1 for r in results if r["zero_answer_match"])
    # zero_context_acc = (
    #     zero_context_acc_count / zero_total_count if zero_total_count != 0 else 0
    # )
    # zero_ans_acc = zero_ans_acc_count / zero_total_count if zero_total_count != 0 else 0
    # # 计算负样本准确率
    # negative_total_count = sum(1 for r in results if r["ground_truth_flag"] == -1)
    # negative_ans_acc_count = sum(1 for r in results if r["negative_answer_match"])
    # negative_ans_acc = (
    #     negative_ans_acc_count / negative_total_count
    #     if negative_total_count != 0
    #     else 0
    # )
    # 打印案例
    logger.info("\n=== Incorrect Cases ===")
    for idx, result in enumerate(results):
        # embeding的上下文是否匹配
        if not result["embedding_context_match"] and result["ground_truth_flag"] == 1:
            logger.info(
                f"\nRecord {idx + 1}:"
                + "\n"
                + f"Keyword: {result['keyword']}"
                + "\n"
                + f"Ground Truth: {result['ground_truth']}"
                + "\n"
                + f"Embedding Context Match: {'?' if result['embedding_context_match'] else '?'}"
            )
        # rerank的上下文是否匹配
        if not result["rerank_context_match"] and result["ground_truth_flag"] == 1:
            logger.info(
                f"\nRecord {idx + 1}:"
                + "\n"
                + f"Keyword: {result['keyword']}"
                + "\n"
                + f"Ground Truth: {result['ground_truth']}"
                + "\n"
                + f"Rerank Context Match: {'?' if result['rerank_context_match'] else '?'}"
            )
        # 正样本没找到对应的context和answer
        if not result["positive_context_match"] and result["ground_truth_flag"] == 1:
            logger.info(
                f"\nRecord {idx + 1}:"
                + "\n"
                + f"Keyword: {result['keyword']}"
                + "\n"
                + f"Ground Truth: {result['ground_truth']}"
                + "\n"
                + f"Rerank Context Match: {'?' if result['positive_context_match'] else '?'}"
            )
        if (
            not result["modify_value_embedding_context_match"]
            and result["ground_truth_flag"] == 1
        ):
            logger.info(
                f"\nRecord {idx + 1}:"
                + "\n"
                + f"Keyword: {result['keyword']}"
                + "\n"
                + f"Ground Truth: {result['ground_truth']}"
                + "\n"
                + f"Modify Value:{result['modify_value']}"
                + "\n"
                + f"Modify Value Embedding Context Match: {'?' if result['modify_value_embedding_context_match'] else '?'}"
            )
        if (
            not result["modify_value_rerank_context_match"]
            and result["ground_truth_flag"] == 1
        ):
            logger.info(
                f"\nRecord {idx + 1}:"
                + "\n"
                + f"Keyword: {result['keyword']}"
                + "\n"
                + f"Ground Truth: {result['ground_truth']}"
                + "\n"
                + f"Modify Value:{result['modify_value']}"
                + "\n"
                + f"Modify Value Rerank Context Match: {'?' if result['modify_value_rerank_context_match'] else '?'}"
            )
        # if not result["positive_answer_match"] and result["ground_truth_flag"] == 1:
        #     logger.info(
        #         f"\nRecord answer {idx + 1}:"
        #         + "\n"
        #         + f"Keyword answer: {result['keyword']}"
        #         + "\n"
        #         + "Answer answer  (no match found):"
        #         + "\n"
        #         + f"{result['answer'][:200] if len(result['answer']) > 200 else result['answer']}"
        #         + "\n"
        #         + f"Answer score:{ result['ans_score']}"
        #     )
        # 零样本没找到对应的context和answer
        # if not result['zero_context_match'] and result["ground_truth_flag"]==0:
        #     logger.info(f"\nRecord {idx + 1}:"+'\n'
        #                 +f"Keyword: {result['keyword']}"+'\n'
        #                 +f"Ground Truth: {result['ground_truth']}"+'\n'
        #                 +f"zero_context_match: {'?' if result['zero_context_match'] else '?'}"
        #                 )
        # if not result['zero_answer_match'] and result["ground_truth_flag"]==0:
        #     logger.info(f"\nRecord {idx + 1}:"+'\n'
        #                 +f"Keyword: {result['keyword']}"+'\n'
        #                 +f"Answer (zero_answer_match no found):"+'\n'
        #                 +result['answer'][:200] + "..." if len(result['answer']) > 200 else result['answer']
        #                 )

        # 负样本没有找到对应的answer
        # if not result["negative_answer_match"] and result["ground_truth_flag"] == -1:
        #     logger.info(
        #         f"\nRecord {idx + 1}:"
        #         + "\n"
        #         + f"Keyword: {result['keyword']}"
        #         + "\n"
        #         + "LLM hallucinations:"
        #         + "\n"
        #         + f"{result['answer'][:200] if len(result['answer']) > 200 else result['answer']}"
        #         + "\n"
        #         + f"Answer score:{ result['ans_score']}"
        #     )
    # 打印总体统计,填表长度
    logger.info(
        "\n=== Overall Statistics ===" + f"Results Total Records: {len(results)}"
    )

    logger.info(
        "\n=== Overall positive Statistics ==="
        + "\n"
        + f"Total Records: {positive_total_count}"
        + "\n"
        + f"Embedding Context Accuracy: {positive_embedding_context_acc:.2%} ({positive_embedding_context_acc_count}/{positive_total_count})"
        + "\n"
        + f"Rerank Context Accuracy(rerank的表格块是大模型总结格式): {positive_rerank_context_acc:.2%} ({positive_rerank_context_acc_count}/{positive_total_count})"
        + "\n"
        + f"LLM Context Accuracy(rerank后表格转成markdown格式): {positive_context_acc:.2%} ({positive_context_acc_count}/{positive_total_count})"
        + "\n"
        + f"Modify Value Embedding Context Accuracy: {modify_value_positive_embedding_context_acc:.2%} ({modify_value_positive_embedding_context_acc_count}/{positive_total_count})"
        + "\n"
        + f"Modify Value Rerank Context Accuracy: {modify_value_positive_rerank_context_acc:.2%} ({modify_value_positive_rerank_context_acc_count}/{positive_total_count})"
        # + "\n"
        # + f"Answer Accuracy: {positive_ans_acc:.2%} ({positive_ans_acc_count}/{positive_total_count})"
    )

    # logger.info("\n=== Overall zero Statistics ==="+'\n'
    #             +f"Total Records: {zero_total_count}"+'\n'
    #             +f"Context Accuracy: {zero_context_acc:.2%} ({zero_context_acc_count}/{zero_total_count})"+'\n'
    #             +f"Answer Accuracy: {zero_ans_acc:.2%} ({zero_ans_acc_count}/{zero_total_count})"
    #             )
    # logger.info(
    #     "\n=== Overall negative Statistics ==="
    #     + "\n"
    #     + f"Total Records: {negative_total_count}"
    #     + "\n"
    #     + f"Answer Accuracy: {negative_ans_acc:.2%} ({negative_ans_acc_count}/{negative_total_count})"
    # )


if __name__ == "__main__":
    file_path = r"./data/LushanReport.docx"
    # doc_text, pydoc = read_docx_file(file_path)
    # query_info, tbl_info = get_appendix(pydoc)
    # print(query_info)
    # print(tbl_info)
    # print_docx_content(document_content)
    pages = [
        Document(
            page_content="hello world, this is a test document",
            metadata={"chunk_id": 1},
        ),
        Document(
            page_content="another test chunk for demonstration",
            metadata={"chunk_id": 2},
        ),
    ]
