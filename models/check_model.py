from langchain_community.vectorstores import Chroma
from langchain_core.output_parsers import StrOutputParser
from utils.utils import filter_gt
from config import CONFIG
from models.new_rerank_model import rerank_docs
from models.prompt_engineering import (
    check_embed_prompt,
    check_rerank_prompt,
    llm_query_keyword,
    rerank_prompt,
    check_prompt,
)
from tqdm import tqdm
from utils.logger_util import logger
import threading
from models.llm_model import get_llm_model
import random
from utils.utils import error_handler
import shutil
from models.embedding_model import get_embedding_model
from chromadb.config import Settings
import os


@error_handler("大模型问答处理出错，请联系技术人员")
def call_check_model(tbl_info, retriever, name_word, progress=None, gt=None):
    ans_list = []
    results = []
    j = 0  # gt的值遍历
    true_docs_list = []
    topk_docs_list = []
    progress_bar = tqdm if progress is None else progress.tqdm
    for i in progress_bar(range(len(tbl_info["need_query"]))):
        if not tbl_info["need_query"][i]:
            continue
        row_info = {
            "need_query": tbl_info["need_query"][i],
            "is_special": tbl_info["is_special"][i],
            "keywords": tbl_info["keywords"][i],
            "answers": tbl_info["answers"][i],  # 单位|数量  or  描述性的短句
        }
        query_keyword = row_info["keywords"]
        answer = row_info["answers"]

        # 更新会话
        model = get_llm_model(model_type="ollama")
        if gt and (gt[j]["flag"] == 1):
            ans, true_docs, topk_docs = query_llm(
                query_keyword,
                retriever,
                name_word,
                row_info["is_special"],
                model,
                answer,
                gt,
            )
            ans_list.append(ans)
            true_docs_list.append(true_docs)
            topk_docs_list.append(topk_docs)
            logger.info(f"LLM回答:{ans}")
        j += 1
        yield ans, [], []

        # !!!!!!!!!!!!!!!
        # if len(ans_list) % 10 == 0:
        #     break

    yield ans_list, true_docs_list, topk_docs_list


# 获取关键词的上下文和llm的回答
def query_llm(
    query_keyword, retriever, name_word, row_info_is_special, model, answer, gt=None
):
    logger.info(f"原始提取出的查询关键词：{query_keyword}")
    ret_query_keyword_retriever = check_embed_prompt(
        query_keyword, row_info_is_special, answer
    )

    logger.info(f"RAG关键词：{ret_query_keyword_retriever}")

    logger.info("RAG开始")
    relevant_docs = retriever.get_relevant_documents(ret_query_keyword_retriever)

    filtered_docs, true_docs = filter_gt(relevant_docs, answer)  # RAG后处理，初步筛选
    logger.info(f"过滤掉的文档：{len(true_docs)}/{len(relevant_docs)}")
    logger.info(f"过滤后进入rerank的文档：{len(filtered_docs)}/{len(relevant_docs)}")

    ret_rerank_keyword = check_rerank_prompt(
        query_keyword, row_info_is_special, name_word, answer
    )
    logger.info(f"RERANK关键词：{ret_query_keyword}")

    logger.info("Rerank开始")
    relevant_docs, relevance_scores_topk = rerank_docs(
        filtered_docs, ret_rerank_keyword, top_k=CONFIG.CHECK_RERANK_TOP_K
    )

    logger.info(f"Rerank结束，相关文档数量：{len(relevant_docs)}")
    context = [doc.page_content for doc in relevant_docs]

    ret_query_keyword = query_keyword.split("|")[1:]
    ret_query_keyword = "的".join(ret_query_keyword)  # 相似度搜索的没有第一级标题
    ret_query_keyword = (
        "在" + query_keyword.split("|")[0] + "中" + ret_query_keyword
    )  # 相似度搜索的没有第一级标题
    logger.info(f"LLM查询字段：{ret_query_keyword}")

    rag_prompt = check_prompt(context, ret_query_keyword, answer)

    rag_chain = {} | rag_prompt | model | StrOutputParser()

    ans = rag_chain.invoke({})
    # ans = f"{ret_query_keyword}"

    return ans, true_docs, relevant_docs


# 检查 ans.split("|")[2] 是否为一个可以转换为浮点数的数字，如果是，再进行转换；如果不是，跳过或者执行其他逻辑。
def try_convert_to_float(value):
    try:
        return float(value)
    except ValueError:
        return value  # 或者返回其他合适的默认值


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
        "ground_truth_flag": gt["flag"],
        "embedding_context_match": False,  # embedding的上下文是否匹配
        "rerank_context_match": False,  # rerank的上下文是否匹配
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
            if is_special:
                if gt["value"] in ans.split("|")[1]:
                    current_result["positive_answer_match"] = True
            else:
                if gt["value"] in ans.split("|")[2] or try_convert_to_float(
                    gt["value"]
                ) == try_convert_to_float(ans.split("|")[2]):
                    current_result["positive_answer_match"] = True

        # 零样本检查context匹配 和 检查answer匹配
        elif current_result["ground_truth_flag"] == 0:
            if gt["value"] in llm_context:
                current_result["zero_context_match"] = True
            if is_special:
                if gt["value"] in ans.split("|")[1]:
                    current_result["zero_answer_match"] = True
            else:
                if gt["value"] in ans.split("|")[2] or try_convert_to_float(
                    gt["value"]
                ) == try_convert_to_float(ans.split("|")[2]):
                    current_result["zero_answer_match"] = True
        # 负样本检查answer是否为无
        else:
            if is_special:
                if "无" in ans.split("|")[1]:
                    current_result["negative_answer_match"] = True
            else:
                if "无" in ans.split("|")[2] and "无" in ans.split("|")[1]:
                    current_result["negative_answer_match"] = True

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
    positive_ans_acc_count = sum(1 for r in results if r["positive_answer_match"])
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
    positive_ans_acc = (
        positive_ans_acc_count / positive_total_count
        if positive_total_count != 0
        else 0
    )
    # 计算零样本准确率
    zero_total_count = sum(1 for r in results if r["ground_truth_flag"] == 0)
    zero_context_acc_count = sum(1 for r in results if r["zero_context_match"])
    zero_ans_acc_count = sum(1 for r in results if r["zero_answer_match"])
    zero_context_acc = (
        zero_context_acc_count / zero_total_count if zero_total_count != 0 else 0
    )
    zero_ans_acc = zero_ans_acc_count / zero_total_count if zero_total_count != 0 else 0
    # 计算负样本准确率
    negative_total_count = sum(1 for r in results if r["ground_truth_flag"] == -1)
    negative_ans_acc_count = sum(1 for r in results if r["negative_answer_match"])
    negative_ans_acc = (
        negative_ans_acc_count / negative_total_count
        if negative_total_count != 0
        else 0
    )
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
                + f"Embedding Context Match: {'✓' if result['embedding_context_match'] else '✗'}"
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
                + f"Rerank Context Match: {'✓' if result['rerank_context_match'] else '✗'}"
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
                + f"Rerank Context Match: {'✓' if result['positive_context_match'] else '✗'}"
            )

        if not result["positive_answer_match"] and result["ground_truth_flag"] == 1:
            logger.info(
                f"\nRecord answer {idx + 1}:"
                + "\n"
                + f"Keyword answer: {result['keyword']}"
                + "\n"
                + "Answer answer  (no match found):"
                + "\n"
                + f"{result['answer'][:200] if len(result['answer']) > 200 else result['answer']}"
                + "\n"
                + f"Answer score:{ result['ans_score']}"
            )
        # 零样本没找到对应的context和answer
        # if not result['zero_context_match'] and result["ground_truth_flag"]==0:
        #     logger.info(f"\nRecord {idx + 1}:"+'\n'
        #                 +f"Keyword: {result['keyword']}"+'\n'
        #                 +f"Ground Truth: {result['ground_truth']}"+'\n'
        #                 +f"zero_context_match: {'✓' if result['zero_context_match'] else '✗'}"
        #                 )
        # if not result['zero_answer_match'] and result["ground_truth_flag"]==0:
        #     logger.info(f"\nRecord {idx + 1}:"+'\n'
        #                 +f"Keyword: {result['keyword']}"+'\n'
        #                 +f"Answer (zero_answer_match no found):"+'\n'
        #                 +result['answer'][:200] + "..." if len(result['answer']) > 200 else result['answer']
        #                 )

        # 负样本没有找到对应的answer
        if not result["negative_answer_match"] and result["ground_truth_flag"] == -1:
            logger.info(
                f"\nRecord {idx + 1}:"
                + "\n"
                + f"Keyword: {result['keyword']}"
                + "\n"
                + "LLM hallucinations:"
                + "\n"
                + f"{result['answer'][:200] if len(result['answer']) > 200 else result['answer']}"
                + "\n"
                + f"Answer score:{ result['ans_score']}"
            )
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
        + f"Answer Accuracy: {positive_ans_acc:.2%} ({positive_ans_acc_count}/{positive_total_count})"
    )

    # logger.info("\n=== Overall zero Statistics ==="+'\n'
    #             +f"Total Records: {zero_total_count}"+'\n'
    #             +f"Context Accuracy: {zero_context_acc:.2%} ({zero_context_acc_count}/{zero_total_count})"+'\n'
    #             +f"Answer Accuracy: {zero_ans_acc:.2%} ({zero_ans_acc_count}/{zero_total_count})"
    #             )
    logger.info(
        "\n=== Overall negative Statistics ==="
        + "\n"
        + f"Total Records: {negative_total_count}"
        + "\n"
        + f"Answer Accuracy: {negative_ans_acc:.2%} ({negative_ans_acc_count}/{negative_total_count})"
    )


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
