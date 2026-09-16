"""
混合检索，彭柯然-2025-1-9
输入：一组文档chucks，一个查询query，top_k,混合比例ratio
输出：top_k个相关文档及其相关度，输出的chuck直接进入rerank环节

补充：
输入为embedding后选出的chucks，中间包含去重处理
ratio为向量检索的比例，[0.0, 1.0]

"""

from typing import List
from langchain.schema import Document
from rank_bm25 import BM25Okapi
import torch
from utils.logger_util import logger
import nltk
from pathlib import Path
import jieba
from utils.utils import chunk_filter
from config import CONFIG
from chromadb.config import Settings
import shutil
import os
import torch
from models.embedding_model import get_embedding_model
from langchain_community.vectorstores import Chroma
from langchain.retrievers.ensemble import EnsembleRetriever


# Improved text preprocessing function
def preprocess_text(text: str) -> List[str]:
    # tokens = word_tokenize(text.lower())
    # stop_words = set(stopwords.words("english"))
    # return [token for token in tokens if token.isalnum() and token not in stop_words]

    tokens = list(jieba.cut(text))  # jieba.cut 返回一个生成器，这里转换成列表
    return [token for token in tokens if token.isalnum()]


def hybrid_search(chucks: List[Document], query: str, top_k: int, ratio: float = 0.5):
    """
    输入：chucks:embedding模型返回的文档列表
    """
    if len(chucks) < top_k:
        logger.warning("混合检索部分，因输入的chucks总数少于topK，不处理直接返回")
        return chucks

    # 防止重复过多，保证能输出不重复的topk，多取几个
    num = len(chucks)

    original_documents = [chuck.page_content for chuck in chucks]

    # token化处理
    tokenized_corpus = [preprocess_text(doc) for doc in original_documents]
    tokenized_query = preprocess_text(query)

    # 计算bm25分数
    bm25 = BM25Okapi(tokenized_corpus)
    bm25_scores = torch.tensor(bm25.get_scores(tokenized_query))

    # 排序取topk及其分数
    res = sorted(
        [(i, score.item()) for i, score in enumerate(bm25_scores)],
        key=lambda x: x[1],
        reverse=True,
    )

    # 稀疏检索（关键词检索）部分的输出
    bm25_res_chucks = [chucks[i] for i, _ in res]
    bm25_res_scores = [score for _, score in res]

    # 与向量相似度检索的结果（chucks）混合
    vector_num = int(ratio * num)
    bm25_num = num - vector_num

    # 初始化交叉插入的结果列表
    mixed_chucks = []

    # 交替插入两个列表
    similarity_ratio = CONFIG.SIMILARITY_RATIO

    for i in range(len(chucks)):
        start_idx = i * similarity_ratio
        end_idx = min(start_idx + similarity_ratio, len(chucks))
        mixed_chucks.extend(chucks[start_idx:end_idx])

        mixed_chucks.append(bm25_res_chucks[i])

    # 如果有剩余，分别添加到混合列表中
    # mixed_chucks.extend(chucks[min(vector_num, bm25_num) : vector_num])
    # mixed_chucks.extend(bm25_res_chucks[min(vector_num, bm25_num) : bm25_num])

    # 去重
    mixed_chucks = chunk_filter(mixed_chucks)

    if len(mixed_chucks) < top_k:
        logger.warning("混合检索并去重后的chucks总数少于topK，debug检查！！！")
        return mixed_chucks

    return mixed_chucks[:top_k]


def getEmbeddingRetriever(pages):
    embeddings_model = get_embedding_model(
        # model_name=CONFIG.EMBED_MODEL, adapter_path=CONFIG.ADAPTER_PATH
        model_name=CONFIG.EMBED_MODEL, adapter_path=CONFIG.ADAPTER_PATH, model_path=CONFIG.EMBEDDING_MODEL_PATH, peft_lora_adapter_path=CONFIG.PEFT_LORA_ADAPTER_PATH
    )
    persist_directory = "./chroma_temp_dir"
    if os.path.isdir(persist_directory):
        shutil.rmtree(persist_directory)

    chroma_settings = Settings(
        is_persistent=True,
        persist_directory=persist_directory,  # 指定临时目录
    )

    db = Chroma(
        collection_name="langchain_store",
        embedding_function=embeddings_model,
        client_settings=chroma_settings,
        persist_directory=persist_directory,
    )

    db.add_documents(documents=pages, embedding=embeddings_model)
    db.persist()
    retriever = db.as_retriever(
        search_type=CONFIG.SEARCH_TYPE, search_kwargs=CONFIG.SEARCH_KWARGS
    )


if __name__ == "__main__":
    # 创建一些模拟文档
    chunks = [
        Document(page_content="我喜欢自然语言处理和机器学习"),
        Document(page_content="我喜欢自然语言处理，但也对深度学习感兴趣"),
        Document(page_content="我喜欢机器学习和自然语言处理技术"),
        Document(page_content="Python是非常强大的编程语言"),
        Document(page_content="自然语言处理技术是人工智能的重要组成部分"),
        Document(page_content="数据科学家通常使用Python进行数据分析"),
        Document(page_content="深度学习是机器学习的一个重要分支"),
        Document(page_content="Python是一门非常强大的编程语言，用途广泛"),
        Document(page_content="人工智能的重要组成部分是自然语言处理技术"),
        Document(page_content="数据科学家经常用Python进行数据分析和可视化"),
        Document(page_content="机器学习的一个重要分支是深度学习"),
        Document(page_content="Python是一种强大的编程语言，常用于AI开发"),
        Document(page_content="自然语言处理技术在人工智能中占据重要地位"),
        Document(page_content="使用Python进行数据分析是数据科学家的常见任务"),
        Document(page_content="深度学习是AI中的重要领域之一"),
    ]

    embeddingRetriever = getEmbeddingRetriever(chunks)

    ensembel = EnsembleRetriever()
    # 定义查询
    query = "数据科学家"

    # 定义 top_k 参数
    top_k = 3

    # 混合比例 ratio
    ratio = 0.5

    # 调用混合检索方法
    results = hybrid_search(chunks, query, top_k, ratio)

    # 输出返回结果
    print("Top-k Documents:")
    for idx, result in enumerate(results):
        print(f"Rank {idx + 1}: {result.page_content}")
