from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
from langchain.schema import Document
from typing import List
from config import CONFIG
import torch

# from rank_bm25 import BM25Okapi
from FlagEmbedding import FlagReranker
from typing import List

from utils.logger_util import logger


# ------------------------
# 1-model: SentenceTransformer
# Yuan-embedding-1.0(new_rerank_model)    zpoint_large_embedding_zh(new_rerank_model) Yinka(new_rerank_model)
# with BM25
# ------------------------
# 加载 SentenceTransformer 模型
# model = SentenceTransformer(CONFIG.RERANK_MODEL_PATH)
# # 可选：设置模型的最大序列长度（与 CONFIG 一致）
# model.max_seq_length = CONFIG.TEXT_MAX_LENGTH
# def embed_text_batch(texts: List[str]) -> torch.Tensor:
#     """
#     使用 SentenceTransformer 模型批量计算嵌入向量。

#     Args:
#         texts (List[str]): 要编码的文本列表。

#     Returns:
#         torch.Tensor: 文本的嵌入向量张量。
#     """
#     # SentenceTransformer 自动完成分词、截断、填充和嵌入生成
#     embeddings = model.encode(texts, normalize_embeddings=True)
#     return torch.tensor(embeddings, device=torch.device("cpu"))  # 转为 PyTorch 张量
# def compute_bm25_scores(documents: List[Document], query: str) -> List[float]:
#     """
#     使用 BM25 算法计算文档与查询的相关性得分。

#     Args:
#         documents (List[Document]): 待评分的文档列表（LangChain 的 Document 对象）。
#         query (str): 查询字符串。

#     Returns:
#         List[float]: 每个文档的 BM25 相关性得分。
#     """
#     # 提取文档内容并分词
#     doc_texts = [doc.page_content for doc in documents]
#     tokenized_docs = [doc.split() for doc in doc_texts]

#     # 初始化 BM25 模型
#     bm25 = BM25Okapi(tokenized_docs)

#     # 查询分词
#     tokenized_query = query.split()

#     # 计算 BM25 相关性得分
#     scores = bm25.get_scores(tokenized_query)
#     return scores

# def rerank_docs(documents: List[Document], query: str, top_k: int = CONFIG.EXTRACT_RERANK_TOP_K) -> List[Document]:
#     """
#     使用 BM25 算法对文档进行重排序。

#     Args:
#         documents (List[Document]): 待重排序的文档列表（LangChain 的 Document 对象）。
#         query (str): 查询字符串。
#         top_k (int): 返回前 k 个相关文档。

#     Returns:
#         List[Document]: 排序后的文档列表。
#     """
#     # 计算 BM25 相关性得分
#     relevance_scores = compute_bm25_scores(documents, query)

#     # 根据相关性得分排序（降序）
#     ranked_indices = sorted(range(len(relevance_scores)), key=lambda i: relevance_scores[i], reverse=True)
#     ranked_docs = [documents[i] for i in ranked_indices[:top_k]]

#     return ranked_docs


# ----------------
# 1.model
# Yuan-embedding-1.0(new_rerank_model)    zpoint_large_embedding_zh(new_rerank_model) Yinka(new_rerank_model)
# with similarity
# ----------------
# def rerank_docs(documents: List[Document], query: str, top_k: int = CONFIG.EXTRACT_RERANK_TOP_K) -> List[Document]:
#     """
#     使用 SentenceTransformer 嵌入模型对文档进行重排序。

#     Args:
#         documents (List[Document]): 待重排序的文档列表（LangChain 的 Document 对象）。
#         query (str): 查询字符串。
#         top_k (int): 返回前 k 个相关文档。

#     Returns:
#         List[Document]: 排序后的文档列表。
#     """
#     # 生成查询的嵌入向量
#     query_embedding = embed_text_batch([query]).numpy()

#     # 提取文档内容
#     doc_texts = [doc.page_content for doc in documents]

#     # 批量计算文档的嵌入向量
#     doc_embeddings = embed_text_batch(doc_texts).numpy()

#     # 计算余弦相似度
#     similarities = cosine_similarity(query_embedding, doc_embeddings)[0]

#     # 根据相似度排序（降序）
#     ranked_indices = similarities.argsort()[::-1]
#     ranked_docs = [documents[i] for i in ranked_indices[:top_k]]

#     return ranked_docs


# ----------------
# 2. model
# baai-bge-reranker
# ----------------
import random
import numpy as np


# # 保证 PyTorch 使用固定随机种子，确保一致性
def set_seed(seed: int = 42):
    torch.manual_seed(seed)
    torch.cuda.manual_seed_all(seed)
    torch.backends.cudnn.deterministic = True  # 禁用 cudnn 的非确定性算法
    torch.backends.cudnn.benchmark = False  # 禁用 cudnn 的自动优化
    # torch.set_deterministic(True)
    np.random.seed(seed)
    random.seed(seed)


# reranker = FlagReranker(CONFIG.RERANK_MODEL_PATH, use_fp16=False)

# def compute_rerank_scores(documents: List[str], query: str) -> List[float]:
#     """
#     使用 BAAI bge-reranker 模型计算文档与查询的相关性得分。

#     Args:
#         documents (List[str]): 待评分的文档列表（字符串列表）。
#         query (str): 查询字符串。

#     Returns:
#         List[float]: 每个文档的相关性得分。
#     """
#     # 确保模型在无随机性的模式下运行
#     reranker.model.eval()  # 设置模型为评估模式，禁用 Dropout 和其他随机性

#     # 提取文档内容
#     doc_texts = [doc.page_content for doc in documents]

#     # 构建查询-文档对
#     pairs = [[query, doc_text] for doc_text in doc_texts]

#     # 使用 BAAI bge-reranker 模型计算得分
#     scores = reranker.compute_score(pairs)

#     return scores


# def rerank_docs(
#     documents: List[Document], query: str, top_k: int = CONFIG.EXTRACT_RERANK_TOP_K
# ) -> tuple[List[Document], List[float]]:
#     """
#     使用 BAAI bge-reranker 模型对文档进行重排序。

#     Args:
#         documents (List[Document]): 待重排序的文档列表（LangChain 的 Document 对象）。
#         query (str): 查询字符串。
#         top_k (int): 返回前 k 个相关文档。

#     Returns:
#         tuple[List[Document], List[float]]: 排序后的文档列表和相关性得分。
#     """
#     # 设置随机种子确保一致性
#     set_seed(42)
#     torch.manual_seed(42)
#     if torch.cuda.is_available():
#         torch.cuda.manual_seed_all(42)

#     # 计算相关性得分
#     relevance_scores = compute_rerank_scores(documents, query)

#     # 根据相关性得分排序（降序）
#     # 使用稳定排序，确保得分相同的文档按原始顺序排序
#     ranked_indices = sorted(
#         range(len(relevance_scores)), key=lambda i: relevance_scores[i], reverse=True
#     )

#     relevance_scores_topk = [relevance_scores[i] for i in ranked_indices[:top_k]]

#     # 返回重排序后的文档以及相关性得分
#     ranked_docs = [documents[i] for i in ranked_indices[: len(relevance_scores_topk)]]

#     return ranked_docs, relevance_scores_topk


# 修改这段代码，目前是直接取top_k，现在修改为相似度分数x以上的
# ------------
# 采取bm25混合的方式
# 如需修改，直接注释并使用上面的代码。
# -----------
from rank_bm25 import BM25Okapi
import torch
from typing import List
from langchain.schema import Document
import jieba
import math
from collections import Counter
from transformers import AutoTokenizer
from concurrent.futures import ThreadPoolExecutor
import threading

# 创建一个全局锁，确保只有一个线程可以访问 `reranker` 模型
lock = threading.Lock()


reranker = FlagReranker(
    CONFIG.RERANK_MODEL_PATH,
    devices=["cuda:0"],
    use_fp16=False
)
reranker.model = reranker.model.to("cuda:0")

print("#==================================================================#")
print("模型设备：", reranker.model.device)
print("是否使用CUDA：", torch.cuda.is_available())
print("设备数量：", torch.cuda.device_count())
print("模型第一个参数所在设备：", next(reranker.model.parameters()).device)
print("#==================================================================#")

logger.info(f"Rerank模型运行设备：{next(reranker.model.parameters()).device}")
# def tokenize(text: str) -> List[str]:
#     # 使用 jieba 分词
#     return list(jieba.cut(text))
# 有问题，后续继续测试，jieba有问题

#
# def compute_bm25_scores(documents: List[str], query: str) -> List[float]:
#     """
#     使用 BM25 计算文档与查询的相关性得分（针对中文文本）。
#     """
#     # 使用 jieba 对文档和查询进行分词
#     tokenized_docs = [tokenize(doc) for doc in documents]
#     tokenized_query = tokenize(query)


#     bm25 = BM25Okapi(tokenized_docs)
#     scores = bm25.get_scores(tokenized_query)
#     return scores
# 计算 BM25 得分
def compute_bm25_scores(documents: List[str], query: str, rewrite_query = True) -> List[float]:
    """
    使用 BM25 计算文档与查询的相关性得分。

    Args:
        documents (List[str]): 待评分的文档列表（字符串列表）。
        query (str): 查询字符串。

    Returns:
        List[float]: 每个文档的相关性得分。
    """
    # 重写 query
    print(query)
    if rewrite_query:
        query = query.split("中，")
        query = query[-1]
    
    print(query)
    print("===============")

    # 将文档和查询转换为分词列表
    tokenized_docs = [list(jieba.cut(doc)) for doc in documents]
    tokenized_query = list(jieba.cut(query))

    # context_paths = [doc.split("\n")[0] for doc in documents]
    # tokenized_docs = [list(jieba.cut(doc)) for doc in context_paths]
    # tokenized_query = list(jieba.cut(query))

    # 使用 BM25 计算得分
    bm25 = BM25Okapi(tokenized_docs)
    scores = bm25.get_scores(tokenized_query)

    # scores = [10 * math.tanh(s) for s in scores]

    return scores


# 计算模型得分
def compute_rerank_scores(documents: List[str], query: str) -> List[float]:
    """
    使用模型计算文档与查询的相关性得分。

    Args:
        documents (List[str]): 待评分的文档列表。
        query (str): 查询字符串。

    Returns:
        List[float]: 每个文档的相关性得分。
    """
    reranker.model.eval()  # 设置模型为评估模式，禁用 Dropout 和其他随机性

    # 构建查询-文档对
    pairs = [[query, doc] for doc in documents]

    # 使用 BAAI bge-reranker 模型计算得分
    with lock:
        # 使用 BAAI bge-reranker 模型计算得分
        scores = reranker.compute_score(pairs)

    return scores


# 计算模型得分
def compute_path_rerank_scores(documents: List[str], query: str) -> List[float]:
    """
    使用模型计算文档与查询的相关性得分。

    Args:
        documents (List[str]): 待评分的文档列表。
        query (str): 查询字符串。

    Returns:
        List[float]: 每个文档的相关性得分。
    """
    reranker.model.eval()  # 设置模型为评估模式，禁用 Dropout 和其他随机性

    # 构建查询-文档对
    context_paths = [doc.split("\n")[0] for doc in documents]
    pairs = [[query, doc] for doc in context_paths]

    # 使用 BAAI bge-reranker 模型计算得分
    with lock:
        # 使用 BAAI bge-reranker 模型计算得分
        scores = reranker.compute_score(pairs)

    return scores


def compute_rrf_scores(model_ranks: List[int], bm25_ranks: List[int], k=60):
    n = len(model_ranks)
    rrf_scores = [0.0] * n
    for i in range(n):
        rrf_scores[i] = 1 / (k + model_ranks[i]) + 1 / (k + bm25_ranks[i])
    return rrf_scores


# zhy2/23测试
def only_rerank_docs(
    doc_texts: List[Document], query: str, top_k: int = 5
) -> tuple[List[Document], List[float]]:
    """
    使用 BM25 和模型得分对文档进行重排序。

    Args:
        documents (List[Document]): 待重排序的文档列表。
        query (str): 查询字符串。
        top_k (int): 返回前 k 个相关文档。

    Returns:
        tuple[List[Document], List[float]]: 排序后的文档列表和相关性得分。
    """
    # 提取文档内容
    # doc_texts = [doc.page_content for doc in documents]

    # 计算 BM25 相关性得分
    bm25_scores = compute_bm25_scores(doc_texts, query)

    # 使用 FlagReranker 计算模型得分
    model_scores = compute_rerank_scores(doc_texts, query)

    # 合并得分（加权平均）
    # 在这里我们采用 50% 加权方式，你可以调整权重
    combined_scores = [
        # 0.55 0.5
        0.55* bm25 + 0.45* model for bm25, model in zip(bm25_scores, model_scores)
    ]

    # 根据合并得分进行排序（降序）
    ranked_indices = sorted(
        range(len(combined_scores)), key=lambda i: combined_scores[i], reverse=True
    )

    # 返回排名前 top_k 的文档及得分
    ranked_docs = [doc_texts[i] for i in ranked_indices[:top_k]]
    combined_scores_topk = [combined_scores[i] for i in ranked_indices[:top_k]]

    return ranked_docs, combined_scores_topk

def rerank_docs(
    documents: List[Document], query: str, top_k: int = 5,time_check=0
) -> tuple[List[Document], List[float]]:
    """
    Args:
        documents (List[Document]): 待重排序的文档列表。
        query (str): 查询字符串。
        top_k (int): 返回前 k 个相关文档。

    Returns:
        tuple[List[Document], List[float]]: 排序后的文档列表和相关性得分。
    """
    doc_texts = [doc.page_content for doc in documents]

    # alpha = 0.9
    # scores1 = compute_rerank_scores(doc_texts, query)
    # scores2 = compute_path_rerank_scores(doc_texts, query)
    # model_scores = [alpha * s1 + (1 - alpha) * s2 for s1, s2 in zip(scores1, scores2)]

    model_scores = compute_rerank_scores(doc_texts, query)
    ranked_indices = sorted(
        range(len(model_scores)), key=lambda i: model_scores[i], reverse=True
    )
    # ranked_docs = [documents[i] for i in ranked_indices[:top_k]]
    # model_scores_topk = [model_scores[i] for i in ranked_indices[:top_k]]
    ranked_docs = [documents[i] for i in ranked_indices]
    model_scores_topk = [model_scores[i] for i in ranked_indices]
    for doc,score in zip(ranked_docs,model_scores_topk):
        doc.metadata["rerank_score"]=score
    return ranked_docs, model_scores_topk


def rrf_docs(
    documents: List[Document], query: str, top_k: int = 5, time_check=0, k: int = 60
) -> tuple[List[Document], List[float]]:
    """
    Args:
        documents (List[Document]): 待重排序的文档列表。
        query (str): 查询字符串。
        top_k (int): 返回前 k 个相关文档。

    Returns:
        tuple[List[Document], List[float]]: 排序后的文档列表和相关性得分。
    """
    doc_texts = [doc.page_content for doc in documents]

    # alpha = 0.9
    # scores1 = compute_rerank_scores(doc_texts, query)
    # scores2 = compute_path_rerank_scores(doc_texts, query)
    # model_scores = [alpha * s1 + (1 - alpha) * s2 for s1, s2 in zip(scores1, scores2)]
    
    bm25_query = query.split("，")[-1]
    # bm25_query = query
    rerank_query = query

    print(f"RRF参数k：{k}")
    print(f"BM25 query：{bm25_query}")
    print(f"Rerank query：{rerank_query}")

    # 计算 rerank 得分
    model_scores = compute_rerank_scores(doc_texts, rerank_query)
    # 计算 bm25 得分
    bm25_scores = compute_bm25_scores(doc_texts, bm25_query)

    # 计算 rerank 排序
    ranked_indices = sorted(
        range(len(model_scores)), key=lambda i: model_scores[i], reverse=True
    )
    # 计算 bm25 排序
    bm25_indices = sorted(
        range(len(bm25_scores)), key=lambda i: bm25_scores[i], reverse=True
    )

    # 计算 rerank 排名
    model_ranks = [0] * len(model_scores)
    for rank, idx in enumerate(ranked_indices, start=1):
        model_ranks[idx] = rank
    # 计算 bm25 排名
    bm25_ranks = [0] * len(bm25_scores)
    for rank, idx in enumerate(bm25_indices, start=1):
        bm25_ranks[idx] = rank

    # 计算 rrf 得分
    rrf_scores = compute_rrf_scores(model_ranks, bm25_ranks, k=k)
    # 计算 rrf 排序
    rrf_indices = sorted(
        range(len(rrf_scores)), key=lambda i: rrf_scores[i], reverse=True
    )
    # 计算 rrf 排名
    rrf_ranks = [0] * len(rrf_scores)
    for rank, idx in enumerate(rrf_indices, start=1):
        rrf_ranks[idx] = rank

    # for rerank_score, bm25_score, rrf_score, rerank_rank, bm25_rank, rrf_rank in zip(
    #     model_scores, bm25_scores, rrf_scores, model_ranks, bm25_ranks, rrf_ranks 
    #     ):
    #     print("===========================")
    #     print(f"rerank得分：{rerank_score}")
    #     print(f"rerank排名：{rerank_rank}")
    #     print(f" bm25 得分：{bm25_score}")
    #     print(f" bm25 排名：{bm25_rank}")
    #     print(f" rrf  得分：{rrf_score}")
    #     print(f" rrf  排名：{rrf_rank}")
    #     print("===========================")

    rrf_docs = [documents[i] for i in rrf_indices]
    rrf_scores_topk = [rrf_scores[i] for i in rrf_indices]
    for doc, score in zip(rrf_docs, rrf_scores_topk):
        doc.metadata["rerank_score"]=score
    return rrf_docs, rrf_scores_topk
