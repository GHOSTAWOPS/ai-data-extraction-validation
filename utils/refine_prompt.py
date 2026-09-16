# bge-m3
import chromadb
import ollama
import json
import time
import os
from config import CONFIG
from utils.logger_util import logger
from sentence_transformers import SentenceTransformer

def load_jsonl(file_path):
    """读取 JSONL 数据并返回数据列表"""
    data = []
    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            item = json.loads(line)
            data.append(item)
    return data


# def encode_text(text, model_name="bge-m3"):
#     """使用 Ollama 计算文本嵌入"""
#     response = ollama.embeddings(model=model_name, prompt=text)
#     return response["embedding"]
def encode_text(text, model_name="bge-m3"):
    """计算文本嵌入，支持 Ollama 或 Sentence Transformers"""
    if CONFIG.USETRANSFORMER:
        model = SentenceTransformer(CONFIG.EMBEDDING_MODEL_PATH)
        embedding = model.encode(text)
        return embedding.tolist()
    else:
        if ollama is None:
            raise ImportError("Ollama library is not installed.")
        response = ollama.embeddings(model=CONFIG.EMBED_MODEL, prompt=text)
        return response["embedding"]

def build_or_load_chroma_index(data, collection_name, persist_directory):
    """构建或加载 ChromaDB 索引"""
    chroma_client = chromadb.PersistentClient(path=persist_directory)
    collection = chroma_client.get_or_create_collection(name=collection_name)

    if collection.count() == 0:
        # logger.info("正在构建refine_prompt的索引...")
        queries = [item["query"] for item in data]
        query_vectors = [encode_text(q) for q in queries]  # 使用 Ollama 计算所有向量

        collection.add(
            ids=[str(i) for i in range(len(data))],
            embeddings=query_vectors,
            metadatas=data
        )
        logger.info("refine_prompt的索引构建完成！")
    else:
        logger.info("refine_prompt的索引已加载。")

    return collection

# 去重并合并的构建方法
# def build_or_load_chroma_index(data, collection_name, persist_directory):
#     """构建或加载 ChromaDB 索引"""
#     chroma_client = chromadb.PersistentClient(path=persist_directory)
#     collection = chroma_client.get_or_create_collection(
#         name=collection_name, 
#         metadata={"hnsw:space": "cosine"}
#     )

#     if collection.count() == 0:
#         print("正在构建索引...")
#         queries = [item["query"] for item in data]
#         query_vectors = [encode_text(q) for q in queries]  # 使用 Ollama 计算所有向量

#         # 只存储包含必要字段的条目
#         metadatas = []
#         ids = []
#         embeddings = []

#         for i, item in enumerate(data):
#             # 构建 metadata
#             metadata = {
#                 "values": ', '.join(item["values"]),
#                 "sources": ', '.join(item["sources"]),
#                 "prompt_example": item["prompt_example"]
#             }

#             # 如果 'units' 存在，添加到 metadata
#             if "units" in item:
#                 metadata["units"] = ', '.join(item["units"])

#             # 将有效数据添加到列表中
#             metadatas.append(metadata)
#             ids.append(str(i))
#             embeddings.append(query_vectors[i])

#         if metadatas:
#             collection.add(
#                 ids=ids,
#                 embeddings=embeddings,
#                 metadatas=metadatas  # 只添加有效的 metadata
#             )
#             print("索引构建完成！")
#         else:
#             print("没有符合条件的数据条目，索引未构建。")
#     else:
#         print("索引已加载。")

#     return collection


def search_query(query_text, top_k, collection):
    """使用 ChromaDB 搜索与 query 最相似的 top_k 结果"""
    query_vector = encode_text(query_text)  # Ollama 计算查询向量
    results = collection.query(
        query_embeddings=[query_vector],
        n_results=top_k
    )

    return results["metadatas"][0]  # 返回 top_k 相关结果


def refine_prompt(query):
    '''
        描述：提高prompt的质量，添加变量对应的示例
        输入：当前询问变量名字
        输出：相关示例
    '''
    logger.info("Refining prompt")
    logger.info("Refining prompt for query: {}".format(query))
    # jsonl_file = r"./utils/413merged_new_data_20250318_152830.jsonl"
    jsonl_file = r"./utils/data4_20_3_only2_with_shangxiawen.jsonl"
    persist_directory = r"./utils/chroma_data2"
    collection_name = "query_collection2"

    data = load_jsonl(jsonl_file)
    collection = build_or_load_chroma_index(data, collection_name, persist_directory)

    results = search_query(query, top_k=CONFIG.REFIE_PROMPT_TOP_K, collection=collection)
    
    
    # 提取 'prompt_example' 并合并为一个字符串
    prompt_examples = "\n".join([f"{i+3}."+item["prompt_example"] for i,item in enumerate(results)])

    return prompt_examples