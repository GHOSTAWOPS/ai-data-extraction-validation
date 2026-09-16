import re
import time

import math
import sys
# sys.path.append('E:/lushan2_23/LuShan')  # 添加项目根目录
from docx import Document as pyDocument
from docx.table import Table
from langchain_community.vectorstores import Chroma, FAISS, Weaviate

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
from models.new_rerank_model import only_rerank_docs, rerank_docs, set_seed  # dcx 修改

# from models.rerank_model import rerank_docs  #dcx 修改于2024/10/31

from models.prompt_engineering import embed_prompt, extract_prompt, llm_query_keyword, rerank_prompt
from langchain_community.llms.tongyi import Tongyi
from utils.logger_util import logger

from models.llm_model import get_llm_model
import os
from data_processor import html_table_to_markdown, is_table_caption

# from models.embedding_model import MultiOutputParser, get_embedding_model
from models.new_embed_model import MultiOutputParser, get_embedding_model
from utils.utils import error_handler
from chromadb.config import Settings
import shutil
import os
from datetime import datetime
import torch
from models.hybrid_search import hybrid_search
from utils.utils import chunk_filter

import asyncio
from concurrent.futures import ThreadPoolExecutor
from functools import partial

import matplotlib.pyplot as plt
from matplotlib import rcParams

# 设置字体为 SimHei（支持中文）
# rcParams['font.family'] = ['SimHei']

def process_rerank_file(file, output_dir):  
    with open(file, 'r', encoding='utf-8') as f:
        txt_content = f.read()

    
    os.makedirs(output_dir, exist_ok=True)
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"output_{current_time}.txt"

    llm_sections = txt_content.split("rerank的询问的字段:")
    llm_sections = llm_sections[1:]  # 去掉第一部分（包含retriever和rerank的部分）

    total_queries = 0  # 总查询数量
    rerank_count = 0  # rerank匹配正确的查询数量
    rerank_block_match_counts = [0] * 62  # rerank每个块的匹配总次数，长度为 62
    rerank_unmatched_queries = 0  # rerank所有块都没有匹配的问题数量
    first_three_blocks_query_match = 0  # rerank前1、2、3块中有匹配的问题数
    # embedding_count = 0  # embedding匹配正确的查询数量
    # embedding_block_match_counts = [0] * 61  # embedding每个块的匹配总次数，长度为 61
    # embedding_unmatched_queries = 0  # embedding所有块都没有匹配的问题数量
    last_non_zero_block = 0
    for section in llm_sections:
        # 提取查询
        query = section.split("\n")[0].strip()
        answer = section.split("ground truth flag")[0].strip()
        answer = answer.split("ground truth value:")[1].split()[0].strip()
        section = section.split("\nrerank_context:\n")[1].strip() # 去掉前面的得分和问题
        rerank_section = section.split("\nembedding_context:\n")[0].strip()
        doc_texts = re.findall(r'第\d+块："([^"]*)"', rerank_section)
        relevant_docs, rerank_scores_topk = only_rerank_docs(
                doc_texts, query, top_k=CONFIG.EXTRACT_RERANK_TOP_K
            )
        rerank_context = "\n".join(
            [f'第{i}块："{doc}"' for i, doc in enumerate(relevant_docs)]
        )
        # with open(os.path.join(output_dir, output_filename), 'a', encoding='utf-8') as output_file:
        #     output_file.write(f"rerank的询问的字段: {query}\n")
        #     output_file.write(f"ground truth value: {answer}\n")
        #     output_file.write(f"rerank_context:\n {rerank_context}\n\n")
        print(f"\n查询: {query}")
        print(f"\n答案: {answer}")
        total_queries += 1

        rerank_found_correct = False  # 为整个查询添加一个标志
        # embedding_found_correct = False  # 为整个查询添加一个标志
        first_three_match = False  # 为整个查询添加一个标志

        for i, block in enumerate(relevant_docs):
            if answer in block:
                rerank_block_match_counts[i + 1] += 1
                print(f"第{i}块: 匹配正确答案")
                rerank_found_correct = True
                if i < 3:
                    first_three_match = True
            else:
                print(f"第{i}块: 未匹配到正确答案")
        
        if first_three_match:
            first_three_blocks_query_match += 1
        else:
            print(f"{query}top3_rerank未匹配正确答案")
        if rerank_found_correct:
            rerank_count += 1
        else:
            print("-" * 3000)
            rerank_unmatched_queries += 1
    accuracy = (first_three_blocks_query_match / total_queries) * 100
    print(f"rerank top3精度为{accuracy:.2f}%({first_three_blocks_query_match}/{total_queries})")


    # 绘制rerank柱状图
    plt.figure(figsize=(12, 6))
    bars = plt.bar(range(0, 62), rerank_block_match_counts, color='skyblue', edgecolor='black')
    # 在每个柱的顶部添加数字标注
    for i, bar in enumerate(bars):
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2, height, f'{height}', ha='center', va='bottom', fontsize=10)

    plt.title('每个块的匹配次数柱状图', fontsize=16)
    plt.xlabel('块编号', fontsize=16)
    plt.ylabel('匹配次数', fontsize=16)
    plt.xticks(range(0, 62, 1))
    text = f"rerank\ntop3匹配查询数: {first_three_blocks_query_match}\n总查询数: {total_queries}\n精度: {accuracy:.2f}%"
    plt.text(0.85, 0.95, text, ha='left', va='top', transform=plt.gca().transAxes, fontsize=12,
            bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.5'))
    plt.tight_layout()

    current_time2 = datetime.now().strftime("%Y%m%d_%H%M%S")
    save_path = os.path.join(output_dir, f"rerank_block_match_distribution_{current_time2}.png")

    plt.savefig(save_path)
    # plt.show()

if __name__ == "__main__":
    start_time = time.time()
    files = [
        r"/home/ljc/project/lushan_test/LuShan/data/test_data/embeddingtop60/下尔呷水电站_rerank_context_output_20250223_185328.txt",
        # r"/home/ljc/project/lushan_test/LuShan/data/test_data/embeddingtop60/芦山抽水蓄能电站_rerank_context_output_20250224_061041.txt",
        # r"/home/ljc/project/lushan_test/LuShan/data/test_data/embeddingtop60/猴子岩水电站_rerank_context_output_20250303_085255.txt"
    ]
    output_dir = r"/home/ljc/project/lushan_test/LuShan/data/test_data/only_rerank_output"

    for file in files:
        print(f"processsed:{file}")
        process_rerank_file(file, output_dir)

    end_time = time.time()
    # 计算程序运行时间
    execution_time = end_time - start_time
    # 输出程序运行时间
    print(f"程序运行时间: {execution_time/60}分钟")
