from langchain_community.vectorstores import Chroma
from models.embedding_model import get_embedding_model
from config import CONFIG
import random
from langchain.schema import Document
import torch
from models.new_rerank_model import set_seed
from utils.utils import error_handler
from datetime import datetime


# 随机生成集合名称
def random_collection_name():
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    random_number = random.randint(100000, 999999)
    return f"collection_{timestamp}_{random_number}"


# 将文档添加到知识库
@error_handler("文档处理出错，请联系技术人员")
def add_doc_to_kb(pages, collection_name="name"):
    # 设置随机种子确保一致性
    set_seed(42)
    torch.manual_seed(42)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(42)

    embeddings_model = get_embedding_model(
        # model_name=CONFIG.EMBED_MODEL, adapter_path=CONFIG.ADAPTER_PATH
        model_name=CONFIG.EMBED_MODEL,
        adapter_path=CONFIG.ADAPTER_PATH,
        model_path=CONFIG.EMBEDDING_MODEL_PATH,
        peft_lora_adapter_path=CONFIG.PEFT_LORA_ADAPTER_PATH,
    )

    db = Chroma(
        collection_name=collection_name,
        embedding_function=embeddings_model,
        persist_directory="./chroma_data",  # 需在配置中定义路径，如 "./chroma_data"
    )

    # 添加文档并持久化
    # db.add_documents(documents=pages, embedding=embeddings_model)
    batch_size = 5000
    for i in range(0, len(pages), batch_size):
        batch = pages[i:i+batch_size]
        db.add_documents(documents=batch)

    # db.add_documents(documents=pages)
    # db.persist()

    # 创建 retriever
    retriever = db.as_retriever(
        search_type=CONFIG.SEARCH_TYPE, search_kwargs=CONFIG.SEARCH_KWARGS
    )

    return retriever


# 执行测试
def test_add_documents_and_retrieve():
    # 生成随机集合名称
    collection_name1 = random_collection_name()
    collection_name2 = random_collection_name()

    # 创建两个文档集
    pages1 = [Document(page_content="This is the first page.")]
    pages2 = [Document(page_content="This is the second page.")]

    # 将文档添加到不同的集合中
    retriever1 = add_doc_to_kb(pages1, collection_name1)
    retriever2 = add_doc_to_kb(pages2, collection_name2)

    # 假设你有一个方法来进行检索并输出结果
    result1 = retriever1.get_relevant_documents("first page")  # 在集合1中检索
    result2 = retriever2.get_relevant_documents("second page")  # 在集合2中检索

    # 输出检索结果，验证文档是否正确
    print(f"Results from collection {collection_name1}: {result1}")
    print(f"Results from collection {collection_name2}: {result2}")


if __name__ == "__main__":
    # 执行测试
    test_add_documents_and_retrieve()
