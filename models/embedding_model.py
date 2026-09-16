from langchain.pydantic_v1 import BaseModel, Extra, Field
from langchain.schema import BaseOutputParser
from langchain.embeddings.base import Embeddings
from typing import Any, Dict
from llama_index.embeddings.adapter import (
    LinearAdapterEmbeddingModel,
    AdapterEmbeddingModel,
)
from llama_index.embeddings.ollama import OllamaEmbedding
from langchain_ollama import OllamaEmbeddings
from llama_index.embeddings.adapter.utils import TwoLayerNN

# zhy3_6
from sentence_transformers import SentenceTransformer
import os
from typing import List
from peft import LoraConfig, TaskType
import torch
from config import CONFIG



# 模拟一个自定义输出解析器，可以处理字典(dict)（自我评估回答处理结果（result）和得分（score））
class MultiOutputParser(BaseOutputParser):
    def parse(self, output: str):
        # 假设输出格式是两部分字符串，用分隔符分开
        parts = output.split("###")
        result = {
            "result": parts[0].strip(),
            "score": parts[1].strip() if len(parts) > 1 else "",
        }
        return result  # dict


class WrappedEmbeddingModel(BaseModel, Embeddings):
    client: Any  #: :meta private:
    base_embed_model: Any
    model_name: str
    adapter_path: str
    model_kwargs: Dict[str, Any] = Field(default_factory=dict)

    def __init__(self, **kwargs: Any):
        super().__init__(**kwargs)
        # self.client = LinearAdapterEmbeddingModel(base_embed_model=self.base_embed_model, adapter_path=self.adapter_path)#单层线性层
        self.client = AdapterEmbeddingModel(
            base_embed_model=self.base_embed_model,
            adapter_path=self.adapter_path,
            adapter_cls=TwoLayerNN,
        )  # 双层线性层

    def embed_documents(self, texts):
        if isinstance(texts, list):
            # embeddings = [self.client._get_query_embedding(text) for text in texts]## 微调了document and query
            embeddings = [
                self.client._get_text_embedding(text) for text in texts
            ]  ##只微调了only quert
        elif isinstance(texts, str):
            # embeddings = self.client._get_query_embedding(texts)## 微调了document and query
            embeddings = self.client._get_text_embedding(texts)  ##只微调了only quert
        else:
            raise ValueError("Invalid input type: expected str or List[str]")
        return embeddings

    def embed_query(self, text: str):  ##只微调了only query
        text = text.replace("\n", " ")
        embeddings = self.client._get_query_embedding(text)
        return embeddings


# zhy3_6添加
class LangchainSentenceTransformer(Embeddings):
    def __init__(self, model: SentenceTransformer):  # 接受 SentenceTransformer 对象
        self.model = model

    def embed_documents(self, texts):
        embedding = self.model.encode(
            texts, convert_to_numpy=True, show_progress_bar=True
        )
        return embedding.tolist()

    def embed_query(self, text):
        embedding = self.model.encode([text], convert_to_numpy=True)[0]
        return embedding.tolist()


# zhy3_6
def get_embedding_model(
    model_name, adapter_path=None, model_path=None, peft_lora_adapter_path=None
):
    # def get_embedding_model(model_name, adapter_path=None):
    ## ollama下载或者huggingface下载的embedding模型
    # embeddings_model = embeddings.OllamaEmbeddings(model=CONFIG.EMBED_MODEL)
    # embeddings_model = OllamaEmbeddings(model=CONFIG.EMBED_MODEL)
    # embeddings_model = SentenceTransformer("/home/pkr/projects/Models/Yuan-embedding-1.0").encode
    # embeddings_model = HuggingFaceEmbeddings(model_name=CONFIG.EMBED_MODEL)

    # fine_tune的embedding模型
    if adapter_path is not None:
        embeddings_model = OllamaEmbedding(model_name=model_name)
        embeddings_model = WrappedEmbeddingModel(
            base_embed_model=embeddings_model,
            model_name="fine_tuned_model",
            adapter_path=adapter_path,
        )
    # zhy3_6
    elif peft_lora_adapter_path is not None:
        embeddings_model = SentenceTransformer(model_path)
        adapter_name = os.path.basename(os.path.normpath(peft_lora_adapter_path))
        print(f"正在加载适配器: {peft_lora_adapter_path}, adapter_name={adapter_name}")
        embeddings_model.load_adapter(peft_lora_adapter_path, adapter_name=adapter_name)
        embeddings_model.set_adapter(adapter_name)
        embeddings_model.enable_adapters()
        embeddings_model.active_adapters()
        embeddings_model = LangchainSentenceTransformer(embeddings_model)

    else:
        if CONFIG.USETRANSFORMER:
            embeddings_model = SentenceTransformer(model_path, device="cuda" if torch.cuda.is_available() else "cpu")
            embeddings_model = LangchainSentenceTransformer(embeddings_model)
        else:
        # ollama的embedding模型
            embeddings_model = OllamaEmbeddings(model=model_name)
    return embeddings_model
