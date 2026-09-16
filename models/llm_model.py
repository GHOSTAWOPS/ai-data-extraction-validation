import os

from langchain_community.llms import Ollama
from langchain_openai import ChatOpenAI
from langchain_community.llms.tongyi import Tongyi
from config import CONFIG
import subprocess
from utils.logger_util import logger
def get_llm_model(model_type,query=None):
    if model_type == "ollama":
        if ('(' in query or ')' in query )and CONFIG.LLM_MODEL_inference is not None:##如果关键词中有（）就用推理型模型
            model = Ollama(
            base_url="http://localhost:11434",
            model=CONFIG.LLM_MODEL_inference,
            temperature=0,  # 设置 temperature 值为 0
            top_k=1,
            top_p=0,
            num_predict=1024,  # 设置 输出token 数量 值为 
            num_ctx=4096,# 设置 输入token 数量 值为 
        )
            if CONFIG.LLM_MODEL is not None:
                logger.info(f"Terminated process {CONFIG.LLM_MODEL}")
                subprocess.run(['ollama', 'stop',CONFIG.LLM_MODEL], check=True)  
        else:
            if CONFIG.LLM_MODEL == 'tongyiqianwen72b':
                model = ChatOpenAI(
                    base_url='http://192.168.30.25:20010/v1',
                    api_key=os.environ["QWEN_API_KEY"],
                    model=CONFIG.LLM_MODEL,
                    temperature=0,
                ) 
            else:
                model = Ollama(
                base_url="http://localhost:11434",
                model=CONFIG.LLM_MODEL,
                temperature=0,  # 设置 temperature 值为 0
                top_k=1,
                top_p=0,
                num_predict=1024,  # 设置 输出token 数量 值为 
                num_ctx=4096,# 设置 输入token 数量 值为 
            )
            if CONFIG.LLM_MODEL_inference is not None:
                logger.info(f"Terminated process {CONFIG.LLM_MODEL}")
                subprocess.run(['ollama', 'stop',CONFIG.LLM_MODEL_inference], check=True)   
    elif model_type == "tongyi":
        # 通义api调用,max_retries网络延迟,最大重试次数.
        api_key = os.environ["DASHSCOPE_API_KEY"]
        api_base = "https://dashscope.aliyuncs.com/compatible-mode/v1"
        model = Tongyi(
            api_base=api_base,
            api_key=api_key,
            model="qwen-plus",
            temperature=0,
            max_retries=50,
        )
    elif model_type == "ollama_table":

        if CONFIG.TABLE_SUMMARY_MODEL == 'tongyiqianwen72b':
                model = ChatOpenAI(
                    base_url='http://192.168.30.25:20010/v1',
                    api_key=os.environ["QWEN_API_KEY"],
                    model=CONFIG.TABLE_SUMMARY_MODEL,
                    temperature=0,
                ) 
        else:
            model = Ollama(
                base_url="http://localhost:11434",
                model=CONFIG.TABLE_SUMMARY_MODEL,
                temperature=0,  # 设置 temperature 值为 0
                top_k=1,
                top_p=0,
                num_predict=2048,  # 设置 输出token 数量 值为 
                num_ctx=4096,# 设置 输入token 数量 值为 
            )

    else:
        raise ValueError("Invalid model type: {}".format(model_type))

    return model
