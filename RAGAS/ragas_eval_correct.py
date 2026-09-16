import json
import os
from pathlib import Path

from dotenv import load_dotenv
import requests
import time
from typing import Dict, Any
from dataclasses import dataclass

@dataclass
class EvaluationInput:
    query: str
    context: str
    response: str

class _QwenAPIClient:
    __BASE_URL = "http://192.168.30.25:20010/v1/chat/completions"

    def __init__(self, api_key: str):  # 确保API端点路径正确
        self.__api_key = api_key
        self.__headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
    
    def _extract_json_from_text(self, text: str) -> dict:
        """
        从API响应文本中提取JSON内容
        """
        try:
            # 处理带代码块标记的响应
            if "```json" in text and "```" in text:
                start = text.find("```json") + 7
                end = text.find("```", start)
                json_str = text[start:end].strip()
                return json.loads(json_str)
            elif text.strip().startswith("{") and text.strip().endswith("}"):
                return json.loads(text.strip())
            else:
                # 如果不是JSON格式，尝试返回默认结构
                return {"score": 0, "analysis": text}
        except Exception as e:
            # 解析失败时返回默认结构
            return {"score": 0, "analysis": text}

    def call_api(self, prompt: str, max_tokens: int = 2000) -> str:
        # 修正请求结构：直接使用messages字段，去除多余的嵌套
        payload = {
            "model": "tongyiqianwen72b",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0,
            "max_tokens": max_tokens,
            "stream": False,
        }
        
        try:
            response = requests.post(
                self.__BASE_URL,
                headers=self.__headers,
                json=payload,
                timeout=30
            )
            response.raise_for_status()
            result = response.json()
            
            # 兼容不同的响应格式
            if "choices" in result and result["choices"]:
                return result["choices"][0]["message"]["content"]
            elif "output" in result and "text" in result["output"]:
                return result["output"]["text"]
            else:
                return str(result)
        except Exception as e:
            print(f"API调用错误: {e}")
            print(f"请求URL: {self.base_url}")
            print(f"请求载荷: {json.dumps(payload, ensure_ascii=False, indent=2)}")
            if 'response' in locals():
                print(f"响应状态: {response.status_code}")
                print(f"响应内容: {response.text}")
            else:
                print("无响应")
            return ""

class _RAGASEvaluator:
    def __init__(self, api_client: _QwenAPIClient):
        self.__api_client = api_client

    def evaluate_faithfulness(self, eval_input: EvaluationInput) -> Dict[str, Any]:
        prompt = f"""
请使用RAGAS方法评估回答的忠实度（faithfulness）。

问题: {eval_input.query}
上下文: {eval_input.context}
回答: {eval_input.response}

评估标准：
1. 回答中的信息是否都能在上下文中找到支持
2. 是否存在与上下文矛盾的信息
3. 是否有无法从上下文验证的内容

请给出0-1之间的分数（1分最高，0分最低）和详细分析。
返回JSON格式：{{"score": 分数, "analysis": "分析"}}
"""
        result_text = self.__api_client.call_api(prompt)
        parsed_result = self.__api_client._extract_json_from_text(result_text)
        score = parsed_result.get("score", 0)
        parsed_result["score"] = min(max(score, 0), 1)
        return parsed_result

class _RAGEvaluationSuite:
    def __init__(self, api_key: str):
        self.__api_client = _QwenAPIClient(api_key)
        self.__ragas = _RAGASEvaluator(self.__api_client)

    def evaluate_ragas(self, eval_input: EvaluationInput) -> Dict[str, Any]:
        # print("开始RAGAS评估...")
        start_time = time.perf_counter()
        faithfulness_result = self.__ragas.evaluate_faithfulness(eval_input)
        duration_sec = time.perf_counter() - start_time
        # 只保留到小数点后2位即可
        duration_sec_rounded = round(duration_sec, 2)
        print(f"RAGAS 耗时: {duration_sec_rounded:.2f} 秒")
        return {
            "faithfulness": faithfulness_result,
            "duration_seconds": duration_sec_rounded
        }


# 外部调用接口 1：初始化评估器
def init_ragas(api_key: str) -> object:
    """
    初始化 RAGAS 评估器对象，用于后续的回答评估。

    参数：
        api_key (str): 用于访问大模型 API 的授权密钥，可从环境变量 QWEN_API_KEY 读取。

    返回：
        object: 返回一个 `_RAGEvaluationSuite` 实例，封装了 RAGAS 的评估方法（如回答忠实度评估）。
    """
    return _RAGEvaluationSuite(api_key)


# 外部调用接口 2：传入文本获得得分
def get_ragas_score(evaluator: object, query: str, context: str, response: str) -> Dict[str, Any]:
    """
    使用 RAGAS 评估器对模型生成的回答进行评估，返回忠实度分数及分析信息。

    参数：
        evaluator (object): 由 init_ragas(api_key) 初始化得到的评估器对象。
        query (str): 用户提出的问题，即原始查询。
        context (str): 用于回答该问题的上下文（检索结果、文档片段等）。
        response (str): 模型生成的回答。

    返回：
        Dict[str, Any]: 评估结果字典，格式如下：
            {
                "faithfulness": {
                    "score": float,        # 忠实度得分（范围：0 到 1）
                    "analysis": str        # 模型输出的分析解释
                },
                "duration_seconds": float  # 模型评估耗时（单位：秒）
            }
    """
    eval_input = EvaluationInput(query=query, context=context, response=response)
    return evaluator.evaluate_ragas(eval_input)


# def evaluate_and_save(samples, prefix, evaluator, batch_size: int = 10, save_inputs: bool = False):
#     """批量评估：每 batch_size 条写入一个文件。

#     Args:
#         samples: 样本列表
#         prefix: 文件名前缀（例如 'p' 或 'n'）
#         evaluator: 评估器
#         batch_size: 每多少条合并保存一次
#         save_inputs: 是否把原始 query/context/response 也写进批量结果里
#     输出文件命名：{prefix}{start_index+1}-{end_index}.json，例如 p1-10.json
#     单条出错会记录在 errors 列表中，不影响后续。
#     """
#     batch_results = []
#     errors = []
#     start_index_of_batch = 0

#     def flush_batch(last_idx_inclusive: int):
#         nonlocal batch_results, start_index_of_batch
#         if not batch_results:
#             return
#         start_no = start_index_of_batch + 1
#         end_no = last_idx_inclusive + 1
#         filename = f"{prefix}{start_no}-{end_no}.json"
#         payload = {
#             "prefix": prefix,
#             "range": [start_no, end_no],
#             "count": len(batch_results),
#             "results": batch_results,
#             "errors": errors  # 仅记录本批内出现的错误
#         }
#         with open(filename, "w", encoding="utf-8") as f:
#             json.dump(payload, f, ensure_ascii=False, indent=2)
#         print(f"💾 已保存批次文件: {filename} (包含 {len(batch_results)} 条，有错误 {len(errors)} 条)")
#         batch_results = []
#         errors.clear()
#         start_index_of_batch = last_idx_inclusive + 1

#     for idx, sample in enumerate(samples):
#         global_idx = idx  # 相对于本列表
#         try:
#             eval_input = EvaluationInput(
#                 query=sample.get("query", ""),
#                 context=sample.get("context", ""),
#                 response=sample.get("response", "")
#             )
#             print(f"正在评估 {prefix}{global_idx+1}...")
#             result = evaluator.evaluate_ragas(eval_input)
#             record = {
#                 "id": f"{prefix}{global_idx+1}",
#                 "metrics": result
#             }
#             if save_inputs:
#                 record["input"] = {
#                     "query": eval_input.query,
#                     "context": eval_input.context,
#                     "response": eval_input.response
#                 }
#             batch_results.append(record)
#         except Exception as e:
#             err_msg = f"{prefix}{global_idx+1} 出错: {e}"
#             print(f"❌ {err_msg}")
#             errors.append(err_msg)
#         # 判断是否到达批次边界
#         if (global_idx + 1) % batch_size == 0:
#             flush_batch(global_idx)

#     # 剩余不足一个 batch 的部分
#     if batch_results:
#         flush_batch(len(samples)-1)

def main():
    load_dotenv(Path(__file__).resolve().parents[1] / ".env")
    API_KEY = os.environ["QWEN_API_KEY"]  # 建议改为环境变量
    evaluator = _RAGEvaluationSuite(API_KEY)
    with open("zzy123/positive_samples_cleaned.json", "r", encoding="utf-8") as f:
        positive_samples = json.load(f)
    with open("zzy123/negative_samples_cleaned.json", "r", encoding="utf-8") as f:
        negative_samples = json.load(f)
    # 每 20 条保存一次，不存输入文本（避免文件过大）；如需保存输入设置 save_inputs=True
    # evaluate_and_save(positive_samples, prefix="p1st", evaluator=evaluator, batch_size=20, save_inputs=False)
    # evaluate_and_save(negative_samples, prefix="n1st", evaluator=evaluator, batch_size=20, save_inputs=False)

    # evaluate_and_save(positive_samples, prefix="p2nd", evaluator=evaluator, batch_size=20, save_inputs=False)
    # evaluate_and_save(negative_samples, prefix="n2nd", evaluator=evaluator, batch_size=20, save_inputs=False)

    # evaluate_and_save(positive_samples, prefix="p3rd", evaluator=evaluator, batch_size=20, save_inputs=False)
    # evaluate_and_save(negative_samples, prefix="n3rd", evaluator=evaluator, batch_size=20, save_inputs=False)

if __name__ == "__main__":
    main()
