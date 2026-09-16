from pathlib import Path

from dotenv import load_dotenv

# 显式定位项目配置，避免启动目录影响凭据加载；已有环境变量优先。
load_dotenv(Path(__file__).with_name(".env"))


class CONFIG:
    # --------------------------------extract_model相关模型开关---------------------------------
    ## rerank 开关（1：开 ; 0：关）
    EXTRACT_MODEL_RERANK_SWITCH = 1
    ## llm 开关（1：开(没有llm自评估得分的prompt,deepseek模型的prompt); 0：关 ;
    ##          10：基于self-consistency策略的prompt; 11：有llm自评估得分的prompt）
    EXTRACT_MODEL_LLM_SWITCH = 1
    ## 抽取结果附带答案来源（1：开 ; 0：关）
    EXTRACT_MODEL_SOURCE_SWITCH = 1
    ## 复核结果附带答案来源（1：开 ; 0：关）
    CHECK_MODEL_SOURCE_SWITCH = 1
    ## hybrid_search开关（1：开(hybrid_search) ; 0：关）
    EXTRACT_MODEL_HYBRID_SEARCH_SWITCH = 0
    ## embedding_summary 开关选择（0：关闭embedding_summary ; 1：打开embedding_summary）
    EXTRACT_MODEL_SUMMARY_EMBEDDING_TABLE_SWITCH = 1
    ## rerank_summary 开关选择（0：关闭rerank_summary ; 1：打开rerank_summary）
    EXTRACT_MODEL_SUMMARY_RERANK_TABLE_SWITCH = 1
    ## 表格分块器选择（0：用和文本一样的分块器； 1：用更小的分块器）
    TABLE_SUMMARY_SPLITTER = 0

    ## 提取结果评估器（1：开 ; 0：关）
    RAGAS_SWITCH = 0
    RAGAS_THRESHOLD = 0.79
    # --------------------------------extract_model相关模型开关---------------------------------

    # --------------------------------CHUNK----------------------------------------------------
    CHUNK_SEPARATORS = ["\n\n\n", "\n\n", "\n", "。", "！", "？", "，", " "]
    CHUNK_SIZE = 512
    CHUNK_OVERLAP_ratio = 0.2
    # --------------------------------CHUNK----------------------------------------------------

    # --------------------------------LLM_MODEL------------------------------------------------
    LLM_MODEL = "tongyiqianwen72b"  # 普通变量
    # LLM_MODEL = "did100/qwen2.5-32B-Instruct-Q4_K_M:latest"
    LLM_MODEL_inference = (
        None  # 带有推理的变量如p、概率等参数的 不使用括号型的就（None）
    )
    TABLE_SUMMARY_MODEL = "tongyiqianwen72b"
    # LLM_MODEL = "qwen2.5:7B " deepseek-r1:7B deepseek-r1:14b qwen2.5:32b
    SPLIT_LLM_INVOKE = 1
    PROMPT_SELF_CONSISTENCY = False  # 核验任务prompt是否使用自一致性机制
    # --------------------------------LLM_MODEL------------------------------------------------

    # --------------------------------EMBED_MODEL----------------------------------------------

    EMBED_MODEL = "bge-m3_fine:fp16"
    ADAPTER_PATH = None  # 不使用与训练embedding
    # zhy3_6
    USETRANSFORMER = False
    EMBEDDING_MODEL_PATH = r"check_points/bge-m3"

    PEFT_LORA_ADAPTER_PATH = None
    # --------------------------------EMBED_MODEL----------------------------------------------

    # --------------------------------RERANK_MODEL----------------------------------------------
    RERANK_MODEL_PATH = "check_points/ft_bge-reranker-large_20250219_003840"

    EXTRACT_RERANK_TOP_K = 5
    CHECK_RERANK_TOP_K = 60
    TEXT_MAX_LENGTH = 512  # 超过这个长度会被直接裁剪，以防报错（超出的情况往往是目录部分的文本，无伤大雅）

    RRF_SWITCH = False
    RRF_K = 60
    # --------------------------------RERANK_MODEL----------------------------------------------

    # --------------------------------Retrieval------------------------------------------------
    # 1、向量数据库搜寻文档方案
    # SEARCH_TYPE = "similarity"
    # SEARCH_KWARGS = {
    #     "k": 60,
    # }
    # 2、混合检索
    SEARCH_TYPE = "similarity"
    SEARCH_KWARGS = {
        "k": 120,  # 特意取大，保证正确chuck被包含
    }
    EMBEDDING_TOP_k = 60  # 真实embedding环节会输出的topk
    SIMILARITY_RATIO = 3  # 相似度部分的比例，X:1（X=1,2,3...）, X越大，相似度占比更高，关键词检索占比越低

    # CHECK_RESULT_CHUNK_NUM = 10  # 检索结果中，展示的最大 chunk 数量

    # --------------------------------Retrieval------------------------------------------------

    # --------------------------------文件路径和清理-------------------------------- #
    extract_output_folder_path = "output_excels/extract"
    check_output_folder_path = "output_excels/check"
    json_output_folder_path = "output_jsons/table_jsons"  # table存储总结后的json
    json_rerank_folder_path = "output_jsons/rerank_jsons"  # rerank后块存储路径
    logs_output_folder_path = "logs"
    templates_path = "extract_templates"
    template_user_path = "extract_templates/客户级"
    template_system_path = "extract_templates/系统级"
    download_path = "downloads/辅助工具.zip"
    max_age_hours = 24 * 7  # 文件保存时间
    interval_hours = 2  # 清理间隔时间
    # --------------------------------文件路径和清理-------------------------------- #

    # ---------------------------------refine_prompt- --------#
    REFIE_PROMPT_TOP_K = 3

    ## 表格切分开关(True:切分;False:不切分)
    TABLE_CHUNK_SWITCH = True
