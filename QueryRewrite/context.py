import jieba


HYDRO_KEYWORDS = {
    "洪水", "设计标准", "重现期", "防洪", "设计流量", 
    "洪水位", "水文", "年一遇", "设计洪峰", "频率"
}

def is_hydrology_related(text: str) -> bool:
    """
    判断文本中是否包含水利相关关键词，用于判断概率语境是否为洪水频率。
    """
    words = set(jieba.lcut(text))
    return any(kw in words for kw in HYDRO_KEYWORDS)