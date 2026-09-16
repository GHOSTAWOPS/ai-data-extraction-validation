import re
from .utils import *
from .context import *

def annotate_hydro_frequency(text: str) -> str:
    """
    对输入文本中的“年一遇”或“p=...%”格式进行识别与转换，
    在末尾添加等价备注（例如：200年一遇 等价于 p=0.5%）。
    """
    comment = ""

    # 模式1：识别“200年一遇”或“两百年发生一次”
    match1 = re.search(r"(?:(\d+)|([一二三四五六七八九十百千两]{1,8}))年(?:一遇|发生一次|发生1次|发生过一次)", text)
    if match1:
        num = match1.group(1)
        cn = match1.group(2)
        if cn:
            T = chinese_to_number(cn)
        else:
            T = int(num)
        if T:
            p = round(compute_p(T), 4)
            comment = f"（{T}年一遇 等价于 p={p}%）"

    # 模式2：识别“p = 0.05%”类表达
    match2 = re.search(r"[pP]\s*=\s*(\d+\.?\d*)\s*%", text)
    if match2:
        p_val = float(match2.group(1))
        if is_hydrology_related(text):
            T = int(round(compute_T(p_val)))
            comment = f"（p={p_val}% 等价于 {T}年一遇）"

    if comment and not text.strip().endswith(comment):
        return text.strip() + comment
    return text