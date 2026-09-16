def chinese_to_number(cn: str) -> int:
    """
    将中文数字字符串（如“两千五百”）转换为阿拉伯数字（如2500）。
    支持千、百、十的组合，仅处理自然数。
    """
    cn_num = {'零': 0, '一': 1, '二': 2, '两': 2, '三': 3, '四': 4,
              '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
    cn_unit = {'千': 1000, '百': 100, '十': 10}

    result = 0
    unit = 1
    num = 0

    i = len(cn) - 1
    while i >= 0:
        c = cn[i]
        if c in cn_unit:
            unit = cn_unit[c]
            if i > 0 and cn[i-1] in cn_num:
                num = cn_num[cn[i-1]]
                result += num * unit
                i -= 1
            else:
                result += 1 * unit
        elif c in cn_num:
            result += cn_num[c] * unit
        i -= 1
    return result

def compute_p(T: int) -> float:
    """
    根据重现期 T（年）计算年超越概率 p（百分数）。
    例如：T=200 → p=0.5
    """
    return 100 / T


def compute_T(p: float) -> float:
    """
    根据年超越概率 p（百分数）计算重现期 T（年）。
    例如：p=0.5 → T=200
    """
    return 100 / p