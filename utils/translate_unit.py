'''
find_matching_unit:
输入：单位列的字符串
输出：根据文件unit.txt中的标准单位格式，删除单位列多余的文本
'''

import random
import re
import time


def find_matching_unit(input_string, units):
    """根据给定的单位列表识别字符串中的单位"""
    matched_units = []

    for unit in units:
        # 构建正则表达式模式
        pattern = ''
        current_type = None  # 用于跟踪当前字符类型（ASCII或非ASCII）

        for char in unit:
            is_ascii = char.isascii()

            # 如果字符类型发生变化，不添加可选的标点符号
            if current_type is not None and current_type != is_ascii:
                pattern += f"\\W?"

            if is_ascii:
                pattern += f"\\W?{re.escape(char.lower())}"
            else:
                pattern += re.escape(char)

            current_type = is_ascii

        # 编译正则表达式，忽略大小写
        regex = re.compile(pattern, re.IGNORECASE)
        if regex.search(input_string):
            matched_units.append(unit)

    if matched_units:
        # 选择最长的匹配单位
        selected_unit = max(matched_units, key=lambda x: len(x))
        return selected_unit
    else:
        return None  # 匹配失败返回 None


def generate_test_case(true_value, units):
    """生成包含随机前后缀的测试用例"""

    # 前缀后缀随机产生
    prefix = generate_random_text(random.randint(0, 10))  # 0~10 个随机字符
    suffix = generate_random_text(random.randint(0, 10))  # 0~10 个随机字符


    # 只有前缀
    # prefix = generate_random_text(random.randint(1, 10))  # 1~10 个随机字符
    # 只有后缀
    # suffix = generate_random_text(random.randint(1, 10))  # 1~10 个随机字符


    return f"{prefix}{true_value}{suffix}"
    # return f"{prefix}{true_value}"
    # return f"{true_value}{suffix}"




def generate_random_text(length=5):
    """生成随机长度的无意义文字"""
    keyboard_characters = (
        "的一我他这中大为上国发以会可主就来打到别看安静和方便是否可见发布违法无法反对符号表示服用情话同们时说现向相同外长要过本先着可也里后但还提就可以些你们非比于代只着长车过高中总错产受科些进从用向声解所关得行办体个看听平知识方身自员留专没调合近得后明面高质通没关动非新外相合使自过日标精东行无户消专更是知感长架否心全真基管让全设例化空交上映完张体转会保温决同施行集见传定达根考时统也以术都支聚眼性例愿低强西位关于正线完客近内方口空跟下元口条百门销政制来民给产示上工达论张拓或基中工同点听突看老客立已访项走放度从光氏世如点当复段进速前点子期知性新访光区声所管压则三安否话问提料量科启天德设候户示广投分设必机保举好价齐务社年该位补通面造电错提启听达议形应开持向销节达受知列新本远头身本项言"
    )
    return "".join(random.choices(keyboard_characters, k=length))


def process_test_cases(units, num_tests=1000):
    """进行测试并计算正确率"""
    correct_count = 0
    start_time = time.time()
    for _ in range(num_tests):
        # 从单位表中随机选择一个单位作为真值
        true_value = random.choice(units)

        # 生成包含真值的测试用例
        test_case = generate_test_case(true_value, units)

        # 识别测试用例中的单位
        
        result = find_matching_unit(test_case, units)

        # 判断识别结果是否与真值相同
        if result == true_value:
            correct_count += 1
        # else: print(true_value,test_case,result)
        else: print(f"True Value: {true_value} *** Test Case: {test_case} *** Result: {result}")
    


    # 计算正确率
    accuracy = (correct_count / num_tests) * 100
    print(f"程序的正确率: {accuracy:.2f}%")
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"代码运行时间: {elapsed_time:.6f} 秒")
    # return accuracy


if __name__ == "__main__":
    try:
        # 读取单位列表
        with open('unit_2.txt', 'r', encoding='utf-8') as file:
            units = [line.strip() for line in file]

        # 进行测试并计算正确率

        # 初始化列表用于存储 accuracy 值
        # accuracies = []

        # # 运行 100 次 process_test_cases
        # for _ in range(100):
        #     accuracy = process_test_cases(units)  # 调用函数
        #     accuracies.append(accuracy)          # 将结果存入列表

        # # 计算最大值和最小值
        # max_accuracy = max(accuracies)
        # min_accuracy = min(accuracies)

        # # 打印结果
        # print("Maximum accuracy:", max_accuracy)
        # print("Minimum accuracy:", min_accuracy)

        process_test_cases(units)
    except FileNotFoundError:
        print("unit.txt 文件未找到。")

