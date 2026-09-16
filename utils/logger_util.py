import logging
import os
from datetime import datetime

def setup_logger():
    # 创建日志目录
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)

    # 使用时间戳生成唯一日志文件名
    log_file = os.path.join(log_dir, f"log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log")

    # 配置日志记录器
    logger = logging.getLogger("my_logger")
    logger.setLevel(logging.DEBUG)  # 设置日志级别

    # 日志格式
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    # 文件处理器
    file_handler = logging.FileHandler(log_file)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # 控制台处理器
    # console_handler = logging.StreamHandler()
    # console_handler.setFormatter(formatter)
    # logger.addHandler(console_handler) 

    return logger

# 全局日志实例
logger = setup_logger()
