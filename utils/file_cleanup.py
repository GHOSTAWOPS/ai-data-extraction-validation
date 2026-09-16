import os
import time
import datetime
import threading
from config import CONFIG
from utils.logger_util import logger


def clean_output_folder(output_folder_path, max_age_hours=24):
    """
    根据文件名中的时间戳清理文件夹

    :param output_folder_path: 输出文件夹路径
    :param max_age_hours: 文件最大保存时间（小时）
    """
    current_time = datetime.datetime.now()

    for filename in os.listdir(output_folder_path):
        try:
            full_path = os.path.join(output_folder_path, filename)

            # 处理日志文件（log_YYYY-MM-DD_HH-MM-SS.log 格式）
            if filename.startswith("log_") and filename.endswith(".log"):
                # 提取日志文件名中的时间戳部分
                timestamp_str = filename.split(".")[0]  # 去掉扩展名
                timestamp_str = timestamp_str[4:]  # 去掉 'log_'

                # 提取日期时间部分
                try:
                    file_time = datetime.datetime.strptime(
                        timestamp_str, "%Y-%m-%d_%H-%M-%S"
                    )
                except ValueError as e:
                    logger.error(f"文件名时间戳解析失败: {filename} - 错误: {e}")
                    continue

                # 计算文件存在时间
                time_diff = current_time - file_time
                age_hours = time_diff.total_seconds() / 3600

                # 如果文件超过指定时间，删除
                if age_hours > max_age_hours:
                    try:
                        if os.path.isfile(full_path):
                            os.remove(full_path)
                        else:
                            os.rmdir(full_path)
                        logger.info(f"清理线程删除: {filename}")
                    except Exception as e:
                        logger.error(f"清理线程删除 {filename} 时发生错误: {e}")

            # 处理其他文件类型
            elif filename.startswith("output_") and filename.endswith(
                (".xlsx", ".docx", ".pdf", ".doc", ".txt")
            ):
                # 移除文件扩展名
                timestamp_str = filename.split(".")[0]

                # 提取时间部分：20241119_203943
                timestamp_str = (
                    timestamp_str.split("_")[1] + "_" + timestamp_str.split("_")[2]
                )

                # 解析时间
                try:
                    file_time = datetime.datetime.strptime(
                        timestamp_str, "%Y%m%d_%H%M%S"
                    )
                except ValueError as e:
                    logger.error(f"文件名时间戳解析失败: {filename} - 错误: {e}")
                    continue

                # 计算文件存在时间
                time_diff = current_time - file_time
                age_hours = time_diff.total_seconds() / 3600

                # 如果文件超过指定时间，删除
                if age_hours > max_age_hours:
                    try:
                        if os.path.isfile(full_path):
                            os.remove(full_path)
                        else:
                            os.rmdir(full_path)
                        logger.info(f"清理线程删除: {filename}")
                    except Exception as e:
                        logger.error(f"清理线程删除 {filename} 时发生错误: {e}")
        except Exception as e:
            logger.error(f"清理线程处理 {filename} 时发生错误: {e}")


def continuous_cleanup(max_age_hours=24, interval_hours=1):
    """
    在后台线程中启动文件清理

    :param output_folder_path: 输出文件夹路径
    :param max_age_hours: 文件最大保存时间（小时）
    :param interval_hours: 清理间隔（小时）
    #"""
    while True:
        try:
            clean_output_folder(CONFIG.check_output_folder_path, max_age_hours)
            clean_output_folder(CONFIG.extract_output_folder_path, max_age_hours)
            clean_output_folder(CONFIG.logs_output_folder_path, max_age_hours)
        except Exception as e:
            logger.error(f"清理线程清理过程中发生错误: {e}")

        # 休眠指定的时间间隔
        time.sleep(interval_hours * 3600)


if __name__ == "__main__":
    # 在Gradio启动前调用
    output_folder_path = "/home/pkr/projects/luShan/logs"  # 替换为你的输出文件夹路径
    continuous_cleanup(output_folder_path, max_age_hours=72, interval_hours=1)
