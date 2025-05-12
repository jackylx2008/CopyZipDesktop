import logging
import os
import sys
import zipfile
from pathlib import Path

import yaml

# 动态添加项目根目录到 sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from logging_config import setup_logger  # 引入日志配置函数

# 初始化日志记录器
logger = setup_logger(log_level=logging.INFO, log_file="./logs/copy_zip_desktop.log")


def read_keyword(filename="./filename.txt"):
    """读取文件中的多个关键词"""
    logger.info(f"读取关键词文件: {filename}")
    with open(filename, "r", encoding="utf-8") as f:
        keywords = [line.strip() for line in f.readlines()]
        logger.info(f"找到 {len(keywords)} 个关键词: {keywords}")
        return keywords


def read_config(config_path="./path_config.yaml"):
    """读取 YAML 配置文件"""
    logger.info(f"读取配置文件: {config_path}")
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
        logger.info("配置文件加载成功")
        return config


def find_matching_directories(root_paths, keywords):
    """
    在多个根目录下查找所有包含关键词的最上级目录，避免收集子目录
    """

    matching_dirs = set()  # 使用集合避免重复
    logger.info(f"开始在目录 {root_paths} 中查找包含关键词: {keywords} 的目录")

    for root_path in root_paths:
        for keyword in keywords:
            for path in Path(root_path).rglob("*"):
                if path.is_dir():
                    # 获取路径的最后一个部分
                    last_part = path.parts[-1]
                    if keyword in last_part:
                        matching_dir = str(path.resolve())
                        # 检查是否已经包含了这个目录的任何父目录
                        if not any(
                            matching_dir.startswith(str(existing))
                            for existing in matching_dirs
                        ):
                            # 检查是否已经包含了这个目录的任何子目录
                            matching_dirs = {
                                existing
                                for existing in matching_dirs
                                if not str(existing).startswith(matching_dir)
                            }
                            matching_dirs.add(matching_dir)
                            logger.info(f"找到包含关键词的最上级目录: {matching_dir}")

    if not matching_dirs:
        logger.warning(f"未找到包含任何关键词的目录: {keywords}")

    return list(matching_dirs)


def zip_directory(directory_path, output_path):
    """将目录压缩为 ZIP 文件"""
    zip_filename = os.path.join(output_path, os.path.basename(directory_path) + ".zip")
    logger.info(f"开始压缩目录: {directory_path} 到 {zip_filename}")
    with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(directory_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, start=directory_path)
                zipf.write(file_path, arcname)
    logger.info(f"目录已成功压缩到: {zip_filename}")
    return zip_filename


def main():

    is_B24 = "Yes"
    is_B25B26 = "No"

    if is_B24 == "Yes":
        config = read_config("./path_config_B24.yaml")

    if is_B25B26 == "Yes":
        config = read_config("./path_config_B25B26.yaml")
    try:
        keywords = read_keyword()  # 获取多个关键词

        base_directories = config["base_directories"]  # 读取多个目录
        desktop_output = config["desktop_output"]

        # 确保输出目录存在
        os.makedirs(desktop_output, exist_ok=True)
        logger.info(f"确保输出目录存在: {desktop_output}")

        # 在多个目录中查找包含任何一个关键词的所有最上级目录
        matching_dirs = find_matching_directories(base_directories, keywords)

        if matching_dirs:
            # 对所有找到的最上级目录进行压缩
            for target_directory in matching_dirs:
                zip_path = zip_directory(target_directory, desktop_output)
                logger.info(f"目录 {target_directory} 已压缩到 {zip_path}")
        else:
            logger.warning(f"未找到任何匹配的目录")

    except Exception as e:
        logger.error(f"发生错误: {e}", exc_info=True)


if __name__ == "__main__":
    main()
