import logging
import os
import re
import shutil
import sys
from pathlib import Path

import yaml

# 动态添加项目根目录到 sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from logging_config import setup_logger  # 引入日志配置函数

# 初始化日志记录器
logger = setup_logger(
    log_level=logging.INFO, log_file="./logs/copy_pdf_docx_desktop.log"
)


def read_config(config_path="./path_config.yaml"):
    """读取 YAML 配置文件"""
    logger.info(f"读取配置文件: {config_path}")
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
        logger.info("配置文件加载成功")
        return config


def extract_strings_from_filename(filename, regex_patterns):
    """从文件名中提取匹配的字符串并返回一个set"""
    matches = set()  # 使用set保存匹配结果，避免重复
    for pattern in regex_patterns:  # 遍历每个正则表达式
        try:
            found_matches = re.findall(pattern, filename)
            if found_matches:
                # 将元组的元素合并成一个完整字符串
                for match in found_matches:
                    full_match = "-".join(match)
                    matches.add(full_match)  # 将合并后的完整字符串添加到set中
                    logger.info(f"匹配到的完整字符: {full_match}  在文件名: {filename}")
        except re.error as e:
            logger.error(f"正则表达式错误: {e}")
    return matches


def process_docx_files(config):
    """遍历 docx 文件并提取文件名中的字符串"""
    docx_directories = config.get("docx_directories", [])
    regex_patterns = config.get("regex_pattern", [])

    if not regex_patterns:
        logger.error("配置文件中未指定正则表达式！")
        return set()  # 如果没有配置正则表达式，则返回空的set

    all_matches = set()  # 用于存储所有匹配到的字符串

    for directory in docx_directories:
        logger.info(f"开始遍历目录: {directory}")
        for file_path in Path(directory).rglob("*.docx"):  # 查找所有 docx 文件
            logger.info(f"处理文件: {file_path}")
            filename = file_path.name
            matches = extract_strings_from_filename(filename, regex_patterns)
            all_matches.update(matches)  # 将文件匹配的结果添加到all_matches集合中

    return all_matches


def copy_files(config, matches, file_extension, file_type, desktop_output, depth=None):
    """根据匹配到的字符串，从指定目录中查找并拷贝文件到desktop_output"""
    directories = config.get(file_type, [])

    if not desktop_output:
        logger.error("配置文件中未指定目标目录 (desktop_output)！")
        return

    if not os.path.exists(desktop_output):
        logger.info(f"目标目录 {desktop_output} 不存在，正在创建...")
        os.makedirs(desktop_output)

    all_copied_files = set()  # 用于存储已拷贝的文件路径

    for directory in directories:
        logger.info(f"开始遍历{file_type}目录: {directory}")

        for root, dirs, files in os.walk(directory):  # 使用os.walk()遍历所有目录和文件
            if depth is not None:
                # 如果设置了深度，限制遍历的深度
                current_depth = root[len(directory) :].count(os.sep)
                if current_depth >= depth:
                    continue

            for file in files:
                if file.endswith(f".{file_extension}"):  # 只处理指定扩展名的文件
                    file_path = Path(root) / file
                    for match in matches:
                        if match in file:
                            destination_path = os.path.join(desktop_output, file)
                            if not os.path.exists(destination_path):
                                logger.info(
                                    f"正在拷贝文件: {file_path} 到 {destination_path}"
                                )
                                shutil.copy(file_path, destination_path)
                                all_copied_files.add(destination_path)
                            else:
                                logger.info(f"文件已存在: {destination_path}")
                            break

    logger.info(f"共拷贝了 {len(all_copied_files)} 个{file_type}文件到目标目录.")
    return all_copied_files


def copy_pdf_files(config, matches, depth=None):
    """根据匹配的字符串拷贝 PDF 文件"""
    desktop_output = config.get("desktop_output", "")
    return copy_files(config, matches, "pdf", "pdf_directories", desktop_output, depth)


def copy_docx_files(config, matches, depth=None):
    """根据匹配的字符串拷贝 DOCX 文件"""
    desktop_output = config.get("desktop_output", "")
    return copy_files(
        config, matches, "docx", "docx_directories", desktop_output, depth
    )


if __name__ == "__main__":
    is_B24 = "Yes"
    is_B25B26 = "No"

    if is_B24 == "Yes":
        config = read_config("./path_config_B24.yaml")

    if is_B25B26 == "Yes":
        config = read_config("./path_config_B25B26.yaml")
    matches = process_docx_files(config)  # 处理 docx 文件并返回所有匹配项
    logger.info(f"所有匹配到的完整字符串: {matches}")

    copied_pdfs = copy_pdf_files(
        config, matches, depth=3
    )  # 根据匹配的字符串拷贝 PDF 文件，设置最大深度为3

    copied_docx = copy_docx_files(
        config, matches, depth=3
    )  # 根据匹配的字符串拷贝 DOCX 文件，设置最大深度为3
    logger.info(f"拷贝的DOCX文件: {copied_docx}")
