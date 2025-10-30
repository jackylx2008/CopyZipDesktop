"""
自动拷贝 PDF 和 DOCX 文件工具

该模块用于根据文件名中的关键词匹配来拷贝 PDF 和 DOCX 文件。
主要功能：
1. 从配置文件读取源目录列表和目标目录
2. 从文本文件读取关键词用于匹配
3. 自动递归查找并拷贝匹配的 PDF 和 DOCX 文件
4. 支持配置拷贝深度和匹配模式

用法示例：
    python copy_pdf_docx_desktop.py

配置文件 (path_config*.yaml) 格式：
    pdf_directories:  # PDF文件搜索目录
        - "D:/path/to/pdf1"
        - "D:/path/to/pdf2"
    docx_directories:  # DOCX文件搜索目录
        - "D:/path/to/docx1"
    desktop_output: "D:/output/path"  # 文件输出目录
"""

import logging
import os
import re
import shutil
import sys
from pathlib import Path

import yaml

# 动态添加项目根目录到 sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# logging_config 可能在项目根目录，通过 E402 注释避免导入顺序检查报告
from logging_config import setup_logger  # noqa: E402  # 引入日志配置函数

# 初始化日志记录器
logger = setup_logger(
    log_level=logging.INFO, log_file="./logs/copy_pdf_docx_desktop.log"
)


def read_config(config_path="./path_config.yaml"):
    """读取 YAML 配置文件"""
    logger.info("读取配置文件: %s", config_path)
    with open(config_path, "r", encoding="utf-8") as f:
        loaded_config = yaml.safe_load(f)
        logger.info("配置文件加载成功")
        return loaded_config


def extract_strings_from_filename(filename, regex_patterns):
    """从文件名中提取匹配的字符串并返回一个 set"""
    result_set = set()
    for pattern in regex_patterns:
        try:
            found_matches = re.findall(pattern, filename)
            if found_matches:
                for match in found_matches:
                    if isinstance(match, tuple):
                        # 过滤空组，然后用 - 连接
                        parts = [m for m in match if m]
                        full_match = "-".join(parts)
                    else:
                        full_match = match
                    full_match = full_match.strip()
                    result_set.add(full_match)
                    logger.info(
                        "匹配到的完整字符: %s 在文件名: %s",
                        full_match,
                        filename,
                    )
        except re.error as e:
            logger.error("正则表达式错误: %s", str(e))
    return result_set


def process_docx_files(config):
    """遍历 docx 文件并提取文件名中的字符串"""
    docx_directories = config.get("docx_directories", [])
    regex_patterns = config.get("regex_pattern", [])

    if not regex_patterns:
        logger.error("配置文件中未指定正则表达式！")
        return set()  # 如果没有配置正则表达式，则返回空的set

    all_matches = set()  # 用于存储所有匹配到的字符串

    for directory in docx_directories:
        logger.info("开始遍历目录: %s", directory)
        for file_path in Path(directory).rglob("*.docx"):  # 查找所有 docx 文件
            logger.info("处理文件: %s", file_path)
            filename = file_path.name
            found = extract_strings_from_filename(filename, regex_patterns)
            all_matches.update(found)  # 将匹配结果添加到集合中

    return all_matches


def copy_files(
    config,
    file_matches,
    file_extension,
    file_type,
    desktop_output,
    depth=None,
):
    """复制匹配的文件到目标目录

    Args:
        config: 配置字典
        file_matches: 要匹配的字符串集合
        file_extension: 文件扩展名
        file_type: 文件类型配置项名称
        desktop_output: 输出目录路径
        depth: 递归深度限制
    """
    directories = config.get(file_type, [])
    # 匹配模式: contains|startswith|regex
    match_mode = config.get("match_mode", "contains")

    if not desktop_output:
        logger.error("配置文件中未指定目标目录 (desktop_output)！")
        return

    if not os.path.exists(desktop_output):
        logger.info("目标目录 %s 不存在，正在创建...", desktop_output)
        os.makedirs(desktop_output)

    all_copied_files = set()

    for directory in directories:
        logger.info("开始遍历%s目录: %s", file_type, directory)
        for root, _, files in os.walk(directory):
            if depth is not None:
                depth_calc = len(directory)
                current_depth = root[depth_calc:].count(os.sep)
                if current_depth >= depth:
                    continue

            for file in files:
                if not file.lower().endswith(f".{file_extension}"):
                    continue
                file_path = Path(root) / file

                for match_str in file_matches:
                    do_copy = False
                    if match_mode == "contains":
                        if match_str in file:
                            do_copy = True
                    elif match_mode == "startswith":
                        # 要求文件名以 match_str 开头（忽略扩展名）
                        name_only = os.path.splitext(file)[0]
                        if name_only.startswith(match_str):
                            do_copy = True
                    elif match_mode == "regex":
                        try:
                            if re.search(match_str, file):
                                do_copy = True
                        except re.error:
                            # 如果 match_str 不是合法正则则回退到 contains
                            if match_str in file:
                                do_copy = True

                    if do_copy:
                        destination_path = os.path.join(desktop_output, file)
                        if not os.path.exists(destination_path):
                            logger.info(
                                "正在拷贝文件: %s 到 %s",
                                file_path,
                                destination_path,
                            )
                            shutil.copy(file_path, destination_path)
                            all_copied_files.add(destination_path)
                        else:
                            logger.info("文件已存在: %s", destination_path)
                        break

    logger.info(
        "共拷贝了 %d 个%s文件到目标目录.",
        len(all_copied_files),
        file_type,
    )
    return all_copied_files


def copy_pdf_files(config, match_list, depth=None):
    """根据匹配的字符串拷贝 PDF 文件"""
    desktop_output = config.get("desktop_output", "")
    return copy_files(
        config,
        match_list,
        "pdf",
        "pdf_directories",
        desktop_output,
        depth,
    )


def copy_docx_files(config, match_list, depth=None):
    """根据匹配的字符串拷贝 DOCX 文件"""
    desktop_output = config.get("desktop_output", "")
    return copy_files(
        config,
        match_list,
        "docx",
        "docx_directories",
        desktop_output,
        depth,
    )


if __name__ == "__main__":
    IS_B24 = "Yes"
    IS_B25B26 = "No"

    # 根据选择读取对应的配置文件
    CONFIG_FILE = "./path_config_B24.yaml" if IS_B24 == "Yes" else None
    if not CONFIG_FILE:
        CONFIG_FILE = "./path_config_B25B26.yaml"
    YAML_PATH = CONFIG_FILE
    cfg = read_config(YAML_PATH)

    # 处理文件并记录匹配结果
    matches = process_docx_files(cfg)
    logger.info("所有匹配到的完整字符串: %s", matches)

    # 拷贝文件，最大深度设置为3
    copied_pdfs = copy_pdf_files(cfg, matches, depth=3)
    logger.info("拷贝的PDF文件: %s", copied_pdfs)

    copied_docx = copy_docx_files(cfg, matches, depth=3)
    logger.info("拷贝的DOCX文件: %s", copied_docx)
