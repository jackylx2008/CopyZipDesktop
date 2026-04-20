"""
自动拷贝 PDF 文件工具

该模块用于根据文件名列表中的关键词或文件名前缀，从配置的 PDF 源目录中拷贝 PDF 文件到桌面输出目录。

用法示例：
    python copy_pdf_desktop.py

配置文件 (path_config*.yaml) 格式：
    pdf_directories:
        - "D:/path/to/pdf1"
        - "D:/path/to/pdf2"
    desktop_output: "D:/output/path"
    input_txt: "./filename.txt"
    match_mode: "startswith"  # 可选 contains|startswith|regex
"""

import logging
import os
import re
import shutil
import sys
from pathlib import Path

import yaml

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from logging_config import setup_logger  # noqa: E402

logger = setup_logger(log_level=logging.INFO, log_file="./logs/copy_pdf_desktop.log")


def read_config(config_path="./path_config.yaml"):
    """读取 YAML 配置文件"""
    logger.info("读取配置文件: %s", config_path)
    with open(config_path, "r", encoding="utf-8") as f:
        loaded_config = yaml.safe_load(f)
        logger.info("配置文件加载成功")
        return loaded_config


def read_filename_list(input_txt_path):
    """读取目标 PDF 名称或匹配关键字列表"""
    path = Path(input_txt_path)
    if not path.exists():
        logger.error("目标文件列表不存在: %s", input_txt_path)
        return []

    with path.open("r", encoding="utf-8") as f:
        lines = [line.strip() for line in f if line.strip()]

    logger.info("从文件读取到 %d 条目标名称", len(lines))
    return lines


def copy_pdf_files(config, targets, depth=None):
    """复制匹配的 PDF 文件到目标目录"""
    directories = config.get("pdf_directories", [])
    desktop_output = config.get("desktop_output", "")
    match_mode = config.get("match_mode", "startswith")

    if not desktop_output:
        logger.error("配置文件中未指定目标目录 (desktop_output)！")
        return set()

    if not os.path.exists(desktop_output):
        logger.info("目标目录 %s 不存在，正在创建...", desktop_output)
        os.makedirs(desktop_output, exist_ok=True)

    copied_files = set()

    for directory in directories:
        logger.info("开始遍历 PDF 源目录: %s", directory)
        for root, _, files in os.walk(directory):
            if depth is not None:
                depth_calc = len(directory)
                current_depth = root[depth_calc:].count(os.sep)
                if current_depth >= depth:
                    continue

            for file in files:
                if not file.lower().endswith(".pdf"):
                    continue
                file_path = Path(root) / file
                name_only = os.path.splitext(file)[0]

                for target in targets:
                    if not target:
                        continue

                    do_copy = False
                    if match_mode == "contains":
                        if target in file:
                            do_copy = True
                    elif match_mode == "regex":
                        try:
                            if re.search(target, file):
                                do_copy = True
                        except re.error:
                            if target in file:
                                do_copy = True
                    else:
                        # 默认使用 startswith
                        if name_only.startswith(target):
                            do_copy = True

                    if do_copy:
                        destination_path = Path(desktop_output) / file
                        if not destination_path.exists():
                            logger.info(
                                "正在拷贝: %s -> %s", file_path, destination_path
                            )
                            shutil.copy(file_path, destination_path)
                            copied_files.add(str(destination_path))
                        else:
                            logger.info("目标文件已存在: %s", destination_path)
                        break

    logger.info("共拷贝了 %d 个 PDF 文件到目标目录", len(copied_files))
    return copied_files


if __name__ == "__main__":
    IS_B24 = "Yes"
    IS_B25B26 = "No"

    CONFIG_FILE = "./path_config_B24.yaml" if IS_B24 == "Yes" else None
    if not CONFIG_FILE:
        CONFIG_FILE = "./path_config_B25B26.yaml"

    cfg = read_config(CONFIG_FILE)
    input_txt = cfg.get("input_txt", "./filename.txt")
    if not Path(input_txt).exists():
        input_txt = "./filename.txt"

    targets = read_filename_list(input_txt)
    copied_pdfs = copy_pdf_files(cfg, targets, depth=3)
    logger.info("拷贝完成，PDF 文件列表: %s", copied_pdfs)
