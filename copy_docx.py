from logging_config import setup_logger  # 引入日志配置函数
import logging
import os
import shutil
import sys

import yaml

# 动态添加项目根目录到 sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))


# 初始化日志记录器
logger = setup_logger(
    log_level=logging.INFO, log_file="./logs/copy_pdf_docx_desktop.log"
)


def read_config(config_path="./docx.yaml"):
    """读取 YAML 配置文件"""
    logger.info(f"读取配置文件: {config_path}")
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
        logger.info("配置文件加载成功")
        return config


def read_input_txt(input_txt_path):
    """读取 input_txt 文件，每行作为关键字"""
    if not os.path.exists(input_txt_path):
        logger.error(f"输入文件 {input_txt_path} 不存在")
        return []
    with open(input_txt_path, "r", encoding="utf-8") as f:
        keywords = [line.strip() for line in f if line.strip()]
    logger.info(f"读取 {len(keywords)} 个关键字")
    return keywords


def find_and_copy_docx(keywords, docx_dirs, output_dir, match_mode="startswith"):
    """遍历 docx_dirs，找到包含关键字的 docx 文件，并拷贝到 output_dir"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        logger.info(f"创建目标目录: {output_dir}")

    copied_files = 0
    unmatched_keywords = set(keywords)  # 记录未匹配的关键字

    for directory in docx_dirs:
        if not os.path.exists(directory):
            logger.warning(f"目录 {directory} 不存在，跳过")
            continue

        for root, _, files in os.walk(directory):
            for file in files:
                if file.endswith(".docx"):
                    for keyword in keywords:
                        is_match = False
                        if match_mode == "contains":
                            if keyword in file:
                                is_match = True
                        elif match_mode == "startswith":
                            if file.startswith(keyword):
                                is_match = True

                        if is_match:
                            src_path = os.path.join(root, file)
                            dest_path = os.path.join(output_dir, file)
                            shutil.copy2(src_path, dest_path)
                            logger.info(f"复制文件: {src_path} -> {dest_path}")
                            copied_files += 1
                            if keyword in unmatched_keywords:
                                unmatched_keywords.remove(keyword)
                            break  # 找到一个匹配就处理下一个文件

    if unmatched_keywords:
        logger.warning(f"未匹配到的关键字: {', '.join(unmatched_keywords)}")

    logger.info(f"总共复制 {copied_files} 个文件")


def main():
    config = read_config()
    desktop_output = config.get("desktop_output")
    docx_dirs = config.get("docx_directories", [])
    input_txt_path = config.get("input_txt")

    if not desktop_output or not docx_dirs or not input_txt_path:
        logger.error(
            "配置文件缺少必要的字段: desktop_output, docx_directories, input_txt"
        )
        return

    keywords = read_input_txt(input_txt_path)
    if not keywords:
        logger.warning("输入文件中没有关键字")
        return

    match_mode = config.get("match_mode", "startswith")
    find_and_copy_docx(keywords, docx_dirs, desktop_output, match_mode)


if __name__ == "__main__":
    main()
