"""
自动压缩目录工具

该模块用于在指定的根目录中查找包含关键词的文件夹，并将它们压缩成ZIP文件。
主要功能：
1. 从配置文件读取根目录列表和输出路径
2. 从文本文件读取关键词列表
3. 在根目录中递归查找包含关键词的最上级目录
4. 将找到的目录压缩为ZIP文件并保存到指定位置

用法示例：
    python copy_zip_desktop.py

配置文件 (path_config*.yaml) 格式:
    base_directories:  # 要搜索的根目录列表
        - "D:/path/to/search1"
        - "D:/path/to/search2"
    desktop_output: "D:/output/path"  # ZIP文件输出目录
"""

import logging
import os
import sys
import zipfile
from pathlib import Path

import yaml
from logging_config import setup_logger  # 引入日志配置函数

# 动态添加项目根目录到 sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))


# 初始化日志记录器
logger = setup_logger(
    log_level=logging.INFO,
    log_file="./logs/copy_zip_desktop.log",
)


def read_keyword(filename="./filename.txt"):
    """读取文件中的多个关键词"""
    logger.info("读取关键词文件: %s", filename)
    with open(filename, "r", encoding="utf-8") as f:
        keywords = [line.strip() for line in f.readlines()]
        keywords = [s for s in keywords if s != ""]
        logger.info("找到 %d 个关键词: %s", len(keywords), keywords)
        return keywords


def read_config(config_path="./path_config.yaml"):
    """读取 YAML 配置文件"""
    logger.info("读取配置文件: %s", config_path)
    with open(config_path, "r", encoding="utf-8") as f:
        loaded_config = yaml.safe_load(f)
        logger.info("配置文件加载成功")
        return loaded_config


def find_matching_directories(root_paths, keywords, match_mode="startswith"):
    """
    在多个根目录下查找所有包含关键词的最上级目录，避免收集子目录
    """
    matching_dirs = set()  # 使用集合避免重复
    logger.info(
        "开始在目录 %s 中查找包含关键词: %s 的目录 (模式: %s)",
        root_paths,
        keywords,
        match_mode,
    )

    for root_path in root_paths:
        for keyword in keywords:
            for path in Path(root_path).rglob("*"):
                if path.is_dir():
                    # 获取路径的最后一个部分
                    last_part = path.parts[-1]

                    is_match = False
                    if match_mode == "contains":
                        if keyword in last_part:
                            is_match = True
                    elif match_mode == "startswith":
                        if last_part.startswith(keyword):
                            is_match = True

                    if is_match:
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
                            logger.info("找到包含关键词的最上级目录: %s", matching_dir)

    if not matching_dirs:
        logger.warning("未找到包含任何关键词的目录: %s", keywords)

    return list(matching_dirs)


def zip_directory(directory_path, output_path):
    """将目录压缩为 ZIP 文件"""
    base_name = os.path.basename(directory_path)
    zip_filename = os.path.join(output_path, f"{base_name}.zip")
    logger.info("开始压缩目录: %s 到 %s", directory_path, zip_filename)

    with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(directory_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, start=directory_path)
                zipf.write(file_path, arcname)

    logger.info("目录已成功压缩到: %s", zip_filename)
    return zip_filename


def main():
    """主函数：读取配置、查找目录并执行压缩"""
    try:
        # 根据需要选择配置文件
        is_b24 = "Yes"
        is_b25b26 = "No"

        # 加载配置文件
        app_config = (
            read_config("./path_config_B24.yaml")
            if is_b24 == "Yes"
            else (
                read_config("./path_config_B25B26.yaml")
                if is_b25b26 == "Yes"
                else read_config("./path_config.yaml")
            )
        )

        keywords = read_keyword()  # 获取多个关键词
        base_directories = app_config["base_directories"]  # 读取多个目录
        desktop_output = app_config["desktop_output"]
        match_mode = app_config.get("match_mode", "startswith")  # 默认使用 startswith

        # 确保输出目录存在
        os.makedirs(desktop_output, exist_ok=True)
        logger.info("确保输出目录存在: %s", desktop_output)

        # 在多个目录中查找包含任何一个关键词的所有最上级目录
        matching_dirs = find_matching_directories(
            base_directories, keywords, match_mode
        )

        if matching_dirs:
            # 对所有找到的最上级目录进行压缩
            for target_directory in matching_dirs:
                zip_path = zip_directory(target_directory, desktop_output)
                logger.info("目录 %s 已压缩到 %s", target_directory, zip_path)
        else:
            logger.warning("未找到任何匹配的目录")

    except (OSError, yaml.YAMLError) as e:
        # 指定具体的异常类型，避免捕获过于宽泛的 Exception
        logger.error("发生错误: %s", str(e), exc_info=True)


if __name__ == "__main__":
    main()
