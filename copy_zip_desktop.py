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
        raw_keywords = [line.strip() for line in f.readlines()]
        raw_keywords = [s for s in raw_keywords if s != ""]

        duplicate_keywords = []
        deduplicated_keywords = []
        seen_keywords = set()

        for keyword in raw_keywords:
            if keyword in seen_keywords:
                if keyword not in duplicate_keywords:
                    duplicate_keywords.append(keyword)
                continue

            seen_keywords.add(keyword)
            deduplicated_keywords.append(keyword)

        if duplicate_keywords:
            logger.warning(
                "filename.txt 存在重复编号，已按行去重。重复的字符串: %s",
                ", ".join(duplicate_keywords),
            )
        else:
            logger.info("filename.txt 未发现重复编号")

        logger.info("filename.txt 原始非空编号数量: %d", len(raw_keywords))
        logger.info("filename.txt 去重后真实编号数量: %d", len(deduplicated_keywords))
        logger.info("去重后关键词列表: %s", deduplicated_keywords)
        return deduplicated_keywords


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


def find_directory_for_keyword(root_paths, keyword, match_mode="startswith"):
    """为单个关键词查找匹配目录。"""
    matching_dirs = find_matching_directories(root_paths, [keyword], match_mode)
    if not matching_dirs:
        return None

    if len(matching_dirs) > 1:
        logger.warning(
            "关键词 %s 匹配到多个目录，将使用第一个目录: %s",
            keyword,
            matching_dirs[0],
        )

    return matching_dirs[0]


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

        input_txt = app_config.get("input_txt", "./filename.txt")
        keywords = read_keyword(input_txt)  # 获取多个关键词
        base_directories = app_config["base_directories"]  # 读取多个目录
        desktop_output = app_config["desktop_output"]
        match_mode = app_config.get("match_mode", "startswith")  # 默认使用 startswith

        # 确保输出目录存在
        os.makedirs(desktop_output, exist_ok=True)
        logger.info("确保输出目录存在: %s", desktop_output)

        with open(input_txt, "r", encoding="utf-8") as f:
            raw_keywords = [line.strip() for line in f.readlines()]
            raw_keywords = [s for s in raw_keywords if s != ""]

        total_keywords = len(raw_keywords)
        unique_keywords_count = len(keywords)
        success_keywords = []
        failed_keywords = []
        generated_zip_paths = []
        processed_directories = set()

        for keyword in keywords:
            target_directory = find_directory_for_keyword(
                base_directories, keyword, match_mode
            )
            if not target_directory:
                failed_keywords.append(keyword)
                logger.warning("编号 %s 未找到匹配目录，未生成 ZIP", keyword)
                continue

            try:
                if target_directory in processed_directories:
                    zip_path = os.path.join(
                        desktop_output,
                        f"{os.path.basename(target_directory)}.zip",
                    )
                    logger.info(
                        "编号 %s 对应目录已处理，复用现有 ZIP: %s",
                        keyword,
                        zip_path,
                    )
                else:
                    zip_path = zip_directory(target_directory, desktop_output)
                    processed_directories.add(target_directory)
                    generated_zip_paths.append(zip_path)
                    logger.info("目录 %s 已压缩到 %s", target_directory, zip_path)

                success_keywords.append(keyword)
            except OSError as e:
                failed_keywords.append(keyword)
                logger.error(
                    "编号 %s 生成 ZIP 失败: %s", keyword, str(e), exc_info=True
                )

        if not success_keywords and not failed_keywords:
            logger.warning("未找到任何匹配的目录")

        logger.info("========== ZIP 生成统计 ==========")
        logger.info("filename.txt 原始非空行数量: %d", total_keywords)
        logger.info("filename.txt 去重后需要生成 ZIP 的数量: %d", unique_keywords_count)
        logger.info("成功的编号数量: %d", len(success_keywords))
        logger.info("目标生成 ZIP 的数量: %d", len(generated_zip_paths))
        logger.info("未成功的数量: %d", len(failed_keywords))
        if failed_keywords:
            logger.warning("未成功的编号: %s", ", ".join(failed_keywords))
        else:
            logger.info("未成功的编号: 无")
        logger.info("================================")

    except (OSError, yaml.YAMLError) as e:
        # 指定具体的异常类型，避免捕获过于宽泛的 Exception
        logger.error("发生错误: %s", str(e), exc_info=True)


if __name__ == "__main__":
    main()
