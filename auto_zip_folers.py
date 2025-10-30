import logging
import os
import sys
import zipfile
from logging.handlers import RotatingFileHandler
from pathlib import Path

import yaml

if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8")
# 动态添加项目根目录到 sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import win32com.client
from logging_config import setup_logger  # 引入日志配置函数


# 初始化日志记录器
def setup_logger(log_level=logging.INFO, log_file=None):
    logger = logging.getLogger()
    logger.setLevel(log_level)
    logger.handlers.clear()  # 清除旧的 handler，防止重复输出

    formatter = logging.Formatter(
        fmt="%(asctime)s - %(levelname)s - %(name)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # 控制台输出（UTF-8）
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # 文件输出（UTF-8）
    if log_file:
        file_handler = RotatingFileHandler(
            log_file, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger


logger = setup_logger(log_file="./logs/auto_zip_folders.log")


def read_config(config_path="./path_config.yaml"):
    """读取 YAML 配置文件"""
    logger.info(f"读取配置文件: {config_path}")
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
        logger.info("配置文件加载成功")
        return config


# ✅ 新增：压缩一个文件夹为 zip
def zip_directory(source_dir: Path, zip_path: Path):
    logger.info(f"压缩文件夹: {source_dir} -> {zip_path}")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(source_dir):
            for file in files:
                file_path = Path(root) / file
                arcname = file_path.relative_to(source_dir)
                zipf.write(file_path, arcname)


# ✅ 新增：压缩子目录并生成最终 zip
def zip_subdirectories(base_dir: Path):
    logger.info(f"开始压缩目录: {base_dir}")
    temp_zip_dir = base_dir / "_zips"
    temp_zip_dir.mkdir(exist_ok=True)

    zip_files = []

    for item in base_dir.iterdir():
        if item.is_dir():
            zip_path = temp_zip_dir / f"{item.name}.zip"
            zip_directory(item, zip_path)
            zip_files.append(zip_path)

    # 最终打包为一个总 zip
    final_zip_path = base_dir.parent / f"{base_dir.name}.zip"
    zip_all_zips(zip_files, final_zip_path)

    logger.info(f"总 zip 文件生成完成: {final_zip_path}")
    return final_zip_path


# ✅ 新增：把多个 zip 文件压成一个大 zip
def zip_all_zips(zip_files, output_zip_path):
    logger.info(f"开始生成总 ZIP 文件: {output_zip_path}")
    with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as big_zip:
        for zip_file in zip_files:
            arcname = zip_file.name
            big_zip.write(zip_file, arcname)


if __name__ == "__main__":
    is_B24 = "Yes"
    is_B25B26 = "No"

    config = read_config("./path_config_B24.yaml")
    if is_B24 == "Yes":
        config = read_config("./path_config_B24.yaml")

    if is_B25B26 == "Yes":
        config = read_config("./path_config_B25B26.yaml")

    # ✅ 执行压缩操作
    base_dirs = config.get("base_directories", [])
    for base_dir_str in base_dirs:
        base_dir = Path(base_dir_str)
        if base_dir.exists() and base_dir.is_dir():
            zip_subdirectories(base_dir)
        else:
            logger.warning(f"无效目录: {base_dir}")
