import logging
import os
import re
import shutil
import sys
import tempfile
from logging.handlers import RotatingFileHandler
from pathlib import Path

import pypandoc
import yaml
from PyPDF2 import PdfMerger

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


logger = setup_logger(log_file="./logs/combine_pdf.log")


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


def extract_base_id(filename, regex_patterns):
    """
    从文件名中用多个正则尝试提取编号前缀，返回第一个成功匹配的编号
    """
    for pattern in regex_patterns:
        try:
            match = re.match(pattern, filename)
            if match:
                if isinstance(match.group(0), tuple):
                    return "-".join(match.groups())
                else:
                    return match.group(0)
        except re.error as e:
            logger.error(f"正则表达式错误: {e}")
    return None


def validate_docx_pdf_pairs(config):
    """
    校验每个 docx 文件是否有且仅有一个对应 pdf 文件（基于编号前缀匹配）
    """
    desktop_output = config.get("desktop_output", "")
    if isinstance(desktop_output, str):
        desktop_output = [desktop_output]
    regex_patterns = config.get("regex_pattern", [])
    if not regex_patterns:
        logger.error("配置文件中未提供正则表达式")
        return

    errors = 0
    for directory in desktop_output:
        logger.info(f"开始校验目录: {directory}")
        docx_files = list(Path(directory).rglob("*.docx"))
        pdf_files = list(Path(directory).rglob("*.pdf"))

        # 构建 PDF 索引
        pdf_map = {}
        for pdf in pdf_files:
            base_id = extract_base_id(pdf.stem, regex_patterns)
            if base_id:
                pdf_map.setdefault(base_id, []).append(pdf)

        # 遍历 .docx 文件
        for docx in docx_files:
            base_id = extract_base_id(docx.stem, regex_patterns)
            if not base_id:
                logger.warning(f"无法从文件名提取编号: {docx.name}")
                continue

            matched_pdfs = pdf_map.get(base_id, [])
            if len(matched_pdfs) == 0:
                logger.error(f"❌ {docx.name} 缺少对应 PDF 文件（编号: {base_id}）")
                errors += 1
            elif len(matched_pdfs) > 1:
                pdf_names = ", ".join([p.name for p in matched_pdfs])
                logger.error(f"❌ {docx.name} 对应多个 PDF 文件: {pdf_names}")
                errors += 1
            else:
                logger.info(f"✅ 匹配成功: {docx.name} ↔ {matched_pdfs[0].name}")

    if errors == 0:
        logger.info("✅ 所有 docx 文件都有唯一对应的 pdf 文件")
    else:
        logger.info(f"❌ 校验完成，共发现 {errors} 个问题")


def convert_docx_to_pdf(docx_path, output_pdf_path):
    """使用 Microsoft Word 将 docx 转换为 PDF"""
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        logger.info(f"打开 Word 文档: {docx_path}")
        doc = word.Documents.Open(str(docx_path))
        doc.SaveAs(str(output_pdf_path), FileFormat=17)  # 17 表示 PDF
        doc.Close(False)
        word.Quit()
    except Exception as e:
        logger.exception(f"使用 Word 转换 {docx_path} 为 PDF 失败")
        raise


def merge_docx_pdf(config, output_path="./final_merged.pdf"):
    """将所有 docx 和其对应的 pdf，统一合并成一个 PDF 文件"""
    desktop_output = config.get("desktop_output", [])
    if isinstance(desktop_output, str):
        desktop_output = [desktop_output]
    regex_patterns = config.get("regex_pattern", [])

    merger = PdfMerger()
    temp_files = []  # 用于保存临时 Word 转 PDF 的路径，后续删除

    for directory in desktop_output:
        docx_files = list(Path(directory).rglob("*.docx"))
        pdf_files = list(Path(directory).rglob("*.pdf"))

        # 构建 PDF 索引
        pdf_map = {}
        for pdf in pdf_files:
            base_id = extract_base_id(pdf.stem, regex_patterns)
            if base_id:
                pdf_map.setdefault(base_id, []).append(pdf)

        for docx in docx_files:
            base_id = extract_base_id(docx.stem, regex_patterns)
            if not base_id:
                logger.warning(f"无法从 docx 提取编号: {docx.name}")
                continue

            matched_pdfs = pdf_map.get(base_id, [])
            if len(matched_pdfs) != 1:
                logger.warning(f"{docx.name} 对应的 PDF 数量不正确，跳过合并")
                continue

            pdf_file = matched_pdfs[0]
            try:
                # 1. 将 docx 转为临时 PDF
                with tempfile.NamedTemporaryFile(
                    suffix=".pdf", delete=False
                ) as tmp_docx_pdf:
                    tmp_docx_pdf_path = Path(tmp_docx_pdf.name)
                convert_docx_to_pdf(str(docx), str(tmp_docx_pdf_path))
                temp_files.append(tmp_docx_pdf_path)

                # 2. 添加到合并对象
                merger.append(str(tmp_docx_pdf_path))
                merger.append(str(pdf_file))
                logger.info(f"✅ 添加合并项: {docx.name} + {pdf_file.name}")

            except Exception as e:
                logger.error(f"合并 {docx.name} 和 {pdf_file.name} 失败: {e}")

    # 输出最终合并文件
    if merger.pages:
        merger.write(output_path)
        merger.close()
        logger.info(f"🎉 所有文档合并完成: {output_path}")
    else:
        logger.warning("⚠️ 没有成功合并任何内容，未生成合并文件。")

    # 清理临时文件
    for tmp_file in temp_files:
        try:
            tmp_file.unlink()
        except Exception as e:
            logger.warning(f"临时文件删除失败: {tmp_file} - {e}")


if __name__ == "__main__":
    is_B24 = "Yes"
    is_B25B26 = "No"

    config = read_config("./path_config_B24.yaml")
    if is_B24 == "Yes":
        config = read_config("./path_config_B24.yaml")

    if is_B25B26 == "Yes":
        config = read_config("./path_config_B25B26.yaml")
    validate_docx_pdf_pairs(config)
    merge_docx_pdf(
        config, output_path=config.get("desktop_output", "") + "/final_merged.pdf"
    )
