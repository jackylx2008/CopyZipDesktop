# -*- coding: utf-8 -*-
"""
combine_pdf.py
- 合并 DOCX（转为 PDF）与对应原 PDF
- 将合并后的 PDF 重新“打印”到竖向 A4：
  * 横向页面自动旋转 90° 放到竖向 A4（可关）
  * 仅对大页进行等比缩小，不放大小页（可关）
  * 居中放置，保留矢量质量（pdfrw + reportlab）

依赖:
  pip install PyPDF2 reportlab pdfrw pypiwin32 pyyaml
"""

import logging
import re
import sys
import tempfile
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any  # 新增

import yaml
from PyPDF2 import PdfMerger

# ========== 可选依赖：Windows Word COM ==========
_HAS_PYWIN32 = False
try:
    import win32com.client
    import pythoncom
    import win32api  # 用于验证安装
    import pywintypes

    # 验证 pywin32 是否正确安装
    _HAS_PYWIN32 = win32api.GetSystemMetrics(0) > 0  # 简单的验证调用
    if _HAS_PYWIN32:
        COM_ERROR = pywintypes.com_error
except ImportError:
    win32com = None
    pythoncom = None
    COM_ERROR = Exception

# ========== 可选依赖：reprint_to_a4 使用 ==========
try:
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import A4

    _HAS_REPORTLAB = True
except ImportError:
    # 设置默认值并标记为 Any 以消除“未绑定”警告
    rl_canvas: Any = None
    A4: Any = None  # type: ignore
    _HAS_REPORTLAB = False

try:
    from pdfrw import PdfReader as PdfrwReader
    from pdfrw.buildxobj import pagexobj
    from pdfrw.toreportlab import makerl

    _HAS_PDFRW = True
except ImportError:
    # 设置默认值并标记为 Any
    PdfrwReader: Any = None
    pagexobj: Any = None
    makerl: Any = None


# ---------- 基础设置 ----------
if sys.platform.startswith("win"):
    # 避免中文日志乱码（需 Python 3.7+）
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass

# 如需引用上级目录模块，可打开
# 将项目根目录添加到 sys.path
# sys.path.append(
#     os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
# )


# ---------- 日志 ----------
def setup_logger(
    name: str = __name__, log_level=logging.INFO, log_file: str | None = None
) -> logging.Logger:
    """
    设置并返回一个日志记录器
    """
    logger_instance = logging.getLogger(name)
    if not logger_instance.hasHandlers():
        logger_instance.setLevel(log_level)

        fmt = logging.Formatter(
            fmt="%(asctime)s - %(levelname)s - %(name)s - %(message)s",
            datefmt=(
                "%-Y-%m-%d %H:%M:%S"
                if not sys.platform.startswith("win")
                else "%Y-%m-%d %H:%M:%S"
            ),
        )

        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(fmt)
        logger_instance.addHandler(ch)

        if log_file:
            Path(log_file).parent.mkdir(parents=True, exist_ok=True)
            max_bytes = 5 * 1024 * 1024  # 5MB
            fh = RotatingFileHandler(
                log_file, maxBytes=max_bytes, backupCount=3, encoding="utf-8"
            )
            fh.setFormatter(fmt)
            logger_instance.addHandler(fh)

    return logger_instance


logger = setup_logger(log_file="./logs/combine_pdf.log")


# ---------- 读取配置 ----------
def read_config(config_path="./path_config.yaml") -> dict:
    """
    读取 YAML 配置文件
    """
    logger.info("读取配置文件: %s", config_path)
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    logger.info("配置文件加载成功")
    return cfg


# ---------- 从文件名提取编号 ----------
def extract_base_id(filename_stem: str, regex_patterns) -> str | None:
    """
    尝试用多个正则从“文件名(不含扩展名)”里提取编号前缀。
    使用 re.match；若需更宽松可改为 re.search。
    """
    for pattern in regex_patterns or []:
        try:
            m = re.match(pattern, filename_stem)
            if m:
                return "-".join(m.groups()) if m.groups() else m.group(0)
        except re.error as e:
            logger.error("正则表达式错误: %s", e)
    return None


# ---------- 校验 DOCX 与 PDF 的配对 ----------
def validate_docx_pdf_pairs(cfg_dict: dict) -> dict[str, Path]:
    """
    校验每个 docx 是否唯一匹配一个 pdf（按编号前缀），
    返回 { base_id: Path(pdf) } 便于后续快速查找。
    """
    output_dirs = cfg_dict.get("desktop_output", [])
    if isinstance(output_dirs, str):
        output_dirs = [output_dirs]
    regex_patterns = cfg_dict.get("regex_pattern", [])

    pdf_index: dict[str, Path] = {}
    if not regex_patterns:
        logger.error("配置文件中未提供正则表达式 regex_pattern")
        return {}

    pdf_index: dict[str, Path] = {}

    for directory in output_dirs:
        directory = Path(directory)
        logger.info("开始校验目录: %s", directory)
        docx_files = list(directory.rglob("*.docx"))
        pdf_files = list(directory.rglob("*.pdf"))

        # 建 PDF 索引：base_id -> [pdf paths]
        pdf_map: dict[str, list[Path]] = {}
        for pdf in pdf_files:
            base_id = extract_base_id(pdf.stem, regex_patterns)
            if base_id:
                pdf_map.setdefault(base_id, []).append(pdf)

        errors = 0
        for docx in docx_files:
            base_id = extract_base_id(docx.stem, regex_patterns)
            if not base_id:
                logger.warning("无法从文件名提取编号: %s", docx.name)
                continue

            matched = pdf_map.get(base_id, [])
            if len(matched) == 0:
                logger.error("❌ %s 缺少对应 PDF（编号: %s）", docx.name, base_id)
                errors += 1
            elif len(matched) > 1:
                pdf_names = ", ".join(p.name for p in matched)
                logger.error("❌ %s 对应多个 PDF: %s", docx.name, pdf_names)
                errors += 1
            else:
                logger.info("✅ 匹配成功: %s ↔ %s", docx.name, matched[0].name)
                pdf_index[base_id] = matched[0]

        if errors == 0:
            logger.info("✅ 本目录内所有 docx 均有唯一 pdf")
        else:
            logger.info("❌ 本目录校验完成，共发现 %d 个问题", errors)

    return pdf_index


# ---------- Word 转 PDF（仅 Windows） ----------
def convert_docx_to_pdf(docx_path: Path, output_pdf_path: Path):
    """使用 Microsoft Word COM 转换（仅 Windows 可用）"""
    if not sys.platform.startswith("win"):
        raise RuntimeError("此功能仅支持 Windows 系统")
    if not _HAS_PYWIN32:
        raise RuntimeError(
            "未能正确初始化 pywin32。请尝试以下步骤：\n"
            "1. 确保已安装: pip install --upgrade pywin32\n"
            "2. 以管理员权限运行: python -m win32com.client.makepy"
        )

    word = None
    doc = None
    try:
        # 初始化 COM 安全级别
        if pythoncom is not None:
            pythoncom.CoInitialize()
        # 创建 Word 应用实例
        try:
            if win32com is None or not hasattr(win32com, "client"):
                raise RuntimeError("win32com 未正确导入，请检查 pywin32 是否安装。")
            word = win32com.client.DispatchEx("Word.Application")
        except Exception as e:
            raise RuntimeError(
                f"无法启动 Word: {e}。请确保 Microsoft Word 已正确安装。"
            ) from e
        word.DisplayAlerts = 0
        logger.info("打开 Word 文档: %s", docx_path)
        doc = word.Documents.Open(str(docx_path))
        # 17 = wdFormatPDF
        doc.SaveAs(str(output_pdf_path), FileFormat=17)
    except (AttributeError, COM_ERROR) as e:
        logger.exception("使用 Word 将 %s 转为 PDF 失败: %s", docx_path, str(e))
        raise
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except (AttributeError, COM_ERROR):
            pass
        try:
            if word is not None:
                word.Quit()
        except (AttributeError, COM_ERROR):
            pass
        # 清理 COM
        if pythoncom is not None:
            pythoncom.CoUninitialize()


# ---------- 合并 DOCX-PDF 对 ----------
def merge_docx_pdf(cfg_dict: dict, output_path: Path):
    """
    对每个 docx：先转成临时 pdf，再按"docx.pdf + 对应原 pdf"追加到总合并里。
    """
    output_dirs = cfg_dict.get("desktop_output", [])
    if isinstance(output_dirs, str):
        output_dirs = [output_dirs]
    regex_patterns = cfg_dict.get("regex_pattern", []) or []

    # 先做一次索引校验，拿到 base_id -> pdf 的映射
    pdf_index = validate_docx_pdf_pairs(cfg_dict)

    merger = PdfMerger()
    temp_files: list[Path] = []
    appended_count = 0

    try:
        for directory in output_dirs:
            directory = Path(directory)
            docx_files = list(directory.rglob("*.docx"))

            for docx in docx_files:
                base_id = extract_base_id(docx.stem, regex_patterns)
                if not base_id:
                    logger.warning("无法从 docx 提取编号: %s", docx.name)
                    continue

                pdf_file = pdf_index.get(base_id)
                if not pdf_file:
                    logger.warning("%s 未找到唯一对应 PDF，跳过合并", docx.name)
                    continue

                try:
                    # 1) docx -> 临时 pdf
                    with tempfile.NamedTemporaryFile(
                        suffix=".pdf", delete=False
                    ) as tmp_file:
                        tmp_docx_pdf = Path(tmp_file.name)
                    convert_docx_to_pdf(docx, tmp_docx_pdf)
                    temp_files.append(tmp_docx_pdf)

                    # 2) 追加到合并
                    merger.append(str(tmp_docx_pdf))
                    merger.append(str(pdf_file))
                    appended_count += 2
                    logger.info("✅添加合并项: %s + %s", docx.name, pdf_file.name)
                except (IOError, ValueError) as e:
                    logger.error("合并%s与%s失败: %s", docx.name, pdf_file.name, e)

        if appended_count > 0:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            merger.write(str(output_path))
            logger.info("🎉 合并完成: %s", output_path)
        else:
            logger.warning("⚠️ 没有成功合并任何内容，未生成合并文件。")
    finally:
        try:
            merger.close()
        except (IOError, ValueError):
            pass
        # 清理临时文件
        for p in temp_files:
            try:
                p.unlink(missing_ok=True)
            except (IOError, PermissionError) as e:
                logger.warning("临时文件删除失败: %s - %s", p, e)


# ---------- 工具：mm → pt ----------
def mm_to_pt(mm: float) -> float:
    """
    毫米转磅（PostScript Points）
    """
    return mm * 72.0 / 25.4


# ---------- 关键：重新打印成竖向 A4 ----------
def reprint_to_a4(
    input_pdf: Path | str,
    output_pdf: Path | str,
    margin_mm: float = 10.0,
    shrink_only: bool = True,
    rotate_landscape_to_portrait: bool = True,
    # 兼容旧版本参数名（如果传了旧名，则覆盖新值）
    auto_rotate_landscape: bool | None = None,
    log: logging.Logger | None = None,
):
    """
    将任意 PDF 重新“打印”到**竖向 A4**。
      - 若源页为横向（宽>高），可选地旋转 90° 后以竖向 A4 输出（rotate_landscape_to_portrait=True）。
      - 若源页尺寸大于可用内容区，按等比缩小，保证不超出 A4；
        当 shrink_only=True 时，不会放大小页。
      - 内容居中放置，尽量保留矢量质量（pdfrw + reportlab）。

    参数:
      input_pdf:  输入 PDF 路径
      output_pdf: 输出 PDF 路径
      margin_mm:  四边统一页边距（毫米）
      shrink_only: 仅缩小不放大（True），若 False 则小页也会放大至占满可用区
      rotate_landscape_to_portrait: 横向页是否旋转 90° 后排版到竖向 A4
      auto_rotate_landscape: 兼容老参数名；若传入则覆盖 rotate_landscape_to_portrait
      log: Logger，不传则使用 logging.getLogger(__name__)

    依赖:
      pip install reportlab pdfrw
    """
    logger_local = log or logging.getLogger(__name__)

    # 依赖检查与友好提示
    if not _HAS_REPORTLAB:
        raise RuntimeError("缺少 reportlab，请先安装：pip install reportlab")
    if not _HAS_PDFRW:
        raise RuntimeError("缺少 pdfrw，请先安装：pip install pdfrw")

    # 确保 reportlab 已导入
    if not _HAS_REPORTLAB:
        raise RuntimeError("无法导入 reportlab.lib.pagesizes.A4，请检查 reportlab 安装")

    # 兼容旧参数名
    if auto_rotate_landscape is not None:
        rotate_landscape_to_portrait = auto_rotate_landscape

    input_pdf = Path(input_pdf)
    output_pdf = Path(output_pdf)
    if not input_pdf.is_file():
        raise FileNotFoundError(f"未找到输入文件: {input_pdf}")

    # 目标画布统一为竖向 A4
    a4_w, a4_h = A4
    margin_pt = mm_to_pt(margin_mm)
    content_w = max(1.0, a4_w - 2 * margin_pt)
    content_h = max(1.0, a4_h - 2 * margin_pt)

    # 确保 PdfrwReader 已定义
    reader = PdfrwReader(str(input_pdf))

    # 修复：显式转换为 list 解决 Pylance 对 len() 和 enumerate() 的报错
    pages = list(reader.pages) if reader.pages is not None else []
    total = len(pages)

    logger_local.info(
        "开始重新排版到竖向 A4: %s → %s，总页数 %d，边距 %.1f mm，横向页旋转: %s，仅缩小: %s",
        input_pdf,
        output_pdf,
        total,
        margin_mm,
        rotate_landscape_to_portrait,
        shrink_only,
    )

    c = rl_canvas.Canvas(str(output_pdf), pagesize=A4)

    for idx, page in enumerate(pages, start=1):
        # 将源页转为 XObject（可保持矢量）
        xobj = pagexobj(page)

        # 源页原始尺寸（pt）
        src_w = float(xobj.BBox[2] - xobj.BBox[0])
        src_h = float(xobj.BBox[3] - xobj.BBox[1])
        is_landscape = src_w > src_h

        # 是否对横向页做 90° 旋转后排版到竖向 A4
        rotate90 = rotate_landscape_to_portrait and is_landscape

        # 计算在“目标可用区”内的等比缩放比例
        # 注意：若旋转，则放置后的“宽=src_h”“高=src_w”
        target_w = src_h if rotate90 else src_w
        target_h = src_w if rotate90 else src_h

        scale_raw = min(content_w / target_w, content_h / target_h)
        scale = min(1.0, scale_raw) if shrink_only else scale_raw

        # 旋转后的占位尺寸（在 A4 坐标系里）
        placed_w = target_w * scale
        placed_h = target_h * scale

        # 居中偏移（以 A4 画布左下角为原点）
        offset_x = margin_pt + (content_w - placed_w) / 2.0
        offset_y = margin_pt + (content_h - placed_h) / 2.0

        orient = (
            "横向→旋转90°排到竖向"
            if rotate90
            else ("竖向" if not is_landscape else "横向(未旋转)")
        )
        logger_local.info(
            "第 %d/%d 页 | 源: %.1f×%.1f pt | 放置: %.1f×%.1f pt | "
            "比例: %.4f | 模式: %s",
            idx,
            total,
            src_w,
            src_h,
            placed_w,
            placed_h,
            scale,
            orient,
        )

        # 开始绘制到竖向 A4
        c.setPageSize(A4)
        c.saveState()

        if rotate90:
            # 旋转 90°（逆时针）：
            # 放置点取 (offset_x, offset_y + placed_h)，再 rotate(90),
            # 此时内容沿正 X 向右、沿负 Y 向下，能恰好落入放置框。
            c.translate(offset_x, offset_y + placed_h)
            c.rotate(90)
            c.scale(scale, scale)
            c.doForm(makerl(c, xobj))
        else:
            # 不旋转，常规放置，左下角对齐
            c.translate(offset_x, offset_y)
            c.scale(scale, scale)
            c.doForm(makerl(c, xobj))

        c.restoreState()
        c.showPage()

    c.save()
    logger_local.info("🎉 竖向 A4 重新排版完成: %s（总页数 %d）", output_pdf, total)


# ---------- 脚本入口 ----------
if __name__ == "__main__":
    # 选择配置（示例开关）
    IS_B24 = "Yes"
    IS_B25B26 = "No"

    if IS_B24 == "Yes":
        config = read_config("./path_config_B24.yaml")
    elif IS_B25B26 == "Yes":
        config = read_config("./path_config_B25B26.yaml")
    else:
        config = read_config("./path_config.yaml")

    # 输出路径：优先用第一个目录
    desktop_output = config.get("desktop_output", "")
    if isinstance(desktop_output, list):
        out_dir = Path(desktop_output[0]) if desktop_output else Path(".")
    else:
        out_dir = Path(desktop_output) if desktop_output else Path(".")

    # 如果需要先合并：
    merge_docx_pdf(config, output_path=out_dir / "final_merged.pdf")

    merged_pdf = out_dir / "final_merged.pdf"
    a4_pdf = out_dir / "final_merged_A4.pdf"

    # 将 merged_pdf 重新打印为竖向 A4：
    # - rotate_landscape_to_portrait=True: 横向页旋转 90°，统一竖向 A4
    # - shrink_only=True: 仅缩小不放大；如需让小页也放大铺满可改 False
    # - auto_rotate_landscape 兼容旧调用名，按需保留/删除
    reprint_to_a4(
        merged_pdf,
        a4_pdf,
        margin_mm=10.0,
        shrink_only=True,
        rotate_landscape_to_portrait=True,
        # 兼容老代码调用（如果你之前写的是 auto_rotate_landscape=True）
        # auto_rotate_landscape=True,
    )

    logger.info("全部处理完成。")
