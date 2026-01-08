# CopyZipDesktop

 winget install pandoc

<https://blog.csdn.net/qq_44824148/article/details/128601430>

Latex 安装教程

1. 下载texlive.iso
1. 下载texlive.iso
  <https://mirrors.tuna.tsinghua.edu.cn/ctan/systems/texlive/Images/>
2. 点击装载
3. 运行 install-tl-windows.bat 进行安装
4. 验证是否安装成功
5. 安装开发工具

## combine_pdf.py only work in windows system

## 修改日志

### 2026-01-08
- **修复问题**：解决了 `filename.txt` 中不包含 `JZ-` 前缀但仍误压缩/拷贝 `JZ-` 开头文件夹的问题。
- **功能增强**：
  - 在 `copy_zip_desktop.py` 和 `copy_docx.py` 中引入了 `match_mode` 配置支持。
  - 将所有相关脚本（`copy_zip_desktop.py`, `copy_docx.py`, `copy_pdf_docx_desktop.py`）的默认匹配模式统一为 `startswith`。
  - 允许通过配置文件修改匹配模式为 `contains`（子字符串匹配）、`startswith`（前缀匹配）或 `regex`（正则表达式，仅限部分脚本）。
