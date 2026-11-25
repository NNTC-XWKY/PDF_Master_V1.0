# PDF Master

一款 Windows 桌面 PDF / Word 双向转换工具，使用 PySide6 构建现代化图形界面。

## 当前功能
- PDF 批量转换为 DOCX（基于 pdf2docx）
- DOC / DOCX 批量转换为 PDF（调用 LibreOffice）
- 拖拽导入文件
- 自定义输出目录
- 转换进度和错误提示
- GitHub Actions 自动构建 Windows EXE

## 本地运行
建议使用 Python 3.11。

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python src/main.py
```

## Word 转 PDF 依赖
Word 转 PDF 通过 LibreOffice 的无界面模式完成。运行软件的 Windows 电脑需要安装 LibreOffice：
https://www.libreoffice.org/download/download-libreoffice/

PDF 转 Word 使用 pdf2docx。复杂排版、扫描件、特殊字体的转换效果取决于源文件；扫描版 PDF 通常需要额外 OCR，当前版本尚未集成 OCR。

## 上传 GitHub 并生成 EXE
1. 在 GitHub 新建一个空仓库，例如 `pdf-master`。
2. 解压本项目，将其中所有文件上传到仓库根目录。
3. 打开仓库的 **Actions** 页面。
4. 如果提示启用 Actions，点击启用。
5. 选择 **Build Windows EXE**，点击 **Run workflow**；也可以直接 push 到 `main` 分支自动触发。
6. 构建成功后，打开对应运行记录，在页面底部 **Artifacts** 下载 `PDFMaster-Windows-EXE`。
7. 解压下载文件即可得到 `PDFMaster.exe`。

## 注意
- 这是首个可运行版本，建议先用副本文件测试。
- Windows 可能对未签名的自编译 EXE 显示 SmartScreen 提示。
- 当前 Word 转 PDF 依赖目标电脑安装 LibreOffice；PyInstaller 打包的 EXE 不包含 LibreOffice。
- 转换结果不保证与原文件 100% 一致，尤其是复杂 PDF、扫描件、嵌入字体和复杂表格。
