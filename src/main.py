import os
import sys
import shutil
import subprocess
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QColor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFrame, QFileDialog, QListWidget, QListWidgetItem,
    QProgressBar, QMessageBox, QComboBox, QLineEdit
)
from pdf2docx import Converter


APP_NAME = "PDF Master"


def find_libreoffice():
    candidates = [
        shutil.which("soffice"),
        shutil.which("soffice.exe"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for item in candidates:
        if item and Path(item).exists():
            return item
    return None


class ConvertWorker(QThread):
    progress = Signal(int)
    status = Signal(str)
    finished_one = Signal(str)
    failed = Signal(str)

    def __init__(self, files, mode, output_dir):
        super().__init__()
        self.files = files
        self.mode = mode
        self.output_dir = Path(output_dir)

    def run(self):
        total = len(self.files)
        for index, file_name in enumerate(self.files):
            source = Path(file_name)
            try:
                self.status.emit(f"正在处理：{source.name}")
                if self.mode == "pdf_to_word":
                    if source.suffix.lower() != ".pdf":
                        raise ValueError("PDF 转 Word 模式只接受 PDF 文件。")
                    target = self.output_dir / f"{source.stem}.docx"
                    converter = Converter(str(source))
                    try:
                        converter.convert(str(target), start=0, end=None)
                    finally:
                        converter.close()
                else:
                    if source.suffix.lower() not in (".doc", ".docx"):
                        raise ValueError("Word 转 PDF 模式只接受 DOC 或 DOCX 文件。")
                    soffice = find_libreoffice()
                    if not soffice:
                        raise RuntimeError(
                            "没有检测到 LibreOffice。\n\n"
                            "请安装 LibreOffice，并确保 soffice.exe 可用，"
                            "或将其安装在默认目录后重新启动软件。"
                        )
                    result = subprocess.run(
                        [soffice, "--headless", "--convert-to", "pdf",
                         "--outdir", str(self.output_dir), str(source)],
                        capture_output=True, text=True, timeout=180
                    )
                    target = self.output_dir / f"{source.stem}.pdf"
                    if result.returncode != 0 or not target.exists():
                        detail = (result.stderr or result.stdout or "未知错误").strip()
                        raise RuntimeError(f"LibreOffice 转换失败：{detail}")
                self.finished_one.emit(str(target))
            except Exception as exc:
                self.failed.emit(f"{source.name}：{exc}")
            self.progress.emit(int((index + 1) * 100 / total))


class DropList(QListWidget):
    files_dropped = Signal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.DropOnly)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        self.files_dropped.emit(paths)
        event.acceptProposedAction()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Master")
        self.resize(1050, 720)
        self.setMinimumSize(850, 600)
        self.files = []
        self.output_dir = str(Path.home() / "Documents" / "PDF Master")
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)
        self.mode = "pdf_to_word"
        self.worker = None
        self.build_ui()
        self.apply_style()

    def build_ui(self):
        base = QWidget()
        self.setCentralWidget(base)
        outer = QHBoxLayout(base)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(235)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(22, 28, 22, 24)
        brand = QLabel("✦  PDF Master")
        brand.setObjectName("brand")
        side.addWidget(brand)
        side.addSpacing(35)
        nav_title = QLabel("文件转换")
        nav_title.setObjectName("muted")
        side.addWidget(nav_title)
        self.pdf_btn = QPushButton("▤   PDF 转 Word")
        self.word_btn = QPushButton("▧   Word 转 PDF")
        self.pdf_btn.setCheckable(True)
        self.word_btn.setCheckable(True)
        self.pdf_btn.setChecked(True)
        self.pdf_btn.clicked.connect(lambda: self.set_mode("pdf_to_word"))
        self.word_btn.clicked.connect(lambda: self.set_mode("word_to_pdf"))
        for b in (self.pdf_btn, self.word_btn):
            b.setObjectName("navButton")
            side.addWidget(b)
        side.addStretch()
        side.addWidget(QLabel("PDF Master  •  Desktop"))
        outer.addWidget(sidebar)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(38, 30, 38, 30)
        content_layout.setSpacing(18)
        outer.addWidget(content, 1)

        top = QHBoxLayout()
        heading_box = QVBoxLayout()
        self.heading = QLabel("PDF 转 Word")
        self.heading.setObjectName("heading")
        self.subheading = QLabel("将 PDF 文档转换为可编辑的 Word 文件")
        self.subheading.setObjectName("muted")
        heading_box.addWidget(self.heading)
        heading_box.addWidget(self.subheading)
        top.addLayout(heading_box)
        top.addStretch()
        self.theme_label = QLabel("简洁 · 高效 · 安全")
        self.theme_label.setObjectName("pill")
        top.addWidget(self.theme_label, alignment=Qt.AlignTop)
        content_layout.addLayout(top)

        self.drop_area = QFrame()
        self.drop_area.setObjectName("dropArea")
        drop_layout = QVBoxLayout(self.drop_area)
        drop_layout.setContentsMargins(20, 28, 20, 28)
        drop_layout.setAlignment(Qt.AlignCenter)
        icon = QLabel("⇧")
        icon.setObjectName("uploadIcon")
        icon.setAlignment(Qt.AlignCenter)
        drop_layout.addWidget(icon)
        drop_title = QLabel("拖放文件到这里")
        drop_title.setObjectName("dropTitle")
        drop_title.setAlignment(Qt.AlignCenter)
        drop_layout.addWidget(drop_title)
        drop_hint = QLabel("也可以点击下方按钮选择文件")
        drop_hint.setObjectName("muted")
        drop_hint.setAlignment(Qt.AlignCenter)
        drop_layout.addWidget(drop_hint)
        self.choose_btn = QPushButton("选择文件")
        self.choose_btn.setObjectName("primary")
        self.choose_btn.clicked.connect(self.choose_files)
        drop_layout.addWidget(self.choose_btn, alignment=Qt.AlignCenter)
        content_layout.addWidget(self.drop_area)

        files_header = QHBoxLayout()
        files_title = QLabel("待转换文件")
        files_title.setObjectName("sectionTitle")
        files_header.addWidget(files_title)
        files_header.addStretch()
        self.count_label = QLabel("0 个文件")
        self.count_label.setObjectName("muted")
        files_header.addWidget(self.count_label)
        clear_btn = QPushButton("清空列表")
        clear_btn.setObjectName("textButton")
        clear_btn.clicked.connect(self.clear_files)
        files_header.addWidget(clear_btn)
        content_layout.addLayout(files_header)

        self.file_list = DropList()
        self.file_list.setObjectName("fileList")
        self.file_list.setMinimumHeight(115)
        self.file_list.files_dropped.connect(self.add_files)
        content_layout.addWidget(self.file_list, 1)

        output_row = QHBoxLayout()
        output_row.addWidget(QLabel("输出位置"))
        self.output_edit = QLineEdit(self.output_dir)
        self.output_edit.setReadOnly(True)
        output_row.addWidget(self.output_edit, 1)
        browse_out = QPushButton("更改")
        browse_out.setObjectName("secondary")
        browse_out.clicked.connect(self.choose_output)
        output_row.addWidget(browse_out)
        content_layout.addLayout(output_row)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        content_layout.addWidget(self.progress)
        self.status_label = QLabel("准备就绪")
        self.status_label.setObjectName("muted")
        content_layout.addWidget(self.status_label)

        bottom = QHBoxLayout()
        bottom.addStretch()
        self.convert_btn = QPushButton("开始转换  →")
        self.convert_btn.setObjectName("primary")
        self.convert_btn.setMinimumWidth(180)
        self.convert_btn.clicked.connect(self.start_conversion)
        bottom.addWidget(self.convert_btn)
        content_layout.addLayout(bottom)

    def apply_style(self):
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #F5F7FB; color: #20263A; font-family: "Segoe UI", "Microsoft YaHei"; font-size: 14px; }
            #sidebar { background: #171C32; color: #FFFFFF; }
            #brand { color: white; font-size: 22px; font-weight: 700; }
            #sidebar QLabel { color: #AAB2CC; }
            #sidebar #brand { color: white; }
            #muted { color: #7D859B; }
            #sidebar #muted { color: #8F98B5; }
            #navButton { text-align: left; color: #C6CBE0; background: transparent; border: 0; padding: 14px 12px; border-radius: 9px; margin-top: 5px; }
            #navButton:checked, #navButton:hover { background: #2A3150; color: white; }
            #heading { font-size: 29px; font-weight: 700; color: #1D2540; }
            #pill { color: #6558D9; background: #EEECFF; padding: 8px 13px; border-radius: 12px; }
            #dropArea { background: white; border: 2px dashed #C9CDED; border-radius: 18px; }
            #uploadIcon { color: #695CE7; font-size: 38px; font-weight: 600; }
            #dropTitle { font-size: 19px; font-weight: 650; }
            #sectionTitle { font-size: 17px; font-weight: 650; }
            #fileList { background: white; border: 1px solid #E7EAF2; border-radius: 12px; padding: 8px; }
            QLineEdit { background: white; border: 1px solid #E2E6EF; border-radius: 8px; padding: 10px; color: #424B63; }
            QProgressBar { border: none; border-radius: 6px; background: #E7EAF2; height: 10px; text-align: center; color: transparent; }
            QProgressBar::chunk { background: #6B5CE7; border-radius: 6px; }
            QPushButton { border: none; border-radius: 9px; padding: 11px 17px; font-weight: 600; }
            #primary { background: #6658E8; color: white; }
            #primary:hover { background: #5749D4; }
            #primary:disabled { background: #B8B4E8; }
            #secondary { background: #ECEEFF; color: #5146C8; }
            #textButton { color: #777F96; background: transparent; }
        """)

    def set_mode(self, mode):
        self.mode = mode
        self.pdf_btn.setChecked(mode == "pdf_to_word")
        self.word_btn.setChecked(mode == "word_to_pdf")
        if mode == "pdf_to_word":
            self.heading.setText("PDF 转 Word")
            self.subheading.setText("将 PDF 文档转换为可编辑的 Word 文件")
        else:
            self.heading.setText("Word 转 PDF")
            self.subheading.setText("将 DOC / DOCX 文档快速导出为 PDF")
        self.clear_files()

    def choose_files(self):
        if self.mode == "pdf_to_word":
            paths, _ = QFileDialog.getOpenFileNames(self, "选择 PDF 文件", "", "PDF 文件 (*.pdf)")
        else:
            paths, _ = QFileDialog.getOpenFileNames(self, "选择 Word 文件", "", "Word 文件 (*.doc *.docx)")
        self.add_files(paths)

    def add_files(self, paths):
        allowed = (".pdf",) if self.mode == "pdf_to_word" else (".doc", ".docx")
        for p in paths:
            if Path(p).is_file() and Path(p).suffix.lower() in allowed and p not in self.files:
                self.files.append(p)
                item = QListWidgetItem(f"  ▣   {Path(p).name}")
                item.setToolTip(p)
                self.file_list.addItem(item)
        self.count_label.setText(f"{len(self.files)} 个文件")

    def clear_files(self):
        self.files = []
        self.file_list.clear()
        self.count_label.setText("0 个文件")
        self.progress.setValue(0)
        self.status_label.setText("准备就绪")

    def choose_output(self):
        folder = QFileDialog.getExistingDirectory(self, "选择输出文件夹", self.output_dir)
        if folder:
            self.output_dir = folder
            self.output_edit.setText(folder)

    def start_conversion(self):
        if not self.files:
            QMessageBox.information(self, "提示", "请先添加需要转换的文件。")
            return
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)
        self.convert_btn.setEnabled(False)
        self.progress.setValue(0)
        self.worker = ConvertWorker(self.files[:], self.mode, self.output_dir)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.status.connect(self.status_label.setText)
        self.worker.finished_one.connect(lambda p: self.status_label.setText(f"已完成：{Path(p).name}"))
        self.worker.failed.connect(lambda e: QMessageBox.warning(self, "转换失败", e))
        self.worker.finished.connect(self.conversion_finished)
        self.worker.start()

    def conversion_finished(self):
        self.convert_btn.setEnabled(True)
        self.status_label.setText("转换任务已结束")
        QMessageBox.information(self, "完成", f"转换任务已结束。\n输出目录：\n{self.output_dir}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
