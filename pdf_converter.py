"""
PDF转Word转换器 - 支持批量处理和扫描版PDF
支持OCR识别扫描版PDF
"""

import os
import sys
from pathlib import Path
from typing import List, Tuple
import threading
from datetime import datetime

import pytesseract
from pdf2image import convert_from_path
from pdf2docx import Converter
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QListWidget, QListWidgetItem, QLabel,
    QProgressBar, QCheckBox, QSpinBox, QComboBox, QMessageBox,
    QTabWidget, QTextEdit, QGroupBox
)
from PyQt6.QtCore import Qt, pyqtSignal, QThread
from PyQt6.QtGui import QIcon, QColor, QFont


class ConversionWorker(QThread):
    """后台转换线程"""
    progress = pyqtSignal(int)
    message = pyqtSignal(str)
    finished = pyqtSignal(bool)
    
    def __init__(self, files: List[str], output_dir: str, use_ocr: bool, 
                 ocr_lang: str = 'chi_sim', dpi: int = 200):
        super().__init__()
        self.files = files
        self.output_dir = output_dir
        self.use_ocr = use_ocr
        self.ocr_lang = ocr_lang
        self.dpi = dpi
        
    def run(self):
        """执行转换"""
        try:
            total_files = len(self.files)
            
            for index, pdf_file in enumerate(self.files):
                try:
                    self.message.emit(f"正在处理: {os.path.basename(pdf_file)}")
                    
                    # 生成输出文件名
                    output_path = os.path.join(
                        self.output_dir,
                        Path(pdf_file).stem + '.docx'
                    )
                    
                    if self.use_ocr:
                        self._convert_with_ocr(pdf_file, output_path)
                    else:
                        self._convert_normal(pdf_file, output_path)
                    
                    self.message.emit(f"✓ 完成: {os.path.basename(pdf_file)}")
                    progress_percent = int((index + 1) / total_files * 100)
                    self.progress.emit(progress_percent)
                    
                except Exception as e:
                    self.message.emit(f"✗ 错误 ({os.path.basename(pdf_file)}): {str(e)}")
                    progress_percent = int((index + 1) / total_files * 100)
                    self.progress.emit(progress_percent)
            
            self.finished.emit(True)
            
        except Exception as e:
            self.message.emit(f"转换失败: {str(e)}")
            self.finished.emit(False)
    
    def _convert_normal(self, pdf_file: str, output_path: str):
        """普通转换（适用于数字PDF）"""
        cv = Converter(pdf_file)
        cv.convert(output_path)
        cv.close()
    
    def _convert_with_ocr(self, pdf_file: str, output_path: str):
        """使用OCR转换扫描版PDF"""
        try:
            # 将PDF转换为图片
            self.message.emit(f"  → 提取PDF页面...")
            images = convert_from_path(pdf_file, dpi=self.dpi)
            
            # 创建新的Word文档
            doc = Document()
            
            total_pages = len(images)
            for page_num, image in enumerate(images, 1):
                self.message.emit(f"  → OCR识别第 {page_num}/{total_pages} 页...")
                
                # 使用Tesseract进行OCR识别
                try:
                    text = pytesseract.image_to_string(image, lang=self.ocr_lang)
                    
                    # 添加图片
                    temp_image_path = f"/tmp/temp_page_{page_num}.png"
                    image.save(temp_image_path)
                    doc.add_picture(temp_image_path, width=Inches(6))
                    
                    # 添加识别的文本
                    if text.strip():
                        p = doc.add_paragraph(text)
                        p.style = 'Normal'
                    
                    # 添加页分符
                    if page_num < total_pages:
                        doc.add_page_break()
                    
                    # 清理临时文件
                    if os.path.exists(temp_image_path):
                        os.remove(temp_image_path)
                        
                except Exception as e:
                    self.message.emit(f"  ⚠ 第{page_num}页OCR失败: {str(e)}")
            
            # 保存文档
            doc.save(output_path)
            
        except Exception as e:
            # 如果OCR失败，尝试普通转换
            self.message.emit(f"  ⚠ OCR失败，尝试普通转换...")
            self._convert_normal(pdf_file, output_path)


class PDFConverterUI(QMainWindow):
    """PDF转Word转换器UI"""
    
    def __init__(self):
        super().__init__()
        self.pdf_files: List[str] = []
        self.worker = None
        self.convert_btn = None  # 保存按钮引用
        self.init_ui()
        
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("PDF转Word转换器 v1.0")
        self.setGeometry(100, 100, 900, 700)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout()
        
        # 创建标题
        title = QLabel("PDF转Word批量转换工具")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        main_layout.addWidget(title)
        
        # 创建标签页
        tabs = QTabWidget()
        main_layout.addWidget(tabs)
        
        # 转换标签页
        convert_widget = self._create_convert_tab()
        tabs.addTab(convert_widget, "转换")
        
        # 设置标签页
        settings_widget = self._create_settings_tab()
        tabs.addTab(settings_widget, "设置")
        
        # 日志标签页
        log_widget = self._create_log_tab()
        tabs.addTab(log_widget, "日志")
        
        central_widget.setLayout(main_layout)
    
    def _create_convert_tab(self) -> QWidget:
        """创建转换标签页"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # 文件列表部分
        file_group = QGroupBox("选择PDF文件")
        file_layout = QVBoxLayout()
        
        # 文件列表
        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(250)
        file_layout.addWidget(QLabel("已选择的文件:"))
        file_layout.addWidget(self.file_list)
        
        # 文件操作按钮
        button_layout = QHBoxLayout()
        
        add_btn = QPushButton("添加文件")
        add_btn.clicked.connect(self.add_files)
        button_layout.addWidget(add_btn)
        
        add_folder_btn = QPushButton("添加文件夹")
        add_folder_btn.clicked.connect(self.add_folder)
        button_layout.addWidget(add_folder_btn)
        
        remove_btn = QPushButton("移除选中")
        remove_btn.clicked.connect(self.remove_file)
        button_layout.addWidget(remove_btn)
        
        clear_btn = QPushButton("清空列表")
        clear_btn.clicked.connect(self.clear_files)
        button_layout.addWidget(clear_btn)
        
        file_layout.addLayout(button_layout)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 选项部分
        options_group = QGroupBox("转换选项")
        options_layout = QVBoxLayout()
        
        # OCR选项
        self.ocr_checkbox = QCheckBox("使用OCR识别扫描版PDF")
        self.ocr_checkbox.setChecked(False)
        self.ocr_checkbox.stateChanged.connect(self.on_ocr_changed)
        options_layout.addWidget(self.ocr_checkbox)
        
        # OCR语言选择
        lang_layout = QHBoxLayout()
        lang_layout.addWidget(QLabel("OCR语言:"))
        self.ocr_lang_combo = QComboBox()
        self.ocr_lang_combo.addItems(['中文简体 (chi_sim)', '中文繁体 (chi_tra)', 
                                      '英文 (eng)', '日文 (jpn)'])
        self.ocr_lang_combo.setEnabled(False)
        lang_layout.addWidget(self.ocr_lang_combo)
        lang_layout.addStretch()
        options_layout.addLayout(lang_layout)
        
        # DPI设置
        dpi_layout = QHBoxLayout()
        dpi_layout.addWidget(QLabel("DPI (仅OCR):"))
        self.dpi_spinbox = QSpinBox()
        self.dpi_spinbox.setValue(200)
        self.dpi_spinbox.setRange(100, 600)
        self.dpi_spinbox.setEnabled(False)
        dpi_layout.addWidget(self.dpi_spinbox)
        dpi_layout.addStretch()
        options_layout.addLayout(dpi_layout)
        
        # 输出目录
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出目录:"))
        self.output_label = QLabel(str(Path.home() / "Desktop"))
        output_layout.addWidget(self.output_label)
        
        output_btn = QPushButton("浏览...")
        output_btn.clicked.connect(self.choose_output_dir)
        output_layout.addWidget(output_btn)
        options_layout.addLayout(output_layout)
        
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # 进度部分
        progress_group = QGroupBox("进度")
        progress_layout = QVBoxLayout()
        
        self.progress_label = QLabel("准备就绪")
        progress_layout.addWidget(self.progress_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        progress_group.setLayout(progress_layout)
        layout.addWidget(progress_group)
        
        # 控制按钮
        control_layout = QHBoxLayout()
        control_layout.addStretch()
        
        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setMinimumHeight(40)
        self.convert_btn.setMinimumWidth(120)
        self.convert_btn.clicked.connect(self.start_conversion)
        control_layout.addWidget(self.convert_btn)
        
        layout.addLayout(control_layout)
        
        widget.setLayout(layout)
        return widget
    
    def _create_settings_tab(self) -> QWidget:
        """创建设置标签页"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        info_text = QTextEdit()
        info_text.setReadOnly(True)
        info_text.setText("""
<h3>关于本工具</h3>
<p><b>PDF转Word批量转换工具</b></p>

<h4>功能特性:</h4>
<ul>
  <li>✓ 支持批量转换PDF文件</li>
  <li>✓ 支持普通数字PDF和扫描版PDF</li>
  <li>✓ 支持OCR文字识别（扫描版）</li>
  <li>✓ 支持多种语言识别（中文、英文、日文等）</li>
  <li>✓ 友好的图形化界面</li>
  <li>✓ 实时转换日志显示</li>
</ul>

<h4>使用步骤:</h4>
<ol>
  <li>在"转换"标签页添加要转换的PDF文件</li>
  <li>根据需要选择转换选项（是否使用OCR等）</li>
  <li>设置输出目录</li>
  <li>点击"开始转换"按钮</li>
</ol>

<h4>注意事项:</h4>
<ul>
  <li>• 扫描版PDF转换需要安装Tesseract OCR</li>
  <li>• DPI值越高，识别准确度越高，但处理时间越长</li>
  <li>• OCR识别需要较长的处理时间，请耐心等待</li>
</ul>

<h4>系统要求:</h4>
<p>• Python 3.7+</p>
<p>• Tesseract-OCR (用于扫描版PDF识别)</p>
        """)
        layout.addWidget(info_text)
        
        widget.setLayout(layout)
        return widget
    
    def _create_log_tab(self) -> QWidget:
        """创建日志标签页"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)
        
        clear_log_btn = QPushButton("清空日志")
        clear_log_btn.clicked.connect(lambda: self.log_text.clear())
        layout.addWidget(clear_log_btn)
        
        widget.setLayout(layout)
        return widget
    
    def add_files(self):
        """添加PDF文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择PDF文件", "",
            "PDF文件 (*.pdf);;所有文件 (*.*)"
        )
        for f in files:
            if f not in self.pdf_files:
                self.pdf_files.append(f)
        self.update_file_list()
    
    def add_folder(self):
        """添加文件夹中的所有PDF"""
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            for f in Path(folder).glob("*.pdf"):
                if str(f) not in self.pdf_files:
                    self.pdf_files.append(str(f))
            self.update_file_list()
    
    def remove_file(self):
        """移除选中的文件"""
        for item in self.file_list.selectedItems():
            self.pdf_files.remove(item.text())
        self.update_file_list()
    
    def clear_files(self):
        """清空文件列表"""
        reply = QMessageBox.question(
            self, "确认", "确定要清空所有文件吗?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.pdf_files.clear()
            self.update_file_list()
    
    def update_file_list(self):
        """更新文件列表显示"""
        self.file_list.clear()
        for f in self.pdf_files:
            item = QListWidgetItem(f)
            self.file_list.addItem(item)
    
    def on_ocr_changed(self):
        """OCR复选框状态改变"""
        enabled = self.ocr_checkbox.isChecked()
        self.ocr_lang_combo.setEnabled(enabled)
        self.dpi_spinbox.setEnabled(enabled)
    
    def choose_output_dir(self):
        """选择输出目录"""
        folder = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if folder:
            self.output_label.setText(folder)
    
    def start_conversion(self):
        """开始转换"""
        if not self.pdf_files:
            QMessageBox.warning(self, "警告", "请先添加PDF文件!")
            return
        
        output_dir = self.output_label.text()
        if not os.path.exists(output_dir):
            QMessageBox.warning(self, "错误", "输出目录不存在!")
            return
        
        # 禁用按钮
        self.convert_btn.setEnabled(False)
        
        # 获取OCR选项
        use_ocr = self.ocr_checkbox.isChecked()
        ocr_lang_map = {
            '中文简体 (chi_sim)': 'chi_sim',
            '中文繁体 (chi_tra)': 'chi_tra',
            '英文 (eng)': 'eng',
            '日文 (jpn)': 'jpn'
        }
        ocr_lang = ocr_lang_map.get(self.ocr_lang_combo.currentText(), 'chi_sim')
        dpi = self.dpi_spinbox.value()
        
        # 创建工作线程
        self.worker = ConversionWorker(
            self.pdf_files,
            output_dir,
            use_ocr,
            ocr_lang,
            dpi
        )
        self.worker.progress.connect(self.on_progress)
        self.worker.message.connect(self.on_message)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()
        
        self.progress_label.setText("正在转换...")
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] 开始转换 {len(self.pdf_files)} 个文件\n")
    
    def on_progress(self, value):
        """进度更新"""
        self.progress_bar.setValue(value)
    
    def on_message(self, message):
        """消息更新"""
        self.progress_label.setText(message)
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
    
    def on_finished(self, success):
        """转换完成"""
        self.convert_btn.setEnabled(True)
        if success:
            self.progress_label.setText("转换完成!")
            output_dir = self.output_label.text()
            QMessageBox.information(
                self, "成功",
                f"所有文件已转换完成！\n输出目录: {output_dir}"
            )
        else:
            self.progress_label.setText("转换失败")
            QMessageBox.critical(self, "错误", "转换过程中出现错误!")


def main():
    """主函数"""
    app = QApplication(sys.argv)
    window = PDFConverterUI()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
