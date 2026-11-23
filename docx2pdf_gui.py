#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX2PDF 图形界面转换器
使用 docx2pdf 库实现 DOC/DOCX 到 PDF 的完美转换
支持批量转换和进度显示
基于 PyQt5 实现
"""

import sys
import os
import glob
import time
from pathlib import Path

# 导入 PyQt5
try:
    from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                                QHBoxLayout, QPushButton, QLineEdit, QListWidget,
                                QFileDialog, QLabel, QProgressBar, QTextEdit,
                                QGroupBox, QCheckBox, QListWidgetItem, QMessageBox,
                                QComboBox)
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
    from PyQt5.QtGui import QFont
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False

# 导入各种转换引擎
try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False


class ConvertThread(QThread):
    """文件转换线程"""
    progress_updated = pyqtSignal(int, int)
    log_updated = pyqtSignal(str)
    conversion_complete = pyqtSignal(int, int)

    def __init__(self, files, output_dir, engine):
        super().__init__()
        self.files = files
        self.output_dir = output_dir
        self.engine = engine
        self.running = True
        
    def run(self):
        total = len(self.files)
        success_count = 0
        fail_count = 0
        
        if total == 0:
            self.log_updated.emit("错误: 没有选择任何文件")
            self.conversion_complete.emit(0, 0)
            return
        
        os.makedirs(self.output_dir, exist_ok=True)
        self.log_updated.emit(f"使用 {self.engine} 引擎开始转换 {total} 个文件...")

        for i, file_path in enumerate(self.files):
            if not self.running:
                break
                
            try:
                filename = os.path.basename(file_path)
                output_path = os.path.join(self.output_dir, 
                                          os.path.splitext(filename)[0] + '.pdf')
                
                self.log_updated.emit(f"正在转换: {filename}...")
                
                if self.engine == "docx2pdf":
                    docx2pdf_convert(file_path, output_path)
                elif self.engine == "win32com":
                    word = win32com.client.Dispatch("Word.Application")
                    try:
                        doc = word.Documents.Open(file_path)
                        doc.SaveAs(output_path, FileFormat=17)  # 17 for PDF
                        doc.Close()
                    finally:
                        word.Quit()

                success_count += 1
                self.progress_updated.emit(success_count + fail_count, total)
                self.log_updated.emit(f"成功: {filename} -> {os.path.basename(output_path)}")
                
            except Exception as e:
                fail_count += 1
                self.log_updated.emit(f"失败: {os.path.basename(file_path)} - {str(e)}")
        
        self.conversion_complete.emit(success_count, fail_count)
        self.log_updated.emit(f"\n=== 转换完成 ===")
        self.log_updated.emit(f"成功: {success_count} 个文件")
        self.log_updated.emit(f"失败: {fail_count} 个文件")
        self.log_updated.emit(f"输出目录: {self.output_dir}")


class Docx2PdfGUI(QMainWindow):
    """DOCX 到 PDF 转换器的图形界面"""
    
    def __init__(self):
        super().__init__()
        self.files = []
        self.output_dir = str(Path.home() / 'Documents')
        self.convert_thread = None
        self.engine = "docx2pdf"
        self.init_ui()
        
    def init_ui(self):
        """初始化界面"""
        self.setWindowTitle('DOCX2PDF 转换器')
        self.setGeometry(100, 100, 800, 600)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 标题
        title_label = QLabel('DOCX/DOC 到 PDF 转换器')
        title_label.setFont(QFont('Arial', 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet('color: #1a73e8;')
        main_layout.addWidget(title_label)
        
        # 引擎选择区域
        engine_group = QGroupBox('转换引擎')
        engine_layout = QHBoxLayout()
        engine_layout.addWidget(QLabel('选择引擎:'))

        self.engine_combo = QComboBox()
        if DOCX2PDF_AVAILABLE:
            self.engine_combo.addItem("docx2pdf (默认, 跨平台)", "docx2pdf")
        else:
            self.engine_combo.addItem("docx2pdf (不可用)", "docx2pdf")

        if WIN32COM_AVAILABLE:
            self.engine_combo.addItem("win32com (Word COM)", "win32com")
        else:
            self.engine_combo.addItem("win32com (不可用)", "win32com")
        
        self.engine_combo.currentIndexChanged.connect(self.engine_changed)
        engine_layout.addWidget(self.engine_combo)

        engine_group.setLayout(engine_layout)
        main_layout.addWidget(engine_group)
        
        # 文件选择区域
        file_group = QGroupBox('选择文件')
        file_layout = QVBoxLayout()
        
        file_input_layout = QHBoxLayout()
        self.file_input = QLineEdit()
        self.file_input.setReadOnly(True)
        browse_btn = QPushButton('浏览...')
        browse_btn.clicked.connect(self.browse_file)
        file_input_layout.addWidget(self.file_input)
        file_input_layout.addWidget(browse_btn)
        file_layout.addLayout(file_input_layout)
        
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.MultiSelection)
        file_layout.addWidget(self.file_list)
        
        button_layout = QHBoxLayout()
        add_btn = QPushButton('添加文件')
        add_btn.clicked.connect(self.add_file)
        remove_btn = QPushButton('移除选择')
        remove_btn.clicked.connect(self.remove_selected)
        clear_btn = QPushButton('清空列表')
        clear_btn.clicked.connect(self.clear_files)
        button_layout.addWidget(add_btn)
        button_layout.addWidget(remove_btn)
        button_layout.addWidget(clear_btn)
        file_layout.addLayout(button_layout)
        
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)
        
        # 输出设置区域
        output_group = QGroupBox('输出设置')
        output_layout = QHBoxLayout()
        
        output_layout.addWidget(QLabel('输出目录:'))
        self.output_input = QLineEdit(self.output_dir)
        output_layout.addWidget(self.output_input)
        browse_output_btn = QPushButton('浏览...')
        browse_output_btn.clicked.connect(self.browse_output_dir)
        output_layout.addWidget(browse_output_btn)
        
        self.open_output_checkbox = QCheckBox('转换完成后打开输出目录')
        self.open_output_checkbox.setChecked(True)
        output_layout.addWidget(self.open_output_checkbox)
        
        output_group.setLayout(output_layout)
        main_layout.addWidget(output_group)
        
        # 进度显示区域
        progress_group = QGroupBox('转换进度')
        progress_layout = QVBoxLayout()
        
        self.status_label = QLabel('就绪')
        self.status_label.setStyleSheet('color: green;')
        progress_layout.addWidget(self.status_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        self.count_label = QLabel('0/0 文件已转换')
        self.count_label.setAlignment(Qt.AlignCenter)
        progress_layout.addWidget(self.count_label)
        
        progress_group.setLayout(progress_layout)
        main_layout.addWidget(progress_group)
        
        # 按钮区域
        button_group = QGroupBox('')
        button_layout = QHBoxLayout()
        
        self.convert_btn = QPushButton('开始转换')
        self.convert_btn.setStyleSheet('background-color: #1a73e8; color: white;')
        self.convert_btn.setFont(QFont('Arial', 12, QFont.Bold))
        self.convert_btn.clicked.connect(self.start_conversion)
        button_layout.addWidget(self.convert_btn)
        
        batch_btn = QPushButton('批量转换')
        batch_btn.clicked.connect(self.batch_convert)
        button_layout.addWidget(batch_btn)
        
        exit_btn = QPushButton('退出')
        exit_btn.clicked.connect(self.close)
        button_layout.addWidget(exit_btn)
        
        button_group.setLayout(button_layout)
        main_layout.addWidget(button_group)
        
        # 日志区域
        log_group = QGroupBox('转换日志')
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont('Consolas', 9))
        log_layout.addWidget(self.log_text)
        
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        
        main_layout.setStretch(0, 0)
        main_layout.setStretch(1, 1)
        main_layout.setStretch(2, 1)
        main_layout.setStretch(3, 1)
        main_layout.setStretch(4, 2)
        main_layout.setStretch(5, 3)
        
    def engine_changed(self):
        """引擎改变回调"""
        idx = self.engine_combo.currentIndex()
        self.engine = self.engine_combo.itemData(idx, role=Qt.UserRole)

    def browse_file(self):
        """浏览选择文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self, '选择 Word 文件', '', 
            'Word Files (*.doc *.docx);;All Files (*)'
        )
        if files:
            for file_path in files:
                if file_path not in self.files:
                    self.files.append(file_path)
                    self.file_list.addItem(file_path)
                    
    def add_file(self):
        """添加文件到列表"""
        self.browse_file()
        
    def remove_selected(self):
        """移除选择的文件"""
        selected_items = self.file_list.selectedItems()
        for item in selected_items:
            file_path = item.text()
            self.files.remove(file_path)
            self.file_list.takeItem(self.file_list.row(item))
            
    def clear_files(self):
        """清空文件列表"""
        self.files.clear()
        self.file_list.clear()
        
    def browse_output_dir(self):
        """浏览选择输出目录"""
        directory = QFileDialog.getExistingDirectory(
            self, '选择输出目录', self.output_dir
        )
        if directory:
            self.output_dir = directory
            self.output_input.setText(directory)
            
    def log_message(self, message):
        """添加日志消息"""
        timestamp = time.strftime('%H:%M:%S')
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        
    def update_progress(self, current, total):
        """更新进度条"""
        if total > 0:
            percentage = int((current / total) * 100)
            self.progress_bar.setValue(percentage)
            self.count_label.setText(f"{current}/{total} 文件已转换")
            
    def conversion_complete(self, success, fail):
        """转换完成"""
        self.convert_btn.setEnabled(True)
        self.status_label.setText('转换完成!')
        self.status_label.setStyleSheet('color: green;')
        
        if self.open_output_checkbox.isChecked() and self.output_dir:
            try:
                if sys.platform == 'win32':
                    os.startfile(self.output_dir)
                elif sys.platform == 'darwin':
                    os.system(f'open "{self.output_dir}"')
                else:
                    os.system(f'xdg-open "{self.output_dir}"')
            except Exception as e:
                self.log_message(f"无法打开输出目录: {e}")
                
    def start_conversion(self):
        """开始转换"""
        if not self.files:
            QMessageBox.warning(self, '提示', '请先添加要转换的文件!')
            return
            
        if not self.output_dir:
            QMessageBox.warning(self, '提示', '请选择输出目录!')
            return
            
        if self.engine == "docx2pdf" and not DOCX2PDF_AVAILABLE:
            QMessageBox.critical(self, '错误', 'docx2pdf 未安装！\n请运行: pip install docx2pdf')
            return

        if self.engine == "win32com" and not WIN32COM_AVAILABLE:
            QMessageBox.critical(self, '错误', 'win32com 未安装或无 Microsoft Word！\n请运行: pip install pywin32')
            return

        self.convert_btn.setEnabled(False)
        self.status_label.setText('转换中...')
        self.status_label.setStyleSheet('color: orange;')
        self.progress_bar.setValue(0)
        self.count_label.setText('0/0 文件已转换')

        self.convert_thread = ConvertThread(self.files, self.output_dir, self.engine)
        self.convert_thread.progress_updated.connect(self.update_progress)
        self.convert_thread.log_updated.connect(self.log_message)
        self.convert_thread.conversion_complete.connect(self.conversion_complete)
        self.convert_thread.started.connect(lambda: self.log_message(f"使用 {self.engine} 引擎开始转换 {len(self.files)} 个文件..."))
        self.convert_thread.start()
        
    def batch_convert(self):
        """批量转换"""
        directory = QFileDialog.getExistingDirectory(
            self, '选择要转换的目录', ''
        )
        if not directory:
            return
            
        self.log_message(f"扫描目录: {directory}")
        
        doc_files = []
        for ext in ['*.doc', '*.docx']:
            pattern = os.path.join(directory, ext)
            found_files = glob.glob(pattern)
            doc_files.extend(found_files)
        
        if not doc_files:
            QMessageBox.critical(self, '错误', '未找到任何 DOC/DOCX 文件!')
            self.log_message("错误: 未找到任何 DOC/DOCX 文件")
            return
        
        self.log_message(f"找到 {len(doc_files)} 个文件待转换")
        
        self.clear_files()
        for file_path in doc_files:
            self.files.append(file_path)
            self.file_list.addItem(file_path)
            
        self.output_dir = directory
        self.output_input.setText(directory)
        
        self.start_conversion()


def main():
    """主函数"""
    if not GUI_AVAILABLE:
        print("错误: PyQt5 未安装")
        print("请运行: pip install PyQt5")
        return
        
    print("=" * 60)
    print("DOCX/DOC 到 PDF 转换器")
    print("=" * 60)
    
    try:
        import docx2pdf
        print("✓ docx2pdf 已安装")
    except ImportError:
        print("✗ docx2pdf 未安装")
        print("如需使用 docx2pdf 引擎，请运行: pip install docx2pdf")
    
    try:
        import win32com.client
        print("✓ win32com 已安装")
    except ImportError:
        print("✗ win32com 未安装")
        print("如需使用 Microsoft Word 转换引擎，请运行: pip install pywin32")

    app = QApplication(sys.argv)
    window = Docx2PdfGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
