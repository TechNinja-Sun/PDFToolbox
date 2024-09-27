# -*- coding: utf-8 -*-
# Name：孙圣雷
# Time：2024/9/17 下午11:42
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog, QMessageBox
from PyQt5.QtGui import QIcon
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image
import os

import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog,
    QMessageBox, QFrame
)
from PyQt5.QtGui import QIcon, QFont, QColor
from PyQt5.QtCore import Qt
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image
import os

class PDFToolbox(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 设置窗口标题和图标
        self.setWindowTitle('PDF Toolbox')
        self.setWindowIcon(QIcon('combined_icon.ico'))
        self.setStyleSheet('background-color: #2F3B52; color: white;')  # 设置背景和文字颜色

        # 创建布局
        layout = QVBoxLayout()

        # 添加标题
        title_label = QLabel('PDF Toolbox')
        title_label.setFont(QFont('Arial', 24))
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # 创建一个分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        layout.addWidget(separator)

        # PDF to Word 按钮
        btn_pdf_to_word = QPushButton('PDF to Word', self)
        btn_pdf_to_word.setStyleSheet(self.button_style())
        btn_pdf_to_word.clicked.connect(self.pdf_to_word)
        layout.addWidget(btn_pdf_to_word)

        # Word to PDF 按钮
        btn_word_to_pdf = QPushButton('Word to PDF', self)
        btn_word_to_pdf.setStyleSheet(self.button_style())
        btn_word_to_pdf.clicked.connect(self.word_to_pdf)
        layout.addWidget(btn_word_to_pdf)

        # JPG to PDF 按钮
        btn_jpg_to_pdf = QPushButton('JPG to PDF', self)
        btn_jpg_to_pdf.setStyleSheet(self.button_style())
        btn_jpg_to_pdf.clicked.connect(self.jpg_to_pdf)
        layout.addWidget(btn_jpg_to_pdf)

        # PNG to PDF 按钮
        btn_png_to_pdf = QPushButton('PNG to PDF', self)
        btn_png_to_pdf.setStyleSheet(self.button_style())
        btn_png_to_pdf.clicked.connect(self.png_to_pdf)
        layout.addWidget(btn_png_to_pdf)

        # 显示作者信息
        author_label = QLabel('Author: Shenglei Sun')
        author_label.setFont(QFont('Arial', 12))
        author_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(author_label)

        # 显示GitHub链接
        # github_label = QLabel('<a href="https://github.com/TechNinja-Sun/PDFToolbox">GitHub Repository</a>')
        github_label = QLabel('<a href="https://github.com/TechNinja-Sun/PDFToolbox" style="color:white;">GitHub Repository</a>')
        github_label.setFont(QFont('Arial', 14))
        github_label.setStyleSheet("QLabel { color : white; }")
        github_label.setAlignment(Qt.AlignCenter)
        github_label.setOpenExternalLinks(True)
        layout.addWidget(github_label)

        # 设置布局
        self.setLayout(layout)

    def button_style(self):
        return """
        QPushButton {
            background-color: #405D73;
            color: white;
            font-size: 20px;
            font-weight: bold;
            border-radius: 10px;
            padding: 10px;
        }
        QPushButton:hover {
            background-color: #1E90FF;
        }
        """

    # PDF to Word
    def pdf_to_word(self):
        pdf_file, _ = QFileDialog.getOpenFileName(self, 'Select PDF file', '', 'PDF files (*.pdf)')
        if pdf_file:
            word_file, _ = QFileDialog.getSaveFileName(self, 'Save Word file', '', 'Word files (*.docx)')
            if word_file:
                cv = Converter(pdf_file)
                cv.convert(word_file, start=0, end=None)
                cv.close()
                QMessageBox.information(self, 'Success', 'PDF has been converted to Word successfully.')

    # Word to PDF
    def word_to_pdf(self):
        word_file, _ = QFileDialog.getOpenFileName(self, 'Select Word file', '', 'Word files (*.docx)')
        if word_file:
            convert(word_file)
            QMessageBox.information(self, 'Success', 'Word has been converted to PDF successfully.')

    # JPG to PDF
    def jpg_to_pdf(self):
        image_file, _ = QFileDialog.getOpenFileName(self, 'Select JPG file', '', 'Image files (*.jpg)')
        if image_file:
            pdf_file, _ = QFileDialog.getSaveFileName(self, 'Save PDF file', '', 'PDF files (*.pdf)')
            if pdf_file:
                image = Image.open(image_file)
                rgb_im = image.convert('RGB')
                rgb_im.save(pdf_file)
                QMessageBox.information(self, 'Success', 'JPG has been converted to PDF successfully.')

    # PNG to PDF
    def png_to_pdf(self):
        image_file, _ = QFileDialog.getOpenFileName(self, 'Select PNG file', '', 'Image files (*.png)')
        if image_file:
            pdf_file, _ = QFileDialog.getSaveFileName(self, 'Save PDF file', '', 'PDF files (*.pdf)')
            if pdf_file:
                image = Image.open(image_file)
                rgb_im = image.convert('RGB')
                rgb_im.save(pdf_file)
                QMessageBox.information(self, 'Success', 'PNG has been converted to PDF successfully.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    toolbox = PDFToolbox()
    toolbox.resize(500, 400)  # 调整窗口大小
    toolbox.show()
    sys.exit(app.exec_())


''' V1.0 🐳🐳🐳
class PDFToolbox(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 设置窗口标题和图标
        self.setWindowTitle('PDF Toolbox')
        self.setWindowIcon(QIcon('combined_icon.ico'))

        # 创建布局
        layout = QVBoxLayout()

        # PDF to Word 按钮
        btn_pdf_to_word = QPushButton('PDF to Word', self)
        btn_pdf_to_word.clicked.connect(self.pdf_to_word)
        layout.addWidget(btn_pdf_to_word)

        # Word to PDF 按钮
        btn_word_to_pdf = QPushButton('Word to PDF', self)
        btn_word_to_pdf.clicked.connect(self.word_to_pdf)
        layout.addWidget(btn_word_to_pdf)

        # JPG to PDF 按钮
        btn_jpg_to_pdf = QPushButton('JPG to PDF', self)
        btn_jpg_to_pdf.clicked.connect(self.jpg_to_pdf)
        layout.addWidget(btn_jpg_to_pdf)

        # PNG to PDF 按钮
        btn_png_to_pdf = QPushButton('PNG to PDF', self)
        btn_png_to_pdf.clicked.connect(self.png_to_pdf)
        layout.addWidget(btn_png_to_pdf)

        # 显示作者信息
        author_label = QLabel('Author: Your Name')
        layout.addWidget(author_label)

        # 显示GitHub链接
        github_label = QLabel('<a href="https://github.com/TechNinja-Sun/PDFToolbox">GitHub Repository</a>')
        github_label.setOpenExternalLinks(True)
        layout.addWidget(github_label)

        # 设置布局
        self.setLayout(layout)

    # PDF to Word
    def pdf_to_word(self):
        pdf_file, _ = QFileDialog.getOpenFileName(self, 'Select PDF file', '', 'PDF files (*.pdf)')
        if pdf_file:
            word_file, _ = QFileDialog.getSaveFileName(self, 'Save Word file', '', 'Word files (*.docx)')
            if word_file:
                cv = Converter(pdf_file)
                cv.convert(word_file, start=0, end=None)
                cv.close()
                QMessageBox.information(self, 'Success', 'PDF has been converted to Word successfully.')

    # Word to PDF
    def word_to_pdf(self):
        word_file, _ = QFileDialog.getOpenFileName(self, 'Select Word file', '', 'Word files (*.docx)')
        if word_file:
            convert(word_file)
            QMessageBox.information(self, 'Success', 'Word has been converted to PDF successfully.')

    # JPG to PDF
    def jpg_to_pdf(self):
        image_file, _ = QFileDialog.getOpenFileName(self, 'Select JPG file', '', 'Image files (*.jpg)')
        if image_file:
            pdf_file, _ = QFileDialog.getSaveFileName(self, 'Save PDF file', '', 'PDF files (*.pdf)')
            if pdf_file:
                image = Image.open(image_file)
                rgb_im = image.convert('RGB')
                rgb_im.save(pdf_file)
                QMessageBox.information(self, 'Success', 'JPG has been converted to PDF successfully.')

    # PNG to PDF
    def png_to_pdf(self):
        image_file, _ = QFileDialog.getOpenFileName(self, 'Select PNG file', '', 'Image files (*.png)')
        if image_file:
            pdf_file, _ = QFileDialog.getSaveFileName(self, 'Save PDF file', '', 'PDF files (*.pdf)')
            if pdf_file:
                image = Image.open(image_file)
                rgb_im = image.convert('RGB')
                rgb_im.save(pdf_file)
                QMessageBox.information(self, 'Success', 'PNG has been converted to PDF successfully.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    toolbox = PDFToolbox()
    toolbox.resize(400, 300)
    toolbox.show()
    sys.exit(app.exec_())
'''