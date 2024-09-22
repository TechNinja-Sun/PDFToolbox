# -*- coding: utf-8 -*-
# Name：孙圣雷
# Time：2024/9/17 下午11:42
'''
Packaging instructions :
pyinstaller --onefile --windowed --icon=combined_icon.ico --add-data "D:/PDFToolbox/combined_icon.ico;." PDFToolbox.py
'''
import subprocess
import sys
import wx
from pdf2docx import Converter
from PIL import Image
import os
from docx2pdf import convert
import ctypes
import webbrowser  # 用于打开GitHub链接


def is_admin():
    """检测是否为管理员权限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_as_admin():
    """请求管理员权限运行当前脚本"""
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)


def run_reg_file():
    reg_file_path = resource_path("env.reg")  # 替换为你的 .reg 文件名
    if os.path.exists(reg_file_path):
        try:
            # 使用 regedit 命令执行 .reg 文件
            subprocess.run(['regedit.exe', '/s', reg_file_path], check=True)
            print("Registry file applied successfully.")
        except subprocess.CalledProcessError:
            print("Failed to apply registry file.")
    else:
        print(f"Registry file not found: {reg_file_path}")


def resource_path(relative_path):
    """用于找到资源文件，尤其在打包后"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def pdf_to_word(pdf_file, word_file):
    cv = Converter(pdf_file)
    cv.convert(word_file, start=0, end=None)
    cv.close()


def word_to_pdf(word_file, pdf_file):
    convert(word_file, pdf_file)


def image_to_pdf(image_file, pdf_file):
    image = Image.open(image_file)
    image.convert("RGB").save(pdf_file)


class PDFToolBoxApp(wx.Frame):
    def __init__(self, *args, **kw):
        super(PDFToolBoxApp, self).__init__(*args, **kw)

        self.SetTitle('PDFToolBox')
        self.SetSize((400, 400))  # 调整大小以适应新增按钮

        icon_path = self.resource_path("combined_icon.ico")
        if os.path.exists(icon_path):
            icon = wx.Icon(icon_path, wx.BITMAP_TYPE_ICO)
            self.SetIcon(icon)
        else:
            wx.MessageBox(f"找不到图标文件: {icon_path}", "错误", wx.OK | wx.ICON_ERROR)

        # 获取文件路径 (如果有传递文件)
        self.file_path = sys.argv[1] if len(sys.argv) > 1 else None

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 创建功能按钮
        btn_pdf_to_word = wx.Button(panel, label="PDF to Word")
        btn_word_to_pdf = wx.Button(panel, label="Word to PDF")
        btn_jpg_to_pdf = wx.Button(panel, label="JPG to PDF")
        btn_png_to_pdf = wx.Button(panel, label="PNG to PDF")

        vbox.Add(btn_pdf_to_word, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_word_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_jpg_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_png_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)

        # 新增 "Author" 和 "GITHUB-Link" 按钮
        btn_author = wx.Button(panel, label="Author")
        btn_github = wx.Button(panel, label="GITHUB-Link")

        vbox.Add(btn_author, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_github, flag=wx.EXPAND | wx.ALL, border=10)

        panel.SetSizer(vbox)

        # 绑定按钮事件
        btn_pdf_to_word.Bind(wx.EVT_BUTTON, self.on_pdf_to_word)
        btn_word_to_pdf.Bind(wx.EVT_BUTTON, self.on_word_to_pdf)
        btn_jpg_to_pdf.Bind(wx.EVT_BUTTON, self.on_jpg_to_pdf)
        btn_png_to_pdf.Bind(wx.EVT_BUTTON, self.on_png_to_pdf)

        # 绑定 "Author" 和 "GITHUB-Link" 按钮事件
        btn_author.Bind(wx.EVT_BUTTON, self.on_author)
        btn_github.Bind(wx.EVT_BUTTON, self.on_github)

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def select_file(self, wildcard):
        fileDialog = wx.FileDialog(self, "选择文件", wildcard=wildcard,
                                   style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return None
        return fileDialog.GetPath()

    def save_file_as(self, default_filename, wildcard):
        fileDialog = wx.FileDialog(self, "保存文件", wildcard=wildcard,
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
                                   defaultFile=default_filename)
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return None
        return fileDialog.GetPath()

    def on_pdf_to_word(self, event):
        pdf_file = self.file_path if self.file_path and self.file_path.endswith('.pdf') else self.select_file(
            "PDF files (*.pdf)|*.pdf")
        if pdf_file:
            word_file = self.save_file_as(pdf_file.replace('.pdf', '.docx'), "Word files (*.docx)|*.docx")
            if word_file:
                pdf_to_word(pdf_file, word_file)
                wx.MessageBox(f"转换成功: {word_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_word_to_pdf(self, event):
        word_file = self.file_path if self.file_path and self.file_path.endswith('.docx') else self.select_file(
            "Word files (*.docx)|*.docx")
        if word_file:
            pdf_file = self.save_file_as(word_file.replace('.docx', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                word_to_pdf(word_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_jpg_to_pdf(self, event):
        image_file = self.file_path if self.file_path and self.file_path.endswith('.jpg') else self.select_file(
            "JPEG files (*.jpg)|*.jpg")
        if image_file:
            pdf_file = self.save_file_as(image_file.replace('.jpg', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                image_to_pdf(image_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_png_to_pdf(self, event):
        image_file = self.file_path if self.file_path and self.file_path.endswith('.png') else self.select_file(
            "PNG files (*.png)|*.png")
        if image_file:
            pdf_file = self.save_file_as(image_file.replace('.png', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                image_to_pdf(image_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_author(self, event):
        """显示作者信息"""
        wx.MessageBox("Author: Your Name\nEmail: your.email@example.com", "Author", wx.OK | wx.ICON_INFORMATION)

    def on_github(self, event):
        """打开 GitHub 链接"""
        github_url = "https://github.com/your-repo"
        webbrowser.open(github_url)
        wx.MessageBox(f"Opening GitHub: {github_url}", "GITHUB-Link", wx.OK | wx.ICON_INFORMATION)


if __name__ == '__main__':
    # 检测是否为管理员权限
    if not is_admin():
        run_as_admin()  # 请求管理员权限重新运行
        sys.exit()  # 退出当前脚本，等管理员权限运行

    # 如果有管理员权限则执行
    run_reg_file()

    app = wx.App(False)
    frame = PDFToolBoxApp(None)
    frame.Show()
    app.MainLoop()

''' V3.0 🐳🐳🐳
import subprocess
import sys
import wx
from pdf2docx import Converter
from PIL import Image
import os
from docx2pdf import convert

def run_reg_file():
    reg_file_path = resource_path("env.reg")  # 替换为你的 .reg 文件名
    if os.path.exists(reg_file_path):
        try:
            # 使用 regedit 命令执行 .reg 文件
            subprocess.run(['regedit.exe', '/s', reg_file_path], check=True)
            print("Registry file applied successfully.")
        except subprocess.CalledProcessError:
            print("Failed to apply registry file.")
    else:
        print(f"Registry file not found: {reg_file_path}")

def resource_path(relative_path):
    """用于找到资源文件，尤其在打包后"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def pdf_to_word(pdf_file, word_file):
    cv = Converter(pdf_file)
    cv.convert(word_file, start=0, end=None)
    cv.close()

def word_to_pdf(word_file, pdf_file):
    convert(word_file, pdf_file)

def image_to_pdf(image_file, pdf_file):
    image = Image.open(image_file)
    image.convert("RGB").save(pdf_file)

class PDFToolBoxApp(wx.Frame):
    def __init__(self, *args, **kw):
        super(PDFToolBoxApp, self).__init__(*args, **kw)

        self.SetTitle('PDFToolBox')
        self.SetSize((400, 300))

        icon_path = self.resource_path("combined_icon.ico")
        if os.path.exists(icon_path):
            icon = wx.Icon(icon_path, wx.BITMAP_TYPE_ICO)
            self.SetIcon(icon)
        else:
            wx.MessageBox(f"找不到图标文件: {icon_path}", "错误", wx.OK | wx.ICON_ERROR)

        # 获取文件路径 (如果有传递文件)
        self.file_path = sys.argv[1] if len(sys.argv) > 1 else None
        # if self.file_path:
        #     wx.MessageBox(f"已传入文件: {self.file_path}", "信息", wx.OK | wx.ICON_INFORMATION)

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        btn_pdf_to_word = wx.Button(panel, label="PDF to Word")
        btn_word_to_pdf = wx.Button(panel, label="Word to PDF")
        btn_jpg_to_pdf = wx.Button(panel, label="JPG to PDF")
        btn_png_to_pdf = wx.Button(panel, label="PNG to PDF")

        vbox.Add(btn_pdf_to_word, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_word_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_jpg_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_png_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)

        panel.SetSizer(vbox)

        btn_pdf_to_word.Bind(wx.EVT_BUTTON, self.on_pdf_to_word)
        btn_word_to_pdf.Bind(wx.EVT_BUTTON, self.on_word_to_pdf)
        btn_jpg_to_pdf.Bind(wx.EVT_BUTTON, self.on_jpg_to_pdf)
        btn_png_to_pdf.Bind(wx.EVT_BUTTON, self.on_png_to_pdf)

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def select_file(self, wildcard):
        fileDialog = wx.FileDialog(self, "选择文件", wildcard=wildcard,
                                   style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return None
        return fileDialog.GetPath()

    def save_file_as(self, default_filename, wildcard):
        fileDialog = wx.FileDialog(self, "保存文件", wildcard=wildcard,
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
                                   defaultFile=default_filename)
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return None
        return fileDialog.GetPath()

    def on_pdf_to_word(self, event):
        pdf_file = self.file_path if self.file_path and self.file_path.endswith('.pdf') else self.select_file("PDF files (*.pdf)|*.pdf")
        if pdf_file:
            word_file = self.save_file_as(pdf_file.replace('.pdf', '.docx'), "Word files (*.docx)|*.docx")
            if word_file:
                pdf_to_word(pdf_file, word_file)
                wx.MessageBox(f"转换成功: {word_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_word_to_pdf(self, event):
        word_file = self.file_path if self.file_path and self.file_path.endswith('.docx') else self.select_file("Word files (*.docx)|*.docx")
        if word_file:
            pdf_file = self.save_file_as(word_file.replace('.docx', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                word_to_pdf(word_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_jpg_to_pdf(self, event):
        image_file = self.file_path if self.file_path and self.file_path.endswith('.jpg') else self.select_file("JPEG files (*.jpg)|*.jpg")
        if image_file:
            pdf_file = self.save_file_as(image_file.replace('.jpg', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                image_to_pdf(image_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_png_to_pdf(self, event):
        image_file = self.file_path if self.file_path and self.file_path.endswith('.png') else self.select_file("PNG files (*.png)|*.png")
        if image_file:
            pdf_file = self.save_file_as(image_file.replace('.png', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                image_to_pdf(image_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)


if __name__ == '__main__':
    run_reg_file()

    app = wx.App(False)
    frame = PDFToolBoxApp(None)
    frame.Show()
    app.MainLoop()'''


''' V2.0 🐳🐳🐳
import sys
import wx
from pdf2docx import Converter
import docx
import pdfkit
from PIL import Image
import os
from docx2pdf import convert


def pdf_to_word(pdf_file, word_file):
    cv = Converter(pdf_file)
    cv.convert(word_file, start=0, end=None)
    cv.close()

def word_to_pdf(word_file, pdf_file):
    convert(word_file, pdf_file)

def image_to_pdf(image_file, pdf_file):
    image = Image.open(image_file)
    image.convert("RGB").save(pdf_file)

class PDFToolBoxApp(wx.Frame):
    def __init__(self, *args, **kw):
        super(PDFToolBoxApp, self).__init__(*args, **kw)

        self.SetTitle('PDFToolBox')
        self.SetSize((400, 300))

        if len(sys.argv) > 1:
            file_path = sys.argv[1]
            wx.MessageBox(f"文件路径: {file_path}", "信息", wx.OK | wx.ICON_INFORMATION)

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        icon = wx.Icon("combined_icon.ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon)

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        btn_pdf_to_word = wx.Button(panel, label="PDF to Word")
        btn_word_to_pdf = wx.Button(panel, label="Word to PDF")
        btn_jpg_to_pdf = wx.Button(panel, label="JPG to PDF")
        btn_png_to_pdf = wx.Button(panel, label="PNG to PDF")

        vbox.Add(btn_pdf_to_word, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_word_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_jpg_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_png_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)

        panel.SetSizer(vbox)

        btn_pdf_to_word.Bind(wx.EVT_BUTTON, self.on_pdf_to_word)
        btn_word_to_pdf.Bind(wx.EVT_BUTTON, self.on_word_to_pdf)
        btn_jpg_to_pdf.Bind(wx.EVT_BUTTON, self.on_jpg_to_pdf)
        btn_png_to_pdf.Bind(wx.EVT_BUTTON, self.on_png_to_pdf)

    def select_file(self, wildcard):
        with wx.FileDialog(self, "选择文件", wildcard=wildcard,
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return None
            return fileDialog.GetPath()

    def save_file_as(self, default_filename, wildcard):
        with wx.FileDialog(self, "保存文件", wildcard=wildcard,
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
                           defaultFile=default_filename) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return None
            return fileDialog.GetPath()

    def on_pdf_to_word(self, event):
        pdf_file = self.select_file("PDF files (*.pdf)|*.pdf")
        if pdf_file:
            word_file = self.save_file_as(pdf_file.replace('.pdf', '.docx'), "Word files (*.docx)|*.docx")
            if word_file:
                pdf_to_word(pdf_file, word_file)
                wx.MessageBox(f"转换成功: {word_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_word_to_pdf(self, event):
        word_file = self.select_file("Word files (*.docx)|*.docx")
        if word_file:
            pdf_file = self.save_file_as(word_file.replace('.docx', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                word_to_pdf(word_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_jpg_to_pdf(self, event):
        image_file = self.select_file("JPEG files (*.jpg)|*.jpg")
        if image_file:
            pdf_file = self.save_file_as(image_file.replace('.jpg', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                image_to_pdf(image_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_png_to_pdf(self, event):
        image_file = self.select_file("PNG files (*.png)|*.png")
        if image_file:
            pdf_file = self.save_file_as(image_file.replace('.png', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                image_to_pdf(image_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)

if __name__ == '__main__':
    app = wx.App(False)
    frame = PDFToolBoxApp(None)
    frame.Show()
    app.MainLoop()
'''



''' V1.1 🐳🐳🐳
def pdf_to_word(pdf_file, word_file):
    cv = Converter(pdf_file)
    cv.convert(word_file, start=0, end=None)
    cv.close()

def word_to_pdf(word_file, pdf_file):
    html_file = word_file.replace('.docx', '.html')
    doc = docx.Document(word_file)
    doc.save(html_file)
    pdfkit.from_file(html_file, pdf_file)
    os.remove(html_file)

def image_to_pdf(image_file, pdf_file):
    image = Image.open(image_file)
    image.convert("RGB").save(pdf_file)

class PDFToolBoxApp(wx.Frame):
    def __init__(self, *args, **kw):
        super(PDFToolBoxApp, self).__init__(*args, **kw)

        self.SetTitle('PDFToolBox')
        self.SetSize((400, 300))

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        btn_pdf_to_word = wx.Button(panel, label="PDF to Word")
        btn_word_to_pdf = wx.Button(panel, label="Word to PDF")
        btn_jpg_to_pdf = wx.Button(panel, label="JPG to PDF")
        btn_png_to_pdf = wx.Button(panel, label="PNG to PDF")

        vbox.Add(btn_pdf_to_word, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_word_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_jpg_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)
        vbox.Add(btn_png_to_pdf, flag=wx.EXPAND | wx.ALL, border=10)

        panel.SetSizer(vbox)

        btn_pdf_to_word.Bind(wx.EVT_BUTTON, self.on_pdf_to_word)
        btn_word_to_pdf.Bind(wx.EVT_BUTTON, self.on_word_to_pdf)
        btn_jpg_to_pdf.Bind(wx.EVT_BUTTON, self.on_jpg_to_pdf)
        btn_png_to_pdf.Bind(wx.EVT_BUTTON, self.on_png_to_pdf)

    def select_file(self, wildcard):
        with wx.FileDialog(self, "选择文件", wildcard=wildcard,
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return None
            return fileDialog.GetPath()

    def save_file_as(self, default_filename, wildcard):
        with wx.FileDialog(self, "保存文件", wildcard=wildcard,
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
                           defaultFile=default_filename) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return None
            return fileDialog.GetPath()

    def on_pdf_to_word(self, event):
        pdf_file = self.select_file("PDF files (*.pdf)|*.pdf")
        if pdf_file:
            word_file = self.save_file_as(pdf_file.replace('.pdf', '.docx'), "Word files (*.docx)|*.docx")
            if word_file:
                pdf_to_word(pdf_file, word_file)
                wx.MessageBox(f"转换成功: {word_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_word_to_pdf(self, event):
        word_file = self.select_file("Word files (*.docx)|*.docx")
        if word_file:
            pdf_file = self.save_file_as(word_file.replace('.docx', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                word_to_pdf(word_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_jpg_to_pdf(self, event):
        image_file = self.select_file("JPEG files (*.jpg)|*.jpg")
        if image_file:
            pdf_file = self.save_file_as(image_file.replace('.jpg', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                image_to_pdf(image_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)

    def on_png_to_pdf(self, event):
        image_file = self.select_file("PNG files (*.png)|*.png")
        if image_file:
            pdf_file = self.save_file_as(image_file.replace('.png', '.pdf'), "PDF files (*.pdf)|*.pdf")
            if pdf_file:
                image_to_pdf(image_file, pdf_file)
                wx.MessageBox(f"转换成功: {pdf_file}", "信息", wx.OK | wx.ICON_INFORMATION)

if __name__ == '__main__':
    app = wx.App(False)
    frame = PDFToolBoxApp(None)
    frame.Show()
    app.MainLoop()
    '''





''' 🐳🐳🐳 （V1.0 use tkinter）
def select_file():
    file_path = filedialog.askopenfilename()
    return file_path

def convert_pdf_to_word():
    pdf_file = select_file()
    if pdf_file:
        word_file = pdf_file.replace('.pdf', '.docx')
        pdf_to_word(pdf_file, word_file)

def convert_word_to_pdf():
    word_file = select_file()
    if word_file:
        pdf_file = word_file.replace('.docx', '.pdf')
        word_to_pdf(word_file, pdf_file)

def convert_image_to_pdf():
    image_file = select_file()
    if image_file:
        pdf_file = image_file.replace('.jpg', '.pdf').replace('.png', '.pdf')
        image_to_pdf(image_file, pdf_file)

app = tk.Tk()
app.title("PDFToolBox")

btn_pdf_to_word = tk.Button(app, text="PDF to Word", command=convert_pdf_to_word)
btn_pdf_to_word.pack()

btn_word_to_pdf = tk.Button(app, text="Word to PDF", command=convert_word_to_pdf)
btn_word_to_pdf.pack()

btn_image_to_pdf = tk.Button(app, text="Image to PDF", command=convert_image_to_pdf)
btn_image_to_pdf.pack()

app.mainloop()
'''