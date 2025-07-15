import sys
import subprocess


def install(package):
    """自动安装缺失的包"""
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


print("正在检查并安装必要的依赖库...")
try:
    import numpy as np
except ImportError:
    print("NumPy 未安装，正在安装...")
    install("numpy")
    import numpy as np

try:
    import pandas as pd
except (ImportError, ValueError) as e:
    if "numpy.dtype size changed" in str(e):
        print("检测到 NumPy 版本兼容性问题，正在修复...")
        # 先卸载pandas再重新安装
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "uninstall", "-y", "pandas"])
        except:
            pass
        # 安装兼容版本的pandas和numpy
        install("--upgrade numpy")
        install("--upgrade pandas")
        install("--upgrade pandas")
    else:
        print("pandas 未安装，正在安装...")
        install("pandas")

    import pandas as pd

# 以下是主程序
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import re
import xml.etree.ElementTree as ET
import zipfile
import tempfile
import shutil


class DeduplicationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FTD文件去重工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)

        # 定义配色方案
        self.bg_color = "#2c3e50"
        self.header_color = "#3498db"
        self.text_bg = "#ecf0f1"
        self.btn_color = "#1abc9c"
        self.btn_hover = "#16a085"
        self.btn_remove = "#e74c3c"
        self.btn_remove_hover = "#c0392b"
        self.status_color = "#34495e"
        self.format_highlight = {
            "txt": "#1abc9c", "doc": "#3498db", "docx": "#2980b9",
            "xls": "#9b59b6", "xlsx": "#8e44ad"
        }

        # 创建主框架
        self.root.configure(bg=self.bg_color)

        # 标题框架
        header_frame = tk.Frame(root, bg=self.header_color)
        header_frame.pack(fill=tk.X, pady=0)

        title_label = tk.Label(
            header_frame,
            text="▣ FTD文件去重工具",
            font=("微软雅黑", 16, "bold"),
            bg=self.header_color,
            fg="white",
            pady=10
        )
        title_label.pack(pady=5)

        # 创建主内容框架
        main_frame = tk.Frame(root, bg=self.bg_color, padx=15, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 文件格式说明
        format_frame = tk.Frame(main_frame, bg=self.bg_color)
        format_frame.pack(fill=tk.X, pady=5)

        tk.Label(
            format_frame,
            text="支持格式: ",
            bg=self.bg_color,
            fg="#ecf0f1",
            font=("微软雅黑", 9)
        ).pack(side=tk.LEFT)

        # 创建格式标签
        for fmt, color in self.format_highlight.items():
            tk.Label(
                format_frame,
                text=f"{fmt.upper()}",
                bg=self.bg_color,
                fg=color,
                font=("微软雅黑", 9, "bold"),
                padx=5
            ).pack(side=tk.LEFT, padx=2)

        # 输入文件选择
        input_frame = tk.LabelFrame(
            main_frame,
            text=" 选择输入文件 ",
            font=("微软雅黑", 9),
            bg=self.bg_color,
            fg="#ecf0f1"
        )
        input_frame.pack(fill=tk.X, pady=8)

        tk.Label(
            input_frame,
            text="输入文件:",
            bg=self.bg_color,
            fg="#ecf0f1",
            font=("微软雅黑", 9)
        ).grid(row=0, column=0, padx=5, pady=5, sticky='w')

        self.input_path = tk.StringVar()
        input_entry = tk.Entry(
            input_frame,
            textvariable=self.input_path,
            width=50,
            font=("微软雅黑", 9),
            relief=tk.GROOVE
        )
        input_entry.grid(row=0, column=1, padx=5, sticky='ew')

        button_frame1 = tk.Frame(input_frame, bg=self.bg_color)
        button_frame1.grid(row=0, column=2, padx=5)

        tk.Button(
            button_frame1,
            text="浏览...",
            command=self.browse_input,
            bg="#3498db",
            fg="white",
            relief=tk.FLAT,
            font=("微软雅黑", 9, "bold"),
            padx=10
        ).pack(side=tk.LEFT, padx=2)

        # 输出文件选择
        output_frame = tk.LabelFrame(
            main_frame,
            text=" 设置输出文件 ",
            font=("微软雅黑", 9),
            bg=self.bg_color,
            fg="#ecf0f1"
        )
        output_frame.pack(fill=tk.X, pady=8)

        tk.Label(
            output_frame,
            text="输出文件:",
            bg=self.bg_color,
            fg="#ecf0f1",
            font=("微软雅黑", 9)
        ).grid(row=0, column=0, padx=5, pady=5, sticky='w')

        self.output_path = tk.StringVar()
        output_entry = tk.Entry(
            output_frame,
            textvariable=self.output_path,
            width=50,
            font=("微软雅黑", 9),
            relief=tk.GROOVE
        )
        output_entry.grid(row=0, column=1, padx=5, sticky='ew')

        button_frame2 = tk.Frame(output_frame, bg=self.bg_color)
        button_frame2.grid(row=0, column=2, padx=5)

        tk.Button(
            button_frame2,
            text="浏览...",
            command=self.browse_output,
            bg="#3498db",
            fg="white",
            relief=tk.FLAT,
            font=("微软雅黑", 9, "bold"),
            padx=10
        ).pack(side=tk.LEFT, padx=2)

        # 覆盖选项
        option_frame = tk.Frame(main_frame, bg=self.bg_color)
        option_frame.pack(fill=tk.X, pady=5)

        self.overwrite_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            option_frame,
            text="覆盖原文件（输出文件留空时自动启用）",
            variable=self.overwrite_var,
            command=self.update_overwrite,
            bg=self.bg_color,
            fg="#ecf0f1",
            font=("微软雅黑", 9),
            selectcolor=self.bg_color,
            activebackground=self.bg_color,
            activeforeground="#ecf0f1"
        ).pack(anchor=tk.W)

        # 去重范围选择 (仅适用于Word文档)
        self.scope_var = tk.StringVar(value="all")
        scope_frame = tk.Frame(main_frame, bg=self.bg_color)
        scope_frame.pack(fill=tk.X, pady=5)

        tk.Label(
            scope_frame,
            text="Word文档去重范围:",
            bg=self.bg_color,
            fg="#ecf0f1",
            font=("微软雅黑", 9)
        ).pack(side=tk.LEFT)

        scopes = [
            ("全部内容", "all"),
            ("仅段落", "paragraphs"),
            ("仅表格", "tables")
        ]

        for text, value in scopes:
            tk.Radiobutton(
                scope_frame,
                text=text,
                variable=self.scope_var,
                value=value,
                bg=self.bg_color,
                fg="#ecf0f1",
                selectcolor=self.bg_color,
                activebackground=self.bg_color,
                activeforeground="#ecf0f1",
                font=("微软雅黑", 9)
            ).pack(side=tk.LEFT, padx=10)

        # 操作按钮
        btn_frame = tk.Frame(main_frame, bg=self.bg_color)
        btn_frame.pack(fill=tk.X, pady=10)

        btn_style = {
            "font": ("微软雅黑", 10, "bold"),
            "padx": 15,
            "pady": 8,
            "relief": tk.GROOVE,
            "bd": 0
        }

        tk.Button(
            btn_frame,
            text="✓ 执行去重",
            command=self.process_deduplication,
            bg=self.btn_color,
            fg="white",
            activebackground=self.btn_hover,
            **btn_style
        ).pack(side=tk.LEFT, padx=10)

        tk.Button(
            btn_frame,
            text="👁 预览结果",
            command=self.preview_results,
            bg="#f39c12",
            fg="white",
            activebackground="#e67e22",
            **btn_style
        ).pack(side=tk.LEFT, padx=10)

        tk.Button(
            btn_frame,
            text="✕ 退出",
            command=root.destroy,
            bg=self.btn_remove,
            fg="white",
            activebackground=self.btn_remove_hover,
            **btn_style
        ).pack(side=tk.RIGHT, padx=10)

        # 结果显示框架
        result_frame = tk.LabelFrame(
            main_frame,
            text=" 处理结果 ",
            font=("微软雅黑", 9),
            bg=self.bg_color,
            fg="#ecf0f1"
        )
        result_frame.pack(fill=tk.BOTH, expand=True)

        self.result_text = scrolledtext.ScrolledText(
            result_frame,
            wrap=tk.WORD,
            height=8,
            font=("Consolas", 10),
            bg="#ffffff",
            padx=10,
            pady=10,
            relief=tk.GROOVE
        )
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.result_text.config(state=tk.DISABLED)

        # 状态栏
        status_bar = tk.Frame(
            root,
            bg=self.status_color,
            height=22,
            relief=tk.SUNKEN
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.status_var = tk.StringVar(value="就绪 | 选择一个文件开始处理")
        tk.Label(
            status_bar,
            textvariable=self.status_var,
            bg=self.status_color,
            fg="white",
            anchor=tk.W,
            font=("微软雅黑", 9)
        ).pack(side=tk.LEFT, padx=10)

        # 底部版权信息
        copyright_frame = tk.Frame(root, bg=self.bg_color)
        copyright_frame.pack(side=tk.BOTTOM, fill=tk.X)

        tk.Label(
            copyright_frame,
            text="© 2025 FTD文件去重工具 v1.0 | 支持: TXT, DOC, DOCX, XLS, XLSX",
            bg=self.bg_color,
            fg="#95a5a6",
            font=("微软雅黑", 8)
        ).pack(pady=(0, 5))

        # 绑定事件
        self.bind_hover_events()

    def bind_hover_events(self):
        """绑定按钮的悬停事件"""
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Button):
                widget.bind("<Enter>", lambda e: e.widget.config(
                    bg=e.widget.cget("activebackground")
                ))
                widget.bind("<Leave>", lambda e: e.widget.config(
                    bg=e.widget.cget("bg").replace("activebackground", "").split()[0]
                ))

    def get_file_extension(self, file_path):
        """获取文件扩展名（小写，不带点）"""
        if not file_path:
            return None
        ext = os.path.splitext(file_path)[1]
        if ext.startswith('.'):
            ext = ext[1:]
        return ext.lower()

    def browse_input(self):
        """选择输入文件"""
        file_path = filedialog.askopenfilename(
            title="选择输入文件",
            filetypes=[
                ("所有支持的文件", "*.txt *.doc *.docx *.xls *.xlsx"),
                ("文本文件", "*.txt"),
                ("Word 文档", "*.doc *.docx"),
                ("Excel 文件", "*.xls *.xlsx"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.input_path.set(file_path)
            if self.overwrite_var.get() and not self.output_path.get():
                self.output_path.set(file_path)

            ext = self.get_file_extension(file_path)
            if ext in self.format_highlight:
                color = self.format_highlight[ext]
                self.status_var.set(f"已选择输入文件: {os.path.basename(file_path)}")
            else:
                self.status_var.set(f"警告: {ext.upper()}格式支持有限 - {os.path.basename(file_path)}")

    def browse_output(self):
        """选择输出文件"""
        input_file = self.input_path.get()
        input_ext = self.get_file_extension(input_file)

        default_ext = "txt"
        file_types = []

        if input_ext in ["doc", "docx"]:
            file_types = [("Word 文档", "*.docx"), ("所有文件", "*.*")]
            default_ext = "docx"
        elif input_ext in ["xls", "xlsx"]:
            file_types = [("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
            default_ext = "xlsx"
        else:
            file_types = [("文本文件", "*.txt"), ("所有文件", "*.*")]
            default_ext = "txt"

        file_path = filedialog.asksaveasfilename(
            title="保存输出文件",
            defaultextension=f".{default_ext}",
            filetypes=file_types
        )
        if file_path:
            self.output_path.set(file_path)
            self.status_var.set(f"输出文件设置为: {os.path.basename(file_path)}")

    def update_overwrite(self):
        """更新覆盖选项"""
        if self.overwrite_var.get() and self.input_path.get() and not self.output_path.get():
            self.output_path.set(self.input_path.get())

    def validate_inputs(self):
        """验证输入是否有效"""
        input_file = self.input_path.get()
        output_file = self.output_path.get()

        if not input_file:
            messagebox.showerror("输入错误", "请先选择一个输入文件！")
            return False

        if not os.path.exists(input_file):
            messagebox.showerror("文件错误", f"文件不存在:\n{input_file}")
            return False

        ext = self.get_file_extension(input_file)
        supported_formats = ["txt", "doc", "docx", "xls", "xlsx"]
        if ext not in supported_formats:
            messagebox.showerror("格式错误", f"不支持的文件格式: {ext or '未知'}\n\n"
                                             f"支持格式: {', '.join(supported_formats)}")
            return False

        if not output_file:
            messagebox.showerror("输出错误", "请设置输出文件路径！")
            return False

        output_ext = self.get_file_extension(output_file)
        if output_ext != ext:
            if not messagebox.askyesno("格式不同",
                                       f"输出文件格式({output_ext})与输入格式({ext})不同，\n"
                                       "可能导致格式丢失。是否继续？",
                                       icon="warning"):
                return False

        return True

    def extract_text_from_docx(self, file_path):
        """从DOCX文件中提取文本（无Office依赖）"""
        try:
            # 创建一个临时目录用于解压DOCX文件
            with tempfile.TemporaryDirectory() as tmp_dir:
                # 解压DOCX文件
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(tmp_dir)

                # 解析document.xml文件
                doc_xml_path = os.path.join(tmp_dir, 'word', 'document.xml')
                if not os.path.exists(doc_xml_path):
                    return [], 0

                tree = ET.parse(doc_xml_path)
                root = tree.getroot()

                # 定义XML命名空间
                namespaces = {
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                }

                # 提取文本内容
                text_lines = []
                scope = self.scope_var.get()

                # 提取段落文本
                if scope in ["all", "paragraphs"]:
                    for paragraph in root.findall('.//w:p', namespaces):
                        para_text = []
                        for run in paragraph.findall('.//w:r', namespaces):
                            for text in run.findall('.//w:t', namespaces):
                                if text.text:
                                    para_text.append(text.text.strip())
                        if para_text:
                            text_lines.append(''.join(para_text))

                # 提取表格文本
                if scope in ["all", "tables"]:
                    for table in root.findall('.//w:tbl', namespaces):
                        for row in table.findall('.//w:tr', namespaces):
                            row_text = []
                            for cell in row.findall('.//w:tc', namespaces):
                                cell_text = []
                                for paragraph in cell.findall('.//w:p', namespaces):
                                    para_text = []
                                    for run in paragraph.findall('.//w:r', namespaces):
                                        for text in run.findall('.//w:t', namespaces):
                                            if text.text:
                                                para_text.append(text.text.strip())
                                    if para_text:
                                        cell_text.append(''.join(para_text))
                                if cell_text:
                                    row_text.append(' '.join(cell_text))
                            if row_text:
                                text_lines.append('\t'.join(row_text))

            return text_lines, len(text_lines)
        except Exception as e:
            messagebox.showerror("提取错误", f"从DOCX文件中提取内容失败:\n{str(e)}")
            return [], 0

    def extract_text_from_doc(self, file_path):
        """从DOC文件中提取文本（兼容处理）"""
        # 显示警告信息
        messagebox.showwarning(
            "DOC格式限制",
            "DOC文件是旧格式，处理能力有限。\n\n已将其视为文本文件处理。"
        )

        # 尝试作为文本文件提取内容
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                lines = [line.strip() for line in f.readlines()]
            return lines, len(lines)
        except Exception as e:
            messagebox.showerror("处理错误", f"处理DOC文件失败:\n{str(e)}")
            return [], 0

    def extract_text_from_excel(self, file_path):
        """从Excel文件中提取文本（使用pandas）"""
        try:
            # 确定读取引擎
            if file_path.endswith('.xls'):
                import xlrd
                engine = 'xlrd'
            else:
                engine = 'openpyxl'

            sheets = pd.read_excel(file_path, sheet_name=None, engine=engine)

            all_text = []
            total_lines = 0

            for sheet_name, df in sheets.items():
                # 添加表名标题
                all_text.append(f"\n--- Sheet: {sheet_name} ---")

                # 处理表头
                headers = [str(col) for col in df.columns]
                all_text.append("\t".join(headers))

                # 处理数据行
                for idx, row in df.iterrows():
                    row_values = [str(v) for v in row.values]
                    all_text.append("\t".join(row_values))

                total_lines += len(df) + 2

            return all_text, len(all_text)
        except Exception as e:
            messagebox.showerror("提取错误", f"从Excel文件中提取内容失败:\n{str(e)}")
            return [], 0

    def deduplicate_text(self, lines):
        """去重文本内容（保留顺序）"""
        seen = set()
        unique_lines = []

        for line in lines:
            stripped_line = line.strip()
            # 对于表格行，我们按整行比较
            if '\t' in stripped_line:
                key = stripped_line
            else:
                key = stripped_line.lower()

            if key not in seen:
                seen.add(key)
                unique_lines.append(line)

        return unique_lines, len(unique_lines)

    def save_dedup_result(self, unique_lines, input_file, output_file):
        """保存去重结果到文件"""
        input_ext = self.get_file_extension(input_file)
        output_ext = self.get_file_extension(output_file)

        try:
            # 对于Excel文件，保存为Excel格式
            if output_ext in ["xls", "xlsx"]:
                # 提取表头和数据
                header_line = None
                data_lines = []

                for line in unique_lines:
                    if '--- Sheet:' in line:
                        continue
                    if not header_line:
                        header_line = line
                    else:
                        data_lines.append(line)

                # 解析数据
                if header_line and data_lines:
                    headers = header_line.split('\t')
                    data = [line.split('\t') for line in data_lines]

                    # 创建DataFrame
                    df = pd.DataFrame(data, columns=headers)

                    # 保存到Excel
                    if output_ext == 'xlsx':
                        df.to_excel(output_file, index=False, engine='openpyxl')
                    else:
                        df.to_excel(output_file, index=False, engine='xlwt')
                else:
                    # 如果没有数据，创建空DataFrame
                    pd.DataFrame().to_excel(output_file, index=False)

                return True
            # 对于文本和Word文件，保存为文本格式
            else:
                with open(output_file, 'w', encoding='utf-8') as f:
                    for line in unique_lines:
                        if '--- Sheet:' not in line:  # 跳过sheet标题
                            f.write(line + '\n')
                return True
        except Exception as e:
            messagebox.showerror("保存错误", f"保存去重结果失败:\n{str(e)}")
            return False

    def preview_results(self):
        """预览去重结果"""
        if not self.validate_inputs():
            return

        input_file = self.input_path.get()
        output_file = self.output_path.get()
        ext = self.get_file_extension(input_file)

        try:
            lines = []
            original_count = 0

            # 根据文件格式提取内容
            if ext == "txt":
                with open(input_file, 'r', encoding='utf-8') as f:
                    lines = [line.strip() for line in f.readlines()]
                original_count = len(lines)
            elif ext == "doc":
                lines, original_count = self.extract_text_from_doc(input_file)
            elif ext == "docx":
                lines, original_count = self.extract_text_from_docx(input_file)
            elif ext in ["xls", "xlsx"]:
                lines, original_count = self.extract_text_from_excel(input_file)

            if not lines:
                messagebox.showwarning("内容为空", "未提取到任何内容，文件可能为空或格式不受支持")
                return

            # 去重文本
            unique_lines, unique_count = self.deduplicate_text(lines)

            # 显示预览结果
            self.result_text.config(state=tk.NORMAL)
            self.result_text.delete(1.0, tk.END)

            # 标题
            self.result_text.tag_config("header", foreground="#2980b9", font=("微软雅黑", 10, "bold"))
            self.result_text.insert(tk.END, f"文件预览 ({ext.upper()}, 最多15行)\n", "header")
            self.result_text.insert(tk.END, "=" * 60 + "\n\n")

            # 预览内容
            for i, line in enumerate(unique_lines[:15]):
                self.result_text.tag_config("line_num", foreground="#7f8c8d")
                self.result_text.insert(tk.END, f"{i + 1:>2}. ", "line_num")

                # 表格行特殊处理
                if '\t' in line:
                    self.result_text.tag_config("table_row", foreground="#9b59b6")
                    columns = line.split('\t')
                    truncated = [col[:12] + ('...' if len(col) > 15 else '') for col in columns]
                    self.result_text.insert(tk.END, " | ".join(truncated) + "\n", "table_row")
                else:
                    self.result_text.tag_config("text_line", foreground="#2c3e50")
                    # 对长文本进行截断处理
                    if len(line) > 80:
                        line = line[:77] + "..."
                    self.result_text.insert(tk.END, line + "\n", "text_line")

            if len(unique_lines) > 15:
                self.result_text.insert(tk.END, f"\n...以及另外 {len(unique_lines) - 15} 行\n\n", "line_num")
            else:
                self.result_text.insert(tk.END, "\n")

            # 统计数据
            self.result_text.tag_config("stats", foreground="#27ae60", font=("微软雅黑", 9, "bold"))
            self.result_text.insert(tk.END, "统计信息:\n", "stats")
            self.result_text.insert(tk.END, f"原始行数: {original_count}\n")
            self.result_text.insert(tk.END, f"去重后行数: {len(unique_lines)}\n")
            self.result_text.insert(tk.END, f"移除重复行数: {original_count - len(unique_lines)}\n")

            self.result_text.config(state=tk.DISABLED)

            self.status_var.set(f"预览完成: {ext.upper()}文件, 原始行数 {original_count}, 去重后行数 {len(unique_lines)}")

        except Exception as e:
            messagebox.showerror("处理错误", f"处理文件时发生错误:\n{str(e)}")
            self.status_var.set(f"错误: {str(e)}")

    def process_deduplication(self):
        """执行去重操作"""
        if not self.validate_inputs():
            return

        input_file = self.input_path.get()
        output_file = self.output_path.get()
        ext = self.get_file_extension(input_file)

        # 检查是否覆盖原文件
        if input_file == output_file:
            if not messagebox.askyesno("确认覆盖",
                                       "输出文件与输入文件相同，将覆盖原始文件。\n\n是否继续？",
                                       icon="warning"):
                return

        try:
            lines = []
            original_count = 0

            # 根据文件格式提取内容
            self.status_var.set(f"正在处理 {ext.upper()} 文件...")
            self.root.update()

            if ext == "txt":
                with open(input_file, 'r', encoding='utf-8') as f:
                    lines = [line.strip() for line in f.readlines()]
                original_count = len(lines)
            elif ext == "doc":
                lines, original_count = self.extract_text_from_doc(input_file)
            elif ext == "docx":
                lines, original_count = self.extract_text_from_docx(input_file)
            elif ext in ["xls", "xlsx"]:
                lines, original_count = self.extract_text_from_excel(input_file)

            if not lines:
                messagebox.showwarning("内容为空", "未提取到任何内容，文件可能为空或格式不受支持")
                return

            # 去重文本
            unique_lines, unique_count = self.deduplicate_text(lines)

            # 保存结果
            success = self.save_dedup_result(unique_lines, input_file, output_file)

            if not success:
                return

            # 显示结果
            self.result_text.config(state=tk.NORMAL)
            self.result_text.delete(1.0, tk.END)

            # 结果标题
            self.result_text.tag_config("success", foreground="#27ae60", font=("微软雅黑", 11, "bold"))
            self.result_text.insert(tk.END, "✓ 去重操作成功完成！\n\n", "success")

            # 统计信息
            self.result_text.tag_config("stats", foreground="#e74c3c", font=("微软雅黑", 10))
            self.result_text.insert(tk.END, "处理结果统计:\n", "stats")
            self.result_text.insert(tk.END, f"原始行数: {original_count}\n")
            self.result_text.insert(tk.END, f"去重后行数: {len(unique_lines)}\n")
            self.result_text.insert(tk.END, f"移除重复行数: {original_count - len(unique_lines)}\n\n")

            # 文件信息
            self.result_text.tag_config("file", foreground="#3498db", font=("微软雅黑", 9, "bold"))
            self.result_text.insert(tk.END, "文件信息:\n", "file")
            self.result_text.insert(tk.END, f"输入文件: {os.path.basename(input_file)}\n")
            self.result_text.insert(tk.END, f"输出文件: {os.path.basename(output_file)}\n")
            self.result_text.insert(tk.END, f"输出路径: {os.path.dirname(output_file)}\n")

            self.result_text.config(state=tk.DISABLED)

            self.status_var.set(f"去重完成！移除了 {original_count - len(unique_lines)} 行重复内容")

            # 显示成功对话框
            messagebox.showinfo("操作成功",
                                f"文件去重操作成功完成！\n\n"
                                f"格式: {ext.upper()}\n"
                                f"原始行数: {original_count}\n"
                                f"去重后行数: {len(unique_lines)}\n"
                                f"移除了 {original_count - len(unique_lines)} 行重复内容")

        except Exception as e:
            messagebox.showerror("处理错误", f"处理文件时发生错误:\n{str(e)}")
            self.status_var.set(f"错误: {str(e)}")


def center_window(window, width=None, height=None):
    """居中窗口"""
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    if width is None:
        width = window.winfo_width()
    if height is None:
        height = window.winfo_height()

    x = (screen_width - width) // 2
    y = (screen_height - height) // 2 - 20  # 稍微上移一些

    window.geometry(f"{width}x{height}+{x}+{y}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DeduplicationApp(root)
    center_window(root, 800, 600)
    root.mainloop()
