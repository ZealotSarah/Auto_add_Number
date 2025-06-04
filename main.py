import os
import re
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
from pathlib import Path
from datetime import datetime
import docx
from PyPDF2 import PdfReader
import traceback


class FileRenameApp:
    def __init__(self, root):
        self.root = root
        self.root.title("多功能文件重命名工具")
        self.root.geometry("900x700")
        self.font = ('Microsoft YaHei UI', 10)

        # 存储预览结果
        self.preview_results = []
        self.is_single_file = False  # 标记是否为单个文件模式
        self.current_mode = "batch"  # batch, single, title

        # 创建界面元素
        self.create_widgets()

    def create_widgets(self):
        # 模式选择框架
        mode_frame = tk.Frame(self.root, padx=10, pady=5)
        mode_frame.pack(fill=tk.X)

        self.mode_var = tk.StringVar(value="batch")
        batch_radio = tk.Radiobutton(mode_frame, text="批量重命名",
                                     variable=self.mode_var, value="batch",
                                     font=self.font, command=self.change_mode)
        batch_radio.pack(side=tk.LEFT, padx=10)

        single_radio = tk.Radiobutton(mode_frame, text="单个文件重命名",
                                      variable=self.mode_var, value="single",
                                      font=self.font, command=self.change_mode)
        single_radio.pack(side=tk.LEFT, padx=10)

        title_radio = tk.Radiobutton(mode_frame, text="按文件标题重命名",
                                     variable=self.mode_var, value="title",
                                     font=self.font, command=self.change_mode)
        title_radio.pack(side=tk.LEFT, padx=10)

        # 路径选择框架
        path_frame = tk.Frame(self.root, padx=10, pady=10)
        path_frame.pack(fill=tk.X)

        tk.Label(path_frame, text="文件/文件夹路径:", font=self.font).pack(side=tk.LEFT)

        self.path_entry = tk.Entry(path_frame, width=50, font=self.font)
        self.path_entry.pack(side=tk.LEFT, padx=5)

        self.folder_btn = tk.Button(path_frame, text="浏览文件夹", font=self.font,
                                    command=self.browse_folder)
        self.folder_btn.pack(side=tk.LEFT, padx=5)

        self.file_btn = tk.Button(path_frame, text="浏览文件", font=self.font,
                                  command=self.browse_file)
        self.file_btn.pack(side=tk.LEFT, padx=5)

        self.mode_label = tk.Label(path_frame, text="批量处理模式", font=self.font, fg="blue")
        self.mode_label.pack(side=tk.LEFT, padx=10)

        # 批量/单个重命名选项框架
        self.rename_options_frame = tk.Frame(self.root, padx=10, pady=5)
        self.rename_options_frame.pack(fill=tk.X)

        self.keep_year_var = tk.BooleanVar(value=True)
        keep_year_check = tk.Checkbutton(self.rename_options_frame, text="保留年份前缀",
                                         variable=self.keep_year_var, font=self.font)
        keep_year_check.pack(side=tk.LEFT)

        self.overwrite_existing_var = tk.BooleanVar(value=True)
        overwrite_check = tk.Checkbutton(self.rename_options_frame, text="覆盖已有序号",
                                         variable=self.overwrite_existing_var, font=self.font)
        overwrite_check.pack(side=tk.LEFT, padx=10)

        # 命名格式选项
        format_frame = tk.Frame(self.root, padx=10, pady=5)
        format_frame.pack(fill=tk.X)

        tk.Label(format_frame, text="命名格式:", font=self.font).pack(side=tk.LEFT)

        self.format_var = tk.StringVar(value="{序号}_{原文件名}_{时间}")
        format_options = [
            "{序号}、{原文件名}",
            "{序号}_{原文件名}",
            "{时间}_{原文件名}",
            "{序号}_{时间}_{原文件名}",
            "{原文件名}_{序号}",
            "{原文件名}_{时间}",
            "{序号}_{原文件名}_{时间}"
        ]
        self.format_combo = ttk.Combobox(format_frame, textvariable=self.format_var,
                                         values=format_options, width=25, font=self.font)
        self.format_combo.pack(side=tk.LEFT, padx=5)

        # 时间格式选项
        tk.Label(format_frame, text="时间格式:", font=self.font).pack(side=tk.LEFT, padx=(10, 5))

        self.time_format_var = tk.StringVar(value="%Y%m%d")
        time_options = [
            "%Y%m%d",  # 20230101
            "%Y-%m-%d",  # 2023-01-01
            "%Y%m%d%H%M",  # 202301011200
            "%Y-%m-%d_%H%M",  # 2023-01-01_1200
            "%Y%m%d_%H%M%S",  # 20230101_120000
        ]
        time_combo = ttk.Combobox(format_frame, textvariable=self.time_format_var,
                                  values=time_options, width=15, font=self.font)
        time_combo.pack(side=tk.LEFT, padx=5)

        # 时间源选项
        self.time_source_var = tk.StringVar(value="current")
        current_radio = tk.Radiobutton(format_frame, text="当前时间",
                                       variable=self.time_source_var, value="current", font=self.font)
        current_radio.pack(side=tk.LEFT, padx=5)

        modify_radio = tk.Radiobutton(format_frame, text="修改时间",
                                      variable=self.time_source_var, value="modify", font=self.font)
        modify_radio.pack(side=tk.LEFT, padx=5)

        create_radio = tk.Radiobutton(format_frame, text="创建时间",
                                      variable=self.time_source_var, value="create", font=self.font)
        create_radio.pack(side=tk.LEFT, padx=5)

        # 标题重命名选项框架
        self.title_options_frame = tk.Frame(self.root, padx=10, pady=5)
        self.title_options_frame.pack(fill=tk.X)

        self.auto_rename_title_var = tk.BooleanVar(value=False)
        auto_rename_title_check = tk.Checkbutton(self.title_options_frame, text="自动重命名",
                                                 variable=self.auto_rename_title_var, font=self.font)
        auto_rename_title_check.pack(side=tk.LEFT, padx=5)

        self.title_prefix_var = tk.StringVar(value="")
        prefix_label = tk.Label(self.title_options_frame, text="前缀:", font=self.font)
        prefix_label.pack(side=tk.LEFT, padx=5)
        prefix_entry = tk.Entry(self.title_options_frame, textvariable=self.title_prefix_var, width=10, font=self.font)
        prefix_entry.pack(side=tk.LEFT, padx=5)

        self.title_suffix_var = tk.StringVar(value="")
        suffix_label = tk.Label(self.title_options_frame, text="后缀:", font=self.font)
        suffix_label.pack(side=tk.LEFT, padx=5)
        suffix_entry = tk.Entry(self.title_options_frame, textvariable=self.title_suffix_var, width=10, font=self.font)
        suffix_entry.pack(side=tk.LEFT, padx=5)

        # 预览结果框架
        preview_frame = tk.Frame(self.root, padx=10, pady=10)
        preview_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧 - 原始文件列表
        left_frame = tk.Frame(preview_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.original_label = tk.Label(left_frame, text="原始文件列表:", font=self.font, fg="blue")
        self.original_label.pack(anchor=tk.W)

        self.original_files_text = scrolledtext.ScrolledText(left_frame, width=40, height=15, font=self.font)
        self.original_files_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # 右侧 - 重命名预览
        right_frame = tk.Frame(preview_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        self.preview_label = tk.Label(right_frame, text="重命名预览:", font=self.font, fg="green")
        self.preview_label.pack(anchor=tk.W)

        self.preview_text = scrolledtext.ScrolledText(right_frame, width=40, height=15, font=self.font)
        self.preview_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # 按钮框架
        btn_frame = tk.Frame(self.root, padx=10, pady=10)
        btn_frame.pack(fill=tk.X)

        self.preview_btn = tk.Button(btn_frame, text="预览重命名", font=self.font,
                                     command=self.preview_rename)
        self.preview_btn.pack(side=tk.LEFT, padx=5)

        self.extract_title_btn = tk.Button(btn_frame, text="提取标题", font=self.font,
                                           command=self.extract_title)
        self.extract_title_btn.pack(side=tk.LEFT, padx=5)

        self.execute_btn = tk.Button(btn_frame, text="执行重命名", font=self.font,
                                     command=self.execute_rename, state=tk.DISABLED)
        self.execute_btn.pack(side=tk.LEFT, padx=5)

        # 状态框架
        status_frame = tk.Frame(self.root, padx=10, pady=5)
        status_frame.pack(fill=tk.X)

        self.status_var = tk.StringVar(value="就绪")
        status_label = tk.Label(status_frame, textvariable=self.status_var, font=self.font, fg="blue")
        status_label.pack(anchor=tk.W)

        # 添加变量说明
        help_frame = tk.Frame(self.root, padx=10, pady=5)
        help_frame.pack(fill=tk.X)

        help_text = "支持的变量: {序号} - 自动编号, {原文件名} - 原始文件名, {时间} - 格式化时间, {扩展名} - 文件扩展名"
        tk.Label(help_frame, text=help_text, font=self.font, fg="gray").pack(anchor=tk.W)

        # 初始化界面
        self.change_mode()

    def change_mode(self):
        self.current_mode = self.mode_var.get()

        if self.current_mode == "batch":
            self.is_single_file = False
            self.mode_label.config(text="批量处理模式")
            self.folder_btn.config(text="浏览文件夹", command=self.browse_folder)
            self.rename_options_frame.pack(fill=tk.X)
            self.title_options_frame.pack_forget()
            self.preview_btn.pack(side=tk.LEFT, padx=5)
            self.extract_title_btn.pack_forget()
            self.original_label.config(text="原始文件列表:")
            self.preview_label.config(text="重命名预览:")

        elif self.current_mode == "single":
            self.is_single_file = True
            self.mode_label.config(text="单个文件模式")
            self.folder_btn.config(text="浏览文件", command=self.browse_file)
            self.rename_options_frame.pack(fill=tk.X)
            self.title_options_frame.pack_forget()
            self.preview_btn.pack(side=tk.LEFT, padx=5)
            self.extract_title_btn.pack_forget()
            self.original_label.config(text="原始文件:")
            self.preview_label.config(text="重命名预览:")

        elif self.current_mode == "title":
            self.is_single_file = True
            self.mode_label.config(text="按标题重命名模式")
            self.folder_btn.config(text="浏览文件", command=self.browse_file)
            self.rename_options_frame.pack_forget()
            self.title_options_frame.pack(fill=tk.X)
            self.preview_btn.pack_forget()
            self.extract_title_btn.pack(side=tk.LEFT, padx=5)
            self.original_label.config(text="原始文件名:")
            self.preview_label.config(text="提取结果:")

        self.path_entry.delete(0, tk.END)
        self.original_files_text.delete(1.0, tk.END)
        self.preview_text.delete(1.0, tk.END)
        self.execute_btn.config(state=tk.DISABLED)
        self.status_var.set("就绪")

    def browse_folder(self):
        if self.current_mode == "batch":
            folder_selected = filedialog.askdirectory()
            if folder_selected:
                self.path_entry.delete(0, tk.END)
                self.path_entry.insert(0, folder_selected)
                self.status_var.set(f"已选择文件夹: {os.path.basename(folder_selected)}")
        else:  # single 或 title 模式
            file_selected = filedialog.askopenfilename(
                filetypes=[("所有文件", "*.*")]
            )
            if file_selected:
                self.path_entry.delete(0, tk.END)
                self.path_entry.insert(0, file_selected)
                self.status_var.set(f"已选择文件: {os.path.basename(file_selected)}")

    def browse_file(self):
        if self.current_mode == "title":
            file_selected = filedialog.askopenfilename(
                filetypes=[("Word/PDF文件", "*.docx *.doc *.pdf"),
                           ("Word文件", "*.docx *.doc"),
                           ("PDF文件", "*.pdf"),
                           ("所有文件", "*.*")]
            )
        else:
            file_selected = filedialog.askopenfilename(
                filetypes=[("所有文件", "*.*")]
            )

        if file_selected:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, file_selected)
            self.status_var.set(f"已选择文件: {os.path.basename(file_selected)}")

    def get_file_time(self, file_path, time_type):
        """获取文件的创建时间或修改时间"""
        if time_type == "create":
            return os.path.getctime(file_path)
        elif time_type == "modify":
            return os.path.getmtime(file_path)
        else:  # current
            return datetime.now().timestamp()

    def format_filename(self, filename, index=1, total=1, folder_path=None):
        """根据格式模板生成新文件名"""
        format_str = self.format_var.get()
        time_format = self.time_format_var.get()
        time_source = self.time_source_var.get()

        # 获取文件扩展名
        name_part, ext_part = os.path.splitext(filename)

        # 生成序号（带前导零）
        num_digits = len(str(total)) if total > 1 else 1
        serial_number = f"{index:0{num_digits}d}"

        # 生成时间字符串
        if folder_path and "{时间}" in format_str:
            file_time = self.get_file_time(os.path.join(folder_path, filename), time_source)
            time_str = datetime.fromtimestamp(file_time).strftime(time_format)
        else:
            time_str = datetime.now().strftime(time_format)

        # 替换变量
        new_filename = format_str.replace("{序号}", serial_number)
        new_filename = new_filename.replace("{原文件名}", name_part)
        new_filename = new_filename.replace("{时间}", time_str)
        new_filename = new_filename.replace("{扩展名}", ext_part[1:])  # 去掉点号

        # 确保文件名不包含非法字符
        invalid_chars = r'[\\/:*?"<>|]'
        new_filename = re.sub(invalid_chars, '_', new_filename)

        return new_filename + ext_part

    def preview_rename(self):
        if self.current_mode == "title":
            self.extract_title()
            return

        path = self.path_entry.get().strip()

        if not path:
            messagebox.showerror("错误", "请选择文件或文件夹")
            return

        try:
            if self.is_single_file:
                # 单个文件处理
                if not os.path.exists(path) or not os.path.isfile(path):
                    messagebox.showerror("错误", "所选路径不存在或不是文件")
                    return

                filename = os.path.basename(path)
                folder_path = os.path.dirname(path)

                # 生成新文件名
                new_filename = self.format_filename(filename, folder_path=folder_path)

                # 显示预览
                self.preview_results = [(filename, new_filename)]
                original_text = f"1. {filename}\n"
                preview_text = f"1. 将 '{filename}' → '{new_filename}'\n"

                self.original_files_text.delete(1.0, tk.END)
                self.original_files_text.insert(tk.END, original_text)
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.insert(tk.END, preview_text)

                self.status_var.set(f"预览完成，1个文件将被处理")

            else:
                # 批量处理
                if not os.path.exists(path) or not os.path.isdir(path):
                    messagebox.showerror("错误", "所选路径不存在或不是文件夹")
                    return

                folder_path = path
                files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

                # 年份前缀的正则表达式模式
                year_pattern = re.compile(r'^(\d{4})年')

                # 数字前缀的正则表达式模式（匹配01、02、等格式）
                number_pattern = re.compile(r'^(\d+)(、)\s*')

                # 移除原数字前缀的函数（用于排序）
                def remove_number_prefix(filename):
                    match = number_pattern.search(filename)
                    if match:
                        return filename[match.end():]
                    return filename

                # 按文件名称主体排序（忽略原数字前缀）
                sorted_files = sorted(files, key=lambda x: remove_number_prefix(x))

                # 过滤掉年份前缀文件（如果选择保留）
                regular_files = []
                year_files = []

                for filename in sorted_files:
                    if self.keep_year_var.get() and year_pattern.match(filename):
                        year_files.append(filename)
                    else:
                        regular_files.append(filename)

                # 显示原始文件列表
                original_text = ""
                for i, filename in enumerate(sorted_files, 1):
                    original_text += f"{i:02d}. {filename}\n"

                self.original_files_text.delete(1.0, tk.END)
                self.original_files_text.insert(tk.END, original_text)

                # 生成预览结果
                self.preview_results = []
                preview_text = ""

                total_files = len(regular_files)

                for i, filename in enumerate(regular_files, 1):
                    # 提取文件名主体
                    match = number_pattern.search(filename)
                    if match and self.overwrite_existing_var.get():
                        rest_of_filename = filename[match.end():]
                    else:
                        rest_of_filename = filename

                    # 使用自定义格式生成新文件名
                    new_filename = self.format_filename(rest_of_filename, i, total_files, folder_path)

                    if filename != new_filename:
                        preview_text += f"{i:02d}. 将 '{filename}' → '{new_filename}'\n"
                    else:
                        preview_text += f"{i:02d}. 保持 '{filename}' 不变\n"

                    self.preview_results.append((filename, new_filename))

                if self.keep_year_var.get() and year_files:
                    preview_text += "\n以下文件因包含年份前缀被保留:\n"
                    for j, filename in enumerate(year_files, i + 1):
                        preview_text += f"{j:02d}. {filename}\n"

                # 更新预览文本
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.insert(tk.END, preview_text)

                self.status_var.set(f"预览完成，共 {len(regular_files)} 个文件将被处理")

            # 启用执行按钮
            self.execute_btn.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("错误", f"预览时出错: {str(e)}")
            self.status_var.set("预览失败")

    def extract_title_from_docx(self, file_path):
        """从Word文档中提取标题"""
        try:
            doc = docx.Document(file_path)
            # 尝试从标题样式中提取
            for paragraph in doc.paragraphs:
                if paragraph.style.name.startswith('Heading'):
                    return paragraph.text.strip()

            # 如果没有标题样式，尝试使用第一段
            if doc.paragraphs:
                return doc.paragraphs[0].text.strip()[:30]  # 限制长度
            return "无标题"
        except Exception as e:
            return f"提取失败: {str(e)}"

    def extract_title_from_pdf(self, file_path):
        """从PDF文件中提取标题"""
        try:
            pdf = PdfReader(file_path)
            info = pdf.metadata
            if info and '/Title' in info:
                title = info['/Title']
                if isinstance(title, bytes):
                    title = title.decode('utf-8', errors='replace')
                return title.strip()

            # 如果元数据中没有标题，尝试从内容提取（简单实现）
            if len(pdf.pages) > 0:
                first_page = pdf.pages[0].extract_text()
                # 提取前100个字符中的第一行作为标题
                lines = first_page.split('\n')
                for line in lines:
                    line = line.strip()
                    if line:
                        return line[:30]  # 限制长度
            return "无标题"
        except Exception as e:
            return f"提取失败: {str(e)}"

    def extract_title(self):
        file_path = self.path_entry.get().strip()
        if not file_path:
            messagebox.showerror("错误", "请选择文件")
            return

        if not os.path.exists(file_path):
            messagebox.showerror("错误", "文件不存在")
            return

        file_ext = os.path.splitext(file_path)[1].lower()
        filename = os.path.basename(file_path)

        self.original_files_text.delete(1.0, tk.END)
        self.original_files_text.insert(tk.END, filename)

        self.preview_text.delete(1.0, tk.END)

        try:
            self.status_var.set(f"正在提取标题: {filename}")

            if file_ext in ['.docx', '.doc']:
                title = self.extract_title_from_docx(file_path)
            elif file_ext == '.pdf':
                title = self.extract_title_from_pdf(file_path)
            else:
                title = f"不支持的文件格式: {file_ext}"

            # 处理提取的标题
            if title.startswith("提取失败"):
                self.preview_text.insert(tk.END, f"错误: {title}\n", "error")
                self.status_var.set("提取失败")
                self.execute_btn.config(state=tk.DISABLED)
                return

            # 规范化标题（去除非法字符）
            valid_title = re.sub(r'[\\/:*?"<>|]', '_', title)
            if not valid_title:
                valid_title = "无标题_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S")

            # 生成新文件名
            prefix = self.title_prefix_var.get().strip()
            suffix = self.title_suffix_var.get().strip()
            new_filename = f"{prefix}{valid_title}{suffix}{file_ext}"

            self.preview_text.insert(tk.END, f"原始文件名: {filename}\n\n", "original")
            self.preview_text.insert(tk.END, f"提取的标题: {title}\n\n", "title")
            self.preview_text.insert(tk.END, f"新文件名: {new_filename}\n", "new")

            self.status_var.set(f"标题提取完成: {title}")
            self.preview_results = [(filename, new_filename)]
            self.file_path = file_path

            # 启用重命名按钮
            self.execute_btn.config(state=tk.NORMAL)

            # 如果勾选了自动重命名，直接执行
            if self.auto_rename_title_var.get():
                self.execute_rename()

        except Exception as e:
            self.preview_text.insert(tk.END, f"错误: {str(e)}\n", "error")
            self.status_var.set(f"提取出错: {str(e)}")
            traceback.print_exc()
            self.execute_btn.config(state=tk.DISABLED)

    def execute_rename(self):
        if not self.preview_results:
            messagebox.showinfo("提示", "没有可执行的重命名操作")
            return

        path = self.path_entry.get().strip()

        try:
            if self.current_mode == "title":
                # 按标题重命名
                old_name, new_name = self.preview_results[0]
                old_path = self.file_path
                new_path = os.path.join(os.path.dirname(old_path), new_name)

                os.rename(old_path, new_path)
                self.preview_text.insert(tk.END, f"\n\n重命名成功: {old_name} → {new_name}\n", "success")

                self.status_var.set(f"重命名完成: {new_name}")
                messagebox.showinfo("成功", f"已成功重命名文件:\n{old_name} → {new_name}")

                # 更新路径显示
                self.path_entry.delete(0, tk.END)
                self.path_entry.insert(0, new_path)

            else:
                # 批量或单个重命名
                if self.is_single_file:
                    # 单个文件处理
                    old_name, new_name = self.preview_results[0]
                    old_path = os.path.join(os.path.dirname(path), old_name)
                    new_path = os.path.join(os.path.dirname(path), new_name)

                    os.rename(old_path, new_path)
                    messagebox.showinfo("成功", f"已成功重命名文件:\n{old_name} → {new_name}")

                    # 更新路径显示
                    self.path_entry.delete(0, tk.END)
                    self.path_entry.insert(0, new_path)

                else:
                    # 批量处理
                    folder_path = path
                    processed_count = 0

                    for old_name, new_name in self.preview_results:
                        if old_name != new_name:
                            old_path = os.path.join(folder_path, old_name)
                            new_path = os.path.join(folder_path, new_name)
                            os.rename(old_path, new_path)
                            processed_count += 1

                    messagebox.showinfo("成功", f"已成功处理 {processed_count} 个文件")

            self.status_var.set("操作完成")
            self.execute_btn.config(state=tk.DISABLED)

            # 刷新预览结果
            self.preview_rename()

        except Exception as e:
            messagebox.showerror("错误", f"执行时出错: {str(e)}")
            self.status_var.set("操作失败")


if __name__ == "__main__":
    root = tk.Tk()
    app = FileRenameApp(root)

    # 配置文本标签样式
    app.preview_text.tag_configure("original", foreground="blue")
    app.preview_text.tag_configure("title", foreground="green", font=("Microsoft YaHei UI", 11, "bold"))
    app.preview_text.tag_configure("new", foreground="purple", font=("Microsoft YaHei UI", 10, "underline"))
    app.preview_text.tag_configure("error", foreground="red")
    app.preview_text.tag_configure("success", foreground="dark green", font=("Microsoft YaHei UI", 10, "bold"))

    root.mainloop()