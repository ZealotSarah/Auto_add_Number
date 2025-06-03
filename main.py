import os
import re
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
from pathlib import Path
from datetime import datetime


class FileRenameApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件批量添加数字序号工具")
        self.root.geometry("900x700")  # 增加窗口高度

        # 设置字体确保中文显示正常
        self.font = ('Microsoft YaHei UI', 10)

        # 存储预览结果
        self.preview_results = []

        # 创建界面元素 - 确保在所有方法定义之后调用
        self.create_widgets()

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder_selected)

    def get_file_time(self, file_path, time_type):
        """获取文件的创建时间或修改时间"""
        if time_type == "create":
            return os.path.getctime(file_path)
        elif time_type == "modify":
            return os.path.getmtime(file_path)
        else:  # current
            return datetime.now().timestamp()

    def format_filename(self, filename, index, total, folder_path=None):
        """根据格式模板生成新文件名"""
        format_str = self.format_var.get()
        time_format = self.time_format_var.get()
        time_source = self.time_source_var.get()

        # 获取文件扩展名
        name_part, ext_part = os.path.splitext(filename)

        # 生成序号（带前导零）
        num_digits = len(str(total))
        # 修正缩进和变量名
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
        folder_path = self.folder_entry.get().strip()

        if not folder_path:
            messagebox.showerror("错误", "请选择文件夹")
            return

        if not os.path.exists(folder_path) or not os.path.isdir(folder_path):
            messagebox.showerror("错误", "所选路径不存在或不是文件夹")
            return

        try:
            # 获取文件夹中的所有文件
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
                    # 保留文件名主体部分（去除前缀后的内容）
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

            # 启用执行按钮
            self.execute_btn.config(state=tk.NORMAL)
            self.status_var.set(f"预览完成，共 {len(regular_files)} 个文件将被处理")

        except Exception as e:
            messagebox.showerror("错误", f"预览时出错: {str(e)}")
            self.status_var.set("预览失败")

    def execute_rename(self):
        if not self.preview_results:
            messagebox.showinfo("提示", "没有可执行的重命名操作")
            return

        folder_path = self.folder_entry.get().strip()

        # 确认对话框
        confirm = messagebox.askyesno("确认", f"确定要处理 {len(self.preview_results)} 个文件吗？")
        if not confirm:
            return

        try:
            # 执行重命名
            for old_name, new_name in self.preview_results:
                if old_name != new_name:  # 只处理需要重命名的文件
                    old_path = os.path.join(folder_path, old_name)
                    new_path = os.path.join(folder_path, new_name)
                    os.rename(old_path, new_path)

            messagebox.showinfo("成功", f"已成功处理 {len(self.preview_results)} 个文件")
            self.status_var.set("操作完成")
            self.execute_btn.config(state=tk.DISABLED)

            # 刷新预览结果 - 显示处理后的文件列表
            self.preview_rename()

        except Exception as e:
            messagebox.showerror("错误", f"执行时出错: {str(e)}")
            self.status_var.set("操作失败")

    def create_widgets(self):
        # 选择文件夹框架
        folder_frame = tk.Frame(self.root, padx=10, pady=10)
        folder_frame.pack(fill=tk.X)

        tk.Label(folder_frame, text="文件夹路径:", font=self.font).pack(side=tk.LEFT)

        self.folder_entry = tk.Entry(folder_frame, width=60, font=self.font)
        self.folder_entry.pack(side=tk.LEFT, padx=5)

        browse_btn = tk.Button(folder_frame, text="浏览...", font=self.font, command=self.browse_folder)
        browse_btn.pack(side=tk.LEFT)

        # 选项框架 - 第一行
        options_frame1 = tk.Frame(self.root, padx=10, pady=5)
        options_frame1.pack(fill=tk.X)

        self.keep_year_var = tk.BooleanVar(value=True)
        keep_year_check = tk.Checkbutton(options_frame1, text="保留年份前缀",
                                         variable=self.keep_year_var, font=self.font)
        keep_year_check.pack(side=tk.LEFT)

        self.overwrite_existing_var = tk.BooleanVar(value=True)
        overwrite_check = tk.Checkbutton(options_frame1, text="覆盖已有序号",
                                         variable=self.overwrite_existing_var, font=self.font)
        overwrite_check.pack(side=tk.LEFT, padx=10)

        # 选项框架 - 第二行 (新增命名格式选项)
        options_frame2 = tk.Frame(self.root, padx=10, pady=5)
        options_frame2.pack(fill=tk.X)

        tk.Label(options_frame2, text="命名格式:", font=self.font).pack(side=tk.LEFT)

        # 预设格式下拉菜单 - 添加新的格式选项
        self.format_var = tk.StringVar(value="{序号}、{原文件名}")
        format_options = [
            "{序号}、{原文件名}",
            "{序号}_{原文件名}",
            "{时间}_{原文件名}",
            "{序号}_{时间}_{原文件名}",
            "{原文件名}_{序号}",
            "{原文件名}_{时间}",
            "{序号}_{原文件名}_{时间}"  # 新增的命名格式
        ]
        format_combo = ttk.Combobox(options_frame2, textvariable=self.format_var,
                                    values=format_options, width=25, font=self.font)
        format_combo.pack(side=tk.LEFT, padx=5)

        # 时间格式选项
        tk.Label(options_frame2, text="时间格式:", font=self.font).pack(side=tk.LEFT, padx=(10, 5))

        self.time_format_var = tk.StringVar(value="%Y%m%d")
        time_options = [
            "%Y%m%d",  # 20230101
            "%Y-%m-%d",  # 2023-01-01
            "%Y%m%d%H%M",  # 202301011200
            "%Y-%m-%d_%H%M",  # 2023-01-01_1200
            "%Y%m%d_%H%M%S",  # 20230101_120000
        ]
        time_combo = ttk.Combobox(options_frame2, textvariable=self.time_format_var,
                                  values=time_options, width=15, font=self.font)
        time_combo.pack(side=tk.LEFT, padx=5)

        # 时间源选项
        self.time_source_var = tk.StringVar(value="current")
        current_radio = tk.Radiobutton(options_frame2, text="当前时间",
                                       variable=self.time_source_var, value="current", font=self.font)
        current_radio.pack(side=tk.LEFT, padx=5)

        modify_radio = tk.Radiobutton(options_frame2, text="修改时间",
                                      variable=self.time_source_var, value="modify", font=self.font)
        modify_radio.pack(side=tk.LEFT, padx=5)

        create_radio = tk.Radiobutton(options_frame2, text="创建时间",
                                      variable=self.time_source_var, value="create", font=self.font)
        create_radio.pack(side=tk.LEFT, padx=5)

        # 预览结果框架
        preview_frame = tk.Frame(self.root, padx=10, pady=10)
        preview_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧 - 原始文件列表
        left_frame = tk.Frame(preview_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        tk.Label(left_frame, text="原始文件列表:", font=self.font, fg="blue").pack(anchor=tk.W)

        self.original_files_text = scrolledtext.ScrolledText(left_frame, width=40, height=15, font=self.font)
        self.original_files_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # 右侧 - 重命名预览
        right_frame = tk.Frame(preview_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        tk.Label(right_frame, text="重命名预览:", font=self.font, fg="green").pack(anchor=tk.W)

        self.preview_text = scrolledtext.ScrolledText(right_frame, width=40, height=15, font=self.font)
        self.preview_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # 按钮框架
        btn_frame = tk.Frame(self.root, padx=10, pady=10)
        btn_frame.pack(fill=tk.X)

        preview_btn = tk.Button(btn_frame, text="预览重命名", font=self.font, command=self.preview_rename)
        preview_btn.pack(side=tk.LEFT, padx=5)

        execute_btn = tk.Button(btn_frame, text="执行重命名", font=self.font,
                                command=self.execute_rename, state=tk.DISABLED)
        execute_btn.pack(side=tk.LEFT, padx=5)

        self.execute_btn = execute_btn

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


if __name__ == "__main__":
    root = tk.Tk()
    app = FileRenameApp(root)
    root.mainloop()