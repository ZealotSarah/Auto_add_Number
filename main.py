import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import docx
from PyPDF2 import PdfReader
import datetime
import traceback


class FileRenameByTitleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件标题重命名工具")
        self.root.geometry("600x400")
        self.font = ('Microsoft YaHei UI', 10)
        self.create_widgets()

    def create_widgets(self):
        # 路径选择框架
        path_frame = tk.Frame(self.root, padx=10, pady=10)
        path_frame.pack(fill=tk.X)

        tk.Label(path_frame, text="文件路径:", font=self.font).pack(side=tk.LEFT)
        self.path_entry = tk.Entry(path_frame, width=40, font=self.font)
        self.path_entry.pack(side=tk.LEFT, padx=5)

        browse_btn = tk.Button(path_frame, text="浏览...", font=self.font, command=self.browse_file)
        browse_btn.pack(side=tk.LEFT)

        # 选项框架
        options_frame = tk.Frame(self.root, padx=10, pady=5)
        options_frame.pack(fill=tk.X)

        self.auto_rename_var = tk.BooleanVar(value=False)
        auto_rename_check = tk.Checkbutton(options_frame, text="自动重命名",
                                           variable=self.auto_rename_var, font=self.font)
        auto_rename_check.pack(side=tk.LEFT, padx=5)

        self.add_prefix_var = tk.StringVar(value="")
        prefix_label = tk.Label(options_frame, text="前缀:", font=self.font)
        prefix_label.pack(side=tk.LEFT, padx=5)
        prefix_entry = tk.Entry(options_frame, textvariable=self.add_prefix_var, width=10, font=self.font)
        prefix_entry.pack(side=tk.LEFT, padx=5)

        self.add_suffix_var = tk.StringVar(value="")
        suffix_label = tk.Label(options_frame, text="后缀:", font=self.font)
        suffix_label.pack(side=tk.LEFT, padx=5)
        suffix_entry = tk.Entry(options_frame, textvariable=self.add_suffix_var, width=10, font=self.font)
        suffix_entry.pack(side=tk.LEFT, padx=5)

        # 结果显示框架
        result_frame = tk.Frame(self.root, padx=10, pady=10)
        result_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(result_frame, text="提取结果:", font=self.font).pack(anchor=tk.W)
        self.result_text = tk.Text(result_frame, height=10, width=70, font=self.font)
        self.result_text.pack(fill=tk.BOTH, expand=True, pady=5)
        self.result_text.config(state=tk.DISABLED)

        # 按钮框架
        btn_frame = tk.Frame(self.root, padx=10, pady=10)
        btn_frame.pack(fill=tk.X)

        extract_btn = tk.Button(btn_frame, text="提取标题", font=self.font,
                                command=self.extract_title, width=12)
        extract_btn.pack(side=tk.LEFT, padx=5)

        rename_btn = tk.Button(btn_frame, text="执行重命名", font=self.font,
                               command=self.rename_file, state=tk.DISABLED, width=12)
        rename_btn.pack(side=tk.LEFT, padx=5)

        self.rename_btn = rename_btn

        # 状态框架
        status_frame = tk.Frame(self.root, padx=10, pady=5)
        status_frame.pack(fill=tk.X)

        self.status_var = tk.StringVar(value="就绪")
        status_label = tk.Label(status_frame, textvariable=self.status_var, font=self.font, fg="blue")
        status_label.pack(anchor=tk.W)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Word/PDF文件", "*.docx *.doc *.pdf"),
                       ("Word文件", "*.docx *.doc"),
                       ("PDF文件", "*.pdf"),
                       ("所有文件", "*.*")]
        )
        if file_path:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, file_path)
            self.status_var.set(f"已选择文件: {os.path.basename(file_path)}")

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

        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)

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
                self.result_text.insert(tk.END, f"错误: {title}\n", "error")
                self.status_var.set("提取失败")
            else:
                # 规范化标题（去除非法字符）
                valid_title = re.sub(r'[\\/:*?"<>|]', '_', title)
                if not valid_title:
                    valid_title = "无标题_" + datetime.datetime.now().strftime("%Y%m%d%H%M%S")

                # 生成新文件名
                prefix = self.add_prefix_var.get().strip()
                suffix = self.add_suffix_var.get().strip()
                new_filename = f"{prefix}{valid_title}{suffix}{file_ext}"

                self.result_text.insert(tk.END, f"原始文件名: {filename}\n", "original")
                self.result_text.insert(tk.END, f"提取的标题: {title}\n", "title")
                self.result_text.insert(tk.END, f"新文件名: {new_filename}\n", "new")

                self.status_var.set(f"标题提取完成: {title}")
                self.preview_new_filename = new_filename
                self.original_filename = filename
                self.file_path = file_path

                # 启用重命名按钮
                self.rename_btn.config(state=tk.NORMAL)

                # 如果勾选了自动重命名，直接执行
                if self.auto_rename_var.get():
                    self.rename_file()

        except Exception as e:
            self.result_text.insert(tk.END, f"错误: {str(e)}\n", "error")
            self.status_var.set(f"提取出错: {str(e)}")
            traceback.print_exc()

        self.result_text.config(state=tk.DISABLED)

    def rename_file(self):
        if not hasattr(self, 'file_path'):
            messagebox.showinfo("提示", "请先提取标题")
            return

        try:
            old_path = self.file_path
            new_filename = self.preview_new_filename
            new_path = os.path.join(os.path.dirname(old_path), new_filename)

            os.rename(old_path, new_path)
            self.result_text.config(state=tk.NORMAL)
            self.result_text.insert(tk.END, f"重命名成功: {self.original_filename} → {new_filename}\n", "success")
            self.result_text.config(state=tk.DISABLED)

            self.status_var.set(f"重命名完成: {new_filename}")
            messagebox.showinfo("成功", f"已成功重命名文件:\n{self.original_filename} → {new_filename}")

            # 更新路径显示
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, new_path)

        except Exception as e:
            self.result_text.config(state=tk.NORMAL)
            self.result_text.insert(tk.END, f"重命名失败: {str(e)}\n", "error")
            self.result_text.config(state=tk.DISABLED)

            self.status_var.set(f"重命名失败: {str(e)}")
            messagebox.showerror("错误", f"重命名文件时出错:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = FileRenameByTitleApp(root)

    # 配置文本标签样式
    app.result_text.tag_configure("original", foreground="blue")
    app.result_text.tag_configure("title", foreground="green", font=("Microsoft YaHei UI", 11, "bold"))
    app.result_text.tag_configure("new", foreground="purple", font=("Microsoft YaHei UI", 10, "underline"))
    app.result_text.tag_configure("error", foreground="red")
    app.result_text.tag_configure("success", foreground="dark green", font=("Microsoft YaHei UI", 10, "bold"))

    root.mainloop()