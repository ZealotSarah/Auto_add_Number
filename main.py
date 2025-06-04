import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
from pathlib import Path
from datetime import datetime
import docx
from PyPDF2 import PdfReader
import traceback


def resource_path(relative_path):
    """获取打包后资源的绝对路径"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class FileRenameApp:
    def __init__(self, root):
        self.root = root
        self.root.title("多功能文件重命名工具")
        self.root.geometry("900x700")
        self.font = ('Microsoft YaHei UI', 10)

        # 存储预览结果
        self.preview_results = []
        self.is_single_file = False
        self.current_mode = "title"

        # 创建界面元素
        self.create_widgets()
        print("[INFO] 界面初始化完成")  # 添加日志

    def create_widgets(self):
        # 模式选择框架
        mode_frame = tk.Frame(self.root, padx=10, pady=5)
        mode_frame.pack(fill=tk.X)

        self.mode_var = tk.StringVar(value="title")
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

        self.mode_label = tk.Label(path_frame, text="按标题重命名模式", font=self.font, fg="blue")
        self.mode_label.pack(side=tk.LEFT, padx=10)

        # 标题重命名选项框架
        self.title_options_frame = tk.Frame(self.root, padx=10, pady=5)
        self.title_options_frame.pack(fill=tk.X)

        self.auto_rename_title_var = tk.BooleanVar(value=False)  # 默认不自动重命名
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

        # 左侧 - 原始文件
        left_frame = tk.Frame(preview_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.original_label = tk.Label(left_frame, text="原始文件名:", font=self.font, fg="blue")
        self.original_label.pack(anchor=tk.W)

        self.original_files_text = scrolledtext.ScrolledText(left_frame, width=40, height=5, font=self.font)
        self.original_files_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # 右侧 - 提取结果
        right_frame = tk.Frame(preview_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        self.preview_label = tk.Label(right_frame, text="提取结果:", font=self.font, fg="green")
        self.preview_label.pack(anchor=tk.W)

        self.preview_text = scrolledtext.ScrolledText(right_frame, width=40, height=5, font=self.font)
        self.preview_text.pack(fill=tk.BOTH, expand=True, pady=5)

        # 按钮框架
        btn_frame = tk.Frame(self.root, padx=10, pady=10)
        btn_frame.pack(fill=tk.X)

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

        # 初始化界面
        self.change_mode()
        print("[INFO] 界面组件加载完成")  # 添加日志

    def change_mode(self):
        self.current_mode = self.mode_var.get()
        print(f"[INFO] 切换到{self.current_mode}模式")  # 添加日志

        if self.current_mode == "batch":
            self.is_single_file = False
            self.mode_label.config(text="批量处理模式")
            self.folder_btn.config(text="浏览文件夹", command=self.browse_folder)
            self.title_options_frame.pack_forget()
            self.extract_title_btn.pack_forget()
            self.original_label.config(text="原始文件列表:")
            self.preview_label.config(text="重命名预览:")

        elif self.current_mode == "single":
            self.is_single_file = True
            self.mode_label.config(text="单个文件模式")
            self.folder_btn.config(text="浏览文件", command=self.browse_file)
            self.title_options_frame.pack_forget()
            self.extract_title_btn.pack_forget()
            self.original_label.config(text="原始文件:")
            self.preview_label.config(text="重命名预览:")

        elif self.current_mode == "title":
            self.is_single_file = True
            self.mode_label.config(text="按标题重命名模式")
            self.folder_btn.config(text="浏览文件", command=self.browse_file)
            self.title_options_frame.pack(fill=tk.X)
            self.extract_title_btn.pack(side=tk.LEFT, padx=5)
            self.original_label.config(text="原始文件名:")
            self.preview_label.config(text="提取结果:")

        self.path_entry.delete(0, tk.END)
        self.original_files_text.delete(1.0, tk.END)
        self.preview_text.delete(1.0, tk.END)
        self.execute_btn.config(state=tk.DISABLED)
        self.status_var.set("就绪")
        print("[INFO] 模式切换完成")  # 添加日志

    def browse_folder(self):
        print("[INFO] 打开文件/文件夹选择对话框")  # 添加日志
        if self.current_mode == "batch":
            folder_selected = filedialog.askdirectory()
            if folder_selected:
                self.path_entry.delete(0, tk.END)
                self.path_entry.insert(0, folder_selected)
                self.status_var.set(f"已选择文件夹: {os.path.basename(folder_selected)}")
                print(f"[INFO] 选择文件夹: {folder_selected}")  # 添加日志
        else:  # single 或 title 模式
            file_selected = filedialog.askopenfilename(
                filetypes=[("所有文件", "*.*")]
            )
            if file_selected:
                self.path_entry.delete(0, tk.END)
                self.path_entry.insert(0, file_selected)
                self.status_var.set(f"已选择文件: {os.path.basename(file_selected)}")
                print(f"[INFO] 选择文件: {file_selected}")  # 添加日志

    def browse_file(self):
        print("[INFO] 打开文件选择对话框")  # 添加日志
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
            print(f"[INFO] 选择文件: {file_selected}")  # 添加日志

    def extract_title_from_docx(self, file_path):
        """针对特定文档优化的Word标题提取方法"""
        try:
            print(f"[INFO] 开始从Word文档提取标题: {file_path}")
            doc = docx.Document(file_path)

            # 方法1：优先从文档属性提取（适用于标准文档）
            if doc.core_properties.title:
                title = doc.core_properties.title.strip()
                print(f"[INFO] 从文档属性提取标题: {title}")
                return title

            # 方法2：针对"中华人民共和国劳动合同法"文档的强化提取
            # 特征：标题通常在首行且包含"中华人民共和国"关键词
            for para in doc.paragraphs[:10]:  # 检查前10个段落
                text = para.text.strip()

                # 检查是否包含法律名称特征词
                if ("中华人民共和国" in text or "法" in text) and len(text) < 60:
                    # 过滤目录、章节标题和序号
                    if not re.match(r'^(第[一二三四五六七八九十0-9]+章|目?录|附件|附录|修订说明|前言|序)', text):
                        print(f"[INFO] 从正文提取标题: {text}")
                        return text[:40]  # 限制标题长度

            # 方法3：常规标题提取逻辑（兜底）
            title_candidates = []

            # 尝试提取Heading样式的标题
            for para in doc.paragraphs:
                if para.style.name.startswith('Heading') or para.style.name.lower().startswith('标题'):
                    candidate = para.text.strip()
                    if len(candidate) > 5 and not re.match(r'^(第[0-9一二三四五六七八九十]+)', candidate):
                        title_candidates.append(candidate)
                        print(f"[INFO] 从标题样式提取: {candidate}")
                        break

            # 尝试提取正文前3段中的长文本
            if not title_candidates:
                for para in doc.paragraphs[:3]:
                    text = para.text.strip()
                    if len(text) > 15 and not re.match(r'^(第[0-9一二三四五六七八九十]+|目?录|附件|附录)', text):
                        title_candidates.append(text[:40])
                        print(f"[INFO] 从正文前3段提取: {text[:40]}")
                        break

            # 最后尝试提取文档第一行
            if not title_candidates and doc.paragraphs:
                first_line = doc.paragraphs[0].text.strip()
                if len(first_line) > 5:
                    title_candidates.append(first_line[:40])
                    print(f"[INFO] 从第一行提取: {first_line[:40]}")

            return title_candidates[0] if title_candidates else "无标题"

        except Exception as e:
            error_msg = f"提取失败: {str(e)}"
            print(f"[ERROR] Word标题提取失败: {error_msg}")
            return error_msg

    def extract_title_from_pdf(self, file_path):
        """从PDF文件中提取标题"""
        try:
            print(f"[INFO] 开始从PDF提取标题: {file_path}")
            pdf = PdfReader(file_path)
            info = pdf.metadata
            if info and '/Title' in info:
                title = info['/Title']
                if isinstance(title, bytes):
                    title = title.decode('utf-8', errors='replace')
                print(f"[INFO] 从PDF元数据提取标题: {title}")
                return title.strip()

            # 若元数据无标题，提取正文首行
            if len(pdf.pages) > 0:
                first_page = pdf.pages[0].extract_text()
                if not first_page:
                    print("[INFO] PDF首页无文本内容")
                    return "无标题"

                lines = []
                for line in first_page.split('\n'):
                    stripped = line.strip()
                    if stripped and len(stripped) > 5:
                        lines.append(stripped)

                if lines:
                    # 尝试找到最长的行作为标题
                    longest_line = max(lines, key=len)
                    print(f"[INFO] 从PDF正文提取标题: {longest_line[:40]}")
                    return longest_line[:40]
                else:
                    print("[INFO] PDF首页无有效行")
                    return "无标题"

            print("[INFO] PDF无有效内容，返回无标题")
            return "无标题"
        except Exception as e:
            error_msg = f"提取失败: {str(e)}"
            print(f"[ERROR] PDF标题提取失败: {error_msg}")
            return error_msg

    def extract_title(self):
        file_path = self.path_entry.get().strip()
        if not file_path:
            messagebox.showerror("错误", "请选择文件")
            print("[ERROR] 未选择文件")
            return

        if not os.path.exists(file_path):
            messagebox.showerror("错误", "文件不存在")
            print(f"[ERROR] 文件不存在: {file_path}")
            return

        file_ext = os.path.splitext(file_path)[1].lower()
        filename = os.path.basename(file_path)

        self.original_files_text.delete(1.0, tk.END)
        self.original_files_text.insert(tk.END, filename)

        self.preview_text.delete(1.0, tk.END)

        try:
            self.status_var.set(f"正在提取标题: {filename}")
            print(f"[INFO] 开始提取标题流程: {filename}")

            if file_ext in ['.docx', '.doc']:
                title = self.extract_title_from_docx(file_path)
            elif file_ext == '.pdf':
                title = self.extract_title_from_pdf(file_path)
            else:
                title = f"不支持的文件格式: {file_ext}"
                print(f"[ERROR] 不支持的文件格式: {file_ext}")

            # 处理提取的标题
            if title.startswith("提取失败"):
                self.preview_text.insert(tk.END, f"错误: {title}\n", "error")
                self.status_var.set("提取失败")
                self.execute_btn.config(state=tk.DISABLED)
                print(f"[ERROR] 标题提取失败: {title}")
                return

            # 规范化标题
            valid_title = re.sub(r'[\\/:*?"<>|]', '_', title)
            if not valid_title:
                valid_title = "无标题_" + datetime.now().strftime("%Y%m%d%H%M%S")
                print("[INFO] 生成默认标题: 无标题_时间戳")

            # 生成新文件名
            prefix = self.title_prefix_var.get().strip()
            suffix = self.title_suffix_var.get().strip()
            ext = os.path.splitext(filename)[1] if not suffix else ""
            new_filename = f"{prefix}{valid_title}{suffix}{ext}"

            self.preview_text.insert(tk.END, f"原始文件名: {filename}\n\n", "original")
            self.preview_text.insert(tk.END, f"提取的标题: {title}\n\n", "title")
            self.preview_text.insert(tk.END, f"新文件名: {new_filename}\n", "new")

            self.status_var.set(f"标题提取完成: {title}")
            self.preview_results = [(filename, new_filename)]
            self.file_path = file_path

            # 启用重命名按钮
            self.execute_btn.config(state=tk.NORMAL)
            print(f"[INFO] 标题提取成功，新文件名: {new_filename}")

            # 自动重命名（仅当勾选时执行）
            if self.auto_rename_title_var.get():
                print("[INFO] 自动重命名已启用，执行重命名")
                self.execute_rename()

        except Exception as e:
            self.preview_text.insert(tk.END, f"错误: {str(e)}\n", "error")
            self.status_var.set(f"提取出错: {str(e)}")
            traceback.print_exc()
            self.execute_btn.config(state=tk.DISABLED)
            print(f"[ERROR] 提取标题过程中出错: {str(e)}")

    def execute_rename(self):
        if not self.preview_results:
            messagebox.showinfo("提示", "没有可执行的重命名操作")
            print("[INFO] 没有可执行的重命名操作")
            return

        try:
            old_name, new_name = self.preview_results[0]
            old_path = self.file_path
            new_path = os.path.join(os.path.dirname(old_path), new_name)

            print(f"[INFO] 执行文件重命名: {old_name} → {new_name}")
            os.rename(old_path, new_path)
            self.preview_text.insert(tk.END, f"\n\n重命名成功: {old_name} → {new_name}\n", "success")

            self.status_var.set(f"重命名完成: {new_name}")
            messagebox.showinfo("成功", f"已成功重命名文件:\n{old_name} → {new_name}")

            # 更新路径显示
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, new_path)
            self.original_files_text.delete(1.0, tk.END)
            self.original_files_text.insert(tk.END, new_name)

            print(f"[INFO] 文件重命名成功: {old_path} → {new_path}")

        except Exception as e:
            messagebox.showerror("错误", f"执行时出错: {str(e)}")
            self.status_var.set("操作失败")
            print(f"[ERROR] 文件重命名失败: {str(e)}")


if __name__ == "__main__":
    print("[INFO] 程序启动中...")
    root = tk.Tk()
    app = FileRenameApp(root)
    print("[INFO] 程序初始化完成，进入主事件循环")
    root.mainloop()