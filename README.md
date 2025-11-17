项目简介
这是一款功能强大的文件重命名工具，支持批量重命名、按标题重命名和单文件重命名等多种模式。软件提供直观的预览功能，让用户在实际重命名前确认结果，避免误操作。特别适合需要规范化文件命名的办公、学习和资料整理场景。

主要功能
批量重命名：一次重命名多个文件，提高工作效率
智能标题提取：自动从Word和PDF文档中提取标题作为文件名
多种命名模式：
序号模式 (001, 002...)
日期模式 (YYYYMMDD)
时间戳模式 (YYYYMMDD_HHMMSS)
序号+原文件名/原文件名+序号
日期+原文件名
序号_原文件名_日期
文件后缀修改：批量修改文件扩展名
拖拽排序：通过拖拽调整文件处理顺序
打开文件检测：自动检测并提醒用户关闭正在使用的文件
隐藏文件过滤：自动忽略系统隐藏文件
格式撤销：一键撤销已应用的命名格式
系统要求
Windows 10/11 操作系统
.NET Framework 4.8 或更高版本
2GB 以上内存
100MB 可用磁盘空间
安装方法
方法一：直接使用可执行文件（推荐）
从 发布页面 下载最新版本的 文件批量添加序号工具.exe
双击运行即可，无需额外安装
方法二：从源代码安装
克隆或下载本仓库
安装 Python 3.8 或更高版本
创建虚拟环境并安装依赖：
bash


1
2
3
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
运行程序：
bash


1
python main.py
使用说明
选择操作模式：
批量重命名：处理整个文件夹中的文件
单个文件重命名：处理单个文件
按文件标题重命名：从Word或PDF文档中提取标题
选择文件或文件夹：
点击"浏览文件夹"或"浏览文件"按钮选择目标
设置命名规则：
从预设命名格式中选择合适的模式
设置序号起始值和位数（如需要）
设置日期格式（如需要）
添加前缀或后缀（可选）
预览和执行：
检查右侧预览窗口中的新文件名
点击"执行重命名"完成操作
如需撤销格式，点击"取消格式"按钮
打包成可执行文件
如果你需要从源代码打包成EXE文件，请按以下步骤操作：

安装所需依赖：
powershell


1
pip install pyinstaller python-docx PyPDF2 psutil pywin32
执行打包命令：
powershell


1
2
3
4
5
6
7
8
9
pyinstaller --onefile --noconsole --name "文件批量添加序号工具" `
  --add-data ".venv\Lib\site-packages\docx\templates;docx\templates" `
  --hidden-import psutil `
  --hidden-import docx `
  --hidden-import PyPDF2 `
  --hidden-import win32api `
  --hidden-import win32con `
  --collect-all psutil `
  main.py
生成的可执行文件位于 dist 文件夹中
问题排查
"找不到 psutil 模块"错误：确保在虚拟环境中正确安装了所有依赖
Word模板文件缺失：确认 docx/templates 路径正确添加到打包命令中
重命名失败：检查文件是否被其他程序打开，或是否有权限问题
版本历史
v1.0.0 (2025-04-17)
初始版本发布
支持批量重命名、单文件重命名和按标题重命名
添加文件打开检测和隐藏文件过滤功能
许可证
本项目采用 MIT 许可证 。



1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
Copyright (c) 2025 文件重命名工具开发团队

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
联系我们
如有问题或建议，请通过以下方式联系我们：


注意：使用本软件前请备份重要文件。开发者不对因使用本软件而导致的数据丢失负责。