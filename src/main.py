# Copyright 2023 agenius666
# GitHub: https://github.com/agenius666/Excel-Link-Refresher
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import os
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import time

# 全局变量，用于控制线程是否停止
stop_event = threading.Event()

def process_excel_files(folder_path, skip_file_paths, disable_update_links, log_widget, progress_bar, time_label, root):
    # 启动 Excel 应用程序
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False  # 设置为 False 以隐藏 Excel 窗口
    excel.DisplayAlerts = False  # 禁用警告提示

    # 用于记录处理失败的文件路径
    failed_files = []

    try:
        # 获取所有 Excel 文件的总数
        total_files = sum(
            len([f for f in files if f.endswith('.xlsx') or f.endswith('.xls')])
            for _, _, files in os.walk(folder_path)
        )
        processed_files = 0  # 已处理的文件数

        # 记录开始时间
        start_time = time.time()

        # 遍历指定文件夹及其子文件夹下的所有文件
        for root_dir, dirs, files in os.walk(folder_path):
            for file in files:
                # 检查是否收到停止信号
                if stop_event.is_set():
                    log_widget.insert(tk.END, "处理已中止！\n")
                    log_widget.see(tk.END)  # 自动滚动到最新日志
                    return

                # 检查文件是否是 Excel 文件
                if file.endswith('.xlsx') or file.endswith('.xls'):
                    file_path = os.path.join(root_dir, file)

                    # 如果当前文件在用户指定跳过的文件列表中，则跳过
                    if os.path.abspath(file_path) in skip_file_paths:
                        log_widget.insert(tk.END, f"跳过文件: {file_path}\n")
                        log_widget.see(tk.END)  # 自动滚动到最新日志
                        continue

                    log_widget.insert(tk.END, f"正在处理文件: {file_path}\n")
                    log_widget.see(tk.END)  # 自动滚动到最新日志

                    try:
                        # 根据用户选择设置 UpdateLinks 参数
                        if disable_update_links:
                            workbook = excel.Workbooks.Open(file_path, UpdateLinks=0)
                            log_widget.insert(tk.END, "已关闭更新链接的弹窗。\n")
                            log_widget.see(tk.END)  # 自动滚动到最新日志
                        else:
                            workbook = excel.Workbooks.Open(file_path)

                        # 保存并关闭工作簿
                        workbook.Save()
                        workbook.Close()
                        log_widget.insert(tk.END, f"已保存并关闭文件: {file_path}\n")
                        log_widget.see(tk.END)  # 自动滚动到最新日志
                    except Exception as e:
                        log_widget.insert(tk.END, f"处理文件 {file_path} 时出错: {e}\n")
                        log_widget.see(tk.END)  # 自动滚动到最新日志
                        failed_files.append(file_path)  # 记录失败的文件路径

                    # 更新已处理的文件数和进度条
                    processed_files += 1
                    progress = int((processed_files / total_files) * 100)
                    progress_bar['value'] = progress
                    root.update_idletasks()  # 更新 UI

                    # 更新已使用时间
                    elapsed_time = time.time() - start_time
                    time_label.config(text=f"已用时间: {int(elapsed_time)} 秒")
    finally:
        # 退出 Excel 应用程序
        excel.Quit()
        if not stop_event.is_set():
            log_widget.insert(tk.END, "处理完成！\n")
            log_widget.see(tk.END)  # 自动滚动到最新日志

            # 如果有失败的文件，打印提示信息
            if failed_files:
                log_widget.insert(tk.END, "\n以下文件处理失败：\n")
                for file_path in failed_files:
                    log_widget.insert(tk.END, f"- {file_path}\n")
                log_widget.insert(tk.END, "请检查这些文件并重新处理。\n")
                log_widget.see(tk.END)  # 自动滚动到最新日志

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_path_entry.delete(0, tk.END)
    folder_path_entry.insert(0, folder_path)

def browse_skip_files():
    skip_file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    skip_file_entry.delete(0, tk.END)
    skip_file_entry.insert(0, "; ".join(skip_file_paths))  # 将多个文件路径用分号分隔显示

def start_processing():
    global stop_event
    stop_event.clear()  # 重置停止标志

    folder_path = folder_path_entry.get()
    skip_file_paths = skip_file_entry.get().split("; ")  # 将分号分隔的路径拆分为列表
    skip_file_paths = [os.path.abspath(path.strip()) for path in skip_file_paths if path.strip()]  # 处理路径格式
    disable_update_links = disable_update_links_var.get()

    if not os.path.exists(folder_path):
        messagebox.showerror("错误", "指定的文件夹路径不存在，请检查路径是否正确。")
        return

    # 清空日志区域
    log_text.delete(1.0, tk.END)

    # 显示提示信息
    log_text.insert(tk.END, "作者：agenius666\n")
    log_text.insert(tk.END, "https://github.com/agenius666/Excel-Link-Refresher\n")
    log_text.insert(tk.END, "----------------------------------\n")
    log_text.see(tk.END)  # 自动滚动到最新日志

    # 重置进度条
    progress_bar['value'] = 0
    time_label.config(text="已用时间: 0 秒")

    # 使用线程处理文件，避免 UI 卡住
    processing_thread = threading.Thread(
        target=process_excel_files,
        args=(folder_path, skip_file_paths, disable_update_links, log_text, progress_bar, time_label, root)
    )
    processing_thread.start()

def stop_processing():
    global stop_event
    stop_event.set()  # 设置停止标志
    log_text.insert(tk.END, "正在中止处理...\n")
    log_text.see(tk.END)  # 自动滚动到最新日志

def clear_log():
    log_text.delete(1.0, tk.END)  # 清空日志区域
    log_text.insert(tk.END, "日志已清空！\n")
    log_text.see(tk.END)  # 自动滚动到最新日志

# 创建主窗口
root = tk.Tk()
root.title("Excel Link Refresher - 1.0.0")

# 创建并放置标签和输入框
tk.Label(root, text="文件夹路径:").grid(row=0, column=0, padx=5, pady=5)
folder_path_entry = tk.Entry(root, width=50)
folder_path_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="浏览", command=browse_folder).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="跳过文件路径:").grid(row=1, column=0, padx=5, pady=5)
skip_file_entry = tk.Entry(root, width=50)
skip_file_entry.grid(row=1, column=1, padx=5, pady=5)
tk.Button(root, text="浏览", command=browse_skip_files).grid(row=1, column=2, padx=5, pady=5)

# 创建并放置复选框
disable_update_links_var = tk.BooleanVar()
tk.Checkbutton(root, text="关闭弹窗", variable=disable_update_links_var).grid(row=2, column=1, padx=5, pady=5)

# 创建并放置开始按钮
tk.Button(root, text="开始处理", command=start_processing).grid(row=3, column=1, padx=5, pady=10)

# 创建并放置中止按钮
tk.Button(root, text="中止处理", command=stop_processing).grid(row=3, column=2, padx=5, pady=10)

# 创建并放置清空日志按钮
tk.Button(root, text="清空日志", command=clear_log).grid(row=3, column=0, padx=5, pady=10)

# 创建提示标签
tk.Label(root, text="遍历文件夹（包括子文件夹）下的所有 .xlsx 和 .xls 文件，自动打开并保存，以更新 Excel 链接。\n支持多选跳过文件，选中的文件将不被处理。\n关闭弹窗：关闭 Excel 软件自带的更新链接弹窗。", fg="blue").grid(row=4, column=0, columnspan=3, padx=5, pady=5)

# 创建进度条
progress_bar = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
progress_bar.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

# 创建已用时间标签
time_label = tk.Label(root, text="已用时间: 0 秒")
time_label.grid(row=6, column=0, columnspan=3, padx=5, pady=5)

# 创建日志区域
log_text = scrolledtext.ScrolledText(root, width=70, height=20, wrap=tk.WORD)
log_text.grid(row=7, column=0, columnspan=3, padx=5, pady=5)

# 运行主循环
root.mainloop()
