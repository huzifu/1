import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import logging
import random
import re
import traceback
import time
import numpy as np
import requests
import json
import os
from typing import List, Dict, Any
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException, \
    ElementNotInteractableException
from ttkthemes import ThemedTk

# ================== 配置参数 ==================
# 默认参数值
DEFAULT_CONFIG = {
    "url": "https://www.wjx.cn/vm/mZ3nVoC.aspx# ",
    "target_num": 100,
    "min_duration": 15,
    "max_duration": 180,
    "weixin_ratio": 0.5,
    "min_delay": 2.0,
    "max_delay": 6.0,
    "submit_delay": 3,
    "page_load_delay": 2,
    "per_question_delay": (1.0, 3.0),
    "per_page_delay": (2.0, 6.0),

    "use_ip": False,
    "headless": False,
    "ip_api": "https://service.ipzan.com/core-extract?num=1&no=20250531340212140720&minute=1&pool=quality&secret=c8r9rcegsc8tlo",
    "num_threads": 4,
    "single_prob": {"1": -1, "2": -1, "3": -1, "4": -1, "5": -1, "6": -1, "7": -1},
    "droplist_prob": {"1": [2, 1, 1]},
    "multiple_prob": {"3": [10, 20, 40, 50, 50], "4": [10, 20, 40, 50, 50]},
    "min_selection": {"3": 1, "4": 2},
    "max_selection": {"3": 3, "4": 5},
    "matrix_prob": {"1": [1, 0, 0, 0, 0], "2": -1, "3": [1, 0, 0, 0, 0],
                    "4": [1, 0, 0, 0, 0], "5": [1, 0, 0, 0, 0], "6": [1, 0, 0, 0, 0]},
    "scale_prob": {"7": [0, 2, 3, 4, 1], "12": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1]},
    "texts": {"11": ["多媒体与网络技术的深度融合", "人工智能驱动的个性化学习", "虚拟现实（VR/AR）与元宇宙教育",
                     "跨学科研究与学习科学深化"],
              "12": ["巧用AI工具解放重复劳动", "动态反馈激活课堂参与", "轻量化VR/AR增强认知具象化"]},
    "texts_prob": {"11": [1, 1, 1, 1], "12": [1, 1, 1]},
    "multiple_other": {
        "3": ["其他原因1", "其他原因2", "其他原因3"],
        "4": ["补充说明1", "补充说明2"]
    },
    "multiple_texts": {
        "25": [
            ["教师培训", "教育研讨会", "在线课程"],
            ["教学研究", "课堂观察", "同行评议"],
            ["自我反思", "教学日志", "专业发展计划"]
        ]
    },
    "multiple_texts_prob": {
        "25": [
            [0.4, 0.3, 0.3],
            [0.5, 0.3, 0.2],
            [0.6, 0.2, 0.2]
        ]
    },
    "reorder_prob": {
        "35": [0.4, 0.3, 0.2, 0.1],
        "36": [0.3, 0.4, 0.2, 0.1]
    }
}

class WJXAutoFillApp:
    def __init__(self, root):
        self.root = root
        self.root.title("问卷星自动填写工具")
        self.root.geometry("1200x900")
        self.root.resizable(True, True)

        # 设置现代风格
        style = ttk.Style()
        style.theme_use('default')
        style.configure('TNotebook.Tab', padding=[10, 5])
        style.configure('TButton', padding=[10, 5])
        style.configure('TLabel', padding=[5, 2])
        style.configure('TEntry', padding=[5, 2])

        self.config = DEFAULT_CONFIG.copy()
        self.running = False
        self.paused = False
        self.cur_num = 0
        self.cur_fail = 0
        self.lock = threading.Lock()
        self.pause_event = threading.Event()

        # 初始化输入框列表
        self.single_entries = []
        self.multi_entries = []
        self.droplist_entries = []
        self.scale_entries = []
        self.multi_other_entries = []
        self.min_selection_entries = []
        self.matrix_entries = []
        self.text_entries = []
        self.multiple_text_entries = []
        self.reorder_entries = []

        # 字体设置
        self.font_family = tk.StringVar()
        self.font_size = tk.IntVar()
        self.font_family.set("Arial")
        self.font_size.set(10)

        # 创建主框架
        self.main_paned = ttk.PanedWindow(root, orient=tk.VERTICAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 上半部分：控制区域和标签页
        self.top_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(self.top_frame, weight=1)

        # 下半部分：日志区域
        self.log_frame = ttk.LabelFrame(self.main_paned, text="运行日志")
        self.main_paned.add(self.log_frame, weight=0)

        # === 添加控制按钮区域（顶部）===
        self.control_frame = ttk.Frame(self.top_frame)
        self.control_frame.pack(fill=tk.X, padx=5, pady=(0, 10))

        # 第一行按钮
        top_btn_frame = ttk.Frame(self.control_frame)
        top_btn_frame.pack(fill=tk.X, pady=(0, 5))

        self.start_btn = ttk.Button(top_btn_frame, text="开始填写", command=self.start_filling)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.pause_btn = ttk.Button(top_btn_frame, text="暂停", command=self.toggle_pause, state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(top_btn_frame, text="停止", command=self.stop_filling, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        self.export_config_btn = ttk.Button(top_btn_frame, text="导出配置", command=self.export_config)
        self.export_config_btn.pack(side=tk.LEFT, padx=5)

        self.import_config_btn = ttk.Button(top_btn_frame, text="导入配置", command=self.import_config)
        self.import_config_btn.pack(side=tk.LEFT, padx=5)

        # 状态栏
        status_frame = ttk.Frame(self.control_frame)
        status_frame.pack(fill=tk.X, pady=(5, 0))

        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var)
        self.status_label.pack(side=tk.RIGHT, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)

        # 题目进度
        self.question_progress_var = tk.DoubleVar()
        self.question_progress_bar = ttk.Progressbar(status_frame,
                                                   variable=self.question_progress_var,
                                                   maximum=100,
                                                   length=200)
        self.question_progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)

        self.question_status_var = tk.StringVar(value="题目进度: 0/0")
        self.question_status_label = ttk.Label(status_frame, textvariable=self.question_status_var)
        self.question_status_label.pack(side=tk.RIGHT, padx=5)

        # 创建标签页
        self.notebook = ttk.Notebook(self.top_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 全局设置标签页
        self.global_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.global_frame, text="全局设置")
        self.create_global_settings()

        # 题型设置标签页 - 添加滚动条
        self.question_container = ttk.Frame(self.notebook)
        self.notebook.add(self.question_container, text="题型设置")

        # 创建带滚动条的题型设置框架
        self.canvas = tk.Canvas(self.question_container, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self.question_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # 创建题型设置内容
        self.question_frame = self.scrollable_frame
        self.create_question_settings()

        # 创建日志区域
        self.create_log_area()

        # 设置日志系统
        self.setup_logging()

        # 绑定字体更新事件
        self.font_family.trace_add("write", self.update_font)
        self.font_size.trace_add("write", self.update_font)

        # 初始化字体
        self.update_font()

    def create_log_area(self):
        """创建日志区域"""
        # 日志控制按钮
        log_control_frame = ttk.Frame(self.log_frame)
        log_control_frame.pack(fill=tk.X, padx=5, pady=(5, 0))

        self.clear_log_btn = ttk.Button(log_control_frame, text="清空日志", command=self.clear_log)
        self.clear_log_btn.pack(side=tk.LEFT, padx=5)

        self.export_log_btn = ttk.Button(log_control_frame, text="导出日志", command=self.export_log)
        self.export_log_btn.pack(side=tk.LEFT, padx=5)

        # 日志文本区域
        self.log_area = scrolledtext.ScrolledText(self.log_frame, height=10)
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_area.config(state=tk.DISABLED)

    def setup_logging(self):
        """配置日志系统"""

        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget

            def emit(self, record):
                msg = self.format(record)

                def append():
                    self.text_widget.configure(state='normal')
                    self.text_widget.insert(tk.END, msg + '\n')
                    self.text_widget.configure(state='disabled')
                    self.text_widget.see(tk.END)

                self.text_widget.after(0, append)

        handler = TextHandler(self.log_area)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s',
                                      datefmt='%Y-%m-%d %H:%M:%S')
        handler.setFormatter(formatter)

        logger = logging.getLogger()
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)

    def create_global_settings(self):
        """创建全局设置界面"""
        frame = self.global_frame
        padx, pady = 5, 3

        # 字体设置
        ttk.Label(frame, text="字体选择:").grid(row=0, column=0, padx=padx, pady=pady, sticky=tk.W)
        font_options = tk.font.families()
        font_menu = ttk.Combobox(frame, textvariable=self.font_family, values=font_options)
        font_menu.grid(row=0, column=1, padx=padx, pady=pady, sticky=tk.EW)

        ttk.Label(frame, text="字体大小:").grid(row=0, column=2, padx=padx, pady=pady, sticky=tk.W)
        font_size_spinbox = ttk.Spinbox(frame, from_=8, to=24, increment=1, textvariable=self.font_size)
        font_size_spinbox.grid(row=0, column=3, padx=padx, pady=pady, sticky=tk.W)
        font_size_spinbox.set(10)

        # 问卷链接
        ttk.Label(frame, text="问卷链接:").grid(row=1, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.url_entry = ttk.Entry(frame, width=60)
        self.url_entry.grid(row=1, column=1, columnspan=3, padx=padx, pady=pady, sticky=tk.EW)
        self.url_entry.insert(0, self.config["url"])

        # 目标份数
        ttk.Label(frame, text="目标份数:").grid(row=2, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.target_entry = ttk.Spinbox(frame, from_=1, to=10000, width=10)
        self.target_entry.grid(row=2, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.target_entry.set(self.config["target_num"])

        # 微信作答比率
        ttk.Label(frame, text="微信作答比率:").grid(row=2, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.ratio_scale = ttk.Scale(frame, from_=0, to=1, orient=tk.HORIZONTAL)
        self.ratio_scale.grid(row=2, column=3, padx=padx, pady=pady, sticky=tk.EW)
        self.ratio_scale.set(self.config["weixin_ratio"])
        self.ratio_var = tk.StringVar()
        self.ratio_var.set(f"{self.config['weixin_ratio'] * 100:.0f}%")
        ttk.Label(frame, textvariable=self.ratio_var).grid(row=2, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.ratio_scale.config(command=lambda v: self.ratio_var.set(f"{float(v) * 100:.0f}%"))

        # 作答时长
        ttk.Label(frame, text="作答时长(秒):").grid(row=3, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(frame, text="最短:").grid(row=3, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_duration = ttk.Spinbox(frame, from_=5, to=300, width=5)
        self.min_duration.grid(row=3, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.min_duration.set(self.config["min_duration"])
        ttk.Label(frame, text="最长:").grid(row=3, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_duration = ttk.Spinbox(frame, from_=5, to=300, width=5)
        self.max_duration.grid(row=3, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.max_duration.set(self.config["max_duration"])

        # 延迟设置
        ttk.Label(frame, text="延迟设置(秒):").grid(row=4, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(frame, text="最小:").grid(row=4, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_delay = ttk.Spinbox(frame, from_=0.1, to=10, increment=0.1, width=5)
        self.min_delay.grid(row=4, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.min_delay.set(self.config["min_delay"])
        ttk.Label(frame, text="最大:").grid(row=4, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_delay = ttk.Spinbox(frame, from_=0.1, to=10, increment=0.1, width=5)
        self.max_delay.grid(row=4, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.max_delay.set(self.config["max_delay"])

        # 提交延迟
        ttk.Label(frame, text="提交延迟:").grid(row=5, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.submit_delay = ttk.Spinbox(frame, from_=1, to=10, width=5)
        self.submit_delay.grid(row=5, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.submit_delay.set(self.config["submit_delay"])

        # 窗口数量
        ttk.Label(frame, text="浏览器窗口数量:").grid(row=6, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.num_threads = ttk.Spinbox(frame, from_=1, to=10, width=5)
        self.num_threads.grid(row=6, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.num_threads.set(self.config["num_threads"])

        # IP设置
        self.use_ip_var = tk.BooleanVar(value=self.config["use_ip"])
        ttk.Checkbutton(frame, text="使用代理IP", variable=self.use_ip_var).grid(
            row=7, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(frame, text="IP API:").grid(row=7, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.ip_entry = ttk.Entry(frame, width=40)
        self.ip_entry.grid(row=7, column=2, columnspan=3, padx=padx, pady=pady, sticky=tk.EW)
        self.ip_entry.insert(0, self.config["ip_api"])

        # 无头模式
        self.headless_var = tk.BooleanVar(value=self.config["headless"])
        ttk.Checkbutton(frame, text="无头模式(不显示浏览器)", variable=self.headless_var).grid(
            row=8, column=0, padx=padx, pady=pady, sticky=tk.W)

        # 按钮组
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=9, column=0, columnspan=5, pady=10, sticky=tk.W)

        # 解析问卷按钮
        ttk.Button(button_frame, text="解析问卷", command=self.parse_survey).grid(row=0, column=0, padx=5)

        # 重置默认按钮
        ttk.Button(button_frame, text="重置默认", command=self.reset_defaults).grid(row=0, column=1, padx=5)

    def create_question_settings(self):
        """创建题型设置界面"""
        # 初始化所有题型的输入框列表
        self.single_entries = []
        self.multi_entries = []
        self.matrix_entries = []
        self.text_entries = []
        self.multiple_text_entries = []
        self.reorder_entries = []
        self.droplist_entries = []
        self.scale_entries = []
        self.multi_other_entries = []
        self.min_selection_entries = []

        # 使用Notebook组织不同题型
        self.question_notebook = ttk.Notebook(self.question_frame)
        self.question_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 单选题设置
        if self.config["single_prob"]:
            self.single_frame = ttk.Frame(self.question_notebook)
            self.question_notebook.add(self.single_frame, text=f"单选题({len(self.config['single_prob'])})")
            self.create_single_settings(self.single_frame)

        # 多选题设置
        if self.config["multiple_prob"]:
            self.multi_frame = ttk.Frame(self.question_notebook)
            self.question_notebook.add(self.multi_frame, text=f"多选题({len(self.config['multiple_prob'])})")
            self.create_multi_settings(self.multi_frame)

        # 矩阵题设置
        if self.config["matrix_prob"]:
            self.matrix_frame = ttk.Frame(self.question_notebook)
            self.question_notebook.add(self.matrix_frame, text=f"矩阵题({len(self.config['matrix_prob'])})")
            self.create_matrix_settings(self.matrix_frame)

        # 填空题设置
        if self.config["texts"]:
            self.text_frame = ttk.Frame(self.question_notebook)
            self.question_notebook.add(self.text_frame, text=f"填空题({len(self.config['texts'])})")
            self.create_text_settings(self.text_frame)

        # 多项填空设置
        if self.config["multiple_texts"]:
            self.multiple_text_frame = ttk.Frame(self.question_notebook)
            self.question_notebook.add(self.multiple_text_frame, text=f"多项填空({len(self.config['multiple_texts'])})")
            self.create_multiple_text_settings(self.multiple_text_frame)

        # 排序题设置
        if self.config["reorder_prob"]:
            self.reorder_frame = ttk.Frame(self.question_notebook)
            self.question_notebook.add(self.reorder_frame, text=f"排序题({len(self.config['reorder_prob'])})")
            self.create_reorder_settings(self.reorder_frame)

        # 下拉框设置
        if self.config["droplist_prob"]:
            self.droplist_frame = ttk.Frame(self.question_notebook)
            self.question_notebook.add(self.droplist_frame, text=f"下拉框({len(self.config['droplist_prob'])})")
            self.create_droplist_settings(self.droplist_frame)

        # 量表题设置
        if self.config["scale_prob"]:
            self.scale_frame = ttk.Frame(self.question_notebook)
            self.question_notebook.add(self.scale_frame, text=f"量表题({len(self.config['scale_prob'])})")
            self.create_scale_settings(self.scale_frame)

    def create_single_settings(self, frame):
        """创建单选题设置"""
        padx, pady = 5, 3
        self.single_entries = []

        # 添加全随机按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=0, column=0, columnspan=6, pady=(0, 5), sticky=tk.W)
        ttk.Button(btn_frame, text="全部随机",
                   command=lambda: self.set_all_random("single", frame)).pack(side=tk.LEFT, padx=5)

        # 说明标签
        ttk.Label(frame, text="设置每个选项的概率（-1表示随机选择）").grid(
            row=1, column=0, columnspan=6, padx=padx, pady=pady, sticky=tk.W)

        for q_num, probs in self.config["single_prob"].items():
            row = len(self.single_entries) + 3
            ttk.Label(frame, text=f"第{q_num}题").grid(row=row, column=0, padx=padx, pady=pady)

            option_count = 5 if probs == -1 else len(probs) if isinstance(probs, list) else 5
            entry_row = []

            for col in range(1, option_count + 1):
                entry = ttk.Entry(frame, width=8)
                if probs == -1:
                    entry.insert(0, -1)
                elif isinstance(probs, list) and col <= len(probs):
                    entry.insert(0, probs[col - 1])
                else:
                    entry.insert(0, "")
                entry.grid(row=row, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.single_entries.append(entry_row)

    def create_multi_settings(self, frame):
        """创建多选题设置"""
        padx, pady = 5, 3
        self.multi_entries = []
        self.multi_other_entries = []
        self.min_selection_entries = []

        # 说明标签
        ttk.Label(frame, text="设置每个多选题各选项被选择的概率（0-100之间的数值）").grid(
            row=1, column=0, columnspan=8, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(frame, text="其他选项答案（多个用逗号分隔）").grid(
            row=1, column=8, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(frame, text="最小选择数量").grid(
            row=1, column=9, padx=padx, pady=pady, sticky=tk.W)

        # 表头
        max_options = max(len(probs) for probs in self.config["multiple_prob"].values()) or 5
        headers = ["题号"] + [f"选项{i + 1}" for i in range(max_options)] + ["其他答案", "最小选择数"]
        for col, header in enumerate(headers):
            ttk.Label(frame, text=header, font=("Arial", 9, "bold")).grid(
                row=2, column=col, padx=padx, pady=pady)

        # 创建多选题设置行
        for i, (q_num, probs) in enumerate(self.config["multiple_prob"].items()):
            row = i + 3
            ttk.Label(frame, text=f"第{q_num}题").grid(row=row, column=0, padx=padx, pady=pady)

            entry_row = []
            option_count = len(probs) if isinstance(probs, list) else 5

            for col in range(1, option_count + 1):
                entry = ttk.Entry(frame, width=8)
                if isinstance(probs, list) and col <= len(probs):
                    entry.insert(0, probs[col - 1])
                else:
                    entry.insert(0, 50)
                entry.grid(row=row, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.multi_entries.append(entry_row)

            # 添加"其他"选项输入框
            other_entry = ttk.Entry(frame, width=25)
            if q_num in self.config["multiple_other"]:
                other_text = ", ".join(self.config["multiple_other"][q_num])
                other_entry.insert(0, other_text)
            other_entry.grid(row=row, column=option_count + 1, padx=padx, pady=pady, sticky=tk.EW)
            self.multi_other_entries.append(other_entry)

            # 添加最小选择数量输入框
            min_selection_entry = ttk.Spinbox(frame, from_=1, to=10, width=5)
            if q_num in self.config["min_selection"]:
                min_selection_entry.set(self.config["min_selection"][q_num])
            else:
                min_selection_entry.set(1)
            min_selection_entry.grid(row=row, column=option_count + 2, padx=padx, pady=pady)
            self.min_selection_entries.append(min_selection_entry)

    def create_matrix_settings(self, frame):
        """创建矩阵题设置"""
        padx, pady = 5, 3
        self.matrix_entries = []

        # 说明标签
        ttk.Label(frame, text="设置矩阵题每行选项的概率（-1表示随机选择）").grid(
            row=0, column=0, columnspan=6, padx=padx, pady=pady, sticky=tk.W)

        # 添加全随机按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=1, column=0, columnspan=6, pady=(0, 5), sticky=tk.W)
        ttk.Button(btn_frame, text="全部随机",
                   command=lambda: self.set_all_random("matrix", frame)).pack(side=tk.LEFT, padx=5)

        for i, (q_num, probs) in enumerate(self.config["matrix_prob"].items()):
            row = i + 2
            ttk.Label(frame, text=f"第{q_num}题").grid(row=row, column=0, padx=padx, pady=pady)

            entry_row = []
            option_count = len(probs) if isinstance(probs, list) else 5

            for col in range(1, option_count + 1):
                entry = ttk.Entry(frame, width=8)
                if probs == -1:
                    entry.insert(0, -1)
                elif isinstance(probs, list) and col <= len(probs):
                    entry.insert(0, probs[col - 1])
                else:
                    entry.insert(0, "")
                entry.grid(row=row, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.matrix_entries.append(entry_row)

    def create_text_settings(self, frame):
        """创建填空题设置"""
        padx, pady = 5, 3
        self.text_entries = []

        # 说明标签
        ttk.Label(frame, text="设置填空题可选答案（多个答案用逗号分隔）").grid(
            row=0, column=0, columnspan=2, padx=padx, pady=pady, sticky=tk.W)

        for i, (q_num, texts) in enumerate(self.config["texts"].items()):
            row = i + 1
            ttk.Label(frame, text=f"第{q_num}题").grid(row=row, column=0, padx=padx, pady=pady)

            entry = ttk.Entry(frame, width=60)
            entry.insert(0, ", ".join(texts))
            entry.grid(row=row, column=1, padx=padx, pady=pady, sticky=tk.EW)
            self.text_entries.append(entry)

    def create_multiple_text_settings(self, frame):
        """创建多项填空设置"""
        padx, pady = 5, 3
        self.multiple_text_entries = []

        # 说明标签
        ttk.Label(frame, text="设置每道多项填空题的可选答案（每行答案用逗号分隔）").grid(
            row=0, column=0, columnspan=2, padx=padx, pady=pady, sticky=tk.W)

        row = 1
        for q_num, text_lists in self.config["multiple_texts"].items():
            ttk.Label(frame, text=f"第{q_num}题").grid(row=row, column=0, padx=padx, pady=pady)

            entry_col = []
            for i, texts in enumerate(text_lists):
                entry = ttk.Entry(frame, width=60)
                entry.insert(0, ", ".join(texts))
                entry.grid(row=row + i, column=1, padx=padx, pady=pady, sticky=tk.EW)
                entry_col.append(entry)

            self.multiple_text_entries.append(entry_col)
            row += len(text_lists) + 1

    def create_reorder_settings(self, frame):
        """创建排序题设置"""
        padx, pady = 5, 3
        self.reorder_entries = []

        # 说明标签
        ttk.Label(frame, text="设置排序题每个位置的概率分布（数值和为1）").grid(
            row=0, column=0, columnspan=6, padx=padx, pady=pady, sticky=tk.W)

        for i, (q_num, probs) in enumerate(self.config["reorder_prob"].items()):
            row = i + 1
            ttk.Label(frame, text=f"第{q_num}题").grid(row=row, column=0, padx=padx, pady=pady)

            entry_row = []
            for col, prob in enumerate(probs, 1):
                entry = ttk.Entry(frame, width=8)
                entry.insert(0, prob)
                entry.grid(row=row, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.reorder_entries.append(entry_row)

    def create_droplist_settings(self, frame):
        """创建下拉框设置"""
        padx, pady = 5, 3
        self.droplist_entries = []

        # 说明标签
        ttk.Label(frame, text="设置下拉框各选项的选择概率").grid(
            row=0, column=0, columnspan=6, padx=padx, pady=pady, sticky=tk.W)

        for i, (q_num, probs) in enumerate(self.config["droplist_prob"].items()):
            row = i + 1
            ttk.Label(frame, text=f"第{q_num}题").grid(row=row, column=0, padx=padx, pady=pady)

            entry_row = []
            for col, prob in enumerate(probs, 1):
                entry = ttk.Entry(frame, width=8)
                entry.insert(0, prob)
                entry.grid(row=row, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.droplist_entries.append(entry_row)

    def create_scale_settings(self, frame):
        """创建量表题设置"""
        padx, pady = 5, 3
        self.scale_entries = []

        # 说明标签
        ttk.Label(frame, text="设置量表题各选项的选择概率").grid(
            row=0, column=0, columnspan=12, padx=padx, pady=pady, sticky=tk.W)

        for i, (q_num, probs) in enumerate(self.config["scale_prob"].items()):
            row = i + 1
            ttk.Label(frame, text=f"第{q_num}题").grid(row=row, column=0, padx=padx, pady=pady)

            entry_row = []
            for col, prob in enumerate(probs, 1):
                entry = ttk.Entry(frame, width=8)
                entry.insert(0, prob)
                entry.grid(row=row, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.scale_entries.append(entry_row)

    def save_config(self):
        """保存当前配置"""
        try:
            # 保存全局设置
            self.config["url"] = self.url_entry.get().strip()
            self.config["target_num"] = int(self.target_entry.get())
            self.config["weixin_ratio"] = float(self.ratio_scale.get())
            self.config["min_duration"] = int(self.min_duration.get())
            self.config["max_duration"] = int(self.max_duration.get())
            self.config["min_delay"] = float(self.min_delay.get())
            self.config["max_delay"] = float(self.max_delay.get())
            self.config["submit_delay"] = int(self.submit_delay.get())
            self.config["num_threads"] = int(self.num_threads.get())
            self.config["use_ip"] = self.use_ip_var.get()
            self.config["headless"] = self.headless_var.get()
            self.config["ip_api"] = self.ip_entry.get().strip()

            # 保存单选题设置
            for i, entries in enumerate(self.single_entries):
                q_num = list(self.config["single_prob"].keys())[i]
                probs = []
                for entry in entries:
                    try:
                        value = float(entry.get())
                        if value == -1:
                            self.config["single_prob"][q_num] = -1
                            break
                        probs.append(value)
                    except ValueError:
                        continue
                if probs and -1 not in probs:
                    self.config["single_prob"][q_num] = probs

            # 保存多选题设置
            self.config["multiple_prob"] = {}
            self.config["multiple_other"] = {}
            self.config["min_selection"] = {}
            for i, (entries, other_entry, min_entry) in enumerate(zip(
                    self.multi_entries, self.multi_other_entries, self.min_selection_entries)):
                q_num = list(self.config["multiple_prob"].keys())[i]
                probs = []
                for entry in entries:
                    try:
                        value = float(entry.get())
                        probs.append(value)
                    except ValueError:
                        continue
                if probs:
                    self.config["multiple_prob"][q_num] = probs

                # 保存其他选项答案
                other_text = other_entry.get().strip()
                if other_text:
                    other_answers = [ans.strip() for ans in other_text.split(",")]
                    self.config["multiple_other"][q_num] = other_answers

                # 保存最小选择数量
                try:
                    min_selection = int(min_entry.get())
                    self.config["min_selection"][q_num] = min_selection
                except ValueError:
                    pass

            # 保存矩阵题设置
            for i, entries in enumerate(self.matrix_entries):
                q_num = list(self.config["matrix_prob"].keys())[i]
                probs = []
                for entry in entries:
                    try:
                        value = float(entry.get())
                        if value == -1:
                            self.config["matrix_prob"][q_num] = -1
                            break
                        probs.append(value)
                    except ValueError:
                        continue
                if probs and -1 not in probs:
                    self.config["matrix_prob"][q_num] = probs

            # 保存填空题设置
            self.config["texts"] = {}
            for i, entry in enumerate(self.text_entries):
                q_num = list(self.config["texts"].keys())[i]
                text = entry.get().strip()
                if text:
                    texts = [t.strip() for t in text.split(",")]
                    self.config["texts"][q_num] = texts

            # 保存多项填空设置
            self.config["multiple_texts"] = {}
            self.config["multiple_texts_prob"] = {}
            for i, entries in enumerate(self.multiple_text_entries):
                q_num = list(self.config["multiple_texts"].keys())[i]
                text_lists = []
                prob_lists = []
                for entry in entries:
                    texts = [t.strip() for t in entry.get().strip().split(",")]
                    if texts:
                        text_lists.append(texts)
                        # 为每个答案设置相等概率
                        probs = [1 / len(texts)] * len(texts)
                        prob_lists.append(probs)
                if text_lists:
                    self.config["multiple_texts"][q_num] = text_lists
                    self.config["multiple_texts_prob"][q_num] = prob_lists

            # 保存排序题设置
            for i, entries in enumerate(self.reorder_entries):
                q_num = list(self.config["reorder_prob"].keys())[i]
                probs = []
                for entry in entries:
                    try:
                        value = float(entry.get())
                        probs.append(value)
                    except ValueError:
                        continue
                if probs:
                    self.config["reorder_prob"][q_num] = probs

            return True
        except Exception as e:
            logging.error(f"保存配置时出错: {str(e)}")
            messagebox.showerror("错误", f"保存配置时出错: {str(e)}")
            return False

    def export_config(self):
        """导出配置到文件"""
        if not self.save_config():
            return

        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                initialfile="wjx_config.json"
            )
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.config, f, ensure_ascii=False, indent=4)
                logging.info(f"配置已导出到: {file_path}")
                messagebox.showinfo("成功", "配置导出成功！")
        except Exception as e:
            logging.error(f"导出配置时出错: {str(e)}")
            messagebox.showerror("错误", f"导出配置时出错: {str(e)}")

    def import_config(self):
        """从文件导入配置"""
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            if file_path:
                with open(file_path, 'r', encoding='utf-8') as f:
                    new_config = json.load(f)

                # 验证配置有效性
                required_keys = ["url", "target_num", "min_duration", "max_duration",
                                 "weixin_ratio", "min_delay", "max_delay", "submit_delay"]
                for key in required_keys:
                    if key not in new_config:
                        raise ValueError(f"配置文件缺少必需的字段: {key}")

                self.config = new_config
                self.reset_ui_with_config()
                logging.info(f"配置已从文件导入: {file_path}")
                messagebox.showinfo("成功", "配置导入成功！")
        except Exception as e:
            logging.error(f"导入配置时出错: {str(e)}")
            messagebox.showerror("错误", f"导入配置时出错: {str(e)}")

    def reset_defaults(self):
        """重置为默认配置"""
        if messagebox.askyesno("确认", "确定要重置所有设置为默认值吗？"):
            self.config = DEFAULT_CONFIG.copy()
            self.reset_ui_with_config()
            logging.info("已重置为默认配置")

    def reset_ui_with_config(self):
        """根据当前配置重置UI"""
        # 重置全局设置
        self.url_entry.delete(0, tk.END)
        self.url_entry.insert(0, self.config["url"])

        self.target_entry.delete(0, tk.END)
        self.target_entry.insert(0, str(self.config["target_num"]))

        self.ratio_scale.set(self.config["weixin_ratio"])

        self.min_duration.delete(0, tk.END)
        self.min_duration.insert(0, str(self.config["min_duration"]))

        self.max_duration.delete(0, tk.END)
        self.max_duration.insert(0, str(self.config["max_duration"]))

        self.min_delay.delete(0, tk.END)
        self.min_delay.insert(0, str(self.config["min_delay"]))

        self.max_delay.delete(0, tk.END)
        self.max_delay.insert(0, str(self.config["max_delay"]))

        self.submit_delay.delete(0, tk.END)
        self.submit_delay.insert(0, str(self.config["submit_delay"]))

        self.num_threads.delete(0, tk.END)
        self.num_threads.insert(0, str(self.config["num_threads"]))

        self.use_ip_var.set(self.config["use_ip"])
        self.headless_var.set(self.config["headless"])

        self.ip_entry.delete(0, tk.END)
        self.ip_entry.insert(0, self.config["ip_api"])

        # 重新创建题型设置
        for widget in self.question_frame.winfo_children():
            widget.destroy()
        self.create_question_settings()

    def clear_log(self):
        """清空日志"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        logging.info("日志已清空")

    def export_log(self):
        """导出日志"""
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfile="wjx_log.txt"
            )
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_area.get(1.0, tk.END))
                logging.info(f"日志已导出到: {file_path}")
                messagebox.showinfo("成功", "日志导出成功！")
        except Exception as e:
            logging.error(f"导出日志时出错: {str(e)}")
            messagebox.showerror("错误", f"导出日志时出错: {str(e)}")

    def update_font(self, *args):
        """更新UI字体"""
        try:
            font_family = self.font_family.get()
            font_size = self.font_size.get()
            new_font = (font_family, font_size)
            style = ttk.Style()
            style.configure(".", font=new_font)
            self.log_area.configure(font=new_font)

            def update_widget_font(widget):
                try:
                    if hasattr(widget, 'configure') and 'font' in widget.configure():
                        widget.configure(font=new_font)
                    for child in widget.winfo_children():
                        update_widget_font(child)
                except Exception:
                    pass

            update_widget_font(self.root)
        except Exception as e:
            logging.error(f"更新字体时出错: {str(e)}")

    def parse_survey(self):
        """解析问卷"""
        try:
            url = self.url_entry.get().strip()
            if not url:
                messagebox.showerror("错误", "请输入问卷链接")
                return

            # 创建临时浏览器实例
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')  # 无头模式
            driver = webdriver.Chrome(options=options)

            try:
                driver.get(url)
                time.sleep(2)  # 等待页面加载

                # 获取所有题目
                questions = driver.find_elements(By.CLASS_NAME, "div_question")

                # 重置配置
                self.config["single_prob"] = {}
                self.config["multiple_prob"] = {}
                self.config["matrix_prob"] = {}
                self.config["droplist_prob"] = {}
                self.config["scale_prob"] = {}
                self.config["texts"] = {}
                self.config["multiple_texts"] = {}
                self.config["reorder_prob"] = {}

                for q in questions:
                    try:
                        q_type = q.get_attribute("type")
                        q_num = q.get_attribute("id").replace("div", "")

                        if q_type == "1":  # 单选题
                            options = q.find_elements(By.CSS_SELECTOR, "input[type='radio']")
                            self.config["single_prob"][q_num] = [-1] * len(options)

                        elif q_type == "2":  # 多选题
                            options = q.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                            self.config["multiple_prob"][q_num] = [50] * len(options)
                            self.config["min_selection"][q_num] = 1

                        elif q_type == "3":  # 填空题
                            self.config["texts"][q_num] = [""]

                        elif q_type == "4":  # 矩阵题
                            rows = q.find_elements(By.CLASS_NAME, "matrix-row")
                            self.config["matrix_prob"][q_num] = [-1] * len(rows)

                        elif q_type == "5":  # 量表题
                            options = q.find_elements(By.CSS_SELECTOR, "input[type='radio']")
                            self.config["scale_prob"][q_num] = [1] * len(options)

                        elif q_type == "6":  # 下拉框
                            options = q.find_elements(By.TAG_NAME, "option")
                            self.config["droplist_prob"][q_num] = [1] * (len(options) - 1)  # 减去默认选项

                        elif q_type == "8":  # 多项填空
                            inputs = q.find_elements(By.TAG_NAME, "input")
                            self.config["multiple_texts"][q_num] = [[""] for _ in range(len(inputs))]

                        elif q_type == "7":  # 排序题
                            items = q.find_elements(By.CLASS_NAME, "reorder-item")
                            self.config["reorder_prob"][q_num] = [1 / len(items)] * len(items)

                    except Exception as e:
                        logging.warning(f"解析第{q_num}题时出错: {str(e)}")
                        continue

                # 重新加载UI
                self.reset_ui_with_config()
                logging.info("问卷解析完成")
                messagebox.showinfo("成功", "问卷解析完成！")

            finally:
                driver.quit()

        except Exception as e:
            logging.error(f"解析问卷时出错: {str(e)}")
            messagebox.showerror("错误", f"解析问卷时出错: {str(e)}")

    def start_filling(self):
        """开始填写问卷"""
        try:
            # 保存当前配置
            if not self.save_config():
                return

            # 验证基本参数
            if not self.config["url"]:
                messagebox.showerror("错误", "请输入问卷链接")
                return

            try:
                self.config["target_num"] = int(self.target_entry.get())
                if self.config["target_num"] <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("错误", "目标份数必须是正整数")
                return

            try:
                self.config["num_threads"] = int(self.num_threads.get())
                if self.config["num_threads"] <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("错误", "窗口数量必须是正整数")
                return

            # 更新运行状态
            self.running = True
            self.cur_num = 0
            self.cur_fail = 0
            self.pause_event.clear()

            # 更新按钮状态
            self.start_btn.config(state=tk.DISABLED)
            self.pause_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.NORMAL)

            # 设置进度条初始值
            self.progress_var.set(0)
            self.question_progress_var.set(0)
            self.status_var.set("运行中...")

            # 创建并启动线程
            threads = []
            for i in range(self.config["num_threads"]):
                x = (i % 2) * 600
                y = (i // 2) * 400
                t = threading.Thread(target=self.run_filling, args=(x, y))
                t.daemon = True
                t.start()
                threads.append(t)

            # 启动进度更新线程
            progress_thread = threading.Thread(target=self.update_progress)
            progress_thread.daemon = True
            progress_thread.start()

        except Exception as e:
            logging.error(f"启动失败: {str(e)}")
            messagebox.showerror("错误", f"启动失败: {str(e)}")

    def run_filling(self, x=0, y=0):
        """运行填写任务"""
        options = webdriver.ChromeOptions()
        if self.config["headless"]:
            options.add_argument('--headless')
        else:
            options.add_argument(f'--window-position={x},{y}')

        driver = None
        try:
            while self.running and self.cur_num < self.config["target_num"]:
                # 检查是否暂停
                self.pause_event.wait()

                # 获取代理IP
                if self.config["use_ip"]:
                    try:
                        response = requests.get(self.config["ip_api"])
                        ip = response.text.strip()
                        options.add_argument(f'--proxy-server={ip}')
                    except Exception as e:
                        logging.error(f"获取代理IP失败: {str(e)}")
                        continue

                # 创建浏览器实例
                driver = webdriver.Chrome(options=options)
                try:
                    # 访问问卷
                    driver.get(self.config["url"])
                    time.sleep(self.config["page_load_delay"])

                    # 随机决定是否使用微信作答
                    use_weixin = random.random() < self.config["weixin_ratio"]
                    if use_weixin:
                        weixin_btn = driver.find_element(By.CLASS_NAME, "weixin-answer")
                        weixin_btn.click()
                        time.sleep(2)

                    # 填写问卷
                    self.fill_survey(driver)

                    # 提交问卷
                    submit_btn = driver.find_element(By.ID, "submit_button")
                    time.sleep(self.config["submit_delay"])
                    submit_btn.click()

                    # 等待提交完成
                    time.sleep(2)

                    # 检查是否提交成功
                    if "完成" in driver.title or "提交成功" in driver.page_source:
                        with self.lock:
                            self.cur_num += 1
                        logging.info(f"第 {self.cur_num} 份问卷提交成功")
                    else:
                        with self.lock:
                            self.cur_fail += 1
                        logging.warning(f"第 {self.cur_num + 1} 份问卷提交失败")

                except Exception as e:
                    with self.lock:
                        self.cur_fail += 1
                    logging.error(f"填写问卷时出错: {str(e)}")
                    traceback.print_exc()

                finally:
                    try:
                        driver.quit()
                    except:
                        pass

                # 随机等待
                if self.running:
                    time.sleep(random.uniform(
                        self.config["min_delay"],
                        self.config["max_delay"]))

        except Exception as e:
            logging.error(f"运行任务时出错: {str(e)}")
        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass

    def fill_survey(self, driver):
        """填写问卷内容"""
        questions = driver.find_elements(By.CLASS_NAME, "div_question")
        total_questions = len(questions)

        # 随机总作答时间
        total_time = random.randint(self.config["min_duration"], self.config["max_duration"])
        start_time = time.time()

        for i, q in enumerate(questions):
            if not self.running:
                break

            try:
                q_type = q.get_attribute("type")
                q_num = q.get_attribute("id").replace("div", "")

                # 更新题目进度
                self.question_progress_var.set((i + 1) / total_questions * 100)
                self.question_status_var.set(f"题目进度: {i + 1}/{total_questions}")

                # 随机等待时间
                per_question_delay = random.uniform(*self.config["per_question_delay"])
                time.sleep(per_question_delay)

                # 根据题型填写
                if q_type == "1":  # 单选题
                    self.fill_single_choice(q, q_num)
                elif q_type == "2":  # 多选题
                    self.fill_multiple_choice(q, q_num)
                elif q_type == "3":  # 填空题
                    self.fill_text(q, q_num)
                elif q_type == "4":  # 矩阵题
                    self.fill_matrix(q, q_num)
                elif q_type == "5":  # 量表题
                    self.fill_scale(q, q_num)
                elif q_type == "6":  # 下拉框
                    self.fill_droplist(q, q_num)
                elif q_type == "7":  # 排序题
                    self.fill_reorder(q, q_num)
                elif q_type == "8":  # 多项填空
                    self.fill_multiple_text(q, q_num)

            except Exception as e:
                logging.error(f"填写第{q_num}题时出错: {str(e)}")
                continue

            # 检查是否需要翻页
            try:
                next_page = driver.find_element(By.CLASS_NAME, "next-page")
                if next_page.is_displayed():
                    next_page.click()
                    time.sleep(random.uniform(*self.config["per_page_delay"]))
            except:
                pass

        # 补足剩余时间
        elapsed_time = time.time() - start_time
        if elapsed_time < total_time:
            time.sleep(total_time - elapsed_time)

    def fill_single_choice(self, question, q_num):
        """填写单选题"""
        options = question.find_elements(By.CSS_SELECTOR, "input[type='radio']")
        if not options:
            return

        probs = self.config["single_prob"].get(q_num, [-1])
        if probs == -1 or len(probs) == 1 and probs[0] == -1:
            # 随机选择
            selected = random.choice(options)
        else:
            # 按概率选择
            probs = probs[:len(options)]
            selected = random.choices(options, weights=probs, k=1)[0]

        try:
            selected.click()
        except:
            # 如果直接点击失败，尝试使用JavaScript点击
            driver = options[0].parent
            driver.execute_script("arguments[0].click();", selected)

    def fill_multiple_choice(self, question, q_num):
        """填写多选题"""
        options = question.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
        if not options:
            return

        probs = self.config["multiple_prob"].get(q_num, [50] * len(options))
        min_selection = self.config["min_selection"].get(q_num, 1)

        # 确定选择数量
        max_selection = min(len(options), self.config.get("max_selection", {}).get(q_num, len(options)))
        num_selections = random.randint(min_selection, max_selection)

        # 根据概率选择选项
        selected_indices = []
        for i, prob in enumerate(probs[:len(options)]):
            if random.random() * 100 < prob:
                selected_indices.append(i)

        # 调整选择数量
        if len(selected_indices) < min_selection:
            # 随机添加
            available = [i for i in range(len(options)) if i not in selected_indices]
            add_count = min_selection - len(selected_indices)
            selected_indices.extend(random.sample(available, add_count))
        elif len(selected_indices) > max_selection:
            # 随机移除
            remove_count = len(selected_indices) - max_selection
            selected_indices = random.sample(selected_indices, max_selection)

        # 点击选中的选项
        for i in selected_indices:
            try:
                options[i].click()
            except:
                driver = options[0].parent
                driver.execute_script("arguments[0].click();", options[i])

        # 填写其他选项
        if q_num in self.config["multiple_other"]:
            try:
                other_input = question.find_element(By.CSS_SELECTOR, "input[type='text']")
                other_answers = self.config["multiple_other"][q_num]
                if other_answers:
                    other_input.send_keys(random.choice(other_answers))
            except:
                pass

    def fill_text(self, question, q_num):
        """填写填空题"""
        inputs = question.find_elements(By.TAG_NAME, "input")
        if not inputs:
            inputs = question.find_elements(By.TAG_NAME, "textarea")
        if not inputs:
            return

        texts = self.config["texts"].get(q_num, [""])
        if not texts:
            return

        for input_elem in inputs:
            try:
                selected_text = random.choice(texts)
                # 使用JavaScript设置值，避免输入事件问题
                driver = input_elem.parent
                driver.execute_script(f"arguments[0].value = '{selected_text}';", input_elem)
                # 触发change事件
                driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", input_elem)
                time.sleep(random.uniform(0.1, 0.3))
            except Exception as e:
                logging.error(f"填写文本时出错: {str(e)}")

    def fill_matrix(self, question, q_num):
        """填写矩阵题"""
        rows = question.find_elements(By.CLASS_NAME, "matrix-row")
        if not rows:
            return

        probs = self.config["matrix_prob"].get(q_num, [-1])

        for row in rows:
            options = row.find_elements(By.CSS_SELECTOR, "input[type='radio']")
            if not options:
                continue

            if probs == -1 or (len(probs) == 1 and probs[0] == -1):
                # 随机选择
                selected = random.choice(options)
            else:
                # 按概率选择
                row_probs = probs[:len(options)]
                selected = random.choices(options, weights=row_probs, k=1)[0]

            try:
                selected.click()
            except:
                driver = options[0].parent
                driver.execute_script("arguments[0].click();", selected)

            time.sleep(random.uniform(0.1, 0.3))

    def fill_scale(self, question, q_num):
        """填写量表题"""
        options = question.find_elements(By.CSS_SELECTOR, "input[type='radio']")
        if not options:
            return

        probs = self.config["scale_prob"].get(q_num, [1] * len(options))
        if len(probs) < len(options):
            probs.extend([1] * (len(options) - len(probs)))

        # 按概率选择
        selected = random.choices(options, weights=probs[:len(options)], k=1)[0]

        try:
            selected.click()
        except:
            driver = options[0].parent
            driver.execute_script("arguments[0].click();", selected)

    def fill_droplist(self, question, q_num):
        """填写下拉框"""
        select = question.find_element(By.TAG_NAME, "select")
        options = select.find_elements(By.TAG_NAME, "option")[1:]  # 跳过第一个默认选项
        if not options:
            return

        probs = self.config["droplist_prob"].get(q_num, [1] * len(options))
        if len(probs) < len(options):
            probs.extend([1] * (len(options) - len(probs)))

        # 按概率选择
        selected = random.choices(options, weights=probs[:len(options)], k=1)[0]

        try:
            selected.click()
        except:
            driver = options[0].parent
            driver.execute_script("arguments[0].click();", selected)

    def fill_reorder(self, question, q_num):
        """填写排序题"""
        items = question.find_elements(By.CLASS_NAME, "reorder-item")
        if not items:
            return

        probs = self.config["reorder_prob"].get(q_num, [1 / len(items)] * len(items))

        # 创建排序列表
        order = list(range(len(items)))
        random.shuffle(order)  # 随机打乱顺序

        # 使用JavaScript移动元素
        driver = items[0].parent
        for i, item_index in enumerate(order):
            try:
                # 计算目标位置
                target_y = items[i].location['y']
                current_y = items[item_index].location['y']
                offset = target_y - current_y

                # 使用JavaScript移动元素
                script = f"""
                var item = arguments[0];
                var rect = item.getBoundingClientRect();
                var evt = new MouseEvent('mousedown', {{
                    bubbles: true,
                    clientX: rect.left,
                    clientY: rect.top
                }});
                item.dispatchEvent(evt);

                evt = new MouseEvent('mousemove', {{
                    bubbles: true,
                    clientX: rect.left,
                    clientY: rect.top + {offset}
                }});
                item.dispatchEvent(evt);

                evt = new MouseEvent('mouseup', {{
                    bubbles: true,
                    clientX: rect.left,
                    clientY: rect.top + {offset}
                }});
                item.dispatchEvent(evt);
                """
                driver.execute_script(script, items[item_index])
                time.sleep(random.uniform(0.2, 0.5))
            except Exception as e:
                logging.error(f"移动排序项时出错: {str(e)}")

    def fill_multiple_text(self, question, q_num):
        """填写多项填空题"""
        inputs = question.find_elements(By.TAG_NAME, "input")
        if not inputs:
            return

        text_lists = self.config["multiple_texts"].get(q_num, [])
        if not text_lists:
            return

        for i, input_elem in enumerate(inputs):
            if i < len(text_lists):
                try:
                    selected_text = random.choice(text_lists[i])
                    # 使用JavaScript设置值
                    driver = input_elem.parent
                    driver.execute_script(f"arguments[0].value = '{selected_text}';", input_elem)
                    # 触发change事件
                    driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", input_elem)
                    time.sleep(random.uniform(0.1, 0.3))
                except Exception as e:
                    logging.error(f"填写多项填空时出错: {str(e)}")

    def update_progress(self):
        """更新进度显示"""
        while self.running:
            try:
                if self.config["target_num"] > 0:
                    progress = (self.cur_num / self.config["target_num"]) * 100
                    self.progress_var.set(progress)

                status = "暂停中..." if self.paused else "运行中..."
                status += f" 完成: {self.cur_num}/{self.config['target_num']}"
                if self.cur_fail > 0:
                    status += f" 失败: {self.cur_fail}"
                self.status_var.set(status)

                if self.cur_num >= self.config["target_num"]:
                    self.stop_filling()
                    messagebox.showinfo("完成", "问卷填写完成！")
                    break

            except Exception as e:
                logging.error(f"更新进度时出错: {str(e)}")

            time.sleep(0.5)

    def toggle_pause(self):
        """切换暂停/继续状态"""
        self.paused = not self.paused
        if self.paused:
            self.pause_event.clear()
            self.pause_btn.config(text="继续")
            logging.info("已暂停")
        else:
            self.pause_event.set()
            self.pause_btn.config(text="暂停")
            logging.info("已继续")

    def stop_filling(self):
        """停止填写"""
        self.running = False
        self.pause_event.set()  # 确保所有线程都能退出
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("已停止")
        logging.info("已停止")

    def set_all_random(self, question_type, frame):
        """将所有概率设置为随机"""
        if question_type == "single":
            for entries in self.single_entries:
                for entry in entries:
                    entry.delete(0, tk.END)
                    entry.insert(0, "-1")
        elif question_type == "matrix":
            for entries in self.matrix_entries:
                for entry in entries:
                    entry.delete(0, tk.END)
                    entry.insert(0, "-1")
        logging.info(f"已将所有{question_type}题设置为随机")

    def on_closing(self):
        """关闭窗口时的处理"""
        if self.running:
            if messagebox.askokcancel("确认", "正在运行中，确定要退出吗？"):
                self.stop_filling()
                self.root.destroy()
        else:
            self.root.destroy()


def main():
    """主函数"""
    root = tk.Tk()
    app = WJXAutoFillApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()