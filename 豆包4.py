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
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    ElementNotInteractableException
)

# =================== 默认配置参数 ===================
DEFAULT_CONFIG = {
    "url": "https://www.wjx.cn/vm/sample.aspx",  # 示例URL
    "target_num": 100,  # 目标填写份数
    "min_duration": 15,  # 最短作答时间(秒)
    "max_duration": 180,  # 最长作答时间(秒)
    "weixin_ratio": 0.5,  # 使用微信作答的比例
    "min_delay": 2.0,  # 最小提交间隔(秒)
    "max_delay": 6.0,  # 最大提交间隔(秒)
    "submit_delay": 3,  # 提交前等待时间(秒)
    "page_load_delay": 2,  # 页面加载等待时间(秒)
    "per_question_delay": (1.0, 3.0),  # 每题作答延迟范围(秒)
    "per_page_delay": (2.0, 6.0),  # 翻页延迟范围(秒)

    # 基本设置
    "use_ip": False,  # 是否使用代理IP
    "headless": False,  # 是否使用无头模式
    "ip_api": "",  # 代理IP接口地址
    "num_threads": 4,  # 并发线程数

    # 题型概率配置(示例)
    "single_prob": {},  # 单选题概率
    "multiple_prob": {},  # 多选题概率
    "matrix_prob": {},  # 矩阵题概率
    "scale_prob": {},  # 量表题概率
    "droplist_prob": {},  # 下拉题概率
    "texts": {},  # 填空题答案
    "multiple_texts": {},  # 多项填空答案
    "reorder_prob": {},  # 排序题概率

    # 其他设置
    "min_selection": {},  # 多选题最少选择数
    "max_selection": {},  # 多选题最多选择数
    "multiple_other": {},  # 多选题其他选项
    "multiple_texts_prob": {},  # 多项填空概率
}


class WJXAutoFillApp:
    def __init__(self, root):
        """初始化应用"""
        self.root = root
        self.root.title("问卷星自动填写工具 v2.0")
        self.root.geometry("1200x900")
        self.root.resizable(True, True)

        # 基本变量初始化
        self.config = DEFAULT_CONFIG.copy()
        self.running = False
        self.paused = False
        self.cur_num = 0
        self.cur_fail = 0
        self.lock = threading.Lock()
        self.pause_event = threading.Event()
        self.pause_event.set()

        # 字体设置
        self.font_family = tk.StringVar(value="Microsoft YaHei UI")
        self.font_size = tk.IntVar(value=10)

        # 输入框列表初始化
        self.initialize_entry_lists()

        # 创建界面
        self.create_gui()

        # 设置日志系统
        self.setup_logging()

        # 更新字体
        self.update_font()

        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def initialize_entry_lists(self):
        """初始化所有输入框列表"""
        self.single_entries = []  # 单选题
        self.multi_entries = []  # 多选题
        self.matrix_entries = []  # 矩阵题
        self.scale_entries = []  # 量表题
        self.droplist_entries = []  # 下拉框
        self.text_entries = []  # 填空题
        self.multiple_text_entries = []  # 多项填空
        self.reorder_entries = []  # 排序题
        self.multi_other_entries = []  # 多选其他选项
        self.min_selection_entries = []  # 最小选择数

    def create_gui(self):
        """创建主界面"""
        style = ttk.Style()
        style.theme_use('default')
        style.configure('TNotebook.Tab', padding=[12, 5])
        style.configure('TButton', padding=[10, 5])
        style.configure('Horizontal.TProgressbar',
                        background='#00FF00',
                        troughcolor='#F0F0F0',
                        bordercolor='#999999',
                        lightcolor='#FFFFFF',
                        darkcolor='#999999')

        # 创建主框架
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建上部分框架
        self.top_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(self.top_frame, weight=3)

        # 创建下部分框架
        self.bottom_frame = ttk.LabelFrame(self.main_paned, text="运行日志")
        self.main_paned.add(self.bottom_frame, weight=1)

        # 创建各个部分
        self.create_control_panel()  # 控制面板
        self.create_notebook()  # 设置标签页
        self.create_log_area()  # 日志区域

    def create_control_panel(self):
        """创建控制面板"""
        control_frame = ttk.LabelFrame(self.top_frame, text="控制面板")
        control_frame.pack(fill=tk.X, padx=5, pady=5)

        # 按钮区域
        btn_frame = ttk.Frame(control_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)

        # 第一行按钮
        self.start_btn = ttk.Button(btn_frame, text="开始填写", command=self.start_filling)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.pause_btn = ttk.Button(btn_frame, text="暂停", command=self.toggle_pause, state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(btn_frame, text="停止", command=self.stop_filling, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame, text="导出配置", command=self.export_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="导入配置", command=self.import_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="重置默认", command=self.reset_defaults).pack(side=tk.LEFT, padx=5)

        # 进度显示区域
        progress_frame = ttk.Frame(control_frame)
        progress_frame.pack(fill=tk.X, padx=5, pady=5)

        # 总进度
        ttk.Label(progress_frame, text="总进度:").pack(side=tk.LEFT, padx=5)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame,
                                            variable=self.progress_var,
                                            maximum=100,
                                            length=300)
        self.progress_bar.pack(side=tk.LEFT, padx=5)

        # 题目进度
        ttk.Label(progress_frame, text="题目进度:").pack(side=tk.LEFT, padx=5)
        self.question_progress_var = tk.DoubleVar()
        self.question_progress_bar = ttk.Progressbar(progress_frame,
                                                     variable=self.question_progress_var,
                                                     maximum=100,
                                                     length=200)
        self.question_progress_bar.pack(side=tk.LEFT, padx=5)

        # 状态显示
        self.status_var = tk.StringVar(value="就绪")
        self.question_status_var = tk.StringVar(value="题目进度: 0/0")
        ttk.Label(progress_frame, textvariable=self.status_var).pack(side=tk.RIGHT, padx=5)
        ttk.Label(progress_frame, textvariable=self.question_status_var).pack(side=tk.RIGHT, padx=5)

    def create_notebook(self):
        """创建设置标签页"""
        self.notebook = ttk.Notebook(self.top_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 创建全局设置页
        self.global_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.global_frame, text="全局设置")
        self.create_global_settings()

        # 创建题型设置页
        self.question_container = ttk.Frame(self.notebook)
        self.notebook.add(self.question_container, text="题型设置")

        # 创建带滚动条的题型设置区域
        self.create_scrollable_question_frame()

    def create_scrollable_question_frame(self):
        """创建可滚动的题型设置区域"""
        # 创建Canvas和滚动条
        self.canvas = tk.Canvas(self.question_container)
        scrollbar = ttk.Scrollbar(self.question_container,
                                  orient="vertical",
                                  command=self.canvas.yview)

        # 创建内部框架
        self.question_frame = ttk.Frame(self.canvas)
        self.question_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        # 在Canvas中创建窗口
        self.canvas.create_window((0, 0), window=self.question_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        # 配置Canvas的滚动
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.question_frame.bind("<Enter>", self.bind_mousewheel)
        self.question_frame.bind("<Leave>", self.unbind_mousewheel)

        # 布局
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 创建题型设置内容
        self.create_question_settings()
    def on_canvas_configure(self, event):
        """处理Canvas大小改变事件"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        width = event.width - 10  # 留出滚动条的空间
        self.canvas.itemconfig(self.canvas.find_all()[0], width=width)

    def bind_mousewheel(self, event):
        """绑定鼠标滚轮事件"""
        if os.name == 'nt':  # Windows
            self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)
        else:  # Linux
            self.canvas.bind_all("<Button-4>", self.on_mousewheel)
            self.canvas.bind_all("<Button-5>", self.on_mousewheel)

    def unbind_mousewheel(self, event):
        """解绑鼠标滚轮事件"""
        if os.name == 'nt':  # Windows
            self.canvas.unbind_all("<MouseWheel>")
        else:  # Linux
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")

    def on_mousewheel(self, event):
        """处理鼠标滚轮事件"""
        if os.name == 'nt':  # Windows
            delta = -1 * (event.delta // 120)
        else:  # Linux
            if event.num == 4:
                delta = -1
            else:
                delta = 1
        self.canvas.yview_scroll(delta, "units")

    def create_global_settings(self):
        """创建全局设置页面"""
        frame = self.global_frame
        padx, pady = 5, 3

        # 创建字体设置区域
        font_frame = ttk.LabelFrame(frame, text="字体设置")
        font_frame.pack(fill=tk.X, padx=5, pady=5)

        # 字体选择
        ttk.Label(font_frame, text="字体:").grid(row=0, column=0, padx=padx, pady=pady)
        font_options = sorted(tkfont.families())  # 获取系统字体列表
        font_combo = ttk.Combobox(font_frame, textvariable=self.font_family, values=font_options)
        font_combo.grid(row=0, column=1, padx=padx, pady=pady)

        # 字号选择
        ttk.Label(font_frame, text="字号:").grid(row=0, column=2, padx=padx, pady=pady)
        size_spinbox = ttk.Spinbox(font_frame, from_=8, to=20, width=5,
                                 textvariable=self.font_size)
        size_spinbox.grid(row=0, column=3, padx=padx, pady=pady)

        # 创建基本设置区域
        basic_frame = ttk.LabelFrame(frame, text="基本设置")
        basic_frame.pack(fill=tk.X, padx=5, pady=5)

        # 问卷链接
        ttk.Label(basic_frame, text="问卷链接:").grid(row=0, column=0, padx=padx, pady=pady)
        self.url_entry = ttk.Entry(basic_frame, width=60)
        self.url_entry.grid(row=0, column=1, columnspan=3, padx=padx, pady=pady, sticky=tk.EW)
        self.url_entry.insert(0, self.config["url"])

        # 目标份数和微信比例
        ttk.Label(basic_frame, text="目标份数:").grid(row=1, column=0, padx=padx, pady=pady)
        self.target_entry = ttk.Spinbox(basic_frame, from_=1, to=10000, width=10)
        self.target_entry.grid(row=1, column=1, padx=padx, pady=pady)
        self.target_entry.set(self.config["target_num"])

        ttk.Label(basic_frame, text="微信作答比例:").grid(row=1, column=2, padx=padx, pady=pady)
        self.ratio_scale = ttk.Scale(basic_frame, from_=0, to=1, orient=tk.HORIZONTAL)
        self.ratio_scale.grid(row=1, column=3, padx=padx, pady=pady, sticky=tk.EW)
        self.ratio_scale.set(self.config["weixin_ratio"])

        # 时间设置区域
        time_frame = ttk.LabelFrame(frame, text="时间设置")
        time_frame.pack(fill=tk.X, padx=5, pady=5)

        # 作答时间范围
        ttk.Label(time_frame, text="作答时间(秒):").grid(row=0, column=0, padx=padx, pady=pady)
        ttk.Label(time_frame, text="最短").grid(row=0, column=1, padx=padx, pady=pady)
        self.min_duration = ttk.Spinbox(time_frame, from_=5, to=300, width=5)
        self.min_duration.grid(row=0, column=2, padx=padx, pady=pady)
        self.min_duration.set(self.config["min_duration"])

        ttk.Label(time_frame, text="最长").grid(row=0, column=3, padx=padx, pady=pady)
        self.max_duration = ttk.Spinbox(time_frame, from_=5, to=300, width=5)
        self.max_duration.grid(row=0, column=4, padx=padx, pady=pady)
        self.max_duration.set(self.config["max_duration"])

        # 延迟时间设置
        ttk.Label(time_frame, text="提交间隔(秒):").grid(row=1, column=0, padx=padx, pady=pady)
        ttk.Label(time_frame, text="最短").grid(row=1, column=1, padx=padx, pady=pady)
        self.min_delay = ttk.Spinbox(time_frame, from_=0.1, to=10, increment=0.1, width=5)
        self.min_delay.grid(row=1, column=2, padx=padx, pady=pady)
        self.min_delay.set(self.config["min_delay"])

        ttk.Label(time_frame, text="最长").grid(row=1, column=3, padx=padx, pady=pady)
        self.max_delay = ttk.Spinbox(time_frame, from_=0.1, to=10, increment=0.1, width=5)
        self.max_delay.grid(row=1, column=4, padx=padx, pady=pady)
        self.max_delay.set(self.config["max_delay"])

        # 提交延迟
        ttk.Label(time_frame, text="提交等待(秒):").grid(row=2, column=0, padx=padx, pady=pady)
        self.submit_delay = ttk.Spinbox(time_frame, from_=1, to=10, width=5)
        self.submit_delay.grid(row=2, column=1, padx=padx, pady=pady)
        self.submit_delay.set(self.config["submit_delay"])
        # 高级设置区域
        advanced_frame = ttk.LabelFrame(frame, text="高级设置")
        advanced_frame.pack(fill=tk.X, padx=5, pady=5)

        # 浏览器窗口数
        ttk.Label(advanced_frame, text="并发窗口数:").grid(row=0, column=0, padx=padx, pady=pady)
        self.num_threads = ttk.Spinbox(advanced_frame, from_=1, to=10, width=5)
        self.num_threads.grid(row=0, column=1, padx=padx, pady=pady)
        self.num_threads.set(self.config["num_threads"])

        # 代理IP设置
        self.use_ip_var = tk.BooleanVar(value=self.config["use_ip"])
        ttk.Checkbutton(advanced_frame, text="使用代理IP",
                       variable=self.use_ip_var).grid(row=1, column=0, padx=padx, pady=pady)
        ttk.Label(advanced_frame, text="IP API:").grid(row=1, column=1, padx=padx, pady=pady)
        self.ip_entry = ttk.Entry(advanced_frame, width=50)
        self.ip_entry.grid(row=1, column=2, columnspan=2, padx=padx, pady=pady, sticky=tk.EW)
        self.ip_entry.insert(0, self.config["ip_api"])

        # 无头模式
        self.headless_var = tk.BooleanVar(value=self.config["headless"])
        ttk.Checkbutton(advanced_frame, text="无头模式(不显示浏览器)",
                       variable=self.headless_var).grid(row=2, column=0, columnspan=2,
                                                      padx=padx, pady=pady)

        # 操作按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=10)
        ttk.Button(btn_frame, text="解析问卷",
                  command=self.parse_survey).pack(side=tk.LEFT, padx=5)

    def create_question_settings(self):
        """创建题型设置界面"""
        # 创建Notebook来组织不同题型的设置
        self.question_notebook = ttk.Notebook(self.question_frame)
        self.question_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 根据配置创建各种题型的设置页
        if self.config["single_prob"]:
            self.create_single_choice_settings()
        if self.config["multiple_prob"]:
            self.create_multiple_choice_settings()
        if self.config["matrix_prob"]:
            self.create_matrix_settings()
        if self.config["texts"]:
            self.create_text_settings()
        if self.config["multiple_texts"]:
            self.create_multiple_text_settings()
        if self.config["scale_prob"]:
            self.create_scale_settings()
        if self.config["droplist_prob"]:
            self.create_droplist_settings()
        if self.config["reorder_prob"]:
            self.create_reorder_settings()

    def create_single_choice_settings(self):
        """创建单选题设置页面"""
        frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(frame, text=f"单选题({len(self.config['single_prob'])})")

        # 添加全随机按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(btn_frame, text="全部随机",
                  command=lambda: self.set_all_random("single")).pack(side=tk.LEFT)

        # 创建表格头部
        header_frame = ttk.Frame(frame)
        header_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(header_frame, text="题号", width=10).pack(side=tk.LEFT)
        for i in range(10):  # 最多10个选项
            ttk.Label(header_frame, text=f"选项{i+1}", width=8).pack(side=tk.LEFT)

        # 创建题目设置
        content_frame = ttk.Frame(frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        for q_num, probs in self.config["single_prob"].items():
            row_frame = ttk.Frame(content_frame)
            row_frame.pack(fill=tk.X, pady=2)

            ttk.Label(row_frame, text=f"第{q_num}题", width=10).pack(side=tk.LEFT)
            entry_row = []

            # 确定选项数量
            option_count = len(probs) if isinstance(probs, list) else 5
            for i in range(option_count):
                entry = ttk.Entry(row_frame, width=8)
                if probs == -1:
                    entry.insert(0, "-1")
                elif isinstance(probs, list) and i < len(probs):
                    entry.insert(0, str(probs[i]))
                entry.pack(side=tk.LEFT, padx=2)
                entry_row.append(entry)

            self.single_entries.append(entry_row)

    def create_multiple_choice_settings(self):
        """创建多选题设置页面"""
        frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(frame, text=f"多选题({len(self.config['multiple_prob'])})")

        # 创建说明标签
        ttk.Label(frame, text="设置每个选项被选中的概率(0-100)，最小选择数，及其他选项答案").pack(fill=tk.X, padx=5,
                                                                                                pady=5)

        # 创建表格头部
        header_frame = ttk.Frame(frame)
        header_frame.pack(fill=tk.X, padx=5, pady=2)

        headers = ["题号"] + [f"选项{i + 1}" for i in range(10)] + ["最小选择数", "其他选项答案"]
        for header in headers:
            ttk.Label(header_frame, text=header, width=12 if header == "其他选项答案" else 8).pack(side=tk.LEFT, padx=2)

        # 创建内容区域
        content_frame = ttk.Frame(frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        for q_num, probs in self.config["multiple_prob"].items():
            row_frame = ttk.Frame(content_frame)
            row_frame.pack(fill=tk.X, pady=2)

            # 题号
            ttk.Label(row_frame, text=f"第{q_num}题", width=8).pack(side=tk.LEFT)

            # 选项概率输入框
            entry_row = []
            option_count = len(probs) if isinstance(probs, list) else 5
            for i in range(option_count):
                entry = ttk.Entry(row_frame, width=8)
                if isinstance(probs, list) and i < len(probs):
                    entry.insert(0, str(probs[i]))
                else:
                    entry.insert(0, "50")
                entry.pack(side=tk.LEFT, padx=2)
                entry_row.append(entry)
            self.multi_entries.append(entry_row)

            # 最小选择数
            min_selection = ttk.Spinbox(row_frame, from_=1, to=option_count, width=6)
            min_selection.set(self.config["min_selection"].get(q_num, 1))
            min_selection.pack(side=tk.LEFT, padx=2)
            self.min_selection_entries.append(min_selection)

            # 其他选项答案
            other_entry = ttk.Entry(row_frame, width=15)
            if q_num in self.config["multiple_other"]:
                other_entry.insert(0, ", ".join(self.config["multiple_other"][q_num]))
            other_entry.pack(side=tk.LEFT, padx=2)
            self.multi_other_entries.append(other_entry)

    def create_matrix_settings(self):
        """创建矩阵题设置页面"""
        frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(frame, text=f"矩阵题({len(self.config['matrix_prob'])})")

        # 添加全随机按钮和说明
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(btn_frame, text="全部随机",
                   command=lambda: self.set_all_random("matrix")).pack(side=tk.LEFT)
        ttk.Label(btn_frame, text="  设置每行选项的选择概率(-1表示随机选择)").pack(side=tk.LEFT, padx=5)

        # 创建表格头部
        header_frame = ttk.Frame(frame)
        header_frame.pack(fill=tk.X, padx=5, pady=2)
        headers = ["题号"] + [f"选项{i + 1}" for i in range(10)]
        for header in headers:
            ttk.Label(header_frame, text=header, width=8).pack(side=tk.LEFT, padx=2)

        # 创建内容区域
        content_frame = ttk.Frame(frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        for q_num, probs in self.config["matrix_prob"].items():
            row_frame = ttk.Frame(content_frame)
            row_frame.pack(fill=tk.X, pady=2)

            ttk.Label(row_frame, text=f"第{q_num}题", width=8).pack(side=tk.LEFT)

            entry_row = []
            option_count = len(probs) if isinstance(probs, list) else 5
            for i in range(option_count):
                entry = ttk.Entry(row_frame, width=8)
                if probs == -1:
                    entry.insert(0, "-1")
                elif isinstance(probs, list) and i < len(probs):
                    entry.insert(0, str(probs[i]))
                entry.pack(side=tk.LEFT, padx=2)
                entry_row.append(entry)
            self.matrix_entries.append(entry_row)

    def create_text_settings(self):
        """创建填空题设置页面"""
        frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(frame, text=f"填空题({len(self.config['texts'])})")

        # 说明标签
        ttk.Label(frame, text="设置填空题可选答案，多个答案用逗号分隔").pack(fill=tk.X, padx=5, pady=5)

        # 创建内容区域
        content_frame = ttk.Frame(frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        for q_num, texts in self.config["texts"].items():
            row_frame = ttk.Frame(content_frame)
            row_frame.pack(fill=tk.X, pady=2)

            ttk.Label(row_frame, text=f"第{q_num}题", width=8).pack(side=tk.LEFT)

            entry = ttk.Entry(row_frame, width=60)
            if texts:
                entry.insert(0, ", ".join(texts))
            entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            self.text_entries.append(entry)

    def create_scale_settings(self):
        """创建量表题设置页面"""
        frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(frame, text=f"量表题({len(self.config['scale_prob'])})")

        # 说明标签
        ttk.Label(frame, text="设置量表题每个选项的选择概率").pack(fill=tk.X, padx=5, pady=5)

        # 创建表格头部
        header_frame = ttk.Frame(frame)
        header_frame.pack(fill=tk.X, padx=5, pady=2)
        headers = ["题号"] + [f"选项{i + 1}" for i in range(11)]  # 最多11个选项(0-10)
        for header in headers:
            ttk.Label(header_frame, text=header, width=8).pack(side=tk.LEFT, padx=2)

        # 创建内容区域
        content_frame = ttk.Frame(frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        for q_num, probs in self.config["scale_prob"].items():
            row_frame = ttk.Frame(content_frame)
            row_frame.pack(fill=tk.X, pady=2)

            ttk.Label(row_frame, text=f"第{q_num}题", width=8).pack(side=tk.LEFT)

            entry_row = []
            option_count = len(probs)
            for i in range(option_count):
                entry = ttk.Entry(row_frame, width=8)
                entry.insert(0, str(probs[i]))
                entry.pack(side=tk.LEFT, padx=2)
                entry_row.append(entry)
            self.scale_entries.append(entry_row)

    def create_droplist_settings(self):
        """创建下拉框设置页面"""
        frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(frame, text=f"下拉框({len(self.config['droplist_prob'])})")

        # 说明标签
        ttk.Label(frame, text="设置下拉框每个选项的选择概率").pack(fill=tk.X, padx=5, pady=5)

        # 创建表格头部
        header_frame = ttk.Frame(frame)
        header_frame.pack(fill=tk.X, padx=5, pady=2)
        headers = ["题号"] + [f"选项{i + 1}" for i in range(10)]
        for header in headers:
            ttk.Label(header_frame, text=header, width=8).pack(side=tk.LEFT, padx=2)

        # 创建内容区域
        content_frame = ttk.Frame(frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        for q_num, probs in self.config["droplist_prob"].items():
            row_frame = ttk.Frame(content_frame)
            row_frame.pack(fill=tk.X, pady=2)

            ttk.Label(row_frame, text=f"第{q_num}题", width=8).pack(side=tk.LEFT)

            entry_row = []
            for i, prob in enumerate(probs):
                entry = ttk.Entry(row_frame, width=8)
                entry.insert(0, str(prob))
                entry.pack(side=tk.LEFT, padx=2)
                entry_row.append(entry)
            self.droplist_entries.append(entry_row)

    def create_multiple_text_settings(self):
        """创建多项填空设置页面"""
        frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(frame, text=f"多项填空({len(self.config['multiple_texts'])})")

        # 说明标签
        ttk.Label(frame, text="设置每个填空项的可选答案，多个答案用逗号分隔").pack(fill=tk.X, padx=5, pady=5)

        # 创建内容区域
        content_frame = ttk.Frame(frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        row = 0
        for q_num, text_lists in self.config["multiple_texts"].items():
            # 题号标签
            ttk.Label(content_frame, text=f"第{q_num}题").grid(row=row, column=0, padx=5, pady=5)

            entry_col = []
            for i, texts in enumerate(text_lists):
                ttk.Label(content_frame, text=f"填空{i + 1}:").grid(row=row + i, column=1, padx=5, pady=2)
                entry = ttk.Entry(content_frame, width=60)
                entry.insert(0, ", ".join(texts))
                entry.grid(row=row + i, column=2, padx=5, pady=2, sticky=tk.EW)
                entry_col.append(entry)

            self.multiple_text_entries.append(entry_col)
            row += len(text_lists) + 1  # 为每题之间留出空行

    def create_reorder_settings(self):
        """创建排序题设置页面"""
        frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(frame, text=f"排序题({len(self.config['reorder_prob'])})")

        # 说明标签
        ttk.Label(frame, text="设置每个选项出现在各位置的概率(数值和为1)").pack(fill=tk.X, padx=5, pady=5)

        # 创建表格头部
        header_frame = ttk.Frame(frame)
        header_frame.pack(fill=tk.X, padx=5, pady=2)
        headers = ["题号"] + [f"位置{i + 1}" for i in range(10)]
        for header in headers:
            ttk.Label(header_frame, text=header, width=8).pack(side=tk.LEFT, padx=2)

        # 创建内容区域
        content_frame = ttk.Frame(frame)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        for q_num, probs in self.config["reorder_prob"].items():
            row_frame = ttk.Frame(content_frame)
            row_frame.pack(fill=tk.X, pady=2)

            ttk.Label(row_frame, text=f"第{q_num}题", width=8).pack(side=tk.LEFT)

            entry_row = []
            for i, prob in enumerate(probs):
                entry = ttk.Entry(row_frame, width=8)
                entry.insert(0, str(prob))
                entry.pack(side=tk.LEFT, padx=2)
                entry_row.append(entry)
            self.reorder_entries.append(entry_row)

    def create_log_area(self):
        """创建日志区域"""
        # 日志控制按钮
        control_frame = ttk.Frame(self.bottom_frame)
        control_frame.pack(fill=tk.X, padx=5, pady=(5, 0))

        ttk.Button(control_frame, text="清空日志",
                   command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="导出日志",
                   command=self.export_log).pack(side=tk.LEFT, padx=5)

        # 日志文本区域
        self.log_area = scrolledtext.ScrolledText(self.bottom_frame, height=10)
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

            # 验证参数合法性
            if self.config["min_duration"] > self.config["max_duration"]:
                raise ValueError("最短作答时间不能大于最长作答时间")
            if self.config["min_delay"] > self.config["max_delay"]:
                raise ValueError("最小延迟不能大于最大延迟")
            if self.config["target_num"] <= 0:
                raise ValueError("目标份数必须大于0")

            # 保存单选题设置
            new_single_prob = {}
            for i, entries in enumerate(self.single_entries):
                q_num = list(self.config["single_prob"].keys())[i]
                probs = []
                for entry in entries:
                    try:
                        value = float(entry.get())
                        if value == -1:
                            new_single_prob[q_num] = -1
                            break
                        probs.append(value)
                    except ValueError:
                        continue
                if probs and -1 not in probs:
                    new_single_prob[q_num] = probs
            self.config["single_prob"] = new_single_prob

            # 保存多选题设置
            new_multiple_prob = {}
            new_multiple_other = {}
            new_min_selection = {}
            for i, (entries, other_entry, min_entry) in enumerate(zip(
                    self.multi_entries, self.multi_other_entries, self.min_selection_entries)):
                q_num = list(self.config["multiple_prob"].keys())[i]

                # 保存概率设置
                probs = []
                for entry in entries:
                    try:
                        value = float(entry.get())
                        if not 0 <= value <= 100:
                            raise ValueError(f"第{q_num}题的概率必须在0-100之间")
                        probs.append(value)
                    except ValueError as e:
                        raise ValueError(f"第{q_num}题概率设置错误: {str(e)}")
                if probs:
                    new_multiple_prob[q_num] = probs

                # 保存其他选项
                other_text = other_entry.get().strip()
                if other_text:
                    other_answers = [ans.strip() for ans in other_text.split(",")]
                    new_multiple_other[q_num] = other_answers

                # 保存最小选择数
                try:
                    min_selection = int(min_entry.get())
                    if min_selection < 1:
                        raise ValueError(f"第{q_num}题的最小选择数必须大于0")
                    if min_selection > len(probs):
                        raise ValueError(f"第{q_num}题的最小选择数不能大于选项数")
                    new_min_selection[q_num] = min_selection
                except ValueError as e:
                    raise ValueError(f"第{q_num}题最小选择数设置错误: {str(e)}")

            self.config["multiple_prob"] = new_multiple_prob
            self.config["multiple_other"] = new_multiple_other
            self.config["min_selection"] = new_min_selection

            # 保存矩阵题设置
            new_matrix_prob = {}
            for i, entries in enumerate(self.matrix_entries):
                q_num = list(self.config["matrix_prob"].keys())[i]
                probs = []
                for entry in entries:
                    try:
                        value = float(entry.get())
                        if value == -1:
                            new_matrix_prob[q_num] = -1
                            break
                        if value < 0:
                            raise ValueError(f"第{q_num}题的概率不能为负")
                        probs.append(value)
                    except ValueError as e:
                        raise ValueError(f"第{q_num}题概率设置错误: {str(e)}")
                if probs and -1 not in probs:
                    new_matrix_prob[q_num] = probs
            self.config["matrix_prob"] = new_matrix_prob

            # 保存填空题设置
            new_texts = {}
            for i, entry in enumerate(self.text_entries):
                q_num = list(self.config["texts"].keys())[i]
                text = entry.get().strip()
                if text:
                    texts = [t.strip() for t in text.split(",")]
                    new_texts[q_num] = texts
            self.config["texts"] = new_texts

            # 保存多项填空设置
            new_multiple_texts = {}
            for i, entries in enumerate(self.multiple_text_entries):
                q_num = list(self.config["multiple_texts"].keys())[i]
                text_lists = []
                for entry in entries:
                    texts = [t.strip() for t in entry.get().strip().split(",")]
                    if texts and texts[0]:  # 确保不是空列表或只包含空字符串
                        text_lists.append(texts)
                if text_lists:
                    new_multiple_texts[q_num] = text_lists
            self.config["multiple_texts"] = new_multiple_texts

            # 保存量表题设置
            new_scale_prob = {}
            for i, entries in enumerate(self.scale_entries):
                q_num = list(self.config["scale_prob"].keys())[i]
                probs = []
                for entry in entries:
                    try:
                        value = float(entry.get())
                        if value < 0:
                            raise ValueError(f"第{q_num}题的概率不能为负")
                        probs.append(value)
                    except ValueError as e:
                        raise ValueError(f"第{q_num}题概率设置错误: {str(e)}")
                if probs:
                    new_scale_prob[q_num] = probs
            self.config["scale_prob"] = new_scale_prob

            # 保存下拉框设置
            new_droplist_prob = {}
            for i, entries in enumerate(self.droplist_entries):
                q_num = list(self.config["droplist_prob"].keys())[i]
                probs = []
                for entry in entries:
                    try:
                        value = float(entry.get())
                        if value < 0:
                            raise ValueError(f"第{q_num}题的概率不能为负")
                        probs.append(value)
                    except ValueError as e:
                        raise ValueError(f"第{q_num}题概率设置错误: {str(e)}")
                if probs:
                    new_droplist_prob[q_num] = probs
            self.config["droplist_prob"] = new_droplist_prob

            # 保存排序题设置
            new_reorder_prob = {}
            for i, entries in enumerate(self.reorder_entries):
                q_num = list(self.config["reorder_prob"].keys())[i]
                probs = []
                total = 0
                for entry in entries:
                    try:
                        value = float(entry.get())
                        if value < 0:
                            raise ValueError(f"第{q_num}题的概率不能为负")
                        total += value
                        probs.append(value)
                    except ValueError as e:
                        raise ValueError(f"第{q_num}题概率设置错误: {str(e)}")
                if abs(total - 1.0) > 0.01:  # 允许0.01的误差
                    raise ValueError(f"第{q_num}题的概率之和必须为1")
                if probs:
                    new_reorder_prob[q_num] = probs
            self.config["reorder_prob"] = new_reorder_prob

            logging.info("配置保存成功")
            return True

        except Exception as e:
            logging.error(f"保存配置时出错: {str(e)}")
            messagebox.showerror("错误", str(e))
            return False

    def parse_survey(self):
        """解析问卷内容"""
        try:
            url = self.url_entry.get().strip()
            if not url:
                messagebox.showerror("错误", "请输入问卷链接")
                return

            # 创建临时浏览器实例进行解析
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')

            logging.info("开始解析问卷...")
            driver = None
            try:
                driver = webdriver.Chrome(options=options)
                driver.get(url)
                time.sleep(2)  # 等待页面加载

                # 获取所有题目
                questions = driver.find_elements(By.CLASS_NAME, "div_question")
                if not questions:
                    raise Exception("未找到任何题目，请确认问卷链接是否正确")

                # 重置配置
                self.config["single_prob"] = {}
                self.config["multiple_prob"] = {}
                self.config["matrix_prob"] = {}
                self.config["scale_prob"] = {}
                self.config["droplist_prob"] = {}
                self.config["texts"] = {}
                self.config["multiple_texts"] = {}
                self.config["reorder_prob"] = {}
                self.config["min_selection"] = {}
                self.config["multiple_other"] = {}

                # 解析每个题目
                for q in questions:
                    try:
                        q_type = q.get_attribute("type")
                        q_num = q.get_attribute("id").replace("div", "")
                        q_title = q.find_element(By.CLASS_NAME, "div_title_question").text
                        logging.info(f"解析第{q_num}题: {q_title[:30]}...")

                        if q_type == "1":  # 单选题
                            options = q.find_elements(By.CSS_SELECTOR, "input[type='radio']")
                            self.config["single_prob"][q_num] = [-1] * len(options)

                        elif q_type == "2":  # 多选题
                            options = q.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                            self.config["multiple_prob"][q_num] = [50] * len(options)
                            self.config["min_selection"][q_num] = 1

                            # 检查是否有其他选项
                            try:
                                other_input = q.find_element(By.CSS_SELECTOR, "input[type='text']")
                                if other_input:
                                    self.config["multiple_other"][q_num] = [""]
                            except:
                                pass

                        elif q_type == "3":  # 填空题
                            self.config["texts"][q_num] = [""]

                        elif q_type == "4":  # 矩阵题
                            matrix_rows = q.find_elements(By.CLASS_NAME, "matrix-row")
                            if matrix_rows:
                                options_per_row = len(matrix_rows[0].find_elements(
                                    By.CSS_SELECTOR, "input[type='radio']"))
                                self.config["matrix_prob"][q_num] = [-1] * options_per_row

                        elif q_type == "5":  # 量表题
                            options = q.find_elements(By.CSS_SELECTOR, "input[type='radio']")
                            self.config["scale_prob"][q_num] = [1] * len(options)

                        elif q_type == "6":  # 下拉框
                            options = q.find_elements(By.TAG_NAME, "option")
                            if len(options) > 1:  # 跳过第一个默认选项
                                self.config["droplist_prob"][q_num] = [1] * (len(options) - 1)

                        elif q_type == "7":  # 排序题
                            items = q.find_elements(By.CLASS_NAME, "reorder-item")
                            if items:
                                prob = 1.0 / len(items)
                                self.config["reorder_prob"][q_num] = [prob] * len(items)

                        elif q_type == "8":  # 多项填空
                            inputs = q.find_elements(By.TAG_NAME, "input")
                            if inputs:
                                self.config["multiple_texts"][q_num] = [[""] for _ in range(len(inputs))]

                    except Exception as e:
                        logging.error(f"解析第{q_num}题时出错: {str(e)}")
                        continue

                # 更新界面
                self.reset_ui_with_config()
                logging.info("问卷解析完成")
                messagebox.showinfo("成功", "问卷解析完成！")

            except Exception as e:
                raise Exception(f"解析问卷失败: {str(e)}")

            finally:
                if driver:
                    driver.quit()

        except Exception as e:
            logging.error(str(e))
            messagebox.showerror("错误", str(e))

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
            if not file_path:
                return

            with open(file_path, 'r', encoding='utf-8') as f:
                new_config = json.load(f)

            # 验证配置完整性
            required_keys = [
                "url", "target_num", "min_duration", "max_duration",
                "weixin_ratio", "min_delay", "max_delay", "submit_delay"
            ]
            missing_keys = [key for key in required_keys if key not in new_config]
            if missing_keys:
                raise ValueError(f"配置文件缺少必需的字段: {', '.join(missing_keys)}")

            # 验证配置合理性
            if new_config["min_duration"] > new_config["max_duration"]:
                raise ValueError("最短作答时间不能大于最长作答时间")
            if new_config["min_delay"] > new_config["max_delay"]:
                raise ValueError("最小延迟不能大于最大延迟")
            if new_config["target_num"] <= 0:
                raise ValueError("目标份数必须大于0")

            # 更新配置
            self.config = new_config
            self.reset_ui_with_config()
            logging.info(f"配置已从文件导入: {file_path}")
            messagebox.showinfo("成功", "配置导入成功！")

        except json.JSONDecodeError:
            msg = "配置文件格式错误，请确保是有效的JSON文件"
            logging.error(msg)
            messagebox.showerror("错误", msg)
        except Exception as e:
            logging.error(f"导入配置时出错: {str(e)}")
            messagebox.showerror("错误", f"导入配置时出错: {str(e)}")

    def reset_defaults(self):
        """重置为默认配置"""
        if messagebox.askyesno("确认", "确定要重置所有设置为默认值吗？"):
            self.config = DEFAULT_CONFIG.copy()
            self.reset_ui_with_config()
            logging.info("已重置为默认配置")

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
                    raise ValueError("目标份数必须大于0")
            except ValueError as e:
                messagebox.showerror("错误", str(e))
                return

            try:
                self.config["num_threads"] = int(self.num_threads.get())
                if self.config["num_threads"] <= 0:
                    raise ValueError("并发窗口数必须大于0")
            except ValueError as e:
                messagebox.showerror("错误", str(e))
                return

            # 检查代理IP设置
            if self.config["use_ip"] and not self.config["ip_api"]:
                messagebox.showerror("错误", "启用代理IP时必须提供IP API地址")
                return

            # 更新运行状态
            self.running = True
            self.paused = False
            self.cur_num = 0
            self.cur_fail = 0
            self.pause_event.set()

            # 更新按钮状态
            self.start_btn.config(state=tk.DISABLED)
            self.pause_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.NORMAL)

            # 重置进度条
            self.progress_var.set(0)
            self.question_progress_var.set(0)
            self.status_var.set("运行中...")
            self.question_status_var.set("题目进度: 0/0")

            # 启动工作线程
            logging.info(f"开始运行，目标份数: {self.config['target_num']}")
            threads = []
            for i in range(self.config["num_threads"]):
                x = (i % 2) * 600  # 横向位置
                y = (i // 2) * 400  # 纵向位置
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
            self.stop_filling()

    def run_filling(self, x=0, y=0):
        """运行填写任务"""
        options = webdriver.ChromeOptions()
        if self.config["headless"]:
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
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
                        response = requests.get(self.config["ip_api"], timeout=5)
                        ip = response.text.strip()
                        if not ip:
                            raise Exception("获取到的IP地址为空")
                        options.add_argument(f'--proxy-server={ip}')
                        logging.info(f"使用代理IP: {ip}")
                    except Exception as e:
                        logging.error(f"获取代理IP失败: {str(e)}")
                        continue

                # 创建浏览器实例
                try:
                    driver = webdriver.Chrome(options=options)
                    driver.set_page_load_timeout(30)  # 设置页面加载超时
                    driver.set_script_timeout(30)  # 设置脚本执行超时

                    # 访问问卷
                    driver.get(self.config["url"])
                    time.sleep(self.config["page_load_delay"])

                    # 随机决定是否使用微信作答
                    if random.random() < self.config["weixin_ratio"]:
                        try:
                            weixin_btn = driver.find_element(By.CLASS_NAME, "weixin-answer")
                            weixin_btn.click()
                            time.sleep(2)
                            logging.info("使用微信模式作答")
                        except:
                            logging.warning("切换微信模式失败")

                    # 填写问卷
                    self.fill_survey(driver)

                    # 提交问卷
                    try:
                        submit_btn = driver.find_element(By.ID, "submit_button")
                        time.sleep(self.config["submit_delay"])
                        submit_btn.click()
                        time.sleep(2)

                        # 验证提交结果
                        if "完成" in driver.title or "提交成功" in driver.page_source:
                            with self.lock:
                                self.cur_num += 1
                            logging.info(f"第 {self.cur_num} 份问卷提交成功")
                        else:
                            raise Exception("未检测到提交成功标记")

                    except Exception as e:
                        with self.lock:
                            self.cur_fail += 1
                        logging.error(f"提交问卷失败: {str(e)}")

                except Exception as e:
                    with self.lock:
                        self.cur_fail += 1
                    logging.error(f"填写问卷时出错: {str(e)}")

                finally:
                    try:
                        if driver:
                            driver.quit()
                    except:
                        pass

                # 随机等待
                if self.running:
                    delay = random.uniform(
                        self.config["min_delay"],
                        self.config["max_delay"]
                    )
                    time.sleep(delay)

        except Exception as e:
            logging.error(f"运行任务时出错: {str(e)}")
        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass

    def toggle_pause(self):
        """切换暂停/继续状态"""
        self.paused = not self.paused
        if self.paused:
            self.pause_event.clear()
            self.pause_btn.config(text="继续")
            self.status_var.set("已暂停")
            logging.info("已暂停")
        else:
            self.pause_event.set()
            self.pause_btn.config(text="暂停")
            self.status_var.set("运行中...")
            logging.info("已继续")

    def stop_filling(self):
        """停止填写"""
        self.running = False
        self.pause_event.set()  # 确保所有线程都能退出

        # 更新按钮状态
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)

        # 更新状态显示
        self.status_var.set("已停止")
        logging.info(f"已停止，完成数: {self.cur_num}，失败数: {self.cur_fail}")

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

    def set_all_random(self, question_type):
        """将指定题型的所有概率设置为随机"""
        try:
            if question_type == "single":
                for entries in self.single_entries:
                    for entry in entries:
                        entry.delete(0, tk.END)
                        entry.insert(0, "-1")
                logging.info("已将所有单选题设置为随机")

            elif question_type == "matrix":
                for entries in self.matrix_entries:
                    for entry in entries:
                        entry.delete(0, tk.END)
                        entry.insert(0, "-1")
                logging.info("已将所有矩阵题设置为随机")

        except Exception as e:
            logging.error(f"设置随机概率时出错: {str(e)}")
            messagebox.showerror("错误", f"设置随机概率时出错: {str(e)}")

    def update_font(self, *args):
        """更新界面字体"""
        try:
            font_family = self.font_family.get()
            font_size = self.font_size.get()
            font = (font_family, font_size)

            # 更新样式
            style = ttk.Style()
            style.configure(".", font=font)

            # 更新日志区域字体
            self.log_area.configure(font=font)

            def update_widget_fonts(widget):
                """递归更新所有小部件的字体"""
                try:
                    if isinstance(widget, (tk.Label, tk.Entry, ttk.Entry)):
                        widget.configure(font=font)
                    for child in widget.winfo_children():
                        update_widget_fonts(child)
                except:
                    pass

            update_widget_fonts(self.root)
            logging.info(f"字体已更新: {font_family} {font_size}")

        except Exception as e:
            logging.error(f"更新字体时出错: {str(e)}")

    def clear_log(self):
        """清空日志区域"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        logging.info("日志已清空")

    def export_log(self):
        """导出日志到文件"""
        try:
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            default_filename = f"wjx_log_{timestamp}.txt"

            file_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                initialfile=default_filename
            )

            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_area.get(1.0, tk.END))
                logging.info(f"日志已导出到: {file_path}")
                messagebox.showinfo("成功", "日志导出成功！")

        except Exception as e:
            logging.error(f"导出日志时出错: {str(e)}")
            messagebox.showerror("错误", f"导出日志时出错: {str(e)}")

    def on_closing(self):
        """窗口关闭时的处理"""
        if self.running:
            if messagebox.askokcancel("确认", "正在运行中，确定要退出吗？"):
                self.stop_filling()
                self.root.destroy()
        else:
            self.root.destroy()


def setup_logger():
    """配置全局日志设置"""
    # 创建logs目录
    if not os.path.exists('logs'):
        os.makedirs('logs')

    # 设置日志文件名
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join('logs', f'wjx_auto_{timestamp}.log')

    # 配置日志格式
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )


def main():
    """主函数"""
    # 设置日志
    setup_logger()

    try:
        # 创建主窗口
        root = tk.Tk()
        root.title("问卷星自动填写工具 v2.0")

        # 设置窗口图标
        try:
            if os.path.exists("icon.ico"):
                root.iconbitmap("icon.ico")
        except:
            pass

        # 创建应用实例
        app = WJXAutoFillApp(root)

        # 启动应用
        root.mainloop()

    except Exception as e:
        logging.error(f"程序运行时出错: {str(e)}")
        messagebox.showerror("错误", f"程序运行时出错: {str(e)}")


if __name__ == "__main__":
    main()