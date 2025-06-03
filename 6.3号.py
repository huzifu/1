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
from PIL import Image, ImageTk

# ================== 配置参数 ==================
# 默认参数值
DEFAULT_CONFIG = {
    "url": "https://www.wjx.cn/vm/mZ3nVoC.aspx#",
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
    "ip_api": "https://service.ipzan.com/core-extract?num=1&minute=1&pool=quality&secret=YOUR_SECRET",
    "num_threads": 4,

    # 单选题概率配置
    "single_prob": {
        "1": -1,  # -1表示随机选择
        "2": [0.3, 0.7],  # 数组表示每个选项的选择概率
        "3": [0.2, 0.2, 0.6]
    },

    # 多选题概率配置
    "multiple_prob": {
        "4": [0.4, 0.3, 0.3],  # 数组表示每个选项被选中的概率
        "5": [0.5, 0.5, 0.5, 0.5]
    },

    # 矩阵题概率配置
    "matrix_prob": {
        "6": [0.2, 0.3, 0.5],  # 每行选项的选择概率
        "7": -1  # -1表示随机选择
    },

    # 量表题概率配置
    "scale_prob": {
        "8": [0.1, 0.2, 0.4, 0.2, 0.1],  # 每个刻度的选择概率
        "9": [0.2, 0.2, 0.2, 0.2, 0.2]
    },

    # 填空题答案配置
    "texts": {
        "10": ["示例答案1", "示例答案2", "示例答案3"],
        "11": ["回答A", "回答B", "回答C"]
    },

    # 多项填空配置
    "multiple_texts": {
        "12": [
            ["选项1", "选项2", "选项3"],
            ["选项A", "选项B", "选项C"]
        ]
    },

    # 排序题概率配置
    "reorder_prob": {
        "13": [0.4, 0.3, 0.2, 0.1],  # 每个位置的选择概率
        "14": [0.25, 0.25, 0.25, 0.25]
    },

    # 下拉框概率配置
    "droplist_prob": {
        "15": [0.3, 0.4, 0.3],  # 每个选项的选择概率
        "16": [0.5, 0.5]
    },

    # 题目文本存储
    "question_texts": {}
}


# ToolTip类用于显示题目提示
class ToolTip:
    def __init__(self, widget, text='', delay=500, wraplength=400):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.wraplength = wraplength
        self.tip_window = None
        self.id = None
        self.x = self.y = 0

        # 绑定事件
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<Motion>", self.motion)

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def motion(self, event=None):
        self.x, self.y = event.x, event.y
        self.x += self.widget.winfo_rootx() + 25
        self.y += self.widget.winfo_rooty() + 20
        if self.tip_window:
            self.tip_window.geometry(f"+{self.x}+{self.y}")

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.delay, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self):
        if self.tip_window:
            return
        # 创建提示窗口
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{self.x}+{self.y}")
        label = tk.Label(self.tip_window, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         wraplength=self.wraplength)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()


class WJXAutoFillApp:
    def __init__(self, root):
        self.root = root
        self.root.title("问卷星自动填写工具 v2.0")
        self.root.geometry("1200x900")
        self.root.resizable(True, True)

        # 设置应用图标
        try:
            self.root.iconbitmap("wjx_icon.ico")
        except:
            pass

        # 设置现代风格
        self.root.set_theme("equilux")
        style = ttk.Style()
        style.configure('TNotebook.Tab', padding=[10, 5], font=('Arial', 10, 'bold'))
        style.configure('TButton', padding=[6, 3], font=('Arial', 10))
        style.configure('TLabel', padding=[5, 2], font=('Arial', 10))
        style.configure('TEntry', padding=[5, 2], font=('Arial', 10))
        style.configure('Header.TLabel', font=('Arial', 11, 'bold'), foreground="#2c6fbb")

        self.config = DEFAULT_CONFIG.copy()
        self.running = False
        self.paused = False
        self.cur_num = 0
        self.cur_fail = 0
        self.lock = threading.Lock()
        self.pause_event = threading.Event()
        self.tooltips = []
        self.parsing = False

        # 初始化所有题型的输入框列表
        self.single_entries = []
        self.multi_entries = []
        self.matrix_entries = []
        self.text_entries = []
        self.multiple_text_entries = []
        self.reorder_entries = []
        self.droplist_entries = []
        self.scale_entries = []

        # 字体设置
        self.font_family = tk.StringVar()
        self.font_size = tk.IntVar()
        self.font_family.set("Arial")
        self.font_size.set(10)

        # 创建主框架
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 标题栏
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))

        # 添加logo
        try:
            logo_img = Image.open("wjx_logo.png")
            logo_img = logo_img.resize((40, 40), Image.LANCZOS)
            self.logo = ImageTk.PhotoImage(logo_img)
            logo_label = ttk.Label(title_frame, image=self.logo)
            logo_label.pack(side=tk.LEFT, padx=(0, 10))
        except:
            pass

        title_label = ttk.Label(title_frame, text="问卷星自动填写工具", font=("Arial", 16, "bold"))
        title_label.pack(side=tk.LEFT)

        # 创建主面板
        self.main_paned = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True)

        # 上半部分：控制区域和标签页
        self.top_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(self.top_frame, weight=1)

        # 下半部分：日志区域
        self.log_frame = ttk.LabelFrame(self.main_paned, text="运行日志")
        self.main_paned.add(self.log_frame, weight=0)

        # === 添加控制按钮区域（顶部）===
        control_frame = ttk.LabelFrame(self.top_frame, text="控制面板")
        control_frame.pack(fill=tk.X, padx=5, pady=5)

        # 按钮框架
        btn_frame = ttk.Frame(control_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        # 第一行按钮
        self.start_btn = ttk.Button(btn_frame, text="▶ 开始填写", command=self.start_filling, width=12)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.pause_btn = ttk.Button(btn_frame, text="⏸ 暂停", command=self.toggle_pause, state=tk.DISABLED, width=10)
        self.pause_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(btn_frame, text="⏹ 停止", command=self.stop_filling, state=tk.DISABLED, width=10)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        ttk.Separator(btn_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=10, fill=tk.Y)

        self.export_config_btn = ttk.Button(btn_frame, text="📤 导出配置", command=self.export_config, width=12)
        self.export_config_btn.pack(side=tk.LEFT, padx=5)

        self.import_config_btn = ttk.Button(btn_frame, text="📥 导入配置", command=self.import_config, width=12)
        self.import_config_btn.pack(side=tk.LEFT, padx=5)

        # 状态栏
        status_frame = ttk.Frame(control_frame)
        status_frame.pack(fill=tk.X, pady=(5, 0))

        # 状态指示器
        self.status_indicator = ttk.Label(status_frame, text="●", foreground="green", font=("Arial", 14))
        self.status_indicator.pack(side=tk.LEFT, padx=(5, 0))

        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, font=("Arial", 10))
        self.status_label.pack(side=tk.LEFT, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, maximum=100, length=200)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # 题目进度
        self.question_progress_var = tk.DoubleVar()
        self.question_progress_bar = ttk.Progressbar(status_frame,
                                                     variable=self.question_progress_var,
                                                     maximum=100,
                                                     length=150)
        self.question_progress_bar.pack(side=tk.RIGHT, padx=5)

        self.question_status_var = tk.StringVar(value="题目: 0/0")
        self.question_status_label = ttk.Label(status_frame, textvariable=self.question_status_var, width=12)
        self.question_status_label.pack(side=tk.RIGHT, padx=5)

        # 创建标签页
        self.notebook = ttk.Notebook(self.top_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 创建全局设置和题型设置标签页
        self.global_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.global_frame, text="⚙️ 全局设置")
        self.create_global_settings()

        self.question_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.question_frame, text="📝 题型设置")
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
                color_map = {
                    'INFO': 'black',
                    'WARNING': 'orange',
                    'ERROR': 'red',
                    'CRITICAL': 'red'
                }
                color = color_map.get(record.levelname, 'black')

                def append():
                    self.text_widget.configure(state='normal')
                    self.text_widget.tag_config(record.levelname, foreground=color)
                    self.text_widget.insert(tk.END, msg + '\n', record.levelname)
                    self.text_widget.configure(state='disabled')
                    self.text_widget.see(tk.END)

                self.text_widget.after(0, append)

        handler = TextHandler(self.log_area)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s',
                                      datefmt='%H:%M:%S')
        handler.setFormatter(formatter)

        logger = logging.getLogger()
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        logging.info("应用程序已启动")

    def create_global_settings(self):
        """创建全局设置界面"""
        frame = self.global_frame
        padx, pady = 8, 5

        # 创建滚动条
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ======== 字体设置 ========
        font_frame = ttk.LabelFrame(scrollable_frame, text="显示设置")
        font_frame.grid(row=0, column=0, columnspan=2, padx=padx, pady=pady, sticky=tk.EW)

        ttk.Label(font_frame, text="字体选择:").grid(row=0, column=0, padx=padx, pady=pady, sticky=tk.W)
        font_options = tk.font.families()
        font_menu = ttk.Combobox(font_frame, textvariable=self.font_family, values=font_options, width=15)
        font_menu.grid(row=0, column=1, padx=padx, pady=pady, sticky=tk.W)
        font_menu.set("Arial")

        ttk.Label(font_frame, text="字体大小:").grid(row=0, column=2, padx=padx, pady=pady, sticky=tk.W)
        font_size_spinbox = ttk.Spinbox(font_frame, from_=8, to=24, increment=1, textvariable=self.font_size, width=5)
        font_size_spinbox.grid(row=0, column=3, padx=padx, pady=pady, sticky=tk.W)
        font_size_spinbox.set(10)

        # ======== 问卷设置 ========
        survey_frame = ttk.LabelFrame(scrollable_frame, text="问卷设置")
        survey_frame.grid(row=1, column=0, columnspan=2, padx=padx, pady=pady, sticky=tk.EW)

        # 第一列：问卷链接
        ttk.Label(survey_frame, text="问卷链接:").grid(row=0, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.url_entry = ttk.Entry(survey_frame, width=50)  # 减小宽度
        self.url_entry.grid(row=0, column=1, columnspan=3, padx=padx, pady=pady, sticky=tk.EW)  # 跨3列

        # 第二行：目标份数和微信作答比率
        ttk.Label(survey_frame, text="目标份数:").grid(row=1, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.target_entry = ttk.Spinbox(survey_frame, from_=1, to=10000, width=8)  # 减小宽度
        self.target_entry.grid(row=1, column=1, padx=padx, pady=pady, sticky=tk.W)

        ttk.Label(survey_frame, text="微信作答比率:").grid(row=1, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.ratio_scale = ttk.Scale(survey_frame, from_=0, to=1, orient=tk.HORIZONTAL, length=100)  # 减小长度
        self.ratio_scale.grid(row=1, column=3, padx=padx, pady=pady, sticky=tk.EW)
        self.ratio_var = tk.StringVar()
        self.ratio_var.set(f"{self.config['weixin_ratio'] * 100:.0f}%")
        ratio_label = ttk.Label(survey_frame, textvariable=self.ratio_var, width=4)  # 减小宽度
        ratio_label.grid(row=1, column=4, padx=(0, padx), pady=pady, sticky=tk.W)

        # 作答时长
        ttk.Label(survey_frame, text="作答时长(秒):").grid(row=2, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(survey_frame, text="最短:").grid(row=2, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_duration = ttk.Spinbox(survey_frame, from_=5, to=300, width=5)
        self.min_duration.grid(row=2, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.min_duration.set(self.config["min_duration"])

        ttk.Label(survey_frame, text="最长:").grid(row=2, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_duration = ttk.Spinbox(survey_frame, from_=5, to=300, width=5)
        self.max_duration.grid(row=2, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.max_duration.set(self.config["max_duration"])

        # ======== 延迟设置 ========
        delay_frame = ttk.LabelFrame(scrollable_frame, text="延迟设置")
        delay_frame.grid(row=2, column=0, columnspan=2, padx=padx, pady=pady, sticky=tk.EW)

        # 基础延迟
        ttk.Label(delay_frame, text="基础延迟(秒):").grid(row=0, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(delay_frame, text="最小:").grid(row=0, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_delay = ttk.Spinbox(delay_frame, from_=0.1, to=10, increment=0.1, width=5)
        self.min_delay.grid(row=0, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.min_delay.set(self.config["min_delay"])

        ttk.Label(delay_frame, text="最大:").grid(row=0, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_delay = ttk.Spinbox(delay_frame, from_=0.1, to=10, increment=0.1, width=5)
        self.max_delay.grid(row=0, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.max_delay.set(self.config["max_delay"])

        # 题目延迟
        ttk.Label(delay_frame, text="每题延迟(秒):").grid(row=1, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(delay_frame, text="最小:").grid(row=1, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_q_delay = ttk.Spinbox(delay_frame, from_=0.1, to=5, increment=0.1, width=5)
        self.min_q_delay.grid(row=1, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.min_q_delay.set(self.config["per_question_delay"][0])

        ttk.Label(delay_frame, text="最大:").grid(row=1, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_q_delay = ttk.Spinbox(delay_frame, from_=0.1, to=5, increment=0.1, width=5)
        self.max_q_delay.grid(row=1, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.max_q_delay.set(self.config["per_question_delay"][1])

        # 页面延迟
        ttk.Label(delay_frame, text="页面延迟(秒):").grid(row=2, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(delay_frame, text="最小:").grid(row=2, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_p_delay = ttk.Spinbox(delay_frame, from_=0.1, to=10, increment=0.1, width=5)
        self.min_p_delay.grid(row=2, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.min_p_delay.set(self.config["per_page_delay"][0])

        ttk.Label(delay_frame, text="最大:").grid(row=2, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_p_delay = ttk.Spinbox(delay_frame, from_=0.1, to=10, increment=0.1, width=5)
        self.max_p_delay.grid(row=2, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.max_p_delay.set(self.config["per_page_delay"][1])

        # 提交延迟
        ttk.Label(delay_frame, text="提交延迟:").grid(row=3, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.submit_delay = ttk.Spinbox(delay_frame, from_=1, to=10, width=5)
        self.submit_delay.grid(row=3, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.submit_delay.set(self.config["submit_delay"])

        # ======== 高级设置 ========
        advanced_frame = ttk.LabelFrame(scrollable_frame, text="高级设置")
        advanced_frame.grid(row=3, column=0, columnspan=2, padx=padx, pady=pady, sticky=tk.EW)

        # 窗口数量
        ttk.Label(advanced_frame, text="浏览器窗口数量:").grid(row=0, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.num_threads = ttk.Spinbox(advanced_frame, from_=1, to=10, width=5)
        self.num_threads.grid(row=0, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.num_threads.set(self.config["num_threads"])

        # IP设置
        self.use_ip_var = tk.BooleanVar(value=self.config["use_ip"])
        ttk.Checkbutton(advanced_frame, text="使用代理IP", variable=self.use_ip_var).grid(
            row=1, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(advanced_frame, text="IP API:").grid(row=1, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.ip_entry = ttk.Entry(advanced_frame, width=40)
        self.ip_entry.grid(row=1, column=2, columnspan=3, padx=padx, pady=pady, sticky=tk.EW)
        self.ip_entry.insert(0, self.config["ip_api"])

        # 无头模式
        self.headless_var = tk.BooleanVar(value=self.config["headless"])
        ttk.Checkbutton(advanced_frame, text="无头模式(不显示浏览器)", variable=self.headless_var).grid(
            row=2, column=0, padx=padx, pady=pady, sticky=tk.W)

        # ======== 操作按钮 ========
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10, sticky=tk.W)

        # 解析问卷按钮
        self.parse_btn = ttk.Button(button_frame, text="🔍 解析问卷", command=self.parse_survey, width=15)
        self.parse_btn.grid(row=0, column=0, padx=5)

        # 重置默认按钮
        ttk.Button(button_frame, text="🔄 重置默认", command=self.reset_defaults, width=15).grid(row=0, column=1, padx=5)

        # 提示标签
        tip_label = ttk.Label(scrollable_frame, text="提示: 填写前请先解析问卷以获取题目结构", foreground="blue")
        tip_label.grid(row=5, column=0, columnspan=2, pady=(10, 0))

    def create_question_settings(self):
        """创建题型设置界面"""
        # 创建带滚动条的题型设置框架
        canvas = tk.Canvas(self.question_frame)
        scrollbar = ttk.Scrollbar(self.question_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 创建笔记本控件
        self.question_notebook = ttk.Notebook(scrollable_frame)
        self.question_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 初始化所有题型的输入框列表
        self.single_entries = []
        self.multi_entries = []
        self.matrix_entries = []
        self.text_entries = []
        self.multiple_text_entries = []
        self.reorder_entries = []
        self.droplist_entries = []
        self.scale_entries = []

        # 单选题设置
        self.single_frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(self.single_frame, text=f"单选题({len(self.config['single_prob'])})")
        self.create_single_settings(self.single_frame)

        # 多选题设置
        self.multi_frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(self.multi_frame, text=f"多选题({len(self.config['multiple_prob'])})")
        self.create_multi_settings(self.multi_frame)

        # 矩阵题设置
        self.matrix_frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(self.matrix_frame, text=f"矩阵题({len(self.config['matrix_prob'])})")
        self.create_matrix_settings(self.matrix_frame)

        # 填空题设置
        self.text_frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(self.text_frame, text=f"填空题({len(self.config['texts'])})")
        self.create_text_settings(self.text_frame)

        # 多项填空设置
        self.multiple_text_frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(self.multiple_text_frame, text=f"多项填空({len(self.config['multiple_texts'])})")
        self.create_multiple_text_settings(self.multiple_text_frame)

        # 排序题设置
        self.reorder_frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(self.reorder_frame, text=f"排序题({len(self.config['reorder_prob'])})")
        self.create_reorder_settings(self.reorder_frame)

        # 下拉框设置
        self.droplist_frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(self.droplist_frame, text=f"下拉框({len(self.config['droplist_prob'])})")
        self.create_droplist_settings(self.droplist_frame)

        # 量表题设置
        self.scale_frame = ttk.Frame(self.question_notebook)
        self.question_notebook.add(self.scale_frame, text=f"量表题({len(self.config['scale_prob'])})")
        self.create_scale_settings(self.scale_frame)

        # 添加提示
        tip_frame = ttk.Frame(scrollable_frame)
        tip_frame.pack(fill=tk.X, pady=10)

        ttk.Label(tip_frame, text="提示: 鼠标悬停在题号上可查看题目内容", foreground="blue").pack()

    def create_single_settings(self, frame):
        """创建单选题设置"""
        padx, pady = 8, 5
        self.single_entries = []

        # 添加全随机按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(btn_frame, text="全部随机", width=12,
                   command=lambda: self.set_all_random("single", frame)).pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame, text="全部平均", width=12,
                   command=lambda: self.set_all_average("single", frame)).pack(side=tk.LEFT, padx=5)

        # 说明标签
        ttk.Label(frame, text="设置每个选项的概率（-1表示随机选择）", font=("Arial", 9, "italic")).pack(anchor=tk.W,
                                                                                                      padx=padx)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 表头
        headers = ["题号", "题目预览", "选项1", "选项2", "选项3", "选项4", "选项5"]
        for col, header in enumerate(headers):
            ttk.Label(table_frame, text=header, font=("Arial", 9, "bold")).grid(
                row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, probs) in enumerate(self.config["single_prob"].items(), start=1):
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"单选题 {q_num}")
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row_idx, column=0, padx=padx, pady=pady)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            ttk.Label(table_frame, text=preview_text, width=20).grid(row=row_idx, column=1, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            option_count = 5 if probs == -1 else len(probs) if isinstance(probs, list) else 5
            entry_row = []

            for col in range(2, option_count + 2):
                entry = ttk.Entry(table_frame, width=8)
                if probs == -1:
                    entry.insert(0, -1)
                elif isinstance(probs, list) and col - 2 < len(probs):
                    entry.insert(0, probs[col - 2])
                else:
                    entry.insert(0, "")
                entry.grid(row=row_idx, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.single_entries.append(entry_row)

    def create_multi_settings(self, frame):
        """创建多选题设置"""
        padx, pady = 8, 5
        self.multi_entries = []

        # 添加按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(btn_frame, text="全部50%", width=12,
                   command=lambda: self.set_all_value("multiple", frame, 50)).pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame, text="全部30%", width=12,
                   command=lambda: self.set_all_value("multiple", frame, 30)).pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame, text="全部70%", width=12,
                   command=lambda: self.set_all_value("multiple", frame, 70)).pack(side=tk.LEFT, padx=5)

        # 说明标签
        ttk.Label(frame, text="设置每个选项被选择的概率（0-100之间的数值）", font=("Arial", 9, "italic")).pack(
            anchor=tk.W, padx=padx)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 表头
        headers = ["题号", "题目预览", "选项1", "选项2", "选项3", "选项4", "选项5"]
        for col, header in enumerate(headers):
            ttk.Label(table_frame, text=header, font=("Arial", 9, "bold")).grid(
                row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, probs) in enumerate(self.config["multiple_prob"].items(), start=1):
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"多选题 {q_num}")
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row_idx, column=0, padx=padx, pady=pady)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            ttk.Label(table_frame, text=preview_text, width=20).grid(row=row_idx, column=1, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            entry_row = []
            option_count = len(probs) if isinstance(probs, list) else 5

            for col in range(2, option_count + 2):
                entry = ttk.Entry(table_frame, width=8)
                if isinstance(probs, list) and col - 2 < len(probs):
                    entry.insert(0, probs[col - 2])
                else:
                    entry.insert(0, 50)
                entry.grid(row=row_idx, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.multi_entries.append(entry_row)

    def create_matrix_settings(self, frame):
        """创建矩阵题设置界面"""
        padx, pady = 8, 5
        self.matrix_entries = []

        # 添加按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(btn_frame, text="全部随机", width=12,
                   command=lambda: self.set_all_random("matrix", frame)).pack(side=tk.LEFT, padx=5)

        # 说明标签
        ttk.Label(frame, text="设置每行选项的选择概率（-1表示随机选择）", font=("Arial", 9, "italic")).pack(
            anchor=tk.W, padx=padx)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 表头
        headers = ["题号", "题目预览", "行1", "行2", "行3", "行4", "行5"]
        for col, header in enumerate(headers):
            ttk.Label(table_frame, text=header, font=("Arial", 9, "bold")).grid(
                row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, probs) in enumerate(self.config["matrix_prob"].items(), start=1):
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"矩阵题 {q_num}")
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row_idx, column=0, padx=padx, pady=pady)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            ttk.Label(table_frame, text=preview_text, width=20).grid(row=row_idx, column=1, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            entry_row = []
            row_count = len(probs) if isinstance(probs, list) else 5

            for col in range(2, row_count + 2):
                entry = ttk.Entry(table_frame, width=8)
                if isinstance(probs, list) and col - 2 < len(probs):
                    entry.insert(0, probs[col - 2])
                else:
                    entry.insert(0, -1)
                entry.grid(row=row_idx, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.matrix_entries.append(entry_row)

    def create_text_settings(self, frame):
        """创建填空题设置界面"""
        padx, pady = 8, 5
        self.text_entries = []

        # 说明标签
        ttk.Label(frame, text="设置填空题的答案（多个答案用逗号分隔）", font=("Arial", 9, "italic")).pack(
            anchor=tk.W, padx=padx)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 表头
        headers = ["题号", "题目预览", "答案选项"]
        for col, header in enumerate(headers):
            ttk.Label(table_frame, text=header, font=("Arial", 9, "bold")).grid(
                row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, answers) in enumerate(self.config["texts"].items(), start=1):
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"填空题 {q_num}")
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row_idx, column=0, padx=padx, pady=pady)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            ttk.Label(table_frame, text=preview_text, width=20).grid(row=row_idx, column=1, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            # 答案输入框
            answer_text = ",".join(answers)
            entry = ttk.Entry(table_frame, width=40)
            entry.insert(0, answer_text)
            entry.grid(row=row_idx, column=2, padx=padx, pady=pady, sticky=tk.EW)
            self.text_entries.append(entry)

    def create_multiple_text_settings(self, frame):
        """创建多项填空设置界面"""
        padx, pady = 8, 5
        self.multiple_text_entries = []

        # 说明标签
        ttk.Label(frame, text="设置多项填空的答案（每个输入框的答案用逗号分隔）", font=("Arial", 9, "italic")).pack(
            anchor=tk.W, padx=padx)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 表头
        headers = ["题号", "题目预览", "输入框1", "输入框2", "输入框3", "输入框4"]
        for col, header in enumerate(headers):
            ttk.Label(table_frame, text=header, font=("Arial", 9, "bold")).grid(
                row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, answers_list) in enumerate(self.config["multiple_texts"].items(), start=1):
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"多项填空 {q_num}")
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row_idx, column=0, padx=padx, pady=pady)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            ttk.Label(table_frame, text=preview_text, width=20).grid(row=row_idx, column=1, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            entry_row = []
            for col_idx, answers in enumerate(answers_list, start=2):
                if col_idx < len(headers):  # 确保不超过表头列数
                    answer_text = ",".join(answers)
                    entry = ttk.Entry(table_frame, width=20)
                    entry.insert(0, answer_text)
                    entry.grid(row=row_idx, column=col_idx, padx=padx, pady=pady)
                    entry_row.append(entry)
            self.multiple_text_entries.append(entry_row)

    def create_reorder_settings(self, frame):
        """创建排序题设置界面"""
        padx, pady = 8, 5
        self.reorder_entries = []

        # 说明标签
        ttk.Label(frame, text="设置每个位置的选择概率", font=("Arial", 9, "italic")).pack(
            anchor=tk.W, padx=padx)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 表头
        headers = ["题号", "题目预览", "位置1", "位置2", "位置3", "位置4", "位置5"]
        for col, header in enumerate(headers):
            ttk.Label(table_frame, text=header, font=("Arial", 9, "bold")).grid(
                row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, probs) in enumerate(self.config["reorder_prob"].items(), start=1):
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"排序题 {q_num}")
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row_idx, column=0, padx=padx, pady=pady)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            ttk.Label(table_frame, text=preview_text, width=20).grid(row=row_idx, column=1, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            entry_row = []
            position_count = len(probs) if isinstance(probs, list) else 5

            for col in range(2, position_count + 2):
                entry = ttk.Entry(table_frame, width=8)
                if isinstance(probs, list) and col - 2 < len(probs):
                    entry.insert(0, probs[col - 2])
                else:
                    entry.insert(0, 0.2)  # 默认概率
                entry.grid(row=row_idx, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.reorder_entries.append(entry_row)

    def create_droplist_settings(self, frame):
        """创建下拉框设置界面"""
        padx, pady = 8, 5
        self.droplist_entries = []

        # 说明标签
        ttk.Label(frame, text="设置每个选项的选择概率", font=("Arial", 9, "italic")).pack(
            anchor=tk.W, padx=padx)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 表头
        headers = ["题号", "题目预览", "选项1", "选项2", "选项3", "选项4", "选项5"]
        for col, header in enumerate(headers):
            ttk.Label(table_frame, text=header, font=("Arial", 9, "bold")).grid(
                row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, probs) in enumerate(self.config["droplist_prob"].items(), start=1):
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"下拉框题 {q_num}")
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row_idx, column=0, padx=padx, pady=pady)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            ttk.Label(table_frame, text=preview_text, width=20).grid(row=row_idx, column=1, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            entry_row = []
            option_count = len(probs) if isinstance(probs, list) else 5

            for col in range(2, option_count + 2):
                entry = ttk.Entry(table_frame, width=8)
                if isinstance(probs, list) and col - 2 < len(probs):
                    entry.insert(0, probs[col - 2])
                else:
                    entry.insert(0, 0.3)  # 默认概率
                entry.grid(row=row_idx, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.droplist_entries.append(entry_row)

    def create_scale_settings(self, frame):
        """创建量表题设置界面"""
        padx, pady = 8, 5
        self.scale_entries = []

        # 说明标签
        ttk.Label(frame, text="设置每个刻度的选择概率", font=("Arial", 9, "italic")).pack(
            anchor=tk.W, padx=padx)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 表头
        headers = ["题号", "题目预览", "刻度1", "刻度2", "刻度3", "刻度4", "刻度5"]
        for col, header in enumerate(headers):
            ttk.Label(table_frame, text=header, font=("Arial", 9, "bold")).grid(
                row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, probs) in enumerate(self.config["scale_prob"].items(), start=1):
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"量表题 {q_num}")
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row_idx, column=0, padx=padx, pady=pady)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            ttk.Label(table_frame, text=preview_text, width=20).grid(row=row_idx, column=1, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            entry_row = []
            scale_count = len(probs) if isinstance(probs, list) else 5

            for col in range(2, scale_count + 2):
                entry = ttk.Entry(table_frame, width=8)
                if isinstance(probs, list) and col - 2 < len(probs):
                    entry.insert(0, probs[col - 2])
                else:
                    entry.insert(0, 0.2)  # 默认概率
                entry.grid(row=row_idx, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.scale_entries.append(entry_row)

    def set_all_random(self, q_type, frame):
        """设置当前题型所有题目为随机"""
        if q_type == "single":
            for entry_row in self.single_entries:
                for entry in entry_row:
                    entry.delete(0, tk.END)
                    entry.insert(0, "-1")
            messagebox.showinfo("成功", "所有单选题已设置为随机选择")

        elif q_type == "matrix":
            for entry_row in self.matrix_entries:
                for entry in entry_row:
                    entry.delete(0, tk.END)
                    entry.insert(0, "-1")
            messagebox.showinfo("成功", "所有矩阵题已设置为随机选择")

    def set_all_average(self, q_type, frame):
        """设置当前题型所有题目为平均概率"""
        if q_type == "single":
            for entry_row in self.single_entries:
                option_count = len(entry_row)
                avg_prob = 1.0 / option_count
                for i, entry in enumerate(entry_row):
                    entry.delete(0, tk.END)
                    entry.insert(0, f"{avg_prob:.2f}")
            messagebox.showinfo("成功", "所有单选题已设置为平均概率")

    def set_all_value(self, q_type, frame, value):
        """设置当前题型所有题目为指定值"""
        if q_type == "multiple":
            for entry_row in self.multi_entries:
                for entry in entry_row:
                    entry.delete(0, tk.END)
                    entry.insert(0, str(value))
            messagebox.showinfo("成功", f"所有多选题已设置为{value}%概率")

    def clear_log(self):
        """清空日志"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        logging.info("日志已清空")

    def export_log(self):
        """导出日志到文件"""
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
            try:
                font_size = int(self.font_size.get())
            except (ValueError, TypeError):
                font_size = 10
                self.font_size.set(10)

            # 确保字体族名称有效
            if not font_family or font_family not in tk.font.families():
                font_family = "Arial"
                self.font_family.set(font_family)

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
            self.font_family.set("Arial")
            self.font_size.set(10)

    def parse_survey(self):
        """解析问卷结构并生成配置模板"""
        if self.parsing:
            messagebox.showwarning("警告", "正在解析问卷，请稍候...")
            return

        self.parsing = True
        self.parse_btn.config(state=tk.DISABLED, text="解析中...")
        self.status_var.set("正在解析问卷...")
        self.status_indicator.config(foreground="orange")

        # 在新线程中执行解析
        threading.Thread(target=self._parse_survey_thread).start()

    def _parse_survey_thread(self):
        """解析问卷的线程函数"""
        try:
            url = self.url_entry.get().strip()
            if not url:
                messagebox.showerror("错误", "请输入问卷链接")
                return

            # 创建临时浏览器实例
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)
            options.add_argument('--disable-blink-features=AutomationControlled')

            driver = webdriver.Chrome(options=options)
            driver.implicitly_wait(10)

            try:
                logging.info(f"正在访问问卷: {url}")
                driver.get(url)

                # 显示解析进度
                self.question_progress_var.set(10)
                self.question_status_var.set("加载问卷...")

                # 等待页面加载
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".field.ui-field-contain, .div_question"))
                )

                # 滚动页面确保所有题目加载
                self.question_progress_var.set(30)
                self.question_status_var.set("加载题目...")
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
                driver.execute_script("window.scrollTo(0, 0);")
                time.sleep(1)

                # 获取所有题目
                self.question_progress_var.set(50)
                self.question_status_var.set("解析题目...")
                questions = driver.find_elements(By.CSS_SELECTOR, ".field.ui-field-contain, .div_question")

                if not questions:
                    # 尝试备用选择器
                    questions = driver.find_elements(By.CSS_SELECTOR, ".div_question")
                    if not questions:
                        raise Exception("未找到任何题目，请检查问卷链接是否正确")

                logging.info(f"发现 {len(questions)} 个问题")

                # 重置配置
                parsed_config = {
                    "single_prob": {},
                    "multiple_prob": {},
                    "matrix_prob": {},
                    "droplist_prob": {},
                    "scale_prob": {},
                    "texts": {},
                    "multiple_texts": {},
                    "reorder_prob": {},
                    "question_texts": {}
                }

                # 遍历题目
                for i, q in enumerate(questions):
                    # 更新进度
                    progress = 50 + (i / len(questions)) * 50
                    self.question_progress_var.set(progress)
                    self.question_status_var.set(f"解析题目 {i + 1}/{len(questions)}")

                    try:
                        # 获取题目ID
                        q_id = q.get_attribute("id").replace("div", "")
                        if not q_id:
                            q_id = str(i + 1)

                        # 获取题目文本
                        q_title = ""
                        try:
                            title_elements = q.find_elements(By.CSS_SELECTOR,
                                                             ".div_title_question, .div_topic_question, .topic_title")
                            if title_elements:
                                q_title = title_elements[0].text.strip()
                        except:
                            pass

                        if not q_title:
                            q_title = f"题目{q_id}"

                        parsed_config["question_texts"][q_id] = q_title

                        # 获取题目类型
                        q_type = None

                        # 检测单选题
                        if q.find_elements(By.CSS_SELECTOR, ".ui-radio"):
                            q_type = "3"
                        # 检测多选题
                        elif q.find_elements(By.CSS_SELECTOR, ".ui-checkbox"):
                            q_type = "4"
                        # 检测填空题
                        elif q.find_elements(By.CSS_SELECTOR, "input[type='text']"):
                            q_type = "1"
                        # 检测多行填空题
                        elif q.find_elements(By.CSS_SELECTOR, "textarea"):
                            q_type = "2"
                        # 检测量表题
                        elif q.find_elements(By.CSS_SELECTOR, ".scale-ul"):
                            q_type = "5"
                        # 检测矩阵题
                        elif q.find_elements(By.CSS_SELECTOR, ".matrix"):
                            q_type = "6"
                        # 检测下拉框
                        elif q.find_elements(By.CSS_SELECTOR, "select"):
                            q_type = "7"
                        # 检测排序题
                        elif q.find_elements(By.CSS_SELECTOR, ".sort-ul"):
                            q_type = "11"

                        if not q_type:
                            logging.warning(f"无法确定题目 {q_id} 的类型，跳过")
                            continue

                        logging.info(f"解析第{q_id}题 - 类型:{q_type} - {q_title}")

                        # 根据题型生成配置
                        if q_type in ["1", "2"]:  # 填空题
                            inputs = q.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
                            if len(inputs) > 1:
                                parsed_config["multiple_texts"][q_id] = [["示例答案"] for _ in range(len(inputs))]
                            else:
                                parsed_config["texts"][q_id] = ["示例答案"]

                        elif q_type == "3":  # 单选题
                            options = q.find_elements(By.CSS_SELECTOR, ".ui-radio")
                            if options:
                                parsed_config["single_prob"][q_id] = [-1] * len(options)

                        elif q_type == "4":  # 多选题
                            options = q.find_elements(By.CSS_SELECTOR, ".ui-checkbox")
                            if options:
                                parsed_config["multiple_prob"][q_id] = [50] * len(options)

                        elif q_type == "5":  # 量表题
                            options = q.find_elements(By.CSS_SELECTOR, ".scale-ul li")
                            if options:
                                parsed_config["scale_prob"][q_id] = [1] * len(options)

                        elif q_type == "6":  # 矩阵题
                            rows = q.find_elements(By.CSS_SELECTOR, ".matrix tr")
                            if rows and len(rows) > 1:  # 确保有题目行
                                parsed_config["matrix_prob"][q_id] = [-1] * (len(rows) - 1)  # 减去表头行

                        elif q_type == "7":  # 下拉框
                            try:
                                select = q.find_element(By.TAG_NAME, "select")
                                options = select.find_elements(By.TAG_NAME, "option")
                                if len(options) > 1:  # 排除第一个默认选项
                                    parsed_config["droplist_prob"][q_id] = [1] * (len(options) - 1)
                            except:
                                pass

                        elif q_type == "11":  # 排序题
                            items = q.find_elements(By.CSS_SELECTOR, ".sort-ul li")
                            if items:
                                item_count = len(items)
                                parsed_config["reorder_prob"][q_id] = [1.0 / item_count] * item_count

                    except Exception as e:
                        logging.warning(f"解析题目 {q_id} 时出错: {str(e)}")
                        continue

                # 更新配置
                self.config.update(parsed_config)

                # 重新加载题型设置界面
                self.reload_question_settings()

                # 检查解析结果
                total_questions = (len(parsed_config['single_prob']) +
                                   len(parsed_config['multiple_prob']) +
                                   len(parsed_config['matrix_prob']) +
                                   len(parsed_config['texts']) +
                                   len(parsed_config['multiple_texts']) +
                                   len(parsed_config['reorder_prob']) +
                                   len(parsed_config['droplist_prob']) +
                                   len(parsed_config['scale_prob']))

                if total_questions == 0:
                    logging.warning("解析结束，但未发现任何题目")
                    messagebox.showwarning("警告", "未能解析到任何题目，请检查问卷链接或尝试手动设置题型")
                else:
                    logging.info(f"问卷解析完成，共发现{total_questions}题")
                    messagebox.showinfo("成功", f"问卷解析完成，共发现{total_questions}题！")

            finally:
                driver.quit()
                self.parsing = False
                self.parse_btn.config(state=tk.NORMAL, text="🔍 解析问卷")
                self.status_var.set("就绪")
                self.status_indicator.config(foreground="green")
                self.question_progress_var.set(0)
                self.question_status_var.set("题目: 0/0")

        except Exception as e:
            logging.error(f"解析问卷时出错: {str(e)}")
            messagebox.showerror("错误", f"解析问卷时出错: {str(e)}")
            self.parsing = False
            self.parse_btn.config(state=tk.NORMAL, text="🔍 解析问卷")
            self.status_var.set("解析失败")
            self.status_indicator.config(foreground="red")
            self.question_progress_var.set(0)
            self.question_status_var.set("题目: 0/0")

    def reload_question_settings(self):
        """重新加载题型设置界面"""
        # 清除当前题型设置
        for widget in self.question_frame.winfo_children():
            widget.destroy()

        # 清除旧的Tooltip引用
        self.tooltips = []

        # 重新创建题型设置
        self.create_question_settings()

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

            # 更新运行状态
            self.running = True
            self.paused = False
            self.cur_num = 0
            self.cur_fail = 0
            self.pause_event.clear()

            # 更新按钮状态
            self.start_btn.config(state=tk.DISABLED, text="▶ 运行中")
            self.pause_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.NORMAL)
            self.status_indicator.config(foreground="green")

            # 设置进度条初始值
            self.progress_var.set(0)
            self.question_progress_var.set(0)
            self.status_var.set("运行中...")

            # 创建并启动线程
            self.threads = []
            for i in range(self.config["num_threads"]):
                x = (i % 2) * 600
                y = (i // 2) * 400
                t = threading.Thread(target=self.run_filling, args=(x, y))
                t.daemon = True
                t.start()
                self.threads.append(t)

            # 启动进度更新线程
            progress_thread = threading.Thread(target=self.update_progress)
            progress_thread.daemon = True
            progress_thread.start()

        except Exception as e:
            logging.error(f"启动失败: {str(e)}")
            messagebox.showerror("错误", f"启动失败: {str(e)}")

    # 以下方法保持不变，为了简洁省略...
    # run_filling, fill_survey, auto_detect_question_type, handle_captcha,
    # fill_text, fill_single, fill_multiple, fill_matrix, fill_scale,
    # fill_droplist, fill_reorder, update_progress, toggle_pause, stop_filling,
    # random_delay, export_config, import_config, reset_defaults, save_config, on_closing




    def run_filling(self, x=0, y=0):
        """运行填写任务"""
        options = webdriver.ChromeOptions()
        if self.config["headless"]:
            options.add_argument('--headless')
        else:
            options.add_argument(f'--window-position={x},{y}')

        # 添加反检测参数
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_argument('--disable-blink-features=AutomationControlled')
        # 添加其他必要参数
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        driver = None
        try:
            while self.running and self.cur_num < self.config["target_num"]:
                # 检查是否暂停
                if self.paused:
                    time.sleep(1)
                    continue

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
                        try:
                            weixin_btn = driver.find_element(By.CLASS_NAME, "weixin-answer")
                            weixin_btn.click()
                            time.sleep(2)
                        except:
                            pass

                    # 填写问卷
                    if self.fill_survey(driver):
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
        try:
            # 使用更通用的题目选择器
            questions = driver.find_elements(By.CSS_SELECTOR, ".field.ui-field-contain, .div_question")
            if not questions:
                logging.warning("未找到任何题目，可能页面加载失败")
                return False

            total_questions = len(questions)

            # 随机总作答时间
            total_time = random.randint(self.config["min_duration"], self.config["max_duration"])
            start_time = time.time()

            for i, q in enumerate(questions):
                if not self.running:
                    break

                try:
                    q_type = q.get_attribute("type")
                    q_id = q.get_attribute("id")
                    if q_id:
                        q_num = q_id.replace("div", "")
                    else:
                        q_num = str(i + 1)

                    # 更新题目进度
                    self.question_progress_var.set((i + 1) / total_questions * 100)
                    self.question_status_var.set(f"题目进度: {i + 1}/{total_questions}")

                    # 随机等待时间
                    per_question_delay = random.uniform(*self.config["per_question_delay"])
                    time.sleep(per_question_delay)

                    # 根据题型填写
                    if q_type == "1":  # 填空题
                        self.fill_text(q, q_num)
                    elif q_type == "2":  # 填空题（多行）
                        self.fill_text(q, q_num)
                    elif q_type == "3":  # 单选题
                        self.fill_single(driver, q_num)
                    elif q_type == "4":  # 多选题
                        self.fill_multiple(driver, q_num)
                    elif q_type == "5":  # 量表题
                        self.fill_scale(driver, q_num)
                    elif q_type == "6":  # 矩阵题
                        self.fill_matrix(driver, q_num)
                    elif q_type == "7":  # 下拉框
                        self.fill_droplist(driver, q_num)
                    elif q_type == "11":  # 排序题
                        self.fill_reorder(driver, q_num)
                    else:
                        # 尝试自动检测题型
                        self.auto_detect_question_type(driver, q, q_num)

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

            # 提交问卷
            try:
                submit_btn = driver.find_element(By.ID, "submit_button")
                time.sleep(self.config["submit_delay"])
                submit_btn.click()
                time.sleep(2)

                # 检查提交是否成功
                if "完成" in driver.title or "提交成功" in driver.page_source:
                    return True
                else:
                    # 尝试处理可能的验证码
                    if self.handle_captcha(driver):
                        submit_btn.click()
                        time.sleep(2)
                        return "完成" in driver.title or "提交成功" in driver.page_source
                    return False
            except:
                return False
        except Exception as e:
            logging.error(f"填写问卷过程中出错: {str(e)}")
            return False

    def auto_detect_question_type(self, driver, question, q_num):
        """自动检测题型并填写"""
        try:
            # 尝试检测单选题
            radio_btns = question.find_elements(By.CSS_SELECTOR, ".ui-radio")
            if radio_btns:
                self.fill_single(driver, q_num)
                return

            # 尝试检测多选题
            checkboxes = question.find_elements(By.CSS_SELECTOR, ".ui-checkbox")
            if checkboxes:
                self.fill_multiple(driver, q_num)
                return

            # 尝试检测填空题
            text_inputs = question.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
            if text_inputs:
                self.fill_text(question, q_num)
                return

            # 尝试检测量表题
            scale_items = question.find_elements(By.CSS_SELECTOR, ".scale-ul li")
            if scale_items:
                self.fill_scale(driver, q_num)
                return

            # 尝试检测矩阵题
            matrix_rows = question.find_elements(By.CSS_SELECTOR, ".matrix tr")
            if matrix_rows:
                self.fill_matrix(driver, q_num)
                return

            # 尝试检测下拉框
            dropdowns = question.find_elements(By.CSS_SELECTOR, "select")
            if dropdowns:
                self.fill_droplist(driver, q_num)
                return

            # 尝试检测排序题
            sort_items = question.find_elements(By.CSS_SELECTOR, ".sort-ul li")
            if sort_items:
                self.fill_reorder(driver, q_num)
                return

            logging.warning(f"无法自动检测题目 {q_num} 的类型，跳过")
        except Exception as e:
            logging.error(f"自动检测题目类型时出错: {str(e)}")

    def handle_captcha(self, driver):
        """处理验证码（基础实现）"""
        try:
            # 检查是否有验证码
            captcha = driver.find_element(By.ID, "captcha")
            if captcha.is_displayed():
                logging.warning("检测到验证码，尝试人工处理")
                # 这里可以添加更复杂的验证码处理逻辑
                # 目前只是暂停程序并提示用户
                self.paused = True
                self.pause_btn.config(text="继续")
                messagebox.showwarning("验证码", "检测到验证码，请手动处理并点击继续")
                return True
        except:
            pass
        return False

    def fill_text(self, question, q_num):
        """填写填空题"""
        try:
            q_key = str(q_num)
            if q_key in self.config["texts"]:
                answers = self.config["texts"][q_key]
                if answers:
                    # 随机选择一个答案
                    content = random.choice(answers)
                    # 查找输入框
                    input_elem = question.find_element(By.CSS_SELECTOR, f"#q{q_num}")
                    # 使用JavaScript设置值，避免输入事件问题
                    driver = question.parent
                    driver.execute_script(f"arguments[0].value = '{content}';", input_elem)
                    # 触发change事件
                    driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", input_elem)
            else:
                # 默认答案
                input_elem = question.find_element(By.CSS_SELECTOR, f"#q{q_num}")
                driver = question.parent
                driver.execute_script("arguments[0].value = '已填写';", input_elem)
                driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", input_elem)

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写填空题 {q_num} 时出错: {str(e)}")

    def fill_single(self, driver, q_num):
        """填写单选题"""
        try:
            options = driver.find_elements(By.CSS_SELECTOR, f"#div{q_num} .ui-radio")
            if not options:
                return

            q_key = str(q_num)
            probs = self.config["single_prob"].get(q_key, -1)

            if probs == -1:  # 随机选择
                selected = random.choice(options)
            elif isinstance(probs, list):  # 按概率选择
                # 确保概率列表长度匹配
                probs = probs[:len(options)] if len(probs) > len(options) else probs + [0] * (len(options) - len(probs))
                # 归一化概率
                total = sum(probs)
                if total > 0:
                    probs = [p / total for p in probs]
                    selected = np.random.choice(options, p=probs)
                else:
                    selected = random.choice(options)
            else:  # 默认随机
                selected = random.choice(options)

            try:
                selected.click()
            except:
                # 如果直接点击失败，使用JavaScript点击
                driver.execute_script("arguments[0].click();", selected)

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写单选题 {q_num} 时出错: {str(e)}")

    def fill_multiple(self, driver, q_num):
        """填写多选题"""
        try:
            options = driver.find_elements(By.CSS_SELECTOR, f"#div{q_num} .ui-checkbox")
            if not options:
                return

            q_key = str(q_num)
            probs = self.config["multiple_prob"].get(q_key, [50] * len(options))

            # 确保概率列表长度匹配
            probs = probs[:len(options)] if len(probs) > len(options) else probs + [50] * (len(options) - len(probs))

            selected_indices = []
            for i, prob in enumerate(probs):
                if random.random() * 100 < prob:
                    selected_indices.append(i)

            # 确保至少选择一个选项
            if not selected_indices:
                selected_indices = [random.randint(0, len(options) - 1)]

            for idx in selected_indices:
                try:
                    options[idx].click()
                except:
                    # 使用JavaScript点击
                    driver.execute_script("arguments[0].click();", options[idx])

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写多选题 {q_num} 时出错: {str(e)}")

    def fill_matrix(self, driver, q_num):
        """填写矩阵题"""
        try:
            rows = driver.find_elements(By.CSS_SELECTOR, f"#divRefTab{q_num} tbody tr")
            if not rows:
                return

            q_key = str(q_num)
            probs = self.config["matrix_prob"].get(q_key, -1)

            for i, row in enumerate(rows[1:], 1):  # 跳过表头行
                cols = row.find_elements(By.CSS_SELECTOR, "td")
                if not cols:
                    continue

                if probs == -1:  # 随机选择
                    selected_col = random.randint(1, len(cols) - 1)
                elif isinstance(probs, list):  # 按概率选择
                    # 确保概率列表长度匹配
                    col_probs = probs[:len(cols) - 1] if len(probs) > len(cols) - 1 else probs + [0] * (
                            len(cols) - 1 - len(probs))
                    # 归一化概率
                    total = sum(col_probs)
                    if total > 0:
                        col_probs = [p / total for p in col_probs]
                        selected_col = np.random.choice(range(1, len(cols)), p=col_probs)
                    else:
                        selected_col = random.randint(1, len(cols) - 1)
                else:  # 默认随机
                    selected_col = random.randint(1, len(cols) - 1)

                try:
                    cols[selected_col].click()
                except:
                    # 使用JavaScript点击
                    driver.execute_script("arguments[0].click();", cols[selected_col])

                self.random_delay(0.2, 0.5)  # 每行选择后短暂延迟

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写矩阵题 {q_num} 时出错: {str(e)}")

    def fill_scale(self, driver, q_num):
        """填写量表题"""
        try:
            options = driver.find_elements(By.CSS_SELECTOR, f"#div{q_num} .scale-ul li")
            if not options:
                return

            q_key = str(q_num)
            probs = self.config["scale_prob"].get(q_key, [1] * len(options))

            # 确保概率列表长度匹配
            probs = probs[:len(options)] if len(probs) > len(options) else probs + [1] * (len(options) - len(probs))

            # 归一化概率
            total = sum(probs)
            if total > 0:
                probs = [p / total for p in probs]
                selected = np.random.choice(options, p=probs)
            else:
                selected = random.choice(options)

            try:
                selected.click()
            except:
                driver.execute_script("arguments[0].click();", selected)

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写量表题 {q_num} 时出错: {str(e)}")

    def fill_droplist(self, driver, q_num):
        """填写下拉框题"""
        try:
            # 点击下拉框唤出选项
            dropdown = driver.find_element(By.CSS_SELECTOR, f"#select2-q{q_num}-container")
            dropdown.click()
            time.sleep(0.5)

            # 获取所有选项
            options = driver.find_elements(By.CSS_SELECTOR, f"#select2-q{q_num}-results li")
            if not options:
                return

            q_key = str(q_num)
            probs = self.config["droplist_prob"].get(q_key, [1] * len(options))

            # 确保概率列表长度匹配
            probs = probs[:len(options)] if len(probs) > len(options) else probs + [1] * (len(options) - len(probs))

            # 归一化概率并选择
            total = sum(probs)
            if total > 0:
                probs = [p / total for p in probs]
                selected = np.random.choice(options, p=probs)
            else:
                selected = random.choice(options)

            try:
                selected.click()
            except:
                driver.execute_script("arguments[0].click();", selected)

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写下拉框题 {q_num} 时出错: {str(e)}")

    def fill_reorder(self, driver, q_num):
        """填写排序题"""
        try:
            items = driver.find_elements(By.CSS_SELECTOR, f"#div{q_num} .sort-ul li")
            if not items:
                return

            q_key = str(q_num)
            probs = self.config["reorder_prob"].get(q_key, [1 / len(items)] * len(items))

            # 确保概率列表长度匹配
            probs = probs[:len(items)] if len(probs) > len(items) else probs + [1 / len(items)] * (
                    len(items) - len(probs))

            # 根据概率生成顺序
            order = list(range(len(items)))
            if sum(probs) > 0:
                # 归一化概率
                probs = [p / sum(probs) for p in probs]
                # 使用概率进行排序
                np.random.shuffle(order)  # 先随机打乱
                order.sort(key=lambda x: random.random() * probs[x])  # 按概率权重排序
            else:
                random.shuffle(order)  # 完全随机排序

            # 使用JavaScript移动元素
            script_template = """
            var item = arguments[0];
            var targetY = arguments[1];
            var rect = item.getBoundingClientRect();
            var evt = new MouseEvent('mousedown', {
                bubbles: true,
                clientX: rect.left,
                clientY: rect.top
            });
            item.dispatchEvent(evt);

            evt = new MouseEvent('mousemove', {
                bubbles: true,
                clientX: rect.left,
                clientY: targetY
            });
            item.dispatchEvent(evt);

            evt = new MouseEvent('mouseup', {
                bubbles: true,
                clientX: rect.left,
                clientY: targetY
            });
            item.dispatchEvent(evt);
            """

            for i, target_idx in enumerate(order):
                try:
                    source_item = items[i]
                    target_item = items[target_idx]
                    target_y = target_item.location['y']
                    driver.execute_script(script_template, source_item, target_y)
                    time.sleep(0.3)  # 短暂延迟，等待动画完成
                except Exception as e:
                    logging.error(f"移动排序项时出错: {str(e)}")

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写排序题 {q_num} 时出错: {str(e)}")

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

    def random_delay(self, min_time=None, max_time=None):
        """生成随机延迟时间"""
        if min_time is None:
            min_time = self.config["min_delay"]
        if max_time is None:
            max_time = self.config["max_delay"]
        delay = random.uniform(min_time, max_time)
        time.sleep(delay)

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
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not file_path or not os.path.exists(file_path):
            return

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                imported_config = json.load(f)

            # 更新配置
            self.config.update(imported_config)

            # 更新界面
            # 全局设置
            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, self.config.get("url", ""))

            self.target_entry.delete(0, tk.END)
            self.target_entry.insert(0, str(self.config.get("target_num", 100)))

            self.ratio_scale.set(self.config.get("weixin_ratio", 0.5))
            self.ratio_var.set(f"{self.config.get('weixin_ratio', 0.5) * 100:.0f}%")

            self.min_duration.delete(0, tk.END)
            self.min_duration.insert(0, str(self.config.get("min_duration", 15)))

            self.max_duration.delete(0, tk.END)
            self.max_duration.insert(0, str(self.config.get("max_duration", 180)))

            self.min_delay.delete(0, tk.END)
            self.min_delay.insert(0, str(self.config.get("min_delay", 2.0)))

            self.max_delay.delete(0, tk.END)
            self.max_delay.insert(0, str(self.config.get("max_delay", 6.0)))

            self.submit_delay.delete(0, tk.END)
            self.submit_delay.insert(0, str(self.config.get("submit_delay", 3)))

            self.num_threads.delete(0, tk.END)
            self.num_threads.insert(0, str(self.config.get("num_threads", 4)))

            self.use_ip_var.set(self.config.get("use_ip", False))
            self.headless_var.set(self.config.get("headless", False))

            self.ip_entry.delete(0, tk.END)
            self.ip_entry.insert(0, self.config.get("ip_api", ""))

            # 重新加载题型设置
            self.reload_question_settings()

            logging.info(f"配置已从 {file_path} 导入")
            messagebox.showinfo("成功", "配置导入成功！")
        except Exception as e:
            logging.error(f"导入配置失败: {str(e)}")
            messagebox.showerror("错误", f"导入配置失败: {str(e)}")

    def reset_defaults(self):
        """重置为默认配置"""
        if messagebox.askyesno("确认", "确定要重置所有设置为默认值吗？"):
            # 重置配置为默认值
            self.config = DEFAULT_CONFIG.copy()

            # 更新全局设置界面
            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, self.config["url"])

            self.target_entry.delete(0, tk.END)
            self.target_entry.insert(0, str(self.config["target_num"]))

            self.ratio_scale.set(self.config["weixin_ratio"])
            self.ratio_var.set(f"{self.config['weixin_ratio'] * 100:.0f}%")

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

            # 重新加载题型设置
            self.reload_question_settings()

            logging.info("已重置为默认配置")
            messagebox.showinfo("成功", "已重置为默认配置！")

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
            self.config["use_ip"] = bool(self.use_ip_var.get())
            self.config["headless"] = bool(self.headless_var.get())
            self.config["ip_api"] = self.ip_entry.get().strip()

            # 保存题型设置
            # 单选题
            for i, q_num in enumerate(self.config["single_prob"].keys()):
                try:
                    probs = []
                    for entry in self.single_entries[i]:
                        value = entry.get().strip()
                        if value:
                            probs.append(float(value))
                    if probs and all(p == -1 for p in probs):
                        self.config["single_prob"][q_num] = -1
                    else:
                        self.config["single_prob"][q_num] = probs
                except Exception as e:
                    logging.error(f"保存单选题 {q_num} 配置时出错: {str(e)}")
                    return False

            # 多选题
            for i, q_num in enumerate(self.config["multiple_prob"].keys()):
                try:
                    probs = []
                    for entry in self.multi_entries[i]:
                        value = entry.get().strip()
                        if value:
                            prob = float(value)
                            if not 0 <= prob <= 100:
                                raise ValueError("概率必须在0-100之间")
                            probs.append(prob)
                    self.config["multiple_prob"][q_num] = probs
                except Exception as e:
                    logging.error(f"保存多选题 {q_num} 配置时出错: {str(e)}")
                    return False

            # 矩阵题
            for i, q_num in enumerate(self.config["matrix_prob"].keys()):
                try:
                    probs = []
                    for entry in self.matrix_entries[i]:
                        value = entry.get().strip()
                        if value:
                            probs.append(float(value))
                    if probs and all(p == -1 for p in probs):
                        self.config["matrix_prob"][q_num] = -1
                    else:
                        self.config["matrix_prob"][q_num] = probs
                except Exception as e:
                    logging.error(f"保存矩阵题 {q_num} 配置时出错: {str(e)}")
                    return False

            # 填空题
            for i, q_num in enumerate(self.config["texts"].keys()):
                try:
                    text = self.text_entries[i].get().strip()
                    self.config["texts"][q_num] = [t.strip() for t in text.split(",") if t.strip()]
                except Exception as e:
                    logging.error(f"保存填空题 {q_num} 配置时出错: {str(e)}")
                    return False

            # 多项填空题
            for i, q_num in enumerate(self.config["multiple_texts"].keys()):
                try:
                    text_lists = []
                    for entry in self.multiple_text_entries[i]:
                        text = entry.get().strip()
                        text_lists.append([t.strip() for t in text.split(",") if t.strip()])
                    self.config["multiple_texts"][q_num] = text_lists
                except Exception as e:
                    logging.error(f"保存多项填空题 {q_num} 配置时出错: {str(e)}")
                    return False

            # 排序题
            for i, q_num in enumerate(self.config["reorder_prob"].keys()):
                try:
                    probs = []
                    for entry in self.reorder_entries[i]:
                        value = entry.get().strip()
                        if value:
                            prob = float(value)
                            probs.append(prob)
                    self.config["reorder_prob"][q_num] = probs
                except Exception as e:
                    logging.error(f"保存排序题 {q_num} 配置时出错: {str(e)}")
                    return False

            return True

        except Exception as e:
            logging.error(f"保存配置时出错: {str(e)}")
            messagebox.showerror("错误", f"保存配置时出错: {str(e)}")
            return False

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
    root = ThemedTk(theme="equilux")
    app = WJXAutoFillApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()