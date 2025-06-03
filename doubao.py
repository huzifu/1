import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import logging
import random
import re
import traceback
import time
import numpy
import requests
import json
import os
from typing import List, Dict, Any
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ================== 配置参数 ==================
# 默认参数值
DEFAULT_CONFIG = {
    "url": "https://www.wjx.cn/vm/YItIXSF.aspx",
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
    "matrix_prob": {"1": [1, 0, 0, 0, 0], "2": -1, "3": [1, 0, 0, 0, 0],
                    "4": [1, 0, 0, 0, 0], "5": [1, 0, 0, 0, 0], "6": [1, 0, 0, 0, 0]},
    "scale_prob": {"7": [0, 2, 3, 4, 1], "12": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1]},
    "texts": {"11": ["多媒体与网络技术的深度融合", "人工智能驱动的个性化学习", "虚拟现实（VR/AR）与元宇宙教育",
                     "跨学科研究与学习科学深化"],
              "12": ["巧用AI工具解放重复劳动", "动态反馈激活课堂参与", "轻量化VR/AR增强认知具象化"]},
    "texts_prob": {"11": [1, 1, 1, 1], "12": [1, 1, 1]},
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
    },
    "question_texts": {}  # 新增：存储题目文本
}


# =============================================

class ToolTip:
    """工具提示类，用于显示题目文本提示"""

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
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def showtip(self):
        if not self.text or self.text.strip() == '':
            return

        # 创建提示窗口
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)  # 无边框
        self.tip_window.wm_geometry(f"+{self.x}+{self.y}")

        # 设置样式
        self.tip_window.configure(background="#ffffe0", relief="solid", borderwidth=1)

        # 添加标签
        label = tk.Label(
            self.tip_window,
            text=self.text,
            justify="left",
            background="#ffffe0",
            foreground="black",
            relief="solid",
            borderwidth=1,
            wraplength=self.wraplength,
            font=("Arial", 10)
        )
        label.pack(ipadx=5, ipady=3)

    def hidetip(self):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

    def set_text(self, text):
        self.text = text


class WJXAutoFillApp:

    def __init__(self, root):
        self.root = root
        self.root.title("问卷星自动填写工具")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)

        # 创建样式
        self.create_styles()

        self.config = DEFAULT_CONFIG.copy()
        self.running = False
        self.paused = False
        self.cur_num = 0
        self.cur_fail = 0
        self.lock = threading.Lock()
        self.pause_event = threading.Event()
        self.droplist_entries = []
        # 初始化存储各种问题类型的输入框列表
        self.multi_entries = []
        self.droplist_entries = []
        self.scale_entries = []  # 添加这一行

        # 创建主框架
        self.main_frame = ttk.Frame(root, style="Main.TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建标签页
        self.notebook = ttk.Notebook(self.main_frame, style="TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 全局设置标签页
        self.global_frame = ttk.Frame(self.notebook, style="Settings.TFrame")
        self.notebook.add(self.global_frame, text="全局设置")
        self.create_global_settings()

        # 题型设置标签页
        self.question_frame = ttk.Frame(self.notebook, style="Settings.TFrame")
        self.notebook.add(self.question_frame, text="题型设置")
        self.create_question_settings()

        # 日志区域
        self.log_frame = ttk.LabelFrame(self.main_frame, text="运行日志", style="Log.TLabelframe")
        self.log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 在日志区域上方添加按钮
        log_control_frame = ttk.Frame(self.log_frame, style="Control.TFrame")
        log_control_frame.pack(fill=tk.X, padx=5, pady=(5, 0))

        self.clear_log_btn = ttk.Button(log_control_frame, text="清空日志", command=self.clear_log,
                                        style="Control.TButton")
        self.clear_log_btn.pack(side=tk.LEFT, padx=5)

        self.export_log_btn = ttk.Button(log_control_frame, text="导出日志", command=self.export_log,
                                         style="Control.TButton")
        self.export_log_btn.pack(side=tk.LEFT, padx=5)

        self.log_area = scrolledtext.ScrolledText(self.log_frame, height=10, wrap=tk.WORD)
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_area.config(state=tk.DISABLED)

        # 重定向日志输出
        self.redirect_logging()

        # 控制按钮
        self.control_frame = ttk.Frame(self.main_frame, style="Control.TFrame")
        self.control_frame.pack(fill=tk.X, padx=5, pady=5)

        # 第一行按钮
        top_btn_frame = ttk.Frame(self.control_frame, style="Control.TFrame")
        top_btn_frame.pack(fill=tk.X, pady=(0, 5))

        self.start_btn = ttk.Button(top_btn_frame, text="开始填写", command=self.start_filling, style="Action.TButton")
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.pause_btn = ttk.Button(top_btn_frame, text="暂停", command=self.toggle_pause, state=tk.DISABLED,
                                    style="Action.TButton")
        self.pause_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(top_btn_frame, text="停止", command=self.stop_filling, state=tk.DISABLED,
                                   style="Action.TButton")
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        self.export_config_btn = ttk.Button(top_btn_frame, text="导出配置", command=self.export_config,
                                            style="Control.TButton")
        self.export_config_btn.pack(side=tk.LEFT, padx=5)

        self.import_config_btn = ttk.Button(top_btn_frame, text="导入配置", command=self.import_config,
                                            style="Control.TButton")
        self.import_config_btn.pack(side=tk.LEFT, padx=5)

        # 第二行状态信息
        bottom_info_frame = ttk.Frame(self.control_frame, style="Control.TFrame")
        bottom_info_frame.pack(fill=tk.X, pady=(5, 0))

        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        self.status_label = ttk.Label(bottom_info_frame, textvariable=self.status_var, style="Status.TLabel")
        self.status_label.pack(side=tk.RIGHT, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(bottom_info_frame, variable=self.progress_var, maximum=100,
                                            style="Success.Horizontal.TProgressbar")
        self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)

        # 题目进度
        self.question_progress_var = tk.DoubleVar()
        self.question_progress_bar = ttk.Progressbar(bottom_info_frame,
                                                     variable=self.question_progress_var,
                                                     maximum=100,
                                                     length=200,
                                                     style="Info.Horizontal.TProgressbar")
        self.question_progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)

        self.question_status_var = tk.StringVar()
        self.question_status_var.set("题目进度: 0/0")
        self.question_status_label = ttk.Label(bottom_info_frame, textvariable=self.question_status_var,
                                               style="Status.TLabel")
        self.question_status_label.pack(side=tk.RIGHT, padx=5)

    def create_styles(self):
        """创建自定义样式"""
        style = ttk.Style()

        # 框架样式
        style.configure("Main.TFrame", background="#f0f0f0")
        style.configure("Settings.TFrame", background="#ffffff", borderwidth=1, relief="solid")
        style.configure("Control.TFrame", background="#e0e0e0")

        # 按钮样式
        style.configure("Action.TButton", font=("Arial", 10, "bold"), padding=5)
        style.configure("Control.TButton", padding=3)
        style.map("Action.TButton",
                  background=[("active", "#4CAF50"), ("!disabled", "#2196F3"), ("disabled", "#BBDEFB")],
                  foreground=[("active", "white"), ("!disabled", "white")])
        style.map("Control.TButton",
                  background=[("active", "#E0E0E0"), ("!disabled", "#F5F5F5")])

        # 标签页样式
        style.configure("TNotebook", background="#f0f0f0")
        style.configure("TNotebook.Tab", padding=[10, 5], font=("Arial", 10, "bold"))

        # 标签样式
        style.configure("Status.TLabel", font=("Arial", 9), background="#e0e0e0")

        # 进度条样式
        style.configure("Success.Horizontal.TProgressbar", thickness=15, background="#4CAF50", troughcolor="#E0E0E0")
        style.configure("Info.Horizontal.TProgressbar", thickness=15, background="#2196F3", troughcolor="#E0E0E0")

        # 日志框样式
        style.configure("Log.TLabelframe", font=("Arial", 10, "bold"))
        style.configure("Log.TLabelframe.Label", foreground="#333333")

    def redirect_logging(self):
        """重定向日志输出到GUI"""

        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                logging.Handler.__init__(self)
                self.text_widget = text_widget

            def emit(self, record):
                msg = self.format(record)
                self.text_widget.config(state=tk.NORMAL)
                self.text_widget.insert(tk.END, msg + "\n")
                self.text_widget.see(tk.END)
                self.text_widget.config(state=tk.DISABLED)

        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        text_handler = TextHandler(self.log_area)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(text_handler)

    def create_global_settings(self):
        """创建全局设置界面"""
        frame = self.global_frame
        padx, pady = 5, 3

        # 问卷链接
        ttk.Label(frame, text="问卷链接:").grid(row=0, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.url_entry = ttk.Entry(frame, width=60)
        self.url_entry.grid(row=0, column=1, columnspan=3, padx=padx, pady=pady, sticky=tk.EW)
        self.url_entry.insert(0, self.config["url"])

        # 目标份数
        ttk.Label(frame, text="目标份数:").grid(row=1, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.target_entry = ttk.Spinbox(frame, from_=1, to=10000, width=10)
        self.target_entry.grid(row=1, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.target_entry.set(self.config["target_num"])

        # 微信作答比率
        ttk.Label(frame, text="微信作答比率:").grid(row=1, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.ratio_scale = ttk.Scale(frame, from_=0, to=1, orient=tk.HORIZONTAL)
        self.ratio_scale.grid(row=1, column=3, padx=padx, pady=pady, sticky=tk.EW)
        self.ratio_scale.set(self.config["weixin_ratio"])
        self.ratio_var = tk.StringVar()
        self.ratio_var.set(f"{self.config['weixin_ratio'] * 100:.0f}%")
        ttk.Label(frame, textvariable=self.ratio_var).grid(row=1, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.ratio_scale.config(command=lambda v: self.ratio_var.set(f"{float(v) * 100:.0f}%"))

        # 作答时长
        ttk.Label(frame, text="作答时长(秒):").grid(row=2, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(frame, text="最短:").grid(row=2, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_duration = ttk.Spinbox(frame, from_=5, to=300, width=5)
        self.min_duration.grid(row=2, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.min_duration.set(self.config["min_duration"])
        ttk.Label(frame, text="最长:").grid(row=2, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_duration = ttk.Spinbox(frame, from_=5, to=300, width=5)
        self.max_duration.grid(row=2, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.max_duration.set(self.config["max_duration"])

        # 延迟设置
        ttk.Label(frame, text="延迟设置(秒):").grid(row=3, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(frame, text="最小:").grid(row=3, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_delay = ttk.Spinbox(frame, from_=0.1, to=10, increment=0.1, width=5)
        self.min_delay.grid(row=3, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.min_delay.set(self.config["min_delay"])
        ttk.Label(frame, text="最大:").grid(row=3, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_delay = ttk.Spinbox(frame, from_=0.1, to=10, increment=0.1, width=5)
        self.max_delay.grid(row=3, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.max_delay.set(self.config["max_delay"])

        # 提交延迟
        ttk.Label(frame, text="提交延迟:").grid(row=4, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.submit_delay = ttk.Spinbox(frame, from_=1, to=10, width=5)
        self.submit_delay.grid(row=4, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.submit_delay.set(self.config["submit_delay"])

        # 窗口数量
        ttk.Label(frame, text="浏览器窗口数量:").grid(row=5, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.num_threads = ttk.Spinbox(frame, from_=1, to=10, width=5)
        self.num_threads.grid(row=5, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.num_threads.set(self.config["num_threads"])

        # IP设置
        self.use_ip_var = tk.BooleanVar(value=self.config["use_ip"])
        ttk.Checkbutton(frame, text="使用代理IP", variable=self.use_ip_var).grid(row=6, column=0, padx=padx, pady=pady,
                                                                                 sticky=tk.W)
        ttk.Label(frame, text="IP API:").grid(row=6, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.ip_entry = ttk.Entry(frame, width=40)
        self.ip_entry.grid(row=6, column=2, columnspan=3, padx=padx, pady=pady, sticky=tk.EW)
        self.ip_entry.insert(0, self.config["ip_api"])

        # 无头模式
        self.headless_var = tk.BooleanVar(value=self.config["headless"])
        ttk.Checkbutton(frame, text="无头模式(不显示浏览器)", variable=self.headless_var).grid(row=7, column=0,
                                                                                               padx=padx, pady=pady,
                                                                                               sticky=tk.W)

        # 添加解析问卷按钮
        ttk.Button(frame, text="解析问卷", command=self.parse_survey, style="Control.TButton").grid(row=7, column=1,
                                                                                                    padx=padx, pady=10,
                                                                                                    sticky=tk.W)

        # 添加填充空间
        for i in range(8):
            frame.rowconfigure(i, weight=1)
        for j in range(5):
            frame.columnconfigure(j, weight=1)

    def create_question_settings(self):
        # 新增：初始化所有题型的输入框列表
        self.single_entries = []
        self.multi_entries = []
        self.matrix_entries = []
        self.text_entries = []
        self.multiple_text_entries = []
        self.reorder_entries = []
        self.droplist_entries = []
        self.scale_entries = []

        # 新增：存储Tooltip对象
        self.tooltips = []

        # 原有代码...
        # 使用Notebook组织不同题型
        self.question_notebook = ttk.Notebook(self.question_frame, style="TNotebook")
        self.question_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 单选题设置
        if self.config["single_prob"]:
            self.single_frame = ttk.Frame(self.question_notebook, style="Settings.TFrame")
            self.question_notebook.add(self.single_frame, text=f"单选题({len(self.config['single_prob'])})")
            self.create_single_settings(self.single_frame)

        # 多选题设置
        if self.config["multiple_prob"]:
            self.multi_frame = ttk.Frame(self.question_notebook, style="Settings.TFrame")
            self.question_notebook.add(self.multi_frame, text=f"多选题({len(self.config['multiple_prob'])})")
            self.create_multi_settings(self.multi_frame)

        # 矩阵题设置
        if self.config["matrix_prob"]:
            self.matrix_frame = ttk.Frame(self.question_notebook, style="Settings.TFrame")
            self.question_notebook.add(self.matrix_frame, text=f"矩阵题({len(self.config['matrix_prob'])})")
            self.create_matrix_settings(self.matrix_frame)

        # 填空题设置
        if self.config["texts"]:
            self.text_frame = ttk.Frame(self.question_notebook, style="Settings.TFrame")
            self.question_notebook.add(self.text_frame, text=f"填空题({len(self.config['texts'])})")
            self.create_text_settings(self.text_frame)

        # 多项填空设置
        if self.config["multiple_texts"]:
            self.multiple_text_frame = ttk.Frame(self.question_notebook, style="Settings.TFrame")
            self.question_notebook.add(self.multiple_text_frame, text=f"多项填空({len(self.config['multiple_texts'])})")
            self.create_multiple_text_settings(self.multiple_text_frame)

        # 排序题设置
        if self.config["reorder_prob"]:
            self.reorder_frame = ttk.Frame(self.question_notebook, style="Settings.TFrame")
            self.question_notebook.add(self.reorder_frame, text=f"排序题({len(self.config['reorder_prob'])})")
            self.create_reorder_settings(self.reorder_frame)

        # 下拉框设置
        if self.config["droplist_prob"]:
            self.droplist_frame = ttk.Frame(self.question_notebook, style="Settings.TFrame")
            self.question_notebook.add(self.droplist_frame, text=f"下拉框({len(self.config['droplist_prob'])})")
            self.create_droplist_settings(self.droplist_frame)

        # 量表题设置
        if self.config["scale_prob"]:
            self.scale_frame = ttk.Frame(self.question_notebook, style="Settings.TFrame")
            self.question_notebook.add(self.scale_frame, text=f"量表题({len(self.config['scale_prob'])})")
            self.create_scale_settings(self.scale_frame)

    def create_single_settings(self, frame):
        padx, pady = 5, 3
        self.single_entries = []

        # 添加表头
        ttk.Label(frame, text="题号", font=("Arial", 9, "bold")).grid(row=0, column=0, padx=padx, pady=pady)
        for col in range(1, 6):  # 最多显示5个选项
            ttk.Label(frame, text=f"选项{col}", font=("Arial", 9, "bold")).grid(row=0, column=col, padx=padx, pady=pady)

        for i, (q_num, probs) in enumerate(self.config["single_prob"].items()):
            row = i + 1
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"单选题 {q_num}")
            q_label = ttk.Label(frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row, column=0, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            # 获取实际选项数量
            if probs == -1:
                option_count = 5  # 如果未配置，默认显示5个
            elif isinstance(probs, list):
                option_count = len(probs)  # 使用配置中的选项数量
            else:
                option_count = 5  # 默认

            entry_row = []
            for col in range(1, option_count + 1):
                entry = ttk.Entry(frame, width=8)
                # 设置默认值
                if probs == -1:
                    entry.insert(0, -1)
                elif isinstance(probs, list) and col <= len(probs):
                    entry.insert(0, probs[col - 1])
                else:
                    entry.insert(0, "")  # 空值
                entry.grid(row=row, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.single_entries.append(entry_row)

    def create_multi_settings(self, frame):
        padx, pady = 5, 3
        self.multi_entries = []

        # 说明标签
        ttk.Label(frame, text="设置每个多选题各选项被选择的概率（0-100之间的数值）").grid(row=0, column=0, columnspan=6,
                                                                                        padx=padx, pady=pady,
                                                                                        sticky=tk.W)

        # **动态生成表头（根据配置中的最大选项数）**
        headers = ["题号"] + [f"选项{i + 1}" for i in
                              range(max(len(probs) for probs in self.config["multiple_prob"].values()) or 5)]
        for col, header in enumerate(headers):
            ttk.Label(frame, text=header, font=("Arial", 9, "bold")).grid(row=1, column=col, padx=padx, pady=pady)

        # **遍历多选题配置，动态生成列数**
        for i, (q_num, probs) in enumerate(self.config["multiple_prob"].items()):
            row = i + 2
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"多选题 {q_num}")
            q_label = ttk.Label(frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row, column=0, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            entry_row = []
            # **获取实际选项数量：使用配置中的列表长度，无配置时默认5（可改为动态输入）**
            option_count = len(probs) if isinstance(probs, list) else 5  # 或从全局输入框获取
            for col in range(1, option_count + 1):  # 直接使用实际数量，移除`min(6, ...)`限制
                entry = ttk.Entry(frame, width=8)
                if isinstance(probs, list) and col <= len(probs):
                    entry.insert(0, probs[col - 1])
                else:
                    entry.insert(0, 50)  # 默认概率
                entry.grid(row=row, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.multi_entries.append(entry_row)

    def create_matrix_settings(self, frame):
        """创建矩阵题设置界面"""
        padx, pady = 5, 3

        # 添加全随机按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=0, column=0, columnspan=6, pady=(0, 5), sticky=tk.W)

        ttk.Button(btn_frame, text="全部随机",
                   command=lambda: self.set_all_random("matrix", frame)).pack(side=tk.LEFT, padx=5)

        # 说明标签
        ttk.Label(frame, text="设置矩阵题每个小题的选项概率（-1表示随机，[1,2]表示1:2比例）").grid(row=1, column=0,
                                                                                                columnspan=6, padx=padx,
                                                                                                pady=pady, sticky=tk.W)

        # 创建表格
        headers = ["小题号", "选项1", "选项2", "选项3", "选项4", "选项5"]
        for col, header in enumerate(headers):
            ttk.Label(frame, text=header, font=("Arial", 9, "bold")).grid(row=2, column=col, padx=padx, pady=pady)

        # 添加矩阵题设置行
        self.matrix_entries = []
        for i, (q_num, probs) in enumerate(self.config["matrix_prob"].items()):
            row = i + 3
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"矩阵题小题 {q_num}")
            q_label = ttk.Label(frame, text=f"小题{q_num}", cursor="hand2")
            q_label.grid(row=row, column=0, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            # 获取实际选项数量
            if probs == -1:
                option_count = 5
            elif isinstance(probs, list):
                option_count = len(probs)
            else:
                option_count = 5

            entry_row = []
            for col in range(1, min(6, option_count + 1)):
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
        padx, pady = 5, 3
        # 说明标签
        ttk.Label(frame, text="设置填空题的备选答案和选择概率（题号需与问卷一致）").grid(row=0, column=0, columnspan=9,
                                                                                       padx=padx, pady=pady,
                                                                                       sticky=tk.W)

        # 表头
        headers = ["题号", "答案1", "概率", "答案2", "概率", "答案3", "概率", "答案4", "概率"]
        for col, header in enumerate(headers):
            ttk.Label(frame, text=header, font=("Arial", 9, "bold")).grid(row=1, column=col, padx=padx, pady=pady)

        # 添加填空题设置行
        self.text_entries = []
        for i, (q_num, answers) in enumerate(self.config["texts"].items()):
            row = i + 2
            # 题号输入框
            q_num_entry = ttk.Entry(frame, width=5)
            q_num_entry.insert(0, q_num)
            q_num_entry.grid(row=row, column=0, padx=padx, pady=pady)

            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"填空题 {q_num}")
            q_label = ttk.Label(frame, text=f"题{q_num}", cursor="hand2")
            q_label.grid(row=row, column=0, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            entry_row = [q_num_entry]

            # 获取概率配置
            probs = self.config["texts_prob"].get(q_num, [1] * len(answers))

            # 答案和概率输入框
            for j in range(min(4, len(answers))):  # 最多显示4个答案
                # 答案输入框
                answer_entry = ttk.Entry(frame, width=15)
                answer_entry.insert(0, answers[j])
                answer_entry.grid(row=row, column=1 + j * 2, padx=padx, pady=pady)
                entry_row.append(answer_entry)

                # 概率输入框
                prob_entry = ttk.Entry(frame, width=5)
                prob_value = probs[j] if j < len(probs) else 1
                prob_entry.insert(0, prob_value)
                prob_entry.grid(row=row, column=2 + j * 2, padx=padx, pady=pady)
                entry_row.append(prob_entry)

            # 如果答案不足4个，添加空输入框
            for j in range(len(answers), 4):
                # 空答案输入框
                answer_entry = ttk.Entry(frame, width=15)
                answer_entry.grid(row=row, column=1 + j * 2, padx=padx, pady=pady)
                entry_row.append(answer_entry)

                # 空概率输入框
                prob_entry = ttk.Entry(frame, width=5)
                prob_entry.grid(row=row, column=2 + j * 2, padx=padx, pady=pady)
                entry_row.append(prob_entry)

            self.text_entries.append(entry_row)

    def create_multiple_text_settings(self, frame):
        """创建多项填空设置界面"""
        padx, pady = 5, 3
        # 说明标签
        ttk.Label(frame, text="设置多项填空每个部分的备选答案和选择概率").grid(row=0, column=0, columnspan=9,
                                                                               padx=padx, pady=pady, sticky=tk.W)

        # 添加多项填空设置行
        self.multiple_text_entries = []
        for i, (q_num, parts) in enumerate(self.config["multiple_texts"].items()):
            # 题号标签
            q_text = self.config["question_texts"].get(q_num, f"多项填空题 {q_num}")
            q_label = ttk.Label(frame, text=f"题号: {q_num}", font=("Arial", 9, "bold"), cursor="hand2")
            q_label.grid(row=i * 5 + 1, column=0, columnspan=9, padx=padx, pady=10, sticky=tk.W)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            # 获取该题的概率配置
            prob_config = self.config["multiple_texts_prob"].get(q_num, [])

            for part_idx, part in enumerate(parts):
                # 部分标签
                ttk.Label(frame, text=f"部分 {part_idx + 1}:").grid(row=i * 5 + 2 + part_idx, column=0, padx=padx,
                                                                    pady=pady, sticky=tk.W)

                entry_row = []

                # 答案和概率输入框
                for j in range(min(3, len(part))):  # 最多显示3个答案
                    # 答案输入框
                    answer_entry = ttk.Entry(frame, width=15)
                    answer_entry.insert(0, part[j])
                    answer_entry.grid(row=i * 5 + 2 + part_idx, column=1 + j * 3, padx=padx, pady=pady)
                    entry_row.append(answer_entry)

                    # 概率输入框
                    prob_entry = ttk.Entry(frame, width=5)
                    if j < len(prob_config) and j < len(prob_config[part_idx]):
                        prob_value = prob_config[part_idx][j]
                    else:
                        prob_value = 1.0
                    prob_entry.insert(0, prob_value)
                    prob_entry.grid(row=i * 5 + 2 + part_idx, column=2 + j * 3, padx=padx, pady=pady)
                    entry_row.append(prob_entry)

                # 如果答案不足3个，添加空输入框
                for j in range(len(part), 3):
                    # 空答案输入框
                    answer_entry = ttk.Entry(frame, width=15)
                    answer_entry.grid(row=i * 5 + 2 + part_idx, column=1 + j * 3, padx=padx, pady=pady)
                    entry_row.append(answer_entry)

                    # 空概率输入框
                    prob_entry = ttk.Entry(frame, width=5)
                    prob_entry.grid(row=i * 5 + 2 + part_idx, column=2 + j * 3, padx=padx, pady=pady)
                    entry_row.append(prob_entry)

                self.multiple_text_entries.append(entry_row)

    def create_reorder_settings(self, frame):
        """创建排序题设置界面"""
        padx, pady = 5, 3

        # 添加全随机按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=0, column=0, columnspan=6, pady=(0, 5), sticky=tk.W)

        ttk.Button(btn_frame, text="全部平均概率",
                   command=lambda: self.set_all_random("reorder", frame)).pack(side=tk.LEFT, padx=5)

        # 说明标签
        ttk.Label(frame, text="设置排序题各位置的概率（0-1之间的数值）").grid(row=1, column=0, columnspan=6,
                                                                            padx=padx, pady=pady, sticky=tk.W)

        # 创建表格
        headers = ["题号", "位置1", "位置2", "位置3", "位置4", "位置5"]
        for col, header in enumerate(headers):
            ttk.Label(frame, text=header, font=("Arial", 9, "bold")).grid(row=2, column=col, padx=padx, pady=pady)

        # 添加排序题设置行
        self.reorder_entries = []
        for i, (q_num, probs) in enumerate(self.config["reorder_prob"].items()):
            row = i + 3
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"排序题 {q_num}")
            q_label = ttk.Label(frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row, column=0, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            entry_row = []
            # 根据位置数量创建输入框
            position_count = len(probs) if isinstance(probs, list) else 4

            for col in range(1, min(6, position_count + 1)):
                entry = ttk.Entry(frame, width=8)
                if isinstance(probs, list) and col <= len(probs):
                    entry.insert(0, probs[col - 1])
                else:
                    entry.insert(0, 0.25)  # 默认概率
                entry.grid(row=row, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.reorder_entries.append(entry_row)

    def create_droplist_settings(self, frame):
        """创建下拉框设置界面"""
        padx, pady = 5, 3

        # 添加全随机按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=0, column=0, columnspan=6, pady=(0, 5), sticky=tk.W)

        ttk.Button(btn_frame, text="全部平均概率",
                   command=lambda: self.set_all_random("droplist", frame)).pack(side=tk.LEFT, padx=5)

        # 说明标签
        ttk.Label(frame, text="设置下拉框题各选项的概率（0-1之间的数值）").grid(row=1, column=0, columnspan=6,
                                                                              padx=padx, pady=pady,
                                                                              sticky=tk.W)

        # **动态生成表头**
        headers = ["题号"] + [f"选项{i + 1}" for i in
                              range(max(len(probs) for probs in self.config["droplist_prob"].values()) or 5)]
        for col, header in enumerate(headers):
            ttk.Label(frame, text=header, font=("Arial", 9, "bold")).grid(row=2, column=col, padx=padx, pady=pady)

        for i, (q_num, probs) in enumerate(self.config["droplist_prob"].items()):
            row = i + 3
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"下拉框题 {q_num}")
            q_label = ttk.Label(frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row, column=0, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            entry_row = []
            option_count = len(probs) if isinstance(probs, list) else 5  # 动态获取数量
            for col in range(1, option_count + 1):  # 直接使用实际数量
                entry = ttk.Entry(frame, width=8)
                if isinstance(probs, list) and col <= len(probs):
                    entry.insert(0, probs[col - 1])
                else:
                    entry.insert(0, 0.2)  # 默认概率
                entry.grid(row=row, column=col, padx=padx, pady=pady)
                entry_row.append(entry)
            self.droplist_entries.append(entry_row)

    def create_scale_settings(self, frame):
        """创建量表题设置界面"""
        padx, pady = 5, 3

        # 添加全随机按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=0, column=0, columnspan=6, pady=(0, 5), sticky=tk.W)

        ttk.Button(btn_frame, text="全部平均概率",
                   command=lambda: self.set_all_random("scale", frame)).pack(side=tk.LEFT, padx=5)

        # 说明标签
        ttk.Label(frame, text="设置量表题各选项的概率（0-1之间的数值）").grid(row=1, column=0, columnspan=6,
                                                                            padx=padx, pady=pady,
                                                                            sticky=tk.W)

        # **动态生成表头**
        headers = ["题号"] + [f"选项{i + 1}" for i in
                              range(max(len(probs) for probs in self.config["scale_prob"].values()) or 5)]
        for col, header in enumerate(headers):
            ttk.Label(frame, text=header, font=("Arial", 9, "bold")).grid(row=2, column=col, padx=padx, pady=pady)

        for i, (q_num, probs) in enumerate(self.config["scale_prob"].items()):
            row = i + 3
            # 创建题号标签并添加Tooltip
            q_text = self.config["question_texts"].get(q_num, f"量表题 {q_num}")
            q_label = ttk.Label(frame, text=f"第{q_num}题", cursor="hand2")
            q_label.grid(row=row, column=0, padx=padx, pady=pady)

            # 添加Tooltip
            tooltip = ToolTip(q_label, q_text, wraplength=300)
            self.tooltips.append(tooltip)

            entry_row = []
            option_count = len(probs) if isinstance(probs, list) else 5  # 动态获取数量
            for col in range(1, option_count + 1):  # 直接使用实际数量
                entry = ttk.Entry(frame, width=8)
                if isinstance(probs, list) and col <= len(probs):
                    entry.insert(0, probs[col - 1])
                else:
                    entry.insert(0, 0.2)  # 默认概率
                entry.grid(row=row, column=col, padx=padx, pady=pady)
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

        elif q_type == "multi":
            for entry_row in self.multi_entries:
                for i, entry in enumerate(entry_row):
                    entry.delete(0, tk.END)
                    entry.insert(0, "50")
            messagebox.showinfo("成功", "所有多选题选项概率已设置为50%")

        elif q_type == "matrix":
            for entry_row in self.matrix_entries:
                for entry in entry_row:
                    entry.delete(0, tk.END)
                    entry.insert(0, "-1")
            messagebox.showinfo("成功", "所有矩阵题已设置为随机选择")

        elif q_type == "droplist":
            for entry_row in self.droplist_entries:
                for entry in entry_row:
                    entry.delete(0, tk.END)
                    # 计算平均概率
                    avg_prob = 1.0 / len(entry_row)
                    entry.insert(0, f"{avg_prob:.2f}")
            messagebox.showinfo("成功", "所有下拉框选项概率已设置为平均概率")

        elif q_type == "scale":
            for entry_row in self.scale_entries:
                for entry in entry_row:
                    entry.delete(0, tk.END)
                    # 计算平均概率
                    avg_prob = 1.0 / len(entry_row)
                    entry.insert(0, f"{avg_prob:.2f}")
            messagebox.showinfo("成功", "所有量表题选项概率已设置为平均概率")

        elif q_type == "reorder":
            for entry_row in self.reorder_entries:
                for entry in entry_row:
                    entry.delete(0, tk.END)
                    # 计算平均概率
                    avg_prob = 1.0 / len(entry_row)
                    entry.insert(0, f"{avg_prob:.2f}")
            messagebox.showinfo("成功", "所有排序题位置概率已设置为平均概率")

    def parse_survey(self):
        """解析问卷结构并生成配置模板"""
        if not self.save_config():
            return

        try:
            logging.info("开始解析问卷结构...")
            self.status_var.set("正在解析问卷...")
            self.root.update()

            # 创建浏览器实例
            option = webdriver.ChromeOptions()
            option.add_argument('--headless')  # 无头模式
            option.add_argument('--disable-gpu')
            option.add_argument('--no-sandbox')
            option.add_argument('--disable-dev-shm-usage')
            option.add_experimental_option("excludeSwitches", ["enable-automation"])
            option.add_argument('--disable-blink-features=AutomationControlled')

            driver = webdriver.Chrome(options=option)
            driver.get(self.config["url"])

            # 等待问卷加载
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "divQuestion"))
            )
            time.sleep(2)  # 额外等待

            # 初始化解析结果
            parsed_config = {
                "single_prob": {},
                "droplist_prob": {},
                "multiple_prob": {},
                "matrix_prob": {},
                "scale_prob": {},
                "texts": {},
                "texts_prob": {},
                "multiple_texts": {},
                "multiple_texts_prob": {},
                "reorder_prob": {},
                "question_texts": {}  # 新增：存储题目文本
            }

            # 检测问卷页数和题量
            q_list = self.detect(driver)
            if not q_list or sum(q_list) == 0:
                logging.error("未检测到题目")
                return

            logging.info(f"检测到问卷共 {len(q_list)} 页，总题数: {sum(q_list)}")

            # 遍历所有题目
            current = 0
            for page_idx, page_count in enumerate(q_list):
                for q_idx in range(1, page_count + 1):
                    current += 1
                    try:
                        q_element = driver.find_element(By.CSS_SELECTOR, f"#div{current}")
                        q_type = q_element.get_attribute("type")

                        # 获取题目文本
                        try:
                            # 尝试获取题目文本（可能有多种结构）
                            q_title = q_element.find_element(By.CSS_SELECTOR, ".div_title_question").text.strip()
                        except:
                            try:
                                q_title = q_element.find_element(By.CSS_SELECTOR,
                                                                 "fieldset > ul > li > div").text.strip()
                            except:
                                q_title = f"第{current}题"

                        # 存储题目文本
                        parsed_config["question_texts"][str(current)] = q_title
                        logging.info(
                            f"解析第{current}题 - 题型: {self.get_question_type_name(q_type)} - 题目: {q_title}")

                        if not q_type:
                            continue

                        q_type = int(q_type)

                        # 根据题型生成配置模板
                        if q_type == 1 or q_type == 2:  # 填空题
                            # 检查是否多项填空
                            input_elements = q_element.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
                            if len(input_elements) > 1:
                                # 多项填空
                                part_count = len(input_elements)
                                parsed_config["multiple_texts"][str(current)] = [["示例答案"]] * part_count
                                parsed_config["multiple_texts_prob"][str(current)] = [[1]] * part_count
                                logging.info(f"  多项填空，包含 {part_count} 个部分")
                            else:
                                # 单项填空
                                parsed_config["texts"][str(current)] = ["示例答案1", "示例答案2", "示例答案3"]
                                parsed_config["texts_prob"][str(current)] = [1, 1, 1]

                        elif q_type == 3:  # 单选题
                            options = driver.find_elements(By.XPATH, f'//*[@id="div{current}"]/div[2]/div')
                            if options:
                                option_count = len(options)
                                # 使用实际选项数量创建配置
                                parsed_config["single_prob"][str(current)] = [1] * option_count
                                logging.info(f"  单选题，{option_count}个选项")

                        elif q_type == 4:  # 多选题
                            options = driver.find_elements(By.XPATH, f'//*[@id="div{current}"]/div[2]/div')
                            if options:
                                option_count = len(options)
                                parsed_config["multiple_prob"][str(current)] = [50] * option_count  # 50%概率
                                logging.info(f"  找到 {option_count} 个选项")

                        elif q_type == 5:  # 量表题
                            options = driver.find_elements(By.XPATH, f'//*[@id="div{current}"]/div[2]/div/ul/li')
                            if options:
                                option_count = len(options)
                                parsed_config["scale_prob"][str(current)] = [1] * option_count
                                logging.info(f"  找到 {option_count} 个量表选项")

                        elif q_type == 6:  # 矩阵题
                            rows = driver.find_elements(By.XPATH, f'//*[@id="divRefTab{current}"]/tbody/tr')
                            cols = driver.find_elements(By.XPATH, f'//*[@id="drv{current}_1"]/td')
                            if rows and cols:
                                row_count = len(rows) - 1  # 减去标题行
                                col_count = len(cols) - 1  # 减去题号列

                                # 为每个小题生成配置
                                for i in range(1, row_count + 1):
                                    parsed_config["matrix_prob"][f"{current}_{i}"] = -1

                                logging.info(f"  矩阵题包含 {row_count} 行 {col_count} 列")

                        elif q_type == 7:  # 下拉框
                            # 点击下拉框获取选项
                            try:
                                driver.find_element(By.CSS_SELECTOR, f"#select2-q{current}-container").click()
                                time.sleep(0.5)
                                options = driver.find_elements(By.XPATH, f"//*[@id='select2-q{current}-results']/li")
                                if options:
                                    option_count = len(options)
                                    parsed_config["droplist_prob"][str(current)] = [1] * option_count
                                    logging.info(f"  找到 {option_count} 个下拉选项")

                                # 关闭下拉框
                                driver.find_element(By.CSS_SELECTOR, f"#select2-q{current}-container").click()
                            except:
                                pass

                        elif q_type == 11:  # 排序题
                            options = driver.find_elements(By.XPATH, f'//*[@id="div{current}"]/ul/li')
                            if options:
                                option_count = len(options)
                                parsed_config["reorder_prob"][str(current)] = [1] * option_count
                                logging.info(f"  找到 {option_count} 个排序选项")

                    except Exception as e:
                        logging.warning(f"解析第{current}题时出错: {str(e)}")

            driver.quit()

            # 应用解析结果
            self.apply_parsed_config(parsed_config)
            self.reload_question_settings()

            # 切换到题型设置标签页
            self.notebook.select(1)

            logging.info("问卷解析完成！配置已更新")
            messagebox.showinfo("完成", "问卷解析完成！请检查题型设置标签页中的配置")
            self.status_var.set("就绪")

        except Exception as e:
            logging.error(f"解析问卷时出错: {str(e)}")
            self.status_var.set("解析失败")
            messagebox.showerror("错误", f"解析问卷时出错: {str(e)}\n{traceback.format_exc()}")

    def get_question_type_name(self, q_type: int) -> str:
        """获取题型名称"""
        type_names = {
            1: "填空题",
            2: "填空题",
            3: "单选题",
            4: "多选题",
            5: "量表题",
            6: "矩阵题",
            7: "下拉框",
            8: "滑块题",
            11: "排序题"
        }
        return type_names.get(q_type, f"未知题型({q_type})")

    def apply_parsed_config(self, parsed_config: Dict[str, Any]):
        """应用解析后的配置"""
        # 更新题型配置（确保键为字符串）
        for key in parsed_config:
            if key in self.config:
                self.config[key] = {}
                for q_num, value in parsed_config[key].items():
                    self.config[key][str(q_num)] = value  # 转换为字符串键
        # 保留原有的全局设置
        self.config.update({
            "url": self.config["url"],
            "target_num": self.config["target_num"],
            "min_duration": self.config["min_duration"],
            "max_duration": self.config["max_duration"],
            "weixin_ratio": self.config["weixin_ratio"],
            "min_delay": self.config["min_delay"],
            "max_delay": self.config["max_delay"],
            "submit_delay": self.config["submit_delay"],
            "page_load_delay": self.config["page_load_delay"],
            "per_question_delay": self.config["per_question_delay"],
            "per_page_delay": self.config["per_page_delay"],
            "use_ip": self.config["use_ip"],
            "headless": self.config["headless"],
            "ip_api": self.config["ip_api"],
            "num_threads": self.config["num_threads"]
        })

        # 归一化概率
        for prob_dict in [self.config["single_prob"], self.config["matrix_prob"],
                          self.config["droplist_prob"], self.config["scale_prob"],
                          self.config["reorder_prob"]]:
            for key in prob_dict:
                if isinstance(prob_dict[key], list):
                    prob_sum = sum(prob_dict[key])
                    if prob_sum > 0:
                        prob_dict[key] = [x / prob_sum for x in prob_dict[key]]

        for key in self.config["texts_prob"]:
            if isinstance(self.config["texts_prob"][key], list):
                prob_sum = sum(self.config["texts_prob"][key])
                if prob_sum > 0:
                    self.config["texts_prob"][key] = [x / prob_sum for x in self.config["texts_prob"][key]]

        for key in self.config["multiple_texts_prob"]:
            for i, probs in enumerate(self.config["multiple_texts_prob"][key]):
                if isinstance(probs, list):
                    prob_sum = sum(probs)
                    if prob_sum > 0:
                        self.config["multiple_texts_prob"][key][i] = [x / prob_sum for x in probs]

    def reload_question_settings(self):
        """重新加载题型设置界面（仅更新内部内容，不修改主Notebook标签页）"""
        # 销毁题型设置标签页内的所有控件（保留标签页本身）
        for widget in self.question_frame.winfo_children():
            widget.destroy()
        # 清除旧的Tooltip引用
        self.tooltips = []
        # 重新创建题型设置界面
        self.create_question_settings()
        # 刷新界面
        self.root.update()

    def clear_log(self):
        """清空日志"""
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        logging.info("日志已清空")

    def export_log(self):
        """导出日志到文件"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".log",
            filetypes=[("日志文件", "*.log"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(self.log_area.get(1.0, tk.END))
            logging.info(f"日志已导出到: {file_path}")
            messagebox.showinfo("成功", f"日志已成功导出到:\n{file_path}")
        except Exception as e:
            logging.error(f"导出日志失败: {str(e)}")
            messagebox.showerror("错误", f"导出日志失败: {str(e)}")

    def export_config(self):
        """导出配置到文件"""
        if not self.save_config():
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            logging.info(f"配置已导出到: {file_path}")
            messagebox.showinfo("成功", f"配置已成功导出到:\n{file_path}")
        except Exception as e:
            logging.error(f"导出配置失败: {str(e)}")
            messagebox.showerror("错误", f"导出配置失败: {str(e)}")

    def import_config(self):
        """从文件导入配置"""
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
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

            self.target_entry.set(self.config.get("target_num", 100))

            self.ratio_scale.set(self.config.get("weixin_ratio", 0.5))
            self.ratio_var.set(f"{self.config.get('weixin_ratio', 0.5) * 100:.0f}%")

            self.min_duration.set(self.config.get("min_duration", 15))
            self.max_duration.set(self.config.get("max_duration", 180))

            self.min_delay.set(self.config.get("min_delay", 2.0))
            self.max_delay.set(self.config.get("max_delay", 6.0))

            self.submit_delay.set(self.config.get("submit_delay", 3))

            self.num_threads.set(self.config.get("num_threads", 4))

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

    def save_config(self):
        """保存配置到字典 - 更新以包含所有题型"""
        try:
            # 全局设置
            self.config["url"] = self.url_entry.get()
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
            self.config["ip_api"] = self.ip_entry.get()

            # 单选题设置
            single_prob = {}
            for i, entry_row in enumerate(self.single_entries):
                q_num = str(i + 1)
                values = []
                for entry in entry_row:
                    val = entry.get()
                    if val:
                        values.append(float(val) if '.' in val else int(val))
                if values:
                    single_prob[q_num] = values[0] if len(values) == 1 else values
            self.config["single_prob"] = single_prob

            multiple_prob = {}
            # 遍历所有题目行，包括动态添加的
            for i, entry_row in enumerate(self.multi_entries):
                q_num = str(i + 1)
                values = []
                # 遍历当前题目下的所有选项输入框
                for entry in entry_row:
                    val = entry.get().strip()
                    if val:
                        try:
                            values.append(float(val))
                        except ValueError:
                            # 记录错误但继续处理其他选项
                            print(f"警告: 题目 {q_num} 的选项值 '{val}' 不是有效数字，已忽略")

                # 只有当题目有有效选项值时才添加到配置中
                if values:
                    multiple_prob[q_num] = values
                else:
                    print(f"警告: 题目 {q_num} 没有有效的选项值，已忽略")

            # 更新配置
            self.config["multiple_prob"] = multiple_prob

            # 矩阵题设置
            matrix_prob = {}
            for i, entry_row in enumerate(self.matrix_entries):
                q_num = str(i + 1)
                values = []
                for entry in entry_row:
                    val = entry.get()
                    if val:
                        values.append(float(val) if '.' in val else int(val))
                if values:
                    matrix_prob[q_num] = values[0] if len(values) == 1 else values
            self.config["matrix_prob"] = matrix_prob

            # 填空题设置
            texts = {}
            texts_prob = {}
            for entry_row in self.text_entries:
                q_num = entry_row[0].get().strip()
                if not q_num:
                    continue
                answers = []
                probs = []
                for j in range(1, len(entry_row), 2):
                    if j + 1 < len(entry_row):
                        answer = entry_row[j].get().strip()
                        prob_str = entry_row[j + 1].get().strip() if j + 1 < len(entry_row) else "1"
                        if answer:
                            answers.append(answer)
                            try:
                                prob = float(prob_str)
                            except:
                                prob = 1.0
                            probs.append(prob)
                if answers:
                    total = sum(probs)
                    normalized_probs = [p / total for p in probs] if total > 0 else [1.0 / len(answers)] * len(
                        answers)
                    texts[q_num] = answers
                    texts_prob[q_num] = normalized_probs
            self.config["texts"] = texts
            self.config["texts_prob"] = texts_prob

            # 多项填空设置
            multiple_texts = {}
            multiple_texts_prob = {}
            for i, entry_row in enumerate(self.multiple_text_entries):
                # 每5行对应一个多项填空题
                q_index = i // 5
                part_index = i % 5

                if part_index == 0:  # 每个多项填空题的第一行
                    q_num = f"{q_index + 1}"  # 题号
                    multiple_texts[q_num] = []
                    multiple_texts_prob[q_num] = []

                answers = []
                probs = []
                for j in range(0, len(entry_row), 2):
                    if j + 1 < len(entry_row):
                        answer = entry_row[j].get().strip()
                        prob_str = entry_row[j + 1].get().strip()
                        if answer:
                            answers.append(answer)
                            try:
                                prob = float(prob_str)
                            except:
                                prob = 1.0
                            probs.append(prob)

                if answers:
                    total = sum(probs)
                    normalized_probs = [p / total for p in probs] if total > 0 else [1.0 / len(answers)] * len(answers)
                    multiple_texts[q_num].append(answers)
                    multiple_texts_prob[q_num].append(normalized_probs)

            self.config["multiple_texts"] = multiple_texts
            self.config["multiple_texts_prob"] = multiple_texts_prob

            # 排序题设置
            reorder_prob = {}
            for i, entry_row in enumerate(self.reorder_entries):
                q_num = str(i + 1)
                values = []
                for entry in entry_row:
                    val = entry.get()
                    if val:
                        values.append(float(val))
                if values:
                    reorder_prob[q_num] = values
            self.config["reorder_prob"] = reorder_prob

            # 下拉框设置
            droplist_prob = {}
            for i, entry_row in enumerate(self.droplist_entries):
                q_num = str(i + 1)
                values = []
                for entry in entry_row:
                    val = entry.get()
                    if val:
                        values.append(float(val))
                if values:
                    droplist_prob[q_num] = values
            self.config["droplist_prob"] = droplist_prob

            # 量表题设置
            scale_prob = {}
            for i, entry_row in enumerate(self.scale_entries):
                q_num = str(i + 1)
                values = []
                for entry in entry_row:
                    val = entry.get()
                    if val:
                        values.append(float(val))
                if values:
                    scale_prob[q_num] = values
            self.config["scale_prob"] = scale_prob

            # 归一化概率
            for prob_dict in [self.config["single_prob"], self.config["matrix_prob"],
                              self.config["droplist_prob"], self.config["scale_prob"],
                              self.config["reorder_prob"]]:
                for key in prob_dict:
                    if isinstance(prob_dict[key], list):
                        prob_sum = sum(prob_dict[key])
                        if prob_sum > 0:
                            prob_dict[key] = [x / prob_sum for x in prob_dict[key]]

            for key in self.config["texts_prob"]:
                if isinstance(self.config["texts_prob"][key], list):
                    prob_sum = sum(self.config["texts_prob"][key])
                    if prob_sum > 0:
                        self.config["texts_prob"][key] = [x / prob_sum for x in self.config["texts_prob"][key]]

            for key in self.config["multiple_texts_prob"]:
                for i, probs in enumerate(self.config["multiple_texts_prob"][key]):
                    if isinstance(probs, list):
                        prob_sum = sum(probs)
                        if prob_sum > 0:
                            self.config["multiple_texts_prob"][key][i] = [x / prob_sum for x in probs]

            logging.info("配置保存成功")
            return True
        except Exception as e:
            logging.error(f"保存配置时出错: {str(e)}")
            messagebox.showerror("错误", f"保存配置时出错: {str(e)}")
            return False

    def start_filling(self):
        """开始填写问卷"""
        if self.running:
            logging.warning("程序已在运行中")
            return

        if not self.save_config():
            return

        self.running = True
        self.paused = False
        self.cur_num = 0
        self.cur_fail = 0
        self.progress_var.set(0)
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL)
        self.status_var.set("运行中...")
        self.pause_event.clear()  # 清除暂停事件

        # 启动填写线程
        self.threads = []
        for i in range(self.config["num_threads"]):
            thread = threading.Thread(target=self.run_filling, args=(50 + i * 60, 50))
            thread.daemon = True
            thread.start()
            self.threads.append(thread)

        # 启动进度更新线程
        self.progress_thread = threading.Thread(target=self.update_progress)
        self.progress_thread.daemon = True
        self.progress_thread.start()

    def toggle_pause(self):
        """切换暂停/继续状态"""
        if not self.running:
            return

        if self.paused:
            self.paused = False
            self.pause_event.clear()  # 清除暂停事件，让线程继续
            self.pause_btn.config(text="暂停")
            self.status_var.set("运行中...")
            logging.info("已继续填写问卷")
        else:
            self.paused = True
            self.pause_event.set()  # 设置暂停事件，让线程暂停
            self.pause_btn.config(text="继续")
            self.status_var.set("已暂停...")
            logging.info("已暂停填写问卷")

    def stop_filling(self):
        """停止填写问卷"""
        self.running = False
        self.paused = False
        self.status_var.set("正在停止...")
        self.stop_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.DISABLED)
        logging.info("正在停止问卷填写...")

    def update_progress(self):
        """更新进度条"""
        target_num = self.config["target_num"]
        while self.running and self.cur_num < target_num:
            progress = (self.cur_num / target_num) * 100
            self.progress_var.set(progress)
            self.status_var.set(f"已完成: {self.cur_num}/{target_num} 失败: {self.cur_fail}")
            time.sleep(0.5)

        # 任务完成或停止
        self.progress_var.set(100 if self.cur_num >= target_num else self.progress_var.get())
        status_text = f"已完成: {self.cur_num}/{target_num} 失败: {self.cur_fail}"
        if self.cur_num >= target_num:
            status_text += " - 任务完成!"
        self.status_var.set(status_text)
        self.running = False
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        logging.info("问卷填写任务已结束")

    def random_delay(self, min_time=None, max_time=None):
        """生成随机延迟时间"""
        if min_time is None:
            min_time = self.config["min_delay"]
        if max_time is None:
            max_time = self.config["max_delay"]
        delay = random.uniform(min_time, max_time)
        time.sleep(delay)

    def detect(self, driver: webdriver.Chrome) -> List[int]:
        """检测问卷页数和题量"""
        try:
            q_list = []
            page_num = len(driver.find_elements(By.XPATH, '//*[@id="divQuestion"]/fieldset'))
            for i in range(1, page_num + 1):
                questions = driver.find_elements(By.XPATH, f'//*[@id="fieldset{i}"]/div')
                valid_count = sum(1 for question in questions if
                                  question.get_attribute("topic") and question.get_attribute("topic").isdigit())
                q_list.append(valid_count)
            return q_list
        except:
            return [0]

    def fill_survey(self, driver: webdriver.Chrome):
        try:
            # 等待问卷加载完成
            WebDriverWait(driver, self.config["page_load_delay"]).until(
                EC.presence_of_element_located((By.ID, "divQuestion")))

            q_list = self.detect(driver)  # 检测页数和每一页的题量
            if not q_list or sum(q_list) == 0:
                logging.error("未检测到题目")
                return False
            total_questions = sum(q_list)  # 总题目数
            current_question = 0  # 当前题号
            single_num = 0  # 第num个单选题
            vacant_num = 0  # 第num个填空题
            multiple_text_num = 0  # 第num个多项填空题
            droplist_num = 0  # 第num个下拉框题
            multiple_num = 0  # 第num个多选题
            matrix_num = 0  # 第num个矩阵小题
            scale_num = 0  # 第num个量表题
            reorder_num = 0  # 第num个排序题
            current = 0  # 题号

            # 记录开始时间（用于控制作答时长）
            start_time = time.time()

            for j in q_list:  # 遍历每一页
                # 添加页面开始前的延迟
                self.random_delay(*self.config["per_page_delay"])
                for k in range(1, j + 1):  # 遍历该页的每一题
                    current += 1
                    current_question += 1
                    # 更新题目进度条
                    progress = (current_question / total_questions) * 100
                    self.question_progress_var.set(progress)
                    self.question_status_var.set(f"题目进度: {current_question}/{total_questions}")
                    # 强制GUI刷新
                    self.root.update_idletasks()

                    try:
                        # 获取题型
                        q_element = driver.find_element(By.CSS_SELECTOR, f"#div{current}")
                        q_type = q_element.get_attribute("type")

                        if not q_type:
                            continue

                        q_type = int(q_type)

                        if q_type == 1 or q_type == 2:  # 填空题
                            # 检查是否多项填空
                            input_elements = q_element.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
                            if len(input_elements) > 1:
                                self.fill_multiple_text(driver, current, multiple_text_num)
                                multiple_text_num += 1
                            else:
                                self.fill_text(driver, current)
                                vacant_num += 1
                        elif q_type == 3:  # 单选题
                            self.fill_single(driver, current, single_num)
                            single_num += 1
                        elif q_type == 4:  # 多选题
                            self.fill_multiple(driver, current, multiple_num)
                            multiple_num += 1
                        elif q_type == 5:  # 量表题
                            self.fill_scale(driver, current, scale_num)
                            scale_num += 1
                        elif q_type == 6:  # 矩阵题
                            matrix_num = self.fill_matrix(driver, current, matrix_num)
                        elif q_type == 7:  # 下拉框
                            self.fill_droplist(driver, current, droplist_num)
                            droplist_num += 1
                        elif q_type == 8:  # 滑块题
                            self.fill_slider(driver, current)
                        elif q_type == 11:  # 排序题
                            self.fill_reorder(driver, current, reorder_num)
                            reorder_num += 1
                        else:
                            logging.warning(f"第{current}题为不支持题型: {q_type}")
                    except Exception as e:
                        logging.error(f"填写第{current}题时出错: {str(e)}")
                        traceback.print_exc()

            # 计算已经花费的时间
            elapsed_time = time.time() - start_time
            total_duration = random.uniform(self.config["min_duration"], self.config["max_duration"])

            # 如果已经花费的时间小于随机停留时间，则等待剩余时间
            if elapsed_time < total_duration:
                time.sleep(total_duration - elapsed_time)

            # 提交问卷
            self.submit_survey(driver)

            # 提交后延迟
            time.sleep(self.config["submit_delay"])
            return True
        except Exception as e:
            logging.error(f"填写问卷时出错: {str(e)}")
            traceback.print_exc()
            return False

    def fill_text(self, driver: webdriver.Chrome, current: int):
        try:
            q_key = str(current)
            if q_key in self.config["texts"]:
                answers = self.config["texts"][q_key]
                probs = self.config["texts_prob"][q_key]
                # 修复：使用 probs 作为概率参数
                selected_idx = numpy.random.choice(len(answers), p=probs)
                content = answers[selected_idx]
                driver.find_element(By.CSS_SELECTOR, f"#q{current}").send_keys(content)
            else:
                driver.find_element(By.CSS_SELECTOR, f"#q{current}").send_keys("已填写")
            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写填空题 {current} 时出错: {str(e)}")

    def fill_multiple_text(self, driver: webdriver.Chrome, current: int, index: int):
        """填写多项填空题"""
        try:
            q_key = str(current)
            if q_key in self.config["multiple_texts"]:
                parts = self.config["multiple_texts"][q_key]
                probs = self.config["multiple_texts_prob"][q_key]

                # 获取所有输入框
                input_elements = driver.find_elements(By.CSS_SELECTOR,
                                                      f"#div{current} input[type='text'], #div{current} textarea")

                # 为每个部分填写内容
                for i in range(len(input_elements)):
                    if i < len(parts) and i < len(probs):
                        answers = parts[i]
                        part_probs = probs[i]

                        # 确保概率列表长度与答案一致
                        if len(part_probs) < len(answers):
                            part_probs = part_probs + [1.0] * (len(answers) - len(part_probs))
                        elif len(part_probs) > len(answers):
                            part_probs = part_probs[:len(answers)]

                        # 归一化概率
                        total = sum(part_probs)
                        if total > 0:
                            normalized_probs = [p / total for p in part_probs]
                        else:
                            normalized_probs = [1.0 / len(answers)] * len(answers)

                        # 选择答案
                        selected_idx = numpy.random.choice(len(answers), p=normalized_probs)
                        content = answers[selected_idx]
                        input_elements[i].send_keys(content)
                    else:
                        input_elements[i].send_keys("已填写")
            else:
                # 默认填写
                input_elements = driver.find_elements(By.CSS_SELECTOR,
                                                      f"#div{current} input[type='text'], #div{current} textarea")
                for input_element in input_elements:
                    input_element.send_keys("已填写")

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写多项填空题 {current} 时出错: {str(e)}")

    def fill_single(self, driver: webdriver.Chrome, current: int, index: int):
        try:
            options = driver.find_elements(By.XPATH, f'//*[@id="div{current}"]/div[2]/div')
            if not options:
                return

            q_key = str(current)  # 获取当前题号（如"1"、"2"）
            p = self.config["single_prob"].get(q_key, -1)  # 按题号获取配置概率

            if p == -1:  # 随机选择
                r = random.randint(1, len(options))
            elif isinstance(p, list) and len(p) == len(options):  # 按概率选择
                r = numpy.random.choice(a=numpy.arange(1, len(options) + 1), p=p)
            else:  # 默认随机
                r = random.randint(1, len(options))

            driver.find_element(By.CSS_SELECTOR, f"#div{current} > div.ui-controlgroup > div:nth-child({r})").click()
            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写单选题 {current} 时出错: {str(e)}")

    def fill_multiple(self, driver: webdriver.Chrome, current: int, index: int):
        try:
            options = driver.find_elements(By.XPATH, f'//*[@id="div{current}"]/div[2]/div')
            if not options:
                return

            q_key = str(current)
            p = self.config["multiple_prob"].get(q_key, [50] * len(options))  # 按题号获取概率，默认50%

            # 确保概率列表长度与选项一致
            p = p + [50] * (len(options) - len(p)) if len(p) < len(options) else p[:len(options)]

            selected = []
            for i, prob in enumerate(p):
                if random.random() < prob / 100.0:
                    selected.append(i + 1)

            # 确保至少选择一个
            if not selected:
                selected = [random.randint(1, len(options))]

            for idx in selected:
                driver.find_element(By.CSS_SELECTOR,
                                    f"#div{current} > div.ui-controlgroup > div:nth-child({idx})").click()

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写多选题 {current} 时出错: {str(e)}")

    def fill_matrix(self, driver: webdriver.Chrome, current: int, index: int) -> int:
        """填写矩阵题"""
        try:
            # 获取矩阵的行和列
            rows = driver.find_elements(By.XPATH, f'//*[@id="divRefTab{current}"]/tbody/tr')
            cols = driver.find_elements(By.XPATH, f'//*[@id="drv{current}_1"]/td')

            if not rows or not cols:
                return index

            q_num = len(rows) - 1  # 减去标题行
            col_num = len(cols) - 1  # 减去题号列

            # 遍历每一道小题
            for i in range(1, q_num + 1):
                if index < len(self.config["matrix_prob"]):
                    p = self.config["matrix_prob"][index]
                    index += 1
                else:
                    p = -1

                if p == -1:  # 随机选择
                    opt = random.randint(2, col_num + 1)
                elif isinstance(p, list) and len(p) == col_num:  # 按概率选择
                    opt = numpy.random.choice(a=numpy.arange(2, col_num + 2), p=p)
                else:  # 默认随机
                    opt = random.randint(2, col_num + 1)

                driver.find_element(By.CSS_SELECTOR, f"#drv{current}_{i} > td:nth-child({opt})").click()

            self.random_delay(*self.config["per_question_delay"])
            return index
        except Exception as e:
            logging.error(f"填写矩阵题 {current} 时出错: {str(e)}")
            return index

    def fill_droplist(self, driver: webdriver.Chrome, current: int, index: int):
        """填写下拉框题"""
        try:
            # 点击下拉框
            driver.find_element(By.CSS_SELECTOR, f"#select2-q{current}-container").click()
            time.sleep(0.5)

            # 获取选项
            options = driver.find_elements(By.XPATH, f"//*[@id='select2-q{current}-results']/li")
            if not options:
                return

            # 检查是否配置了该题
            if index < len(self.config["droplist_prob"]):
                p = self.config["droplist_prob"][index]
            else:
                p = -1

            if p == -1:  # 随机选择
                r = random.randint(1, len(options))
            elif isinstance(p, list) and len(p) == len(options):  # 按概率选择
                r = numpy.random.choice(a=numpy.arange(1, len(options) + 1), p=p)
            else:  # 默认随机
                r = random.randint(1, len(options))

            driver.find_element(By.XPATH, f"//*[@id='select2-q{current}-results']/li[{r}]").click()
            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写下拉框题 {current} 时出错: {str(e)}")

    def fill_scale(self, driver: webdriver.Chrome, current: int, index: int):
        """填写量表题"""
        try:
            options = driver.find_elements(By.XPATH, f'//*[@id="div{current}"]/div[2]/div/ul/li')
            if not options:
                return

            # 检查是否配置了该题
            if index < len(self.config["scale_prob"]):
                p = self.config["scale_prob"][index]
            else:
                p = -1

            if p == -1:  # 随机选择
                r = random.randint(1, len(options))
            elif isinstance(p, list) and len(p) == len(options):  # 按概率选择
                r = numpy.random.choice(a=numpy.arange(1, len(options) + 1), p=p)
            else:  # 默认随机
                r = random.randint(1, len(options))

            driver.find_element(By.CSS_SELECTOR,
                                f"#div{current} > div.scale-div > div > ul > li:nth-child({r})").click()
            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写量表题 {current} 时出错: {str(e)}")

    def fill_slider(self, driver: webdriver.Chrome, current: int):
        """填写滑块题"""
        try:
            score = random.randint(1, 100)
            driver.find_element(By.CSS_SELECTOR, f"#q{current}").send_keys(str(score))
            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写滑块题 {current} 时出错: {str(e)}")

    def fill_reorder(self, driver: webdriver.Chrome, current: int, index: int):
        """填写排序题"""
        try:
            options = driver.find_elements(By.XPATH, f'//*[@id="div{current}"]/ul/li')
            if not options:
                return

            # 获取配置的概率
            q_key = str(current)
            if q_key in self.config["reorder_prob"] and index < len(self.config["reorder_prob"][q_key]):
                p = self.config["reorder_prob"][q_key][index]
            else:
                p = -1

            # 随机排序
            if p == -1:  # 完全随机
                for j in range(1, len(options) + 1):
                    r = random.randint(j, len(options))
                    driver.find_element(By.CSS_SELECTOR, f"#div{current} > ul > li:nth-child({r})").click()
                    time.sleep(0.4)
            else:  # 按概率排序
                # 创建一个选项索引列表
                indices = list(range(1, len(options) + 1))
                # 按概率选择排序
                ordered_indices = numpy.random.choice(indices, size=len(options), replace=False, p=p)

                # 按顺序点击选项
                for idx in ordered_indices:
                    driver.find_element(By.CSS_SELECTOR, f"#div{current} > ul > li:nth-child({idx})").click()
                    time.sleep(0.4)

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写排序题 {current} 时出错: {str(e)}")

    def submit_survey(self, driver: webdriver.Chrome):
        """提交问卷"""
        try:
            # 尝试点击提交按钮
            submit_button = driver.find_element(By.XPATH, '//*[@id="ctlNext"]')
            submit_button.click()

            # 处理可能的弹窗
            try:
                confirm_button = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="layui-layer1"]/div[3]/a'))
                )
                confirm_button.click()
            except:
                pass

            # 处理智能验证
            try:
                smart_button = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="SM_BTN_1"]'))
                )
                smart_button.click()
            except:
                pass

            # 处理滑块验证
            try:
                slider = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="nc_1__scale_text"]/span'))
                )
                slider_button = driver.find_element(By.XPATH, '//*[@id="nc_1_n1z"]')
                if str(slider.text).startswith("请按住滑块"):
                    width = slider.size.get("width")
                    ActionChains(driver).drag_and_drop_by_offset(slider_button, width, 0).perform()
            except:
                pass
        except Exception as e:
            logging.error(f"提交问卷时出错: {str(e)}")

    def run_filling(self, xx, yy):
        """实际的问卷填写逻辑"""
        while self.running and self.cur_num < self.config["target_num"]:
            # 检查暂停状态
            if self.pause_event.is_set():
                time.sleep(1)
                continue

            driver = None
            try:
                # 创建浏览器实例
                option = webdriver.ChromeOptions()
                option.add_experimental_option("excludeSwitches", ["enable-automation"])
                option.add_experimental_option("useAutomationExtension", False)
                option.add_argument('--disable-blink-features=AutomationControlled')

                # 设置代理（如果需要）
                if self.config["use_ip"]:
                    ip = self.zanip()
                    if self.validate(ip):
                        option.add_argument(f"--proxy-server={ip}")
                        logging.info(f"使用代理IP: {ip}")

                # 设置无头模式
                if self.config["headless"]:
                    option.add_argument('--headless')
                    option.add_argument('--disable-gpu')

                # 随机决定是否使用微信来源
                use_weixin = random.random() < self.config["weixin_ratio"]
                source_type = "微信" if use_weixin else "其他"

                # 创建浏览器驱动
                driver = webdriver.Chrome(options=option)

                # 设置窗口位置和大小
                if not self.config["headless"]:
                    driver.set_window_position(xx, yy)
                    if use_weixin:
                        driver.set_window_size(375, 812)  # 手机尺寸
                    else:
                        driver.set_window_size(550, 650)  # PC尺寸

                # 设置微信特定的User-Agent
                if use_weixin:
                    driver.execute_cdp_cmd('Network.setUserAgentOverride', {
                        "userAgent": "Mozilla/5.0 (Linux; Android 10; MI 8 Build/QKQ1.190828.002; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/86.0.4240.99 XWEB/4317 MMWEBSDK/20220105 Mobile Safari/537.36 MMWEBID/6170 MicroMessenger/8.0.19.2080(0x28001351) Process/toolsmp WeChat/arm64 Weixin NetType/WIFI Language/zh_CN ABI/arm64"
                    })

                # 访问问卷链接
                driver.get(self.config["url"])
                time.sleep(1)  # 确保页面加载

                # 填写问卷
                if self.fill_survey(driver):
                    # 更新计数
                    with self.lock:
                        self.cur_num += 1
                        logging.info(f"已填写{self.cur_num}份 ({source_type}) - 失败{self.cur_fail}次")
                else:
                    with self.lock:
                        self.cur_fail += 1
                        logging.warning(f"填写失败 ({source_type}) - 总失败{self.cur_fail}次")

                # 关闭浏览器
                if driver:
                    driver.quit()

                # 随机延迟
                self.random_delay(self.config["min_delay"], self.config["max_delay"])

            except Exception as e:
                with self.lock:
                    self.cur_fail += 1
                logging.error(f"填写出错: {str(e)}")
                traceback.print_exc()
                if driver:
                    try:
                        driver.quit()
                    except:
                        pass

                if self.cur_fail >= self.config["target_num"] / 4 + 1:
                    logging.critical("失败次数过多，程序将停止")
                    self.stop_filling()
                    break

                # 失败后延迟
                time.sleep(5)

    def zanip(self):
        """获取代理IP"""
        try:
            return requests.get(self.config["ip_api"]).text.strip()
        except:
            return ""

    def validate(self, ip):
        """校验IP地址合法性"""
        pattern = r"^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?):(\d{1,5})$"
        return re.match(pattern, ip) is not None


if __name__ == "__main__":
    root = tk.Tk()
    app = WJXAutoFillApp(root)
    root.mainloop()