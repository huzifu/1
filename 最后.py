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
import sv_ttk  # 用于现代主题

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

    # 多选题概率配置 - 增强版
    "multiple_prob": {
        "4": {
            "prob": [0.4, 0.3, 0.3],  # 每个选项被选中的概率
            "min_selection": 1,  # 最小选择项数
            "max_selection": 2  # 最大选择项数
        },
        "5": {
            "prob": [0.5, 0.5, 0.5, 0.5],
            "min_selection": 2,
            "max_selection": 3
        }
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
    "question_texts": {
        "1": "您的性别",
        "2": "您的年级",
        "3": "您每月的消费项目",
        "4": "您喜欢的运动",
        "5": "您的兴趣爱好",
        "6": "您对学校的满意度",
        "7": "您的专业课程评价",
        "8": "您的生活满意度",
        "9": "您的学习压力程度",
        "10": "您的姓名",
        "11": "您的联系方式",
        "12": "您的家庭信息",
        "13": "您喜欢的食物排序",
        "14": "您喜欢的电影类型排序",
        "15": "您的出生地",
        "16": "您的职业"
    },

    # 选项文本存储
    "option_texts": {
        "1": ["男", "女"],
        "2": ["大一", "大二", "大三", "大四"],
        "3": ["伙食", "购置衣物", "交通通讯", "生活用品", "日常交际", "学习用品", "娱乐旅游", "其他"],
        "4": ["篮球", "足球", "游泳", "跑步", "羽毛球"],
        "5": ["阅读", "音乐", "游戏", "旅行", "摄影"],
        "6": ["非常满意", "满意", "一般", "不满意", "非常不满意"],
        "7": ["非常满意", "满意", "一般", "不满意", "非常不满意"],
        "8": ["非常满意", "满意", "一般", "不满意", "非常不满意"],
        "9": ["非常大", "较大", "一般", "较小", "没有压力"],
        "13": ["中餐", "西餐", "日料", "快餐"],
        "14": ["科幻", "动作", "喜剧", "爱情"],
        "15": ["北京", "上海", "广州", "深圳"],
        "16": ["学生", "上班族", "自由职业", "退休"]
    }
}


# ToolTip类用于显示题目提示
class ToolTip:
    def __init__(self, widget, text='', delay=300, wraplength=500):  # 减少延迟，增加宽度
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
        # 使用更明显的样式
        label = tk.Label(self.tip_window, text=self.text, justify=tk.LEFT,
                         background="#ffffff", relief=tk.SOLID, borderwidth=1,
                         wraplength=self.wraplength, padx=10, pady=5,
                         font=("Arial", 10))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()


class WJXAutoFillApp:
    def __init__(self, root):
        self.root = root
        self.root.title("问卷星自动填写工具 v4.0")
        self.root.geometry("1200x900")
        self.root.resizable(True, True)

        # 设置应用图标
        try:
            self.root.iconbitmap("wjx_icon.ico")
        except:
            pass

        # 使用现代主题
        sv_ttk.set_theme("light")

        # 自定义样式 - 优化UI设计
        self.style = ttk.Style()
        self.style.theme_use('default')
        self.style.configure('TNotebook.Tab', padding=[10, 5], font=('Arial', 10, 'bold'))
        self.style.configure('TButton', padding=[10, 5], font=('Arial', 10))
        self.style.configure('TLabel', padding=[5, 2], font=('Arial', 10))
        self.style.configure('TEntry', padding=[5, 2])
        self.style.configure('TFrame', background='#f5f5f5')
        self.style.configure('Header.TLabel', font=('Arial', 11, 'bold'), foreground="#2c6fbb")
        self.style.configure('Title.TLabel', font=('Arial', 16, 'bold'), foreground="#2c6fbb")
        self.style.configure('Success.TLabel', foreground='green')
        self.style.configure('Warning.TLabel', foreground='orange')
        self.style.configure('Error.TLabel', foreground='red')
        self.style.configure('Accent.TButton', background='#4a90e2', foreground='white')

        self.config = DEFAULT_CONFIG.copy()
        self.running = False
        self.paused = False
        self.cur_num = 0
        self.cur_fail = 0
        self.lock = threading.Lock()
        self.pause_event = threading.Event()
        self.tooltips = []
        self.parsing = False

        # 初始化字体
        self.font_family = tk.StringVar()
        self.font_size = tk.IntVar()
        self.font_family.set("Arial")
        self.font_size.set(10)

        # 创建主框架
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 标题栏
        title_frame = ttk.Frame(main_frame, style='TFrame')
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

        title_label = ttk.Label(title_frame, text="问卷星自动填写工具", style='Title.TLabel')
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
        self.start_btn = ttk.Button(btn_frame, text="▶ 开始填写", command=self.start_filling, width=12,
                                    style='Accent.TButton')
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.pause_btn = ttk.Button(btn_frame, text="⏸ 暂停", command=self.toggle_pause, state=tk.DISABLED, width=10)
        self.pause_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(btn_frame, text="⏹ 停止", command=self.stop_filling, state=tk.DISABLED, width=10)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        ttk.Separator(btn_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=10, fill=tk.Y)

        # 状态栏
        status_frame = ttk.Frame(control_frame)
        status_frame.pack(fill=tk.X, pady=(5, 0))

        # 状态指示器
        self.status_indicator = ttk.Label(status_frame, text="●", font=("Arial", 14), foreground="green")
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

        # 初始化问卷题型设置的 Notebook
        self.question_notebook = ttk.Notebook(self.question_frame)
        self.question_notebook.pack(fill=tk.BOTH, expand=True)

        # 初始化所有题型的输入框列表 - 移到这里确保在create_question_settings前初始化
        self.single_entries = []
        self.multi_entries = []
        self.min_selection_entries = []
        self.max_selection_entries = []
        self.matrix_entries = []
        self.text_entries = []
        self.multiple_text_entries = []
        self.reorder_entries = []
        self.droplist_entries = []
        self.scale_entries = []

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
        canvas = tk.Canvas(frame, background='#f0f0f0')
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='TFrame')

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
        self.url_entry.insert(0, self.config["url"])

        # 第二行：目标份数和微信作答比率
        ttk.Label(survey_frame, text="目标份数:").grid(row=1, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.target_entry = ttk.Spinbox(survey_frame, from_=1, to=10000, width=8)  # 减小宽度
        self.target_entry.grid(row=1, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.target_entry.set(self.config["target_num"])

        ttk.Label(survey_frame, text="微信作答比率:").grid(row=1, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.ratio_scale = ttk.Scale(survey_frame, from_=0, to=1, orient=tk.HORIZONTAL, length=100)  # 减小长度
        self.ratio_scale.grid(row=1, column=3, padx=padx, pady=pady, sticky=tk.EW)
        self.ratio_scale.set(self.config["weixin_ratio"])
        self.ratio_var = tk.StringVar()
        self.ratio_var.set(f"{self.config['weixin_ratio'] * 100:.0f}%")
        ratio_label = ttk.Label(survey_frame, textvariable=self.ratio_var, width=4)  # 减小宽度
        ratio_label.grid(row=1, column=4, padx=(0, padx), pady=pady, sticky=tk.W)

        # 绑定滑块事件
        self.ratio_scale.bind("<Motion>", self.update_ratio_display)
        self.ratio_scale.bind("<ButtonRelease-1>", self.update_ratio_display)

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
        self.parse_btn = ttk.Button(button_frame, text="解析问卷", command=self.parse_survey, width=15)
        self.parse_btn.grid(row=0, column=0, padx=5)

        # 重置默认按钮
        ttk.Button(button_frame, text="重置默认", command=self.reset_defaults, width=15).grid(row=0, column=1, padx=5)
        # 确保滚动框架有正确的权重分配
        scrollable_frame.columnconfigure(0, weight=1)
        # 提示标签
        tip_label = ttk.Label(scrollable_frame, text="提示: 填写前请先解析问卷以获取题目结构", style='Warning.TLabel')
        tip_label.grid(row=5, column=0, columnspan=2, pady=(10, 0))

    def _process_parsed_questions(self, questions_data):
        """处理解析得到的问卷题目数据"""
        try:
            logging.info(f"解析到的题目数量: {len(questions_data)}")

            # 清空原有配置
            self.config["question_texts"] = {}
            self.config["option_texts"] = {}

            # 初始化题型配置
            self.config["single_prob"] = {}
            self.config["multiple_prob"] = {}
            self.config["matrix_prob"] = {}
            self.config["texts"] = {}
            self.config["multiple_texts"] = {}
            self.config["reorder_prob"] = {}
            self.config["droplist_prob"] = {}
            self.config["scale_prob"] = {}

            # 更新题目和选项信息
            for question in questions_data:
                question_id = str(question.get('id'))
                question_text = question.get('text', f"题目{question_id}")
                options = question.get('options', [])
                q_type = question.get('type', '1')

                # 更新题目文本
                self.config["question_texts"][question_id] = question_text

                # 更新选项文本
                self.config["option_texts"][question_id] = options

                # 根据题型初始化配置
                if q_type == '3':  # 单选题
                    self.config["single_prob"][question_id] = -1  # 默认随机
                elif q_type == '4':  # 多选题
                    self.config["multiple_prob"][question_id] = {
                        "prob": [50] * len(options),
                        "min_selection": 1,
                        "max_selection": min(3, len(options))
                    }
                elif q_type == '6':  # 矩阵题
                    self.config["matrix_prob"][question_id] = -1  # 默认随机
                elif q_type == '1':  # 填空题
                    self.config["texts"][question_id] = ["示例答案"]
                elif q_type == '5':  # 量表题
                    self.config["scale_prob"][question_id] = [0.2] * len(options)
                elif q_type == '7':  # 下拉框
                    self.config["droplist_prob"][question_id] = [0.3] * len(options)
                elif q_type == '11':  # 排序题
                    self.config["reorder_prob"][question_id] = [0.25] * len(options)
                elif q_type == '2':  # 多项填空
                    self.config["multiple_texts"][question_id] = [["示例答案"]] * len(options)

            # 处理完成后，更新题型设置界面
            self.root.after(0, self.reload_question_settings)

        except Exception as e:
            logging.error(f"处理解析的题目时出错: {str(e)}")

    def create_question_settings(self):
        """创建题型设置界面 - 推荐每次完整重建Canvas, Frame, Notebook等所有结构"""
        # 创建滚动框架
        self.question_canvas = tk.Canvas(self.question_frame)
        self.question_scrollbar = ttk.Scrollbar(self.question_frame, orient="vertical",
                                                command=self.question_canvas.yview)
        self.scrollable_question_frame = ttk.Frame(self.question_canvas)
        self.scrollable_question_frame.bind(
            "<Configure>",
            lambda e: self.question_canvas.configure(scrollregion=self.question_canvas.bbox("all"))
        )
        self.question_canvas.create_window((0, 0), window=self.scrollable_question_frame, anchor="nw")
        self.question_canvas.configure(yscrollcommand=self.question_scrollbar.set)
        self.question_scrollbar.pack(side="right", fill="y")
        self.question_canvas.pack(side="left", fill="both", expand=True)

        # 创建Notebook（每次都新建）
        self.question_notebook = ttk.Notebook(self.scrollable_question_frame)
        self.question_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 题型tab配置
        question_types = [
            ('single_prob', "单选题", self.create_single_settings),
            ('multiple_prob', "多选题", self.create_multi_settings),
            ('matrix_prob', "矩阵题", self.create_matrix_settings),
            ('texts', "填空题", self.create_text_settings),
            ('multiple_texts', "多项填空", self.create_multiple_text_settings),
            ('reorder_prob', "排序题", self.create_reorder_settings),
            ('droplist_prob', "下拉框", self.create_droplist_settings),
            ('scale_prob', "量表题", self.create_scale_settings)
        ]
        for config_key, label_text, create_func in question_types:
            count = len(self.config[config_key])
            frame = ttk.Frame(self.question_notebook)
            self.question_notebook.add(frame, text=f"{label_text}({count})")
            desc_frame = ttk.Frame(frame)
            desc_frame.pack(fill=tk.X, padx=8, pady=5)
            if count == 0:
                ttk.Label(desc_frame, text=f"暂无{label_text}题目", font=("Arial", 10, "italic"),
                          foreground="gray").pack(pady=20)
            else:
                create_func(frame)

        # 添加提示
        tip_frame = ttk.Frame(self.scrollable_question_frame)
        tip_frame.pack(fill=tk.X, pady=10)
        ttk.Label(tip_frame, text="提示: 鼠标悬停在题号上可查看题目内容",
                  style='Warning.TLabel').pack()
        self.scrollable_question_frame.update_idletasks()
        self.question_canvas.configure(scrollregion=self.question_canvas.bbox("all"))


    def update_ratio_display(self, event=None):
        """更新微信作答比率显示"""
        ratio = self.ratio_scale.get()
        self.ratio_var.set(f"{ratio * 100:.0f}%")
        self.config["weixin_ratio"] = ratio

    def parse_survey(self):
        """解析问卷结构并生成配置模板 - 优化版本"""
        if self.parsing:
            messagebox.showwarning("警告", "正在解析问卷，请稍候...")
            return

        self.parsing = True
        self.parse_btn.config(state=tk.DISABLED, text="解析中...")
        self.status_var.set("正在解析问卷...")
        self.status_indicator.config(foreground="orange")

        # 在新线程中执行解析
        threading.Thread(target=self._parse_survey_thread, daemon=True).start()

    # ================== 修复解析函数 ==================
    def _parse_survey_thread(self):
        """解析问卷的线程函数 - 优化版本"""
        driver = None
        try:
            url = self.url_entry.get().strip()
            if not url:
                self.root.after(0, lambda: messagebox.showerror("错误", "请输入问卷链接"))
                return

            # 验证URL格式
            if not re.match(r'^https?://(www\.)?wjx\.cn/vm/[\w\d]+\.aspx(#)?$', url):
                self.root.after(0, lambda: messagebox.showerror("错误", "问卷链接格式不正确"))
                return

            # 创建浏览器选项
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--blink-settings=imagesEnabled=false')
            options.add_argument('--disable-extensions')
            options.add_argument('--disable-logging')
            options.add_argument('--log-level=3')
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option("useAutomationExtension", False)
            options.add_argument('--disable-blink-features=AutomationControlled')

            # 添加反检测选项
            options.add_argument('--disable-web-security')
            options.add_argument('--allow-running-insecure-content')
            options.add_argument('--disable-notifications')
            options.add_argument('--disable-popup-blocking')
            options.add_argument('--disable-infobars')
            options.add_argument('--disable-save-password-bubble')
            options.add_argument('--disable-translate')
            options.add_argument('--ignore-certificate-errors')

            # 随机User-Agent
            user_agents = [
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0",
                "Mozilla/5.0 (iPhone; CPU iPhone OS 14_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Mobile/15E148 Safari/604.1"
            ]
            options.add_argument(f'--user-agent={random.choice(user_agents)}')

            prefs = {
                'profile.default_content_setting_values': {
                    'images': 2,
                    'javascript': 1,
                    'css': 2
                }
            }
            options.add_experimental_option('prefs', prefs)

            # 设置加载超时
            driver = webdriver.Chrome(options=options)
            driver.set_page_load_timeout(20)  # 增加超时时间
            driver.implicitly_wait(8)

            try:
                logging.info(f"正在访问问卷: {url}")
                driver.get(url)

                # 显示解析进度
                self.root.after(0, lambda: self.question_progress_var.set(10))
                self.root.after(0, lambda: self.question_status_var.set("加载问卷..."))

                # 等待问卷内容加载完成
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".div_question, .field, .question"))
                )

                # 收集题目数据 - 增强的选择器和逻辑
                questions_data = driver.execute_script("""
                    const getText = (element) => element ? element.textContent.trim() : '';

                    // 使用多种选择器获取题目
                    const questionSelectors = [
                        '.div_question', 
                        '.field', 
                        '.question',
                        '.question-wrapper',
                        '.survey-question'
                    ];

                    let questions = [];
                    for (const selector of questionSelectors) {
                        const elements = document.querySelectorAll(selector);
                        if (elements.length > 0) {
                            questions = Array.from(elements);
                            break;
                        }
                    }

                    // 如果没有找到题目，尝试更通用的选择器
                    if (questions.length === 0) {
                        const potentialQuestions = document.querySelectorAll('div[id^="div"], div[id^="field"]');
                        questions = Array.from(potentialQuestions).filter(q => {
                            return q.querySelector('.question-title, .field-label, .question-text');
                        });
                    }

                    const result = [];

                    questions.forEach((q, index) => {
                        // 获取题目ID
                        let id = q.id.replace('div', '').replace('field', '').replace('question', '') || `${index+1}`;

                        // 获取题目标题 - 尝试多种选择器
                        let titleElement = q.querySelector('.div_title_question, .field-label, .question-title');
                        if (!titleElement) {
                            // 尝试更通用的选择器
                            titleElement = q.querySelector('h2, h3, .title, .question-text');
                        }

                        const title = titleElement ? getText(titleElement) : `题目${id}`;

                        // 检测题型
                        let type = '1'; // 默认为填空题

                        // 检查单选题
                        if (q.querySelector('.ui-radio, input[type="radio"]')) {
                            type = '3'; // 单选题
                        } 
                        // 检查多选题
                        else if (q.querySelector('.ui-checkbox, input[type="checkbox"]')) {
                            type = '4'; // 多选题
                        } 
                        // 检查矩阵题
                        else if (q.querySelector('.matrix, table.matrix')) {
                            type = '6'; // 矩阵题
                        } 
                        // 检查下拉框
                        else if (q.querySelector('select')) {
                            type = '7'; // 下拉框
                        } 
                        // 检查排序题
                        else if (q.querySelector('.sort-ul, .sortable')) {
                            type = '11'; // 排序题
                        } 
                        // 检查量表题
                        else if (q.querySelector('.scale-ul, .scale')) {
                            type = '5'; // 量表题
                        } 
                        // 检查填空题
                        else if (q.querySelector('textarea') || q.querySelector('input[type="text"]')) {
                            if (q.querySelectorAll('input[type="text"]').length > 1) {
                                type = '2'; // 多项填空
                            } else {
                                type = '1'; // 填空题
                            }
                        }

                        // 收集选项 - 增强兼容性
                        const options = [];
                        const optionSelectors = [
                            '.ulradiocheck label', 
                            '.matrix th', 
                            '.scale-ul li', 
                            '.sort-ul li',
                            'select option',
                            '.option-text',
                            '.option-item',
                            '.option-label'
                        ];

                        for (const selector of optionSelectors) {
                            const opts = q.querySelectorAll(selector);
                            if (opts.length > 0) {
                                opts.forEach(opt => {
                                    const text = getText(opt);
                                    if (text) options.push(text);
                                });
                                break; // 找到选项后跳出循环
                            }
                        }

                        // 如果没有获取到选项，尝试其他方式
                        if (options.length === 0) {
                            const dropdownOptions = q.querySelectorAll('option');
                            dropdownOptions.forEach(opt => {
                                if (opt.value && !opt.disabled) {
                                    options.push(getText(opt));
                                }
                            });
                        }

                        // 如果还是没有选项，尝试查找文本节点
                        if (options.length === 0) {
                            const textOptions = q.querySelectorAll('.option-text, .option-content');
                            textOptions.forEach(opt => {
                                const text = getText(opt);
                                if (text) options.push(text);
                            });
                        }

                        result.push({
                            id: id,
                            type: type,
                            text: title,
                            options: options
                        });
                    });

                    return result;
                """)

                # 解析后更新配置
                self._process_parsed_questions(questions_data)

                # 完成解析
                self.root.after(0, lambda: self.question_progress_var.set(100))
                self.root.after(0, lambda: self.question_status_var.set("解析完成"))
                self.root.after(0, lambda: messagebox.showinfo("成功", "问卷解析成功！"))

            except TimeoutException:
                logging.error("问卷加载超时，请检查网络或链接。")
                self.root.after(0, lambda: messagebox.showerror("错误", "问卷加载超时，请检查网络或链接。"))
            except Exception as e:
                logging.error(f"解析问卷时出错: {str(e)}")
                # 将异常信息保存到局部变量
                error_msg = str(e)
                self.root.after(0, lambda: messagebox.showerror("错误", f"解析问卷时出错: {error_msg}"))

        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            self.parsing = False
            self.root.after(0, lambda: self.parse_btn.config(state=tk.NORMAL, text="解析问卷"))
            self.root.after(0, lambda: self.status_var.set("就绪"))
            self.root.after(0, lambda: self.status_indicator.config(foreground="green"))

    def create_single_settings(self, frame):
        """创建单选题设置界面 - 修复输入框显示问题"""
        padx, pady = 8, 5
        # 配置说明框架
        desc_frame = ttk.LabelFrame(frame, text="单选题配置说明")
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)
        ttk.Label(desc_frame, text="配置说明：", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(desc_frame, text="• 输入 -1 表示随机选择\n• 输入正数表示选项的相对权重",
                  justify=tk.LEFT, font=("Arial", 9)).pack(anchor=tk.W, padx=5)

        # 创建表格框架并设置列权重 - 修复列宽问题
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 设置表格列权重，确保输入框区域可扩展 - 增加权重比例
        table_frame.columnconfigure(0, weight=1)  # 题号列
        table_frame.columnconfigure(1, weight=3)  # 题目预览列
        table_frame.columnconfigure(2, weight=8)  # 选项权重配置列（增加权重） ★ 修复点
        table_frame.columnconfigure(3, weight=2)  # 操作列

        # 表头 - 添加列宽设置
        headers = ["题号", "题目预览", "选项权重配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, probs) in enumerate(self.config["single_prob"].items(), start=1):
            base_row = row_idx
            # 获取题目文本
            q_text = self.config["question_texts"].get(q_num, f"单选题 {q_num}")
            # 获取实际选项数量
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0 and isinstance(probs, list):
                option_count = len(probs)

            # 创建题号标签和Tooltip
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)
            # 添加Tooltip
            tooltip_text = f"题目类型: 单选题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)

            # 选项配置容器 - 确保正确布局
            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)  # ★ 修复点

            entry_row = []

            # 生成选项输入框 - 修复布局问题
            for opt_idx in range(option_count):
                opt_container = ttk.Frame(option_frame)
                opt_container.grid(row=opt_idx, column=0, sticky=tk.W, pady=2)  # ★ 修复点

                opt_label = ttk.Label(opt_container, text=f"选项 {opt_idx + 1}: ", width=8)
                opt_label.pack(side=tk.LEFT, padx=(0, 5))

                entry = ttk.Entry(opt_container, width=8)
                if isinstance(probs, list) and opt_idx < len(probs):
                    entry.insert(0, str(probs[opt_idx]))
                elif probs == -1:
                    entry.insert(0, "-1")
                else:
                    entry.insert(0, "1")  # 默认权重为1
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)

            self.single_entries.append(entry_row)

            # 操作按钮 - 布局优化
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)

            # 创建按钮网格
            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)

            # 第一行按钮
            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row1, text="偏左", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("single", "left", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row1, text="偏右", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("single", "right", q, e)).pack(
                side=tk.LEFT, padx=2)

            # 第二行按钮
            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("single", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row2, text="平均", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_average("single", q, e)).pack(
                side=tk.LEFT, padx=2)

            # 添加分隔行
            if row_idx < len(self.config["single_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(
                    row=base_row + 1, column=0, columnspan=4, sticky='ew', pady=10)


    def create_multi_settings(self, frame):
        """创建多选题设置界面 - 完整修复版本"""
        padx, pady = 8, 5

        # 说明标签容器
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)

        # 主要说明
        ttk.Label(desc_frame, text="多选题配置说明：",
                  font=("Arial", 10, "bold")).pack(anchor=tk.W)

        ttk.Label(desc_frame, text="• 每个选项概率范围为0-100，表示该选项被选中的独立概率",
                  font=("Arial", 9)).pack(anchor=tk.W)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 设置列权重 - 确保选项配置列有足够空间
        table_frame.columnconfigure(0, weight=1)  # 题号
        table_frame.columnconfigure(1, weight=3)  # 题目预览
        table_frame.columnconfigure(2, weight=1)  # 最小选择数
        table_frame.columnconfigure(3, weight=1)  # 最大选择数
        table_frame.columnconfigure(4, weight=5)  # 选项概率配置（增加权重）
        table_frame.columnconfigure(5, weight=2)  # 操作

        # 表头
        headers = ["题号", "题目预览", "最小选择数", "最大选择数", "选项概率配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, config) in enumerate(self.config["multiple_prob"].items(), start=1):
            base_row = row_idx

            # 获取题目文本
            q_text = self.config["question_texts"].get(q_num, f"多选题 {q_num}")

            # 获取实际选项数量
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0 and "prob" in config:
                option_count = len(config["prob"])

            # 创建题号标签
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)

            # 添加Tooltips
            tooltip_text = f"题目类型: 多选题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)

            # 最小选择数容器
            min_frame = ttk.Frame(table_frame)
            min_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)

            min_entry = ttk.Spinbox(min_frame, from_=1, to=10, width=4)
            min_entry.set(config.get("min_selection", 1))
            min_entry.pack(fill=tk.X, expand=True)
            self.min_selection_entries.append(min_entry)

            # 最小选择说明
            ttk.Label(min_frame, text="最少选择项数",
                      font=("Arial", 8), foreground="gray").pack(fill=tk.X, expand=True)

            # 最大选择数容器
            max_frame = ttk.Frame(table_frame)
            max_frame.grid(row=base_row, column=3, padx=padx, pady=pady, sticky=tk.NSEW)

            max_entry = ttk.Spinbox(max_frame, from_=1, to=10, width=4)
            max_entry.set(config.get("max_selection", option_count if option_count > 0 else 1))
            max_entry.pack(fill=tk.X, expand=True)
            self.max_selection_entries.append(max_entry)

            # 最大选择说明
            ttk.Label(max_frame, text="最多选择项数",
                      font=("Arial", 8), foreground="gray").pack(fill=tk.X, expand=True)

            # 选项配置容器
            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=4, padx=padx, pady=pady, sticky=tk.NSEW)
            option_frame.columnconfigure(0, weight=1)  # 添加权重配置

            entry_row = []

            # 添加选项 - 根据实际选项数量生成
            for opt_idx in range(option_count):
                # 选项容器框架
                opt_container = ttk.Frame(option_frame)
                opt_container.grid(row=opt_idx, column=0, sticky=tk.W, pady=2)

                # 选项标签
                opt_label = ttk.Label(opt_container, text=f"选项 {opt_idx + 1}: ", width=8)
                opt_label.pack(side=tk.LEFT, padx=(0, 5))

                # 概率输入框
                entry = ttk.Entry(opt_container, width=8)
                if isinstance(config["prob"], list) and opt_idx < len(config["prob"]):
                    entry.insert(0, config["prob"][opt_idx])
                else:
                    entry.insert(0, 50)  # 默认概率50%
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)

            self.multi_entries.append(entry_row)

            # 操作按钮容器
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=5, padx=5, pady=5, sticky=tk.NW)

            # 创建按钮网格
            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)

            # 第一行按钮
            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)

            ttk.Button(btn_row1, text="偏左", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("multiple", "left", q, e)).pack(
                side=tk.LEFT, padx=2)

            ttk.Button(btn_row1, text="偏右", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("multiple", "right", q, e)).pack(
                side=tk.LEFT, padx=2)

            # 第二行按钮
            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)

            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("multiple", q, e)).pack(
                side=tk.LEFT, padx=2)

            ttk.Button(btn_row2, text="50%", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_value("multiple", q, e, 50)).pack(
                side=tk.LEFT, padx=2)

            # 添加分隔行
            if row_idx < len(self.config["multiple_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(
                    row=base_row + 1, column=0, columnspan=6, sticky='ew', pady=10)


    def create_text_settings(self, frame):
        """创建填空题设置界面"""
        padx, pady = 8, 5

        # 说明标签容器
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)

        # 主要说明
        ttk.Label(desc_frame, text="填空题配置说明：",
                  font=("Arial", 10, "bold")).pack(anchor=tk.W)

        ttk.Label(desc_frame,
                  text="• 输入多个答案时用逗号分隔\n• 系统会随机选择一个答案填写",
                  justify=tk.LEFT,
                  font=("Arial", 9)).pack(anchor=tk.W, padx=5)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 设置列权重
        table_frame.columnconfigure(0, weight=1)  # 题号
        table_frame.columnconfigure(1, weight=3)  # 题目预览
        table_frame.columnconfigure(2, weight=5)  # 答案配置
        table_frame.columnconfigure(3, weight=2)  # 操作

        # 表头
        headers = ["题号", "题目预览", "答案配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, answers) in enumerate(self.config["texts"].items(), start=1):
            base_row = row_idx

            # 获取题目文本
            q_text = self.config["question_texts"].get(q_num, f"填空题 {q_num}")

            # 创建题号标签和Tooltip
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)

            # 添加Tooltips
            tooltip_text = f"题目类型: 填空题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)

            # 答案配置容器
            answer_frame = ttk.Frame(table_frame)
            answer_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)

            # 创建答案输入框
            entry = ttk.Entry(answer_frame, width=40)
            entry.pack(fill=tk.X, padx=5, pady=2)

            # 设置初始值（将答案列表转换为逗号分隔的字符串）
            answer_str = ", ".join(answers)
            entry.insert(0, answer_str)
            self.text_entries.append(entry)

            # 操作按钮容器
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)

            # 重置按钮
            reset_btn = ttk.Button(btn_frame, text="重置", width=8,
                                   command=lambda e=entry: self.reset_text_entry(e))
            reset_btn.pack(pady=2)

            # 添加分隔行
            if row_idx < len(self.config["texts"]):
                ttk.Separator(table_frame, orient='horizontal').grid(
                    row=base_row + 1, column=0, columnspan=4, sticky='ew', pady=10)

        def reset_text_entry(self, entry):
            """重置填空题答案为默认值"""
            entry.delete(0, tk.END)
            entry.insert(0, "示例答案")

    def create_multiple_text_settings(self, frame):
        """创建多项填空设置界面"""
        padx, pady = 8, 5

        # 说明标签容器
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)

        # 主要说明
        ttk.Label(desc_frame, text="多项填空配置说明：",
                  font=("Arial", 10, "bold")).pack(anchor=tk.W)

        ttk.Label(desc_frame,
                  text="• 每个输入框对应一个空的答案配置\n• 多个答案用逗号分隔",
                  justify=tk.LEFT,
                  font=("Arial", 9)).pack(anchor=tk.W, padx=5)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 设置列权重
        table_frame.columnconfigure(0, weight=1)  # 题号
        table_frame.columnconfigure(1, weight=3)  # 题目预览
        table_frame.columnconfigure(2, weight=5)  # 答案配置
        table_frame.columnconfigure(3, weight=2)  # 操作

        # 表头
        headers = ["题号", "题目预览", "答案配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, answers_list) in enumerate(self.config["multiple_texts"].items(), start=1):
            base_row = row_idx

            # 获取题目文本
            q_text = self.config["question_texts"].get(q_num, f"多项填空 {q_num}")

            # 创建题号标签和Tooltip
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)

            # 添加Tooltips
            tooltip_text = f"题目类型: 多项填空\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)

            # 答案配置容器
            answer_frame = ttk.Frame(table_frame)
            answer_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)

            entry_row = []

            # 为每个空创建输入框
            for i, answers in enumerate(answers_list):
                # 容器框架
                field_frame = ttk.Frame(answer_frame)
                field_frame.pack(fill=tk.X, pady=2)

                # 标签
                ttk.Label(field_frame, text=f"空 {i + 1}: ", width=6).pack(side=tk.LEFT, padx=(0, 5))

                # 输入框
                entry = ttk.Entry(field_frame, width=40)
                entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

                # 设置初始值（将答案列表转换为逗号分隔的字符串）
                answer_str = ", ".join(answers)
                entry.insert(0, answer_str)
                entry_row.append(entry)

            self.multiple_text_entries.append(entry_row)

            # 操作按钮容器
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)

            # 重置按钮
            reset_btn = ttk.Button(btn_frame, text="重置", width=8,
                                   command=lambda e=entry_row: self.reset_multiple_text_entry(e))
            reset_btn.pack(pady=2)

            # 添加分隔行
            if row_idx < len(self.config["multiple_texts"]):
                ttk.Separator(table_frame, orient='horizontal').grid(
                    row=base_row + 1, column=0, columnspan=4, sticky='ew', pady=10)

        # 添加重置方法
        def reset_multiple_text_entry(self, entries):
            """重置多项填空答案为默认值"""
            for entry in entries:
                entry.delete(0, tk.END)
                entry.insert(0, "示例答案")

    def create_matrix_settings(self, frame):
        """创建矩阵题设置界面 - 完整修复版本"""
        padx, pady = 8, 5

        # 说明标签容器
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)

        # 主要说明
        ttk.Label(desc_frame, text="矩阵题配置说明：",
                  font=("Arial", 10, "bold")).pack(anchor=tk.W)

        ttk.Label(desc_frame,
                  text="• 输入 -1 表示随机选择\n" +
                       "• 输入正数表示选项的相对权重",
                  justify=tk.LEFT,
                  font=("Arial", 9)).pack(anchor=tk.W, padx=5)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 设置列权重
        table_frame.columnconfigure(0, weight=1)  # 题号
        table_frame.columnconfigure(1, weight=3)  # 题目预览
        table_frame.columnconfigure(2, weight=5)  # 选项配置（增加权重）
        table_frame.columnconfigure(3, weight=2)  # 操作

        # 表头
        headers = ["题号", "题目预览", "选项配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, probs) in enumerate(self.config["matrix_prob"].items(), start=1):
            base_row = row_idx

            # 获取题目文本
            q_text = self.config["question_texts"].get(q_num, f"矩阵题 {q_num}")

            # 获取实际选项数量
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0 and isinstance(probs, list):
                option_count = len(probs)

            # 创建题号标签和Tooltip
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)

            # 添加Tooltips
            tooltip_text = f"题目类型: 矩阵题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)

            # 选项配置容器
            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)
            option_frame.columnconfigure(0, weight=1)  # 添加权重配置

            entry_row = []

            # 添加选项 - 根据实际选项数量生成
            for opt_idx in range(option_count):
                # 选项容器框架
                opt_container = ttk.Frame(option_frame)
                opt_container.grid(row=opt_idx, column=0, sticky=tk.W, pady=2)

                # 选项标签
                opt_label = ttk.Label(opt_container, text=f"选项 {opt_idx + 1}: ", width=8)
                opt_label.pack(side=tk.LEFT, padx=(0, 5))

                # 权重输入框
                entry = ttk.Entry(opt_container, width=8)
                if isinstance(probs, list) and opt_idx < len(probs):
                    entry.insert(0, str(probs[opt_idx]))
                elif probs == -1:
                    entry.insert(0, "-1")
                else:
                    entry.insert(0, "1")  # 默认权重为1
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)

            self.matrix_entries.append(entry_row)

            # 操作按钮
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)

            # 创建按钮网格
            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)

            # 第一行按钮
            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)

            ttk.Button(btn_row1, text="偏左", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("matrix", "left", q, e)).pack(
                side=tk.LEFT, padx=2)

            ttk.Button(btn_row1, text="偏右", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("matrix", "right", q, e)).pack(
                side=tk.LEFT, padx=2)

            # 第二行按钮
            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)

            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("matrix", q, e)).pack(
                side=tk.LEFT, padx=2)

            ttk.Button(btn_row2, text="平均", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_average("matrix", q, e)).pack(
                side=tk.LEFT, padx=2)

            # 添加分隔行
            if row_idx < len(self.config["matrix_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(
                    row=base_row + 1, column=0, columnspan=4, sticky='ew', pady=10)

    def create_reorder_settings(self, frame):
        """创建排序题设置界面"""
        padx, pady = 8, 5

        # 说明标签容器
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)

        # 主要说明
        ttk.Label(desc_frame, text="排序题配置说明：",
                  font=("Arial", 10, "bold")).pack(anchor=tk.W)

        ttk.Label(desc_frame,
                  text="• 每个位置的概率表示该位置被选中的相对权重\n" +
                       "• 概率越高，该位置被选中的几率越大",
                  justify=tk.LEFT,
                  font=("Arial", 9)).pack(anchor=tk.W, padx=5)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 设置列权重
        table_frame.columnconfigure(0, weight=1)  # 题号
        table_frame.columnconfigure(1, weight=3)  # 题目预览
        table_frame.columnconfigure(2, weight=5)  # 位置概率配置
        table_frame.columnconfigure(3, weight=2)  # 操作

        # 表头
        headers = ["题号", "题目预览", "位置概率配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, probs) in enumerate(self.config["reorder_prob"].items(), start=1):
            base_row = row_idx

            # 获取题目文本
            q_text = self.config["question_texts"].get(q_num, f"排序题 {q_num}")

            # 获取实际选项数量
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0 and isinstance(probs, list):
                option_count = len(probs)

            # 创建题号标签和Tooltip
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)

            # 添加Tooltips
            tooltip_text = f"题目类型: 排序题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)

            # 位置配置容器
            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)
            option_frame.columnconfigure(0, weight=1)  # 添加权重配置

            entry_row = []

            # 添加位置 - 根据实际位置数量生成
            for pos_idx in range(option_count):
                # 位置容器框架
                pos_container = ttk.Frame(option_frame)
                pos_container.grid(row=pos_idx, column=0, sticky=tk.W, pady=2)

                # 位置标签
                pos_label = ttk.Label(pos_container, text=f"位置 {pos_idx + 1}: ", width=8)
                pos_label.pack(side=tk.LEFT, padx=(0, 5))

                # 权重输入框
                entry = ttk.Entry(pos_container, width=8)
                if isinstance(probs, list) and pos_idx < len(probs):
                    entry.insert(0, str(probs[pos_idx]))
                else:
                    entry.insert(0, f"{1 / option_count:.2f}")  # 平均概率
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)

            self.reorder_entries.append(entry_row)

            # 操作按钮
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)

            # 创建按钮网格
            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)

            # 第一行按钮
            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)

            ttk.Button(btn_row1, text="偏前", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("reorder", "left", q, e)).pack(
                side=tk.LEFT, padx=2)

            ttk.Button(btn_row1, text="偏后", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("reorder", "right", q, e)).pack(
                side=tk.LEFT, padx=2)

            # 第二行按钮
            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)

            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("reorder", q, e)).pack(
                side=tk.LEFT, padx=2)

            ttk.Button(btn_row2, text="平均", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_average("reorder", q, e)).pack(
                side=tk.LEFT, padx=2)

            # 添加分隔行
            if row_idx < len(self.config["reorder_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(
                    row=base_row + 1, column=0, columnspan=4, sticky='ew', pady=10)

    def create_droplist_settings(self, frame):
        """创建下拉框设置界面"""
        padx, pady = 8, 5

        # 说明标签容器
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)

        # 主要说明
        ttk.Label(desc_frame, text="下拉框配置说明：",
                  font=("Arial", 10, "bold")).pack(anchor=tk.W)

        ttk.Label(desc_frame,
                  text="• 每个选项的概率表示该选项被选中的相对权重\n" +
                       "• 概率越高，该选项被选中的几率越大",
                  justify=tk.LEFT,
                  font=("Arial", 9)).pack(anchor=tk.W, padx=5)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 设置列权重
        table_frame.columnconfigure(0, weight=1)  # 题号
        table_frame.columnconfigure(1, weight=3)  # 题目预览
        table_frame.columnconfigure(2, weight=5)  # 选项概率配置
        table_frame.columnconfigure(3, weight=2)  # 操作

        # 表头
        headers = ["题号", "题目预览", "选项概率配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, probs) in enumerate(self.config["droplist_prob"].items(), start=1):
            base_row = row_idx

            # 获取题目文本
            q_text = self.config["question_texts"].get(q_num, f"下拉框题 {q_num}")

            # 获取实际选项数量
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0 and isinstance(probs, list):
                option_count = len(probs)

            # 创建题号标签和Tooltip
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)

            # 添加Tooltips
            tooltip_text = f"题目类型: 下拉框题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)

            # 选项配置容器
            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)
            option_frame.columnconfigure(0, weight=1)  # 添加权重配置

            entry_row = []

            # 添加选项 - 根据实际选项数量生成
            for opt_idx in range(option_count):
                # 选项容器框架
                opt_container = ttk.Frame(option_frame)
                opt_container.grid(row=opt_idx, column=0, sticky=tk.W, pady=2)

                # 选项标签
                opt_label = ttk.Label(opt_container, text=f"选项 {opt_idx + 1}: ", width=8)
                opt_label.pack(side=tk.LEFT, padx=(0, 5))

                # 权重输入框
                entry = ttk.Entry(opt_container, width=8)
                if isinstance(probs, list) and opt_idx < len(probs):
                    entry.insert(0, str(probs[opt_idx]))
                else:
                    entry.insert(0, "0.3")  # 默认概率
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)

            self.droplist_entries.append(entry_row)

            # 操作按钮
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)

            # 创建按钮网格
            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)

            # 第一行按钮
            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)

            ttk.Button(btn_row1, text="偏前", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("droplist", "left", q, e)).pack(
                side=tk.LEFT, padx=2)

            ttk.Button(btn_row1, text="偏后", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("droplist", "right", q, e)).pack(
                side=tk.LEFT, padx=2)

            # 第二行按钮
            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)

            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("droplist", q, e)).pack(
                side=tk.LEFT, padx=2)

            ttk.Button(btn_row2, text="平均", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_average("droplist", q, e)).pack(
                side=tk.LEFT, padx=2)

            # 添加分隔行
            if row_idx < len(self.config["droplist_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(
                    row=base_row + 1, column=0, columnspan=4, sticky='ew', pady=10)

    def create_scale_settings(self, frame):
        """创建量表题设置界面 - 完整修复版本"""
        padx, pady = 8, 5

        # 说明标签容器
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)

        # 主要说明
        ttk.Label(desc_frame, text="量表题配置说明：",
                  font=("Arial", 10, "bold")).pack(anchor=tk.W)

        ttk.Label(desc_frame,
                  text="• 输入概率值表示该刻度被选中的相对概率\n" +
                       "• 概率越高，该刻度被选中的几率越大",
                  justify=tk.LEFT,
                  font=("Arial", 9)).pack(anchor=tk.W, padx=5)

        # 创建表格框架
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 设置列权重
        table_frame.columnconfigure(0, weight=1)  # 题号
        table_frame.columnconfigure(1, weight=3)  # 题目预览
        table_frame.columnconfigure(2, weight=5)  # 刻度概率配置（增加权重）
        table_frame.columnconfigure(3, weight=2)  # 操作

        # 表头
        headers = ["题号", "题目预览", "刻度概率配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        # 添加题目行
        for row_idx, (q_num, probs) in enumerate(self.config["scale_prob"].items(), start=1):
            base_row = row_idx

            # 获取题目文本
            q_text = self.config["question_texts"].get(q_num, f"量表题 {q_num}")

            # 获取实际选项数量
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0 and isinstance(probs, list):
                option_count = len(probs)

            # 创建题号标签和Tooltip
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)

            # 添加题目预览
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)

            # 添加Tooltips
            tooltip_text = f"题目类型: 量表题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)

            # 选项配置容器
            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)
            option_frame.columnconfigure(0, weight=1)  # 添加权重配置

            entry_row = []

            # 添加选项 - 根据实际选项数量生成
            for opt_idx in range(option_count):
                # 选项容器框架
                opt_container = ttk.Frame(option_frame)
                opt_container.grid(row=opt_idx, column=0, sticky=tk.W, pady=2)

                # 选项标签
                opt_label = ttk.Label(opt_container, text=f"刻度 {opt_idx + 1}: ", width=8)
                opt_label.pack(side=tk.LEFT, padx=(0, 5))

                # 权重输入框
                entry = ttk.Entry(opt_container, width=8)
                if isinstance(probs, list) and opt_idx < len(probs):
                    entry.insert(0, str(probs[opt_idx]))
                else:
                    entry.insert(0, "0.2")  # 默认权重为0.2
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)

            self.scale_entries.append(entry_row)

            # 操作按钮
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)

            # 创建按钮网格
            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)

            # 第一行按钮
            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)

            ttk.Button(btn_row1, text="偏左", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("scale", "left", q, e)).pack(
                side=tk.LEFT, padx=2)

            ttk.Button(btn_row1, text="偏右", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("scale", "right", q, e)).pack(
                side=tk.LEFT, padx=2)

            # 第二行按钮
            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)

            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("scale", q, e)).pack(
                side=tk.LEFT, padx=2)

            ttk.Button(btn_row2, text="平均", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_average("scale", q, e)).pack(
                side=tk.LEFT, padx=2)

            # 添加分隔行
            if row_idx < len(self.config["scale_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(
                    row=base_row + 1, column=0, columnspan=4, sticky='ew', pady=10)

    def set_question_bias(self, q_type, direction, q_num, entries):
        """为单个题目设置偏左或偏右分布"""
        bias_factors = {
            "left": [0.4, 0.3, 0.2, 0.1, 0.05],
            "right": [0.05, 0.1, 0.2, 0.3, 0.4]
        }

        factors = bias_factors.get(direction, [0.2, 0.2, 0.2, 0.2, 0.2])

        for i, entry in enumerate(entries):
            if i < len(factors):
                prob = factors[i]
            else:
                prob = factors[-1] * (0.8 ** (i - len(factors) + 1))  # 指数衰减

            # 根据题目类型格式化概率值
            if q_type == "multiple":
                prob_value = int(prob * 100)
            else:
                prob_value = f"{prob:.2f}"

            entry.delete(0, tk.END)
            entry.insert(0, str(prob_value))

        logging.info(f"第{q_num}题已设置为{direction}偏置")

    def set_question_random(self, q_type, q_num, entries):
        """为单个题目设置随机选择"""
        for entry in entries:
            entry.delete(0, tk.END)
            entry.insert(0, "-1")

        logging.info(f"第{q_num}题已设置为随机选择")

    def set_question_average(self, q_type, q_num, entries):
        """为单个题目设置平均概率"""
        option_count = len(entries)
        if option_count == 0:
            return

        avg_prob = 1.0 / option_count

        for entry in entries:
            entry.delete(0, tk.END)
            if q_type == "multiple":
                entry.insert(0, str(int(avg_prob * 100)))
            else:
                entry.insert(0, f"{avg_prob:.2f}")

        logging.info(f"第{q_num}题已设置为平均概率")

    def set_question_value(self, q_type, q_num, entries, value):
        """为单个题目设置指定值（多用于多选题）"""
        for entry in entries:
            entry.delete(0, tk.END)
            entry.insert(0, str(value))

        logging.info(f"第{q_num}题已设置为{value}%概率")

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

            # 更新所有控件的字体
            style = ttk.Style()
            style.configure('.', font=new_font)

            # 更新日志区域字体
            self.log_area.configure(font=new_font)

            # 更新按钮字体
            self.start_btn.configure(style='TButton')
            self.pause_btn.configure(style='TButton')
            self.stop_btn.configure(style='TButton')
            self.parse_btn.configure(style='TButton')

            # 更新标签字体
            for widget in self.root.winfo_children():
                self.update_widget_font(widget, new_font)

        except Exception as e:
            logging.error(f"更新字体时出错: {str(e)}")
            self.font_family.set("Arial")
            self.font_size.set(10)

    def update_widget_font(self, widget, font):
        """递归更新控件的字体"""
        try:
            # 更新当前控件
            if hasattr(widget, 'configure') and 'font' in widget.configure():
                widget.configure(font=font)

            # 递归更新子控件
            for child in widget.winfo_children():
                self.update_widget_font(child, font)
        except Exception as e:
            logging.debug(f"更新控件字体时出错: {str(e)}")

    def reload_question_settings(self):
        """重新加载题型设置界面 - 彻底销毁重建所有控件"""
        # 销毁所有子控件（包括Canvas/Scrollbar/Frame/Notebook）
        for widget in self.question_frame.winfo_children():
            widget.destroy()
        # 清空输入框和tooltip引用
        self.single_entries = []
        self.multi_entries = []
        self.min_selection_entries = []
        self.max_selection_entries = []
        self.matrix_entries = []
        self.text_entries = []
        self.multiple_text_entries = []
        self.reorder_entries = []
        self.droplist_entries = []
        self.scale_entries = []
        self.tooltips = []
        # 重新创建所有内容
        self.create_question_settings()
        # 确保界面刷新
        self.root.update_idletasks()

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

            # 验证URL格式
            if not re.match(r'^https?://(www\.)?wjx\.cn/vm/[\w\d]+\.aspx(#)?$', self.config["url"]):
                messagebox.showerror("错误", "问卷链接格式不正确")
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
                t = threading.Thread(target=self.run_filling, args=(x, y), daemon=True)
                t.start()
                self.threads.append(t)

            # 启动进度更新线程
            progress_thread = threading.Thread(target=self.update_progress, daemon=True)
            progress_thread.start()

        except Exception as e:
            logging.error(f"启动失败: {str(e)}")
            messagebox.showerror("错误", f"启动失败: {str(e)}")

    def run_filling(self, x=0, y=0):
        """运行填写任务 - 优化版本"""
        options = webdriver.ChromeOptions()
        if self.config["headless"]:
            options.add_argument('--headless')
        else:
            options.add_argument(f'--window-position={x},{y}')

        # 添加反检测参数
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument(
            '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
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
                        response = requests.get(self.config["ip_api"], timeout=10)
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
                            # 更健壮的微信作答按钮定位
                            weixin_btn = WebDriverWait(driver, 3).until(
                                EC.element_to_be_clickable((By.CLASS_NAME, "weixin-answer"))
                            )
                            driver.execute_script("arguments[0].scrollIntoView();", weixin_btn)
                            weixin_btn.click()
                            time.sleep(1)
                        except:
                            logging.debug("未找到微信作答按钮或点击失败")

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
        """填写问卷内容 - 优化版本"""
        try:
            questions = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".field.ui-field-contain, .div_question"))
            )

            if not questions:
                questions = driver.find_elements(By.CSS_SELECTOR, ".div_question")
                if not questions:
                    logging.warning("未找到任何题目，可能页面加载失败")
                    return False

            total_questions = len(questions)
            total_time = random.randint(self.config["min_duration"], self.config["max_duration"])
            start_time = time.time()

            # 计算每题平均时间
            avg_time_per_question = total_time / total_questions
            remaining_time = total_time

            for i, q in enumerate(questions):
                if not self.running:
                    break

                # 计算当前题目应花费的时间
                if i == total_questions - 1:
                    # 最后一题使用所有剩余时间
                    question_time = remaining_time
                else:
                    # 为每题分配一个随机时间
                    question_time = min(
                        random.uniform(avg_time_per_question * 0.5, avg_time_per_question * 1.5),
                        remaining_time - (total_questions - i - 1)
                    )

                question_start = time.time()

                try:
                    q_type = q.get_attribute("type")
                    q_id = q.get_attribute("id")
                    q_num = q_id.replace("div", "") if q_id else str(i + 1)

                    # 更新题目进度
                    self.question_progress_var.set((i + 1) / total_questions * 100)
                    self.question_status_var.set(f"题目进度: {i + 1}/{total_questions}")

                    # 填写题目
                    if q_type == "1":
                        self.fill_text(q, q_num)
                    elif q_type == "2":
                        self.fill_text(q, q_num)
                    elif q_type == "3":
                        self.fill_single(driver, q, q_num)
                    elif q_type == "4":
                        self.fill_multiple(driver, q, q_num)
                    elif q_type == "5":
                        self.fill_scale(driver, q, q_num)
                    elif q_type == "6":
                        self.fill_matrix(driver, q, q_num)
                    elif q_type == "7":
                        self.fill_droplist(driver, q, q_num)
                    elif q_type == "11":
                        self.fill_reorder(driver, q, q_num)
                    else:
                        self.auto_detect_question_type(driver, q, q_num)

                    # 计算并等待剩余时间
                    elapsed = time.time() - question_start
                    if elapsed < question_time:
                        time.sleep(question_time - elapsed)

                    remaining_time -= time.time() - question_start

                    # 检查翻页
                    try:
                        next_page = driver.find_element(By.CLASS_NAME, "next-page")
                        if next_page.is_displayed():
                            next_page.click()
                            time.sleep(random.uniform(*self.config["per_page_delay"]))
                    except:
                        pass

                except Exception as e:
                    logging.error(f"填写第{q_num}题时出错: {str(e)}")
                    continue

            # 补足总时长
            elapsed_total = time.time() - start_time
            if elapsed_total < total_time:
                time.sleep(total_time - elapsed_total)

            # 提交问卷
            return self.submit_survey(driver)

        except Exception as e:
            logging.error(f"填写问卷过程中出错: {str(e)}")
            return False

    def submit_survey(self, driver):
        """增强的问卷提交逻辑"""
        max_retries = 3  # 减少重试次数
        retry_count = 0
        success = False

        while retry_count < max_retries and not success and self.running:
            try:
                # 记录提交前URL用于后续比对
                original_url = driver.current_url

                # 多种定位提交按钮的方式
                submit_selectors = [
                    "#submit_button",
                    ".submit-btn",
                    ".submitbutton",
                    "a[id*='submit']",
                    "button[type='submit']",
                    "input[type='submit']",
                    "div.submit",
                    ".btn-submit",
                    ".btn-success",
                    "#ctlNext",
                    "#submit_button",
                    "#submit_btn",
                    "#next_button"
                ]

                submit_btn = None
                for selector in submit_selectors:
                    try:
                        # 增加等待时间，确保按钮可点击
                        submit_btn = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                        if submit_btn:
                            logging.debug(f"使用选择器找到提交按钮: {selector}")
                            break
                    except Exception as e:
                        logging.debug(f"使用选择器 {selector} 未找到提交按钮: {str(e)}")
                        continue

                if not submit_btn:
                    # 尝试通过文本查找提交按钮
                    try:
                        submit_btn = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'提交')]"))
                        )
                        logging.debug("通过文本找到提交按钮")
                    except:
                        logging.error("找不到提交按钮")
                        return False

                # 滚动到按钮并高亮
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submit_btn)
                # 使用复杂点击序列绕过检测
                ActionChains(driver) \
                    .move_to_element(submit_btn) \
                    .pause(0.3) \
                    .click_and_hold() \
                    .pause(0.2) \
                    .release() \
                    .perform()

                logging.info("已尝试提交问卷")

                # 增加提交后的延迟
                time.sleep(self.config["submit_delay"])

                # 等待页面变化 - 检测URL或DOM变化
                try:
                    WebDriverWait(driver, 10).until(
                        lambda d: d.current_url != original_url or
                                  "提交成功" in d.page_source or
                                  "感谢您" in d.page_source or
                                  "/wjx/join/complete.aspx" in d.current_url or
                                  "已完成" in d.page_source
                    )
                except TimeoutException:
                    logging.warning("提交后页面未变化，可能提交失败")

                # 验证提交结果
                success = self.verify_submission(driver)

                if success:
                    logging.info("提交成功验证通过")
                    break

                # 检查验证码
                if self.handle_captcha(driver):
                    logging.info("验证码处理后重新尝试提交")
                    retry_count += 1
                    continue

                logging.warning(f"提交可能失败，准备重试 ({retry_count + 1}/{max_retries})")
                retry_count += 1

                # 重试前刷新页面
                try:
                    driver.refresh()
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".field.ui-field-contain"))
                    )
                except:
                    pass

            except Exception as e:
                logging.error(f"提交出错: {str(e)}")
                retry_count += 1

        return success

    def verify_submission(self, driver):
        """多维度验证提交是否成功"""
        # 1. 检查URL特征
        current_url = driver.current_url
        if any(keyword in current_url for keyword in ["complete", "success", "finish", "end", "thank"]):
            return True

        # 2. 检查页面关键元素
        success_selectors = [
            "div.complete",
            "div.survey-complete",
            "div.text-success",
            "img[src*='success']",
            ".survey-success",
            ".end-page",
            ".endtext",
            ".finish-container",
            ".thank-you-page"
        ]

        for selector in success_selectors:
            try:
                if driver.find_element(By.CSS_SELECTOR, selector):
                    return True
            except:
                continue

        # 3. 检查关键文本
        success_phrases = [
            "提交成功", "问卷已完成", "感谢参与",
            "success", "completed", "thank you",
            "问卷提交成功", "提交成功", "已完成",
            "感谢您的参与", "提交完毕", "finish",
            "问卷结束", "谢谢您的参与"
        ]

        page_text = driver.page_source.lower()
        if any(phrase.lower() in page_text for phrase in success_phrases):
            return True

        # 4. 检查错误消息缺失
        error_phrases = [
            "验证码", "错误", "失败", "未提交",
            "error", "fail", "captcha", "未完成",
            "请检查", "不正确", "需要验证"
        ]

        if not any(phrase in page_text for phrase in error_phrases):
            return True

        return False

    # ================== 增强验证码处理 ==================
    def handle_captcha(self, driver):
        """增强的验证码处理"""
        try:
            # 检查多种验证码形式
            captcha_selectors = [
                "div.captcha-container",
                "div.geetest_panel",
                "iframe[src*='captcha']",
                "div#captcha",
                ".geetest_holder",
                ".nc-container",
                ".captcha-modal"
            ]

            # 检查验证码是否存在
            for selector in captcha_selectors:
                try:
                    captcha = driver.find_element(By.CSS_SELECTOR, selector)
                    if captcha.is_displayed():
                        logging.warning("检测到验证码，尝试自动处理")
                        self.pause_for_captcha()
                        return True
                except:
                    continue

            # 检查页面是否有验证码文本提示
            captcha_phrases = ["验证码", "captcha", "验证", "请完成验证"]
            page_text = driver.page_source.lower()
            if any(phrase in page_text for phrase in captcha_phrases):
                logging.warning("页面检测到验证码提示，暂停程序")
                self.pause_for_captcha()
                return True

        except Exception as e:
            logging.error(f"验证码处理出错: {str(e)}")

        return False

    def pause_for_captcha(self):
        """暂停程序并提醒用户处理验证码"""
        self.paused = True
        self.pause_btn.config(text="继续")

        # 创建提醒窗口
        alert = tk.Toplevel(self.root)
        alert.title("需要验证码")
        alert.geometry("400x200")
        alert.resizable(False, False)

        msg = ttk.Label(alert, text="检测到验证码，请手动处理并点击继续", font=("Arial", 12))
        msg.pack(pady=20)

        # 添加倒计时
        countdown_var = tk.StringVar(value="窗口将在 60 秒后自动继续")
        countdown_label = ttk.Label(alert, textvariable=countdown_var, font=("Arial", 10))
        countdown_label.pack(pady=10)

        def resume_after_timeout(seconds=60):
            if seconds > 0:
                countdown_var.set(f"窗口将在 {seconds} 秒后自动继续")
                alert.after(1000, lambda: resume_after_timeout(seconds - 1))
            else:
                self.paused = False
                self.pause_btn.config(text="暂停")
                alert.destroy()

        # 手动继续按钮
        continue_btn = ttk.Button(alert, text="我已处理验证码",
                                  command=lambda: [alert.destroy(), self.toggle_pause()])
        continue_btn.pack(pady=10)

        # 开始倒计时
        resume_after_timeout()

        # 置顶窗口
        alert.attributes('-topmost', True)
        alert.update()
        alert.attributes('-topmost', False)

    # ================== 增强题目类型检测 ==================
    def auto_detect_question_type(self, driver, question, q_num):
        """自动检测题型并填写 - 增强版本"""
        try:
            # 尝试检测单选题
            radio_btns = question.find_elements(By.CSS_SELECTOR, ".ui-radio, input[type='radio']")
            if radio_btns:
                self.fill_single(driver, question, q_num)
                return

            # 尝试检测多选题
            checkboxes = question.find_elements(By.CSS_SELECTOR, ".ui-checkbox, input[type='checkbox']")
            if checkboxes:
                self.fill_multiple(driver, question, q_num)
                return

            # 尝试检测填空题
            text_inputs = question.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
            if text_inputs:
                self.fill_text(question, q_num)
                return

            # 尝试检测量表题
            scale_items = question.find_elements(By.CSS_SELECTOR, ".scale-ul li, .scale-item")
            if scale_items:
                self.fill_scale(driver, question, q_num)
                return

            # 尝试检测矩阵题
            matrix_rows = question.find_elements(By.CSS_SELECTOR, ".matrix tr, .matrix-row")
            if matrix_rows:
                self.fill_matrix(driver, question, q_num)
                return

            # 尝试检测下拉框
            dropdowns = question.find_elements(By.CSS_SELECTOR, "select")
            if dropdowns:
                self.fill_droplist(driver, question, q_num)
                return

            # 尝试检测排序题
            sort_items = question.find_elements(By.CSS_SELECTOR, ".sort-ul li, .sortable-item")
            if sort_items:
                self.fill_reorder(driver, question, q_num)
                return

            logging.warning(f"无法自动检测题目 {q_num} 的类型，尝试通用方法")

            # 通用方法：尝试查找任何可点击的元素
            clickable_elements = question.find_elements(By.CSS_SELECTOR,
                                                        "li, label, div[onclick], span[onclick], .option")
            if clickable_elements:
                try:
                    # 尝试点击一个随机选项
                    element = random.choice(clickable_elements)
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                                          element)
                    time.sleep(0.2)
                    element.click()
                    self.random_delay(*self.config["per_question_delay"])
                    return
                except:
                    pass

            # 最终尝试：查找输入框
            text_inputs = question.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
            if text_inputs:
                self.fill_text(question, q_num)
                return

            logging.warning(f"完全无法处理题目 {q_num}，跳过")
        except Exception as e:
            logging.error(f"自动检测题目类型时出错: {str(e)}")

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

    def fill_single(self, driver, question, q_num):
        """填写单选题"""
        try:
            options = question.find_elements(By.CSS_SELECTOR, f"#div{q_num} .ui-radio")
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
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", selected)
                time.sleep(0.1)
                selected.click()
            except:
                # 如果直接点击失败，使用JavaScript点击
                driver.execute_script("arguments[0].click();", selected)

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写单选题 {q_num} 时出错: {str(e)}")

    def fill_multiple(self, driver, question, q_num):
        """填写多选题"""
        try:
            options = question.find_elements(By.CSS_SELECTOR, f"#div{q_num} .ui-checkbox")
            if not options:
                return

            q_key = str(q_num)
            config = self.config["multiple_prob"].get(q_key, {"prob": [50] * len(options), "min_selection": 1,
                                                              "max_selection": len(options)})
            probs = config["prob"]
            min_selection = config["min_selection"]
            max_selection = config["max_selection"]

            # 确保max_selection不超过选项总数
            if max_selection > len(options):
                max_selection = len(options)

            # 确保min_selection不超过max_selection
            if min_selection > max_selection:
                min_selection = max_selection

            # 确保概率列表长度匹配
            probs = probs[:len(options)] if len(probs) > len(options) else probs + [50] * (len(options) - len(probs))

            # 确定要选择的选项数量
            num_to_select = random.randint(min_selection, max_selection)

            # 根据概率选择选项
            selected_indices = []
            for i, prob in enumerate(probs):
                if random.random() * 100 < prob:
                    selected_indices.append(i)

            # 如果选择的选项数量不足，补充随机选项
            while len(selected_indices) < min_selection:
                remaining = [i for i in range(len(options)) if i not in selected_indices]
                if not remaining:
                    break
                selected_indices.append(random.choice(remaining))

            # 如果选择的选项数量超过最大值，随机移除一些
            while len(selected_indices) > max_selection:
                selected_indices.pop(random.randint(0, len(selected_indices) - 1))

            for idx in selected_indices:
                try:
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                                          options[idx])
                    time.sleep(0.1)
                    options[idx].click()
                except:
                    # 使用JavaScript点击
                    driver.execute_script("arguments[0].click();", options[idx])

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写多选题 {q_num} 时出错: {str(e)}")

    def fill_matrix(self, driver, question, q_num):
        """填写矩阵题"""
        try:
            rows = question.find_elements(By.CSS_SELECTOR, f"#divRefTab{q_num} tbody tr")
            if not rows:
                return

            q_key = str(q_num)
            probs = self.config["matrix_prob"].get(q_num, -1)

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
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                                          cols[selected_col])
                    time.sleep(0.1)
                    cols[selected_col].click()
                except:
                    # 使用JavaScript点击
                    driver.execute_script("arguments[0].click();", cols[selected_col])

                self.random_delay(0.1, 0.3)  # 每行选择后短暂延迟

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写矩阵题 {q_num} 时出错: {str(e)}")

    def fill_scale(self, driver, question, q_num):
        """填写量表题"""
        try:
            options = question.find_elements(By.CSS_SELECTOR, f"#div{q_num} .scale-ul li")
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
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", selected)
                time.sleep(0.1)
                selected.click()
            except:
                driver.execute_script("arguments[0].click();", selected)

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写量表题 {q_num} 时出错: {str(e)}")

    def fill_droplist(self, driver, question, q_num):
        """填写下拉框题"""
        try:
            # 点击下拉框唤出选项
            dropdown = question.find_element(By.CSS_SELECTOR, f"#select2-q{q_num}-container")
            driver.execute_script("arguments[0].scrollIntoView();", dropdown)
            dropdown.click()
            time.sleep(0.3)

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
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", selected)
                time.sleep(0.1)
                selected.click()
            except:
                driver.execute_script("arguments[0].click();", selected)

            self.random_delay(*self.config["per_question_delay"])
        except Exception as e:
            logging.error(f"填写下拉框题 {q_num} 时出错: {str(e)}")

    def fill_reorder(self, driver, question, q_num):
        """填写排序题"""
        try:
            items = question.find_elements(By.CSS_SELECTOR, f"#div{q_num} .sort-ul li")
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
                    time.sleep(0.2)  # 短暂延迟，等待动画完成
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
            self.status_indicator.config(foreground="orange")
        else:
            self.pause_event.set()
            self.pause_btn.config(text="暂停")
            logging.info("已继续")
            self.status_indicator.config(foreground="green")

    def stop_filling(self):
        """停止填写"""
        self.running = False
        self.pause_event.set()  # 确保所有线程都能退出
        self.start_btn.config(state=tk.NORMAL, text="▶ 开始填写")
        self.pause_btn.config(state=tk.DISABLED, text="⏸ 暂停")
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("已停止")
        self.status_indicator.config(foreground="red")
        logging.info("已停止")

    def reset_defaults(self):
        """重置为默认配置"""
        result = messagebox.askyesno("确认", "确定要重置所有设置为默认值吗？")
        if result:
            self.config = DEFAULT_CONFIG.copy()
            # 更新全局设置界面
            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, self.config["url"])
            self.target_entry.set(self.config["target_num"])
            self.ratio_scale.set(self.config["weixin_ratio"])
            self.ratio_var.set(f"{self.config['weixin_ratio'] * 100:.0f}%")
            self.min_duration.set(self.config["min_duration"])
            self.max_duration.set(self.config["max_duration"])
            self.min_delay.set(self.config["min_delay"])
            self.max_delay.set(self.config["max_delay"])
            self.min_q_delay.set(self.config["per_question_delay"][0])
            self.max_q_delay.set(self.config["per_question_delay"][1])
            self.min_p_delay.set(self.config["per_page_delay"][0])
            self.max_p_delay.set(self.config["per_page_delay"][1])
            self.submit_delay.set(self.config["submit_delay"])
            self.num_threads.set(self.config["num_threads"])
            self.use_ip_var.set(self.config["use_ip"])
            self.headless_var.set(self.config["headless"])
            self.ip_entry.delete(0, tk.END)
            self.ip_entry.insert(0, self.config["ip_api"])

            # 重新加载题型设置
            self.reload_question_settings()
            logging.info("已重置为默认配置")

    def save_config(self):
        """保存当前界面配置到self.config - 优化版本"""
        try:
            # 全局设置
            self.config["url"] = self.url_entry.get().strip()
            try:
                self.config["target_num"] = int(self.target_entry.get())
            except ValueError:
                self.config["target_num"] = DEFAULT_CONFIG["target_num"]

            self.config["weixin_ratio"] = self.ratio_scale.get()

            try:
                self.config["min_duration"] = int(self.min_duration.get())
                self.config["max_duration"] = int(self.max_duration.get())
                self.config["min_delay"] = float(self.min_delay.get())
                self.config["max_delay"] = float(self.max_delay.get())
                self.config["submit_delay"] = int(self.submit_delay.get())

                # 处理元组类型的配置
                self.config["per_question_delay"] = (
                    float(self.min_q_delay.get()),
                    float(self.max_q_delay.get())
                )
                self.config["per_page_delay"] = (
                    float(self.min_p_delay.get()),
                    float(self.max_p_delay.get())
                )

                self.config["num_threads"] = int(self.num_threads.get())
            except ValueError:
                # 使用默认值
                self.config.update({
                    "min_duration": DEFAULT_CONFIG["min_duration"],
                    "max_duration": DEFAULT_CONFIG["max_duration"],
                    "min_delay": DEFAULT_CONFIG["min_delay"],
                    "max_delay": DEFAULT_CONFIG["max_delay"],
                    "submit_delay": DEFAULT_CONFIG["submit_delay"],
                    "per_question_delay": DEFAULT_CONFIG["per_question_delay"],
                    "per_page_delay": DEFAULT_CONFIG["per_page_delay"],
                    "num_threads": DEFAULT_CONFIG["num_threads"]
                })

            self.config["use_ip"] = self.use_ip_var.get()
            self.config["headless"] = self.headless_var.get()
            self.config["ip_api"] = self.ip_entry.get().strip()

            # 单选题配置
            for i, (q_num, _) in enumerate(self.config["single_prob"].items()):
                if i < len(self.single_entries):
                    entries = self.single_entries[i]
                    probs = []
                    for entry in entries:
                        val = entry.get().strip()
                        if val == "-1":
                            probs = -1
                            break
                        else:
                            try:
                                probs.append(float(val))
                            except:
                                probs.append(1.0)
                    self.config["single_prob"][q_num] = probs

            # 多选题配置
            for i, (q_num, _) in enumerate(self.config["multiple_prob"].items()):
                if i < len(self.min_selection_entries) and i < len(self.max_selection_entries):
                    min_val = int(self.min_selection_entries[i].get())
                    max_val = int(self.max_selection_entries[i].get())

                    if i < len(self.multi_entries):
                        entries = self.multi_entries[i]
                        probs = []
                        for entry in entries:
                            try:
                                probs.append(int(entry.get()))
                            except:
                                probs.append(50)

                        self.config["multiple_prob"][q_num] = {
                            "prob": probs,
                            "min_selection": min_val,
                            "max_selection": max_val
                        }

            # 矩阵题配置
            for i, (q_num, _) in enumerate(self.config["matrix_prob"].items()):
                if i < len(self.matrix_entries):
                    entries = self.matrix_entries[i]
                    probs = []
                    for entry in entries:
                        val = entry.get().strip()
                        if val == "-1":
                            probs = -1
                            break
                        else:
                            try:
                                probs.append(float(val))
                            except:
                                probs.append(1.0)
                    self.config["matrix_prob"][q_num] = probs

            # 填空题配置
            for i, (q_num, _) in enumerate(self.config["texts"].items()):
                if i < len(self.text_entries):
                    entry = self.text_entries[i]
                    answers = entry.get().split(",")
                    self.config["texts"][q_num] = [ans.strip() for ans in answers if ans.strip()]

            # 多项填空配置
            for i, (q_num, _) in enumerate(self.config["multiple_texts"].items()):
                if i < len(self.multiple_text_entries):
                    entries = self.multiple_text_entries[i]
                    answers_list = []
                    for entry in entries:
                        answers = entry.get().split(",")
                        answers_list.append([ans.strip() for ans in answers if ans.strip()])
                    self.config["multiple_texts"][q_num] = answers_list

            # 排序题配置
            for i, (q_num, _) in enumerate(self.config["reorder_prob"].items()):
                if i < len(self.reorder_entries):
                    entries = self.reorder_entries[i]
                    probs = []
                    for entry in entries:
                        try:
                            probs.append(float(entry.get()))
                        except:
                            probs.append(0.2)
                    self.config["reorder_prob"][q_num] = probs

            # 下拉框配置
            for i, (q_num, _) in enumerate(self.config["droplist_prob"].items()):
                if i < len(self.droplist_entries):
                    entries = self.droplist_entries[i]
                    probs = []
                    for entry in entries:
                        try:
                            probs.append(float(entry.get()))
                        except:
                            probs.append(0.3)
                    self.config["droplist_prob"][q_num] = probs

            # 量表题配置
            for i, (q_num, _) in enumerate(self.config["scale_prob"].items()):
                if i < len(self.scale_entries):
                    entries = self.scale_entries[i]
                    probs = []
                    for entry in entries:
                        try:
                            probs.append(float(entry.get()))
                        except:
                            probs.append(0.2)
                    self.config["scale_prob"][q_num] = probs

            logging.info("配置已保存")
            return True
        except Exception as e:
            logging.error(f"保存配置时出错: {str(e)}")
            messagebox.showerror("错误", f"保存配置时出错: {str(e)}")
            return False

    def random_delay(self, min_time=None, max_time=None):
        """生成随机延迟时间"""
        if min_time is None:
            min_time = self.config["min_delay"]
        if max_time is None:
            max_time = self.config["max_delay"]
        delay = random.uniform(min_time, max_time)
        time.sleep(delay)


if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    root.geometry("1280x900")  # 增大初始窗口尺寸，宽度≥1200
    app = WJXAutoFillApp(root)
    root.mainloop()
