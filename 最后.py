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
import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
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
    "min_submit_gap": 10,  # 单份提交最小间隔（分钟）
    "max_submit_gap": 20,  # 单份提交最大间隔（分钟）
    "batch_size": 5,  # 每N份后暂停
    "batch_pause": 15,  # 批量暂停M分钟
    "per_page_delay": (2.0, 6.0),
    "enable_smart_gap": True,  # 智能提交间隔开关
    "use_ip": False,
    "headless": False,
    "ip_api": "https://service.ipzan.com/core-extract?num=1&minute=1&pool=quality&secret=YOUR_SECRET",
    "num_threads": 4,
    "use_ip": False,
    "ip_api": "https://service.ipzan.com/core-extract?num=1&minute=1&pool=quality&secret=YOUR_SECRET",
    "ip_change_mode": "per_submit",  # 新增, 可选: per_submit, per_batch
    "ip_change_batch": 5,  # 每N份切换, 仅per_batch有效
    # 单选题概率配置
    "single_prob": {
        "1": -1,  # -1表示随机选择
        "2": [0.3, 0.7],  # 数组表示每个选项的选择概率
        "3": [0.2, 0.2, 0.6]
    },
    "other_texts": {
        # 题号: [可选的其他文本1, 2, 3...]
        "4": ["自定义内容A", "自定义内容B", "自定义内容C"],
        "5": ["随便写点", "哈哈哈", "其他情况"]
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
        """创建全局设置界面，包括智能提交间隔和批量休息设置，并支持鼠标滚轮滚动"""
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

        # 鼠标滚轮支持（跨平台）
        def _on_mousewheel(event):
            if event.delta:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")

        # 鼠标进入canvas时绑定滚轮，离开时解绑，防止全局影响
        def _bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind_mousewheel)
        canvas.bind("<Leave>", _unbind_mousewheel)

        # ======== 字体设置 ========
        font_frame = ttk.LabelFrame(scrollable_frame, text="显示设置")
        font_frame.grid(row=0, column=0, columnspan=2, padx=padx, pady=pady, sticky=tk.EW)

        ttk.Label(font_frame, text="字体选择:").grid(row=0, column=0, padx=padx, pady=pady, sticky=tk.W)
        font_options = tkfont.families()
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

        ttk.Label(survey_frame, text="问卷链接:").grid(row=0, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.url_entry = ttk.Entry(survey_frame, width=50)
        self.url_entry.grid(row=0, column=1, columnspan=3, padx=padx, pady=pady, sticky=tk.EW)
        self.url_entry.insert(0, self.config["url"])

        ttk.Label(survey_frame, text="目标份数:").grid(row=1, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.target_entry = ttk.Spinbox(survey_frame, from_=1, to=10000, width=8)
        self.target_entry.grid(row=1, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.target_entry.set(self.config["target_num"])

        ttk.Label(survey_frame, text="微信作答比率:").grid(row=1, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.ratio_scale = ttk.Scale(survey_frame, from_=0, to=1, orient=tk.HORIZONTAL, length=100)
        self.ratio_scale.grid(row=1, column=3, padx=padx, pady=pady, sticky=tk.EW)
        self.ratio_scale.set(self.config["weixin_ratio"])
        self.ratio_var = tk.StringVar()
        self.ratio_var.set(f"{self.config['weixin_ratio'] * 100:.0f}%")
        ratio_label = ttk.Label(survey_frame, textvariable=self.ratio_var, width=4)
        ratio_label.grid(row=1, column=4, padx=(0, padx), pady=pady, sticky=tk.W)
        self.ratio_scale.bind("<Motion>", self.update_ratio_display)
        self.ratio_scale.bind("<ButtonRelease-1>", self.update_ratio_display)

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
        ttk.Label(delay_frame, text="基础延迟(秒):").grid(row=0, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(delay_frame, text="最小:").grid(row=0, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_delay = ttk.Spinbox(delay_frame, from_=0.1, to=10, increment=0.1, width=5)
        self.min_delay.grid(row=0, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.min_delay.set(self.config["min_delay"])
        ttk.Label(delay_frame, text="最大:").grid(row=0, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_delay = ttk.Spinbox(delay_frame, from_=0.1, to=10, increment=0.1, width=5)
        self.max_delay.grid(row=0, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.max_delay.set(self.config["max_delay"])

        ttk.Label(delay_frame, text="每题延迟(秒):").grid(row=1, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(delay_frame, text="最小:").grid(row=1, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_q_delay = ttk.Spinbox(delay_frame, from_=0.1, to=5, increment=0.1, width=5)
        self.min_q_delay.grid(row=1, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.min_q_delay.set(self.config["per_question_delay"][0])
        ttk.Label(delay_frame, text="最大:").grid(row=1, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_q_delay = ttk.Spinbox(delay_frame, from_=0.1, to=5, increment=0.1, width=5)
        self.max_q_delay.grid(row=1, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.max_q_delay.set(self.config["per_question_delay"][1])

        ttk.Label(delay_frame, text="页面延迟(秒):").grid(row=2, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(delay_frame, text="最小:").grid(row=2, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_p_delay = ttk.Spinbox(delay_frame, from_=0.1, to=10, increment=0.1, width=5)
        self.min_p_delay.grid(row=2, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.min_p_delay.set(self.config["per_page_delay"][0])
        ttk.Label(delay_frame, text="最大:").grid(row=2, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_p_delay = ttk.Spinbox(delay_frame, from_=0.1, to=10, increment=0.1, width=5)
        self.max_p_delay.grid(row=2, column=4, padx=padx, pady=pady, sticky=tk.W)
        self.max_p_delay.set(self.config["per_page_delay"][1])

        ttk.Label(delay_frame, text="提交延迟:").grid(row=3, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.submit_delay = ttk.Spinbox(delay_frame, from_=1, to=10, width=5)
        self.submit_delay.grid(row=3, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.submit_delay.set(self.config["submit_delay"])

        # ======== 智能提交间隔设置 ========
        smart_gap_frame = ttk.LabelFrame(scrollable_frame, text="智能提交间隔")
        smart_gap_frame.grid(row=3, column=0, columnspan=2, padx=padx, pady=pady, sticky=tk.EW)
        self.enable_smart_gap_var = tk.BooleanVar(value=self.config.get("enable_smart_gap", True))
        smart_gap_switch = ttk.Checkbutton(
            smart_gap_frame, text="开启智能提交间隔与批量休息", variable=self.enable_smart_gap_var)
        smart_gap_switch.grid(row=0, column=0, padx=padx, pady=pady, sticky=tk.W, columnspan=5)
        ttk.Label(smart_gap_frame, text="单份提交间隔(分钟):").grid(row=1, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.min_submit_gap = ttk.Spinbox(smart_gap_frame, from_=1, to=120, width=5)
        self.min_submit_gap.grid(row=1, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.min_submit_gap.set(self.config.get("min_submit_gap", 10))
        ttk.Label(smart_gap_frame, text="~").grid(row=1, column=2, padx=2, pady=pady, sticky=tk.W)
        self.max_submit_gap = ttk.Spinbox(smart_gap_frame, from_=1, to=180, width=5)
        self.max_submit_gap.grid(row=1, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.max_submit_gap.set(self.config.get("max_submit_gap", 20))
        ttk.Label(smart_gap_frame, text="每").grid(row=2, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.batch_size = ttk.Spinbox(smart_gap_frame, from_=1, to=100, width=5)
        self.batch_size.grid(row=2, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.batch_size.set(self.config.get("batch_size", 5))
        ttk.Label(smart_gap_frame, text="份后暂停").grid(row=2, column=2, padx=2, pady=pady, sticky=tk.W)
        self.batch_pause = ttk.Spinbox(smart_gap_frame, from_=1, to=120, width=5)
        self.batch_pause.grid(row=2, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.batch_pause.set(self.config.get("batch_pause", 15))
        ttk.Label(smart_gap_frame, text="分钟").grid(row=2, column=4, padx=2, pady=pady, sticky=tk.W)

        # ======== 高级设置 ========
        advanced_frame = ttk.LabelFrame(scrollable_frame, text="高级设置")
        advanced_frame.grid(row=4, column=0, columnspan=2, padx=padx, pady=pady, sticky=tk.EW)
        ttk.Label(advanced_frame, text="浏览器窗口数量:").grid(row=0, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.num_threads = ttk.Spinbox(advanced_frame, from_=1, to=10, width=5)
        self.num_threads.grid(row=0, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.num_threads.set(self.config["num_threads"])
        self.use_ip_var = tk.BooleanVar(value=self.config["use_ip"])
        ttk.Checkbutton(advanced_frame, text="使用代理IP", variable=self.use_ip_var).grid(
            row=1, column=0, padx=padx, pady=pady, sticky=tk.W)
        ttk.Label(advanced_frame, text="IP API:").grid(row=1, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.ip_entry = ttk.Entry(advanced_frame, width=40)
        self.ip_entry.grid(row=1, column=2, columnspan=3, padx=padx, pady=pady, sticky=tk.EW)
        self.ip_entry.insert(0, self.config["ip_api"])
        ttk.Label(advanced_frame, text="代理切换:").grid(row=2, column=0, padx=padx, pady=pady, sticky=tk.W)
        self.ip_change_mode = ttk.Combobox(advanced_frame, values=["per_submit", "per_batch"], width=12)
        self.ip_change_mode.grid(row=2, column=1, padx=padx, pady=pady, sticky=tk.W)
        self.ip_change_mode.set(self.config.get("ip_change_mode", "per_submit"))
        ttk.Label(advanced_frame, text="每N份切换:").grid(row=2, column=2, padx=padx, pady=pady, sticky=tk.W)
        self.ip_change_batch = ttk.Spinbox(advanced_frame, from_=1, to=100, width=5)
        self.ip_change_batch.grid(row=2, column=3, padx=padx, pady=pady, sticky=tk.W)
        self.ip_change_batch.set(self.config.get("ip_change_batch", 5))
        self.headless_var = tk.BooleanVar(value=self.config["headless"])
        ttk.Checkbutton(advanced_frame, text="无头模式(不显示浏览器)", variable=self.headless_var).grid(
            row=3, column=0, padx=padx, pady=pady, sticky=tk.W)

        # ======== 操作按钮 ========
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10, sticky=tk.W)
        self.parse_btn = ttk.Button(button_frame, text="解析问卷", command=self.parse_survey, width=15)
        self.parse_btn.grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="重置默认", command=self.reset_defaults, width=15).grid(row=0, column=1, padx=5)
        scrollable_frame.columnconfigure(0, weight=1)
        tip_label = ttk.Label(scrollable_frame, text="提示: 填写前请先解析问卷以获取题目结构", style='Warning.TLabel')
        tip_label.grid(row=6, column=0, columnspan=2, pady=(10, 0))

    def _process_parsed_questions(self, questions_data):
        """处理解析得到的问卷题目数据，包括自动识别多选题中的“其他”并初始化other_texts"""
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
            # === 新增: 初始化other_texts ===
            if "other_texts" not in self.config:
                self.config["other_texts"] = {}

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
                    # === 新增: 自动检测“其他”选项并初始化other_texts ===
                    for opt in options:
                        if "其他" in opt or "other" in opt.lower():
                            if question_id not in self.config["other_texts"]:
                                # 可以自定义默认内容
                                self.config["other_texts"][question_id] = ["其他：自定义答案1", "其他：自定义答案2",
                                                                           "其他：自定义答案3"]
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
                            '.wjx-option-label',           // 新增，适配问卷星新版
                            '.ui-radio', 
                            '.ui-checkbox', 
                            'label[for]',                  // 适配常见label
                            '.matrix th', 
                            '.scale-ul li', 
                            '.sort-ul li',
                            'select option',
                            '.option-text',
                            '.option-item',
                            '.option-label',
                            'label'
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
        """创建单选题设置界面 - 输入框数量严格等于选项数量"""
        padx, pady = 8, 5
        # 配置说明框架
        desc_frame = ttk.LabelFrame(frame, text="单选题配置说明")
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)
        ttk.Label(desc_frame, text="配置说明：", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(desc_frame, text="• 输入 -1 表示随机选择\n• 输入正数表示选项的相对权重",
                  justify=tk.LEFT, font=("Arial", 9)).pack(anchor=tk.W, padx=5)

        # 创建表格框架并设置列权重
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        table_frame.columnconfigure(0, weight=1)  # 题号列
        table_frame.columnconfigure(1, weight=3)  # 题目预览列
        table_frame.columnconfigure(2, weight=8)  # 选项权重配置列
        table_frame.columnconfigure(3, weight=2)  # 操作列

        # 表头
        headers = ["题号", "题目预览", "选项权重配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)

        for row_idx, (q_num, probs) in enumerate(self.config["single_prob"].items(), start=1):
            base_row = row_idx
            q_text = self.config["question_texts"].get(q_num, f"单选题 {q_num}")

            # 关键：严格用选项数量决定输入框数量
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0:
                option_count = 1  # 防止解析失败时至少有一个输入框

            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)
            tooltip_text = f"题目类型: 单选题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)

            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)

            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)

            entry_row = []
            for opt_idx in range(option_count):
                opt_container = ttk.Frame(option_frame)
                opt_container.grid(row=opt_idx, column=0, sticky=tk.W, pady=2)

                opt_label = ttk.Label(opt_container, text=f"选项 {opt_idx + 1}: ", width=8)
                opt_label.pack(side=tk.LEFT, padx=(0, 5))

                entry = ttk.Entry(opt_container, width=8)
                if isinstance(probs, list) and opt_idx < len(probs):
                    entry.insert(0, str(probs[opt_idx]))
                elif probs == -1:
                    entry.insert(0, "-1")
                else:
                    entry.insert(0, "1")
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)
            self.single_entries.append(entry_row)

            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)

            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)

            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row1, text="偏左", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("single", "left", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row1, text="偏右", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("single", "right", q, e)).pack(
                side=tk.LEFT, padx=2)

            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("single", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row2, text="平均", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_average("single", q, e)).pack(
                side=tk.LEFT, padx=2)

            if row_idx < len(self.config["single_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(
                    row=base_row + 1, column=0, columnspan=4, sticky='ew', pady=10)

    def create_multi_settings(self, frame):
        """
        多选题配置界面 动态生成输入框，并为‘其他’选项生成自定义文本框
        """
        padx, pady = 8, 5
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)
        ttk.Label(desc_frame, text="多选题配置说明：", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(desc_frame, text="• 每个选项概率范围为0-100，表示该选项被选中的独立概率", font=("Arial", 9)).pack(
            anchor=tk.W)
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        table_frame.columnconfigure(0, weight=1)
        table_frame.columnconfigure(1, weight=3)
        table_frame.columnconfigure(2, weight=1)
        table_frame.columnconfigure(3, weight=1)
        table_frame.columnconfigure(4, weight=5)
        table_frame.columnconfigure(5, weight=2)
        headers = ["题号", "题目预览", "最小选择数", "最大选择数", "选项概率配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)
        # 收集“其他”选项自定义内容框
        self.other_entries = {}

        for row_idx, (q_num, config) in enumerate(self.config["multiple_prob"].items(), start=1):
            base_row = row_idx
            q_text = self.config["question_texts"].get(q_num, f"多选题 {q_num}")
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0:
                option_count = 1
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)
            tooltip_text = f"题目类型: 多选题\n\n{q_text}"
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)
            min_frame = ttk.Frame(table_frame)
            min_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)
            min_entry = ttk.Spinbox(min_frame, from_=1, to=option_count, width=4)
            min_entry.set(config.get("min_selection", 1))
            min_entry.pack(fill=tk.X, expand=True)
            self.min_selection_entries.append(min_entry)
            ttk.Label(min_frame, text="最少选择项数", font=("Arial", 8), foreground="gray").pack(fill=tk.X, expand=True)
            max_frame = ttk.Frame(table_frame)
            max_frame.grid(row=base_row, column=3, padx=padx, pady=pady, sticky=tk.NSEW)
            max_entry = ttk.Spinbox(max_frame, from_=1, to=option_count, width=4)
            max_entry.set(config.get("max_selection", option_count))
            max_entry.pack(fill=tk.X, expand=True)
            self.max_selection_entries.append(max_entry)
            ttk.Label(max_frame, text="最多选择项数", font=("Arial", 8), foreground="gray").pack(fill=tk.X, expand=True)
            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=4, padx=padx, pady=pady, sticky=tk.NSEW)
            option_frame.columnconfigure(0, weight=1)
            entry_row = []
            for opt_idx in range(option_count):
                opt_container = ttk.Frame(option_frame)
                opt_container.grid(row=opt_idx, column=0, sticky=tk.W, pady=2)
                opt_label = ttk.Label(opt_container, text=f"选项 {opt_idx + 1}: ", width=8)
                opt_label.pack(side=tk.LEFT, padx=(0, 5))
                entry = ttk.Entry(opt_container, width=8)
                if isinstance(config["prob"], list) and opt_idx < len(config["prob"]):
                    entry.insert(0, config["prob"][opt_idx])
                else:
                    entry.insert(0, 50)
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)
                # ==== 新增：如果是“其他”，加一个可编辑答案的文本框 ====
                option_texts = self.config["option_texts"].get(q_num, [])
                if opt_idx < len(option_texts):
                    if "其他" in option_texts[opt_idx] or "other" in option_texts[opt_idx].lower():
                        # “其他”答案输入框
                        other_edit = ttk.Entry(opt_container, width=30)
                        other_values = self.config.get("other_texts", {}).get(q_num, ["请自定义其他答案"])
                        other_edit.insert(0, ", ".join(other_values))
                        other_edit.pack(side=tk.LEFT, padx=(10, 0))
                        self.other_entries[q_num] = other_edit
                        ttk.Label(opt_container, text="（多个答案用逗号隔开）", font=("Arial", 8),
                                  foreground="gray").pack(side=tk.LEFT)
            self.multi_entries.append(entry_row)
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=5, padx=5, pady=5, sticky=tk.NW)
            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)
            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row1, text="偏左", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("multiple", "left", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row1, text="偏右", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("multiple", "right", q, e)).pack(
                side=tk.LEFT, padx=2)
            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("multiple", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row2, text="50%", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_value("multiple", q, e, 50)).pack(
                side=tk.LEFT, padx=2)
            if row_idx < len(self.config["multiple_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(row=base_row + 1, column=0, columnspan=6,
                                                                     sticky='ew', pady=10)

    def create_text_settings(self, frame):
        """填空题配置界面 动态生成输入框"""
        padx, pady = 8, 5
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)
        ttk.Label(desc_frame, text="填空题配置说明：", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(desc_frame, text="• 输入多个答案时用逗号分隔\n• 系统会随机选择一个答案填写", justify=tk.LEFT,
                  font=("Arial", 9)).pack(anchor=tk.W, padx=5)
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        table_frame.columnconfigure(0, weight=1)
        table_frame.columnconfigure(1, weight=3)
        table_frame.columnconfigure(2, weight=5)
        table_frame.columnconfigure(3, weight=2)
        headers = ["题号", "题目预览", "答案配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)
        for row_idx, (q_num, answers) in enumerate(self.config["texts"].items(), start=1):
            base_row = row_idx
            q_text = self.config["question_texts"].get(q_num, f"填空题 {q_num}")
            # 动态决定有几个空，通常填空题只有一个输入框，特殊问卷可能有多个空
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0:
                option_count = 1
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)
            tooltip_text = f"题目类型: 填空题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)
            answer_frame = ttk.Frame(table_frame)
            answer_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)
            entry_row = []
            for i in range(option_count):
                entry = ttk.Entry(answer_frame, width=40)
                entry.pack(fill=tk.X, padx=5, pady=2)
                answer_str = ", ".join(answers) if i == 0 else ""
                entry.insert(0, answer_str)
                entry_row.append(entry)
            self.text_entries.append(entry_row)
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)
            reset_btn = ttk.Button(btn_frame, text="重置", width=8,
                                   command=lambda e=entry_row: [ent.delete(0, tk.END) or ent.insert(0, "示例答案") for
                                                                ent in e])
            reset_btn.pack(pady=2)
            if row_idx < len(self.config["texts"]):
                ttk.Separator(table_frame, orient='horizontal').grid(row=base_row + 1, column=0, columnspan=4,
                                                                     sticky='ew', pady=10)

    def create_multiple_text_settings(self, frame):
        """多项填空配置界面 动态生成输入框"""
        padx, pady = 8, 5
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)
        ttk.Label(desc_frame, text="多项填空配置说明：", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(desc_frame, text="• 每个输入框对应一个空的答案配置\n• 多个答案用逗号分隔", justify=tk.LEFT,
                  font=("Arial", 9)).pack(anchor=tk.W, padx=5)
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        table_frame.columnconfigure(0, weight=1)
        table_frame.columnconfigure(1, weight=3)
        table_frame.columnconfigure(2, weight=5)
        table_frame.columnconfigure(3, weight=2)
        headers = ["题号", "题目预览", "答案配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)
        for row_idx, (q_num, answers_list) in enumerate(self.config["multiple_texts"].items(), start=1):
            base_row = row_idx
            q_text = self.config["question_texts"].get(q_num, f"多项填空 {q_num}")
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0:
                option_count = len(answers_list)
            if option_count == 0:
                option_count = 1
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)
            tooltip_text = f"题目类型: 多项填空\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)
            answer_frame = ttk.Frame(table_frame)
            answer_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)
            entry_row = []
            for i in range(option_count):
                field_frame = ttk.Frame(answer_frame)
                field_frame.pack(fill=tk.X, pady=2)
                ttk.Label(field_frame, text=f"空 {i + 1}: ", width=6).pack(side=tk.LEFT, padx=(0, 5))
                entry = ttk.Entry(field_frame, width=40)
                entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
                answer_str = ", ".join(answers_list[i]) if i < len(answers_list) else ""
                entry.insert(0, answer_str)
                entry_row.append(entry)
            self.multiple_text_entries.append(entry_row)
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)
            reset_btn = ttk.Button(btn_frame, text="重置", width=8,
                                   command=lambda e=entry_row: [ent.delete(0, tk.END) or ent.insert(0, "示例答案") for
                                                                ent in e])
            reset_btn.pack(pady=2)
            if row_idx < len(self.config["multiple_texts"]):
                ttk.Separator(table_frame, orient='horizontal').grid(row=base_row + 1, column=0, columnspan=4,
                                                                     sticky='ew', pady=10)
    def create_matrix_settings(self, frame):
        """矩阵题配置界面 动态生成输入框"""
        padx, pady = 8, 5
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)
        ttk.Label(desc_frame, text="矩阵题配置说明：", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(desc_frame, text="• 输入 -1 表示随机选择\n• 输入正数表示选项的相对权重", font=("Arial", 9)).pack(
            anchor=tk.W, padx=5)
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        table_frame.columnconfigure(0, weight=1)
        table_frame.columnconfigure(1, weight=3)
        table_frame.columnconfigure(2, weight=5)
        table_frame.columnconfigure(3, weight=2)
        headers = ["题号", "题目预览", "选项配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)
        for row_idx, (q_num, probs) in enumerate(self.config["matrix_prob"].items(), start=1):
            base_row = row_idx
            q_text = self.config["question_texts"].get(q_num, f"矩阵题 {q_num}")
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0:
                option_count = 1
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)
            tooltip_text = f"题目类型: 矩阵题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)
            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)
            option_frame.columnconfigure(0, weight=1)
            entry_row = []
            for opt_idx in range(option_count):
                opt_container = ttk.Frame(option_frame)
                opt_container.grid(row=opt_idx, column=0, sticky=tk.W, pady=2)
                opt_label = ttk.Label(opt_container, text=f"选项 {opt_idx + 1}: ", width=8)
                opt_label.pack(side=tk.LEFT, padx=(0, 5))
                entry = ttk.Entry(opt_container, width=8)
                if isinstance(probs, list) and opt_idx < len(probs):
                    entry.insert(0, str(probs[opt_idx]))
                elif probs == -1:
                    entry.insert(0, "-1")
                else:
                    entry.insert(0, "1")
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)
            self.matrix_entries.append(entry_row)
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)
            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)
            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row1, text="偏左", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("matrix", "left", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row1, text="偏右", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("matrix", "right", q, e)).pack(
                side=tk.LEFT, padx=2)
            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("matrix", q, e)).pack(side=tk.LEFT,
                                                                                                           padx=2)
            ttk.Button(btn_row2, text="平均", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_average("matrix", q, e)).pack(
                side=tk.LEFT, padx=2)
            if row_idx < len(self.config["matrix_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(row=base_row + 1, column=0, columnspan=4,
                                                                     sticky='ew', pady=10)

    def create_reorder_settings(self, frame):
        """排序题配置界面 动态生成输入框"""
        padx, pady = 8, 5
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)
        ttk.Label(desc_frame, text="排序题配置说明：", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(desc_frame, text="• 每个位置的概率表示该位置被选中的相对权重\n• 概率越高，该位置被选中的几率越大",
                  font=("Arial", 9)).pack(anchor=tk.W, padx=5)
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        table_frame.columnconfigure(0, weight=1)
        table_frame.columnconfigure(1, weight=3)
        table_frame.columnconfigure(2, weight=5)
        table_frame.columnconfigure(3, weight=2)
        headers = ["题号", "题目预览", "位置概率配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)
        for row_idx, (q_num, probs) in enumerate(self.config["reorder_prob"].items(), start=1):
            base_row = row_idx
            q_text = self.config["question_texts"].get(q_num, f"排序题 {q_num}")
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0:
                option_count = 1
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)
            tooltip_text = f"题目类型: 排序题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)
            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)
            option_frame.columnconfigure(0, weight=1)
            entry_row = []
            for pos_idx in range(option_count):
                pos_container = ttk.Frame(option_frame)
                pos_container.grid(row=pos_idx, column=0, sticky=tk.W, pady=2)
                pos_label = ttk.Label(pos_container, text=f"位置 {pos_idx + 1}: ", width=8)
                pos_label.pack(side=tk.LEFT, padx=(0, 5))
                entry = ttk.Entry(pos_container, width=8)
                if isinstance(probs, list) and pos_idx < len(probs):
                    entry.insert(0, str(probs[pos_idx]))
                else:
                    entry.insert(0, f"{1 / option_count:.2f}")
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)
            self.reorder_entries.append(entry_row)
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)
            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)
            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row1, text="偏前", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("reorder", "left", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row1, text="偏后", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("reorder", "right", q, e)).pack(
                side=tk.LEFT, padx=2)
            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("reorder", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row2, text="平均", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_average("reorder", q, e)).pack(
                side=tk.LEFT, padx=2)
            if row_idx < len(self.config["reorder_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(row=base_row + 1, column=0, columnspan=4,
                                                                     sticky='ew', pady=10)

    def create_droplist_settings(self, frame):
        """下拉框题配置界面 动态生成输入框"""
        padx, pady = 8, 5
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)
        ttk.Label(desc_frame, text="下拉框配置说明：", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(desc_frame, text="• 每个选项的概率表示该选项被选中的相对权重\n• 概率越高，该选项被选中的几率越大",
                  font=("Arial", 9)).pack(anchor=tk.W, padx=5)
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        table_frame.columnconfigure(0, weight=1)
        table_frame.columnconfigure(1, weight=3)
        table_frame.columnconfigure(2, weight=5)
        table_frame.columnconfigure(3, weight=2)
        headers = ["题号", "题目预览", "选项概率配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)
        for row_idx, (q_num, probs) in enumerate(self.config["droplist_prob"].items(), start=1):
            base_row = row_idx
            q_text = self.config["question_texts"].get(q_num, f"下拉框题 {q_num}")
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0:
                option_count = 1
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)
            tooltip_text = f"题目类型: 下拉框题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)
            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)
            option_frame.columnconfigure(0, weight=1)
            entry_row = []
            for opt_idx in range(option_count):
                opt_container = ttk.Frame(option_frame)
                opt_container.grid(row=opt_idx, column=0, sticky=tk.W, pady=2)
                opt_label = ttk.Label(opt_container, text=f"选项 {opt_idx + 1}: ", width=8)
                opt_label.pack(side=tk.LEFT, padx=(0, 5))
                entry = ttk.Entry(opt_container, width=8)
                if isinstance(probs, list) and opt_idx < len(probs):
                    entry.insert(0, str(probs[opt_idx]))
                else:
                    entry.insert(0, "0.3")
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)
            self.droplist_entries.append(entry_row)
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)
            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)
            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row1, text="偏前", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("droplist", "left", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row1, text="偏后", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("droplist", "right", q, e)).pack(
                side=tk.LEFT, padx=2)
            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("droplist", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row2, text="平均", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_average("droplist", q, e)).pack(
                side=tk.LEFT, padx=2)
            if row_idx < len(self.config["droplist_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(row=base_row + 1, column=0, columnspan=4,
                                                                     sticky='ew', pady=10)

    def create_scale_settings(self, frame):
        """量表题配置界面 动态生成输入框"""
        padx, pady = 8, 5
        desc_frame = ttk.Frame(frame)
        desc_frame.pack(fill=tk.X, padx=padx, pady=pady)
        ttk.Label(desc_frame, text="量表题配置说明：", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(desc_frame, text="• 输入概率值表示该刻度被选中的相对概率\n• 概率越高，该刻度被选中的几率越大",
                  font=("Arial", 9)).pack(anchor=tk.W, padx=5)
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        table_frame.columnconfigure(0, weight=1)
        table_frame.columnconfigure(1, weight=3)
        table_frame.columnconfigure(2, weight=5)
        table_frame.columnconfigure(3, weight=2)
        headers = ["题号", "题目预览", "刻度概率配置", "操作"]
        for col, header in enumerate(headers):
            header_label = ttk.Label(table_frame, text=header, font=("Arial", 9, "bold"))
            header_label.grid(row=0, column=col, padx=padx, pady=pady, sticky=tk.W)
        for row_idx, (q_num, probs) in enumerate(self.config["scale_prob"].items(), start=1):
            base_row = row_idx
            q_text = self.config["question_texts"].get(q_num, f"量表题 {q_num}")
            option_count = len(self.config["option_texts"].get(q_num, []))
            if option_count == 0:
                option_count = 1
            q_label = ttk.Label(table_frame, text=f"第{q_num}题", cursor="hand2", font=("Arial", 10))
            q_label.grid(row=base_row, column=0, padx=padx, pady=pady, sticky=tk.NW)
            tooltip_text = f"题目类型: 量表题\n\n{q_text}"
            tooltip = ToolTip(q_label, tooltip_text, wraplength=400)
            self.tooltips.append(tooltip)
            preview_text = (q_text[:30] + '...') if len(q_text) > 30 else q_text
            preview_label = ttk.Label(table_frame, text=preview_text, width=25, wraplength=200)
            preview_label.grid(row=base_row, column=1, padx=padx, pady=pady, sticky=tk.NW)
            preview_tooltip = ToolTip(preview_label, tooltip_text, wraplength=400)
            self.tooltips.append(preview_tooltip)
            option_frame = ttk.Frame(table_frame)
            option_frame.grid(row=base_row, column=2, padx=padx, pady=pady, sticky=tk.NSEW)
            option_frame.columnconfigure(0, weight=1)
            entry_row = []
            for opt_idx in range(option_count):
                opt_container = ttk.Frame(option_frame)
                opt_container.grid(row=opt_idx, column=0, sticky=tk.W, pady=2)
                opt_label = ttk.Label(opt_container, text=f"刻度 {opt_idx + 1}: ", width=8)
                opt_label.pack(side=tk.LEFT, padx=(0, 5))
                entry = ttk.Entry(opt_container, width=8)
                if isinstance(probs, list) and opt_idx < len(probs):
                    entry.insert(0, str(probs[opt_idx]))
                else:
                    entry.insert(0, "0.2")
                entry.pack(side=tk.LEFT, padx=(0, 10))
                entry_row.append(entry)
            self.scale_entries.append(entry_row)
            btn_frame = ttk.Frame(table_frame)
            btn_frame.grid(row=base_row, column=3, padx=5, pady=5, sticky=tk.NW)
            btn_grid = ttk.Frame(btn_frame)
            btn_grid.pack(fill=tk.BOTH, expand=True)
            btn_row1 = ttk.Frame(btn_grid)
            btn_row1.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row1, text="偏左", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("scale", "left", q, e)).pack(
                side=tk.LEFT, padx=2)
            ttk.Button(btn_row1, text="偏右", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_bias("scale", "right", q, e)).pack(
                side=tk.LEFT, padx=2)
            btn_row2 = ttk.Frame(btn_grid)
            btn_row2.pack(fill=tk.X, pady=2)
            ttk.Button(btn_row2, text="随机", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_random("scale", q, e)).pack(side=tk.LEFT,
                                                                                                          padx=2)
            ttk.Button(btn_row2, text="平均", width=6,
                       command=lambda q=q_num, e=entry_row: self.set_question_average("scale", q, e)).pack(side=tk.LEFT,
                                                                                                           padx=2)
            if row_idx < len(self.config["scale_prob"]):
                ttk.Separator(table_frame, orient='horizontal').grid(row=base_row + 1, column=0, columnspan=4,
                                                                     sticky='ew', pady=10)

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
        """
        运行填写任务 - 可用微信作答比率滑动条控制微信来源填写比例
        """
        import random
        import time
        from selenium import webdriver

        driver = None
        submit_count = 0
        proxy_ip = None

        # 微信和PC UA
        WECHAT_UA = (
            "Mozilla/5.0 (Linux; Android 10; MI 8 Build/QKQ1.190828.002; wv) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/86.0.4240.99 "
            "XWEB/4317 MMWEBSDK/20220105 Mobile Safari/537.36 "
            "MicroMessenger/8.0.18.2040(0x28001235) "
            "Process/toolsmp WeChat/arm64 NetType/WIFI Language/zh_CN ABI/arm64"
        )
        PC_UA = (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/91.0.4472.124 Safari/537.36"
        )

        try:
            while self.running and self.cur_num < self.config["target_num"]:
                if self.paused:
                    time.sleep(1)
                    continue

                # 1. 用滑动条控制微信来源比例
                # self.config["weixin_ratio"]已实时跟随滑动条
                use_weixin = random.random() < float(self.config.get("weixin_ratio", 0.5))

                # 2. 配置chromedriver选项
                options = webdriver.ChromeOptions()
                options.add_experimental_option("excludeSwitches", ["enable-automation"])
                options.add_experimental_option("useAutomationExtension", False)
                options.add_argument('--disable-blink-features=AutomationControlled')
                ua = WECHAT_UA if use_weixin else PC_UA
                options.add_argument(f'--user-agent={ua}')
                options.add_argument('--disable-gpu')
                options.add_argument('--no-sandbox')
                options.add_argument('--disable-dev-shm-usage')
                if self.config["headless"]:
                    options.add_argument('--headless')
                else:
                    options.add_argument(f'--window-position={x},{y}')

                # 3. 代理设置
                use_ip = self.config.get("use_ip", False)
                ip_mode = self.config.get("ip_change_mode", "per_submit")
                ip_batch = self.config.get("ip_change_batch", 5)
                need_new_proxy = False
                if use_ip:
                    if ip_mode == "per_submit":
                        need_new_proxy = True
                    elif ip_mode == "per_batch":
                        if submit_count % ip_batch == 0:
                            need_new_proxy = True
                if use_ip and need_new_proxy:
                    proxy_ip = self.get_new_proxy()
                    if proxy_ip:
                        logging.info(f"使用代理: {proxy_ip}")
                        options.add_argument(f'--proxy-server={proxy_ip}')
                    else:
                        logging.error("本次未获取到有效代理，等待10秒后重试。")
                        time.sleep(10)
                        continue
                elif use_ip and proxy_ip:
                    options.add_argument(f'--proxy-server={proxy_ip}')

                driver = webdriver.Chrome(options=options)
                try:
                    # 4. 设置窗口为手机尺寸以模拟微信端访问
                    if not self.config["headless"]:
                        if use_weixin:
                            driver.set_window_size(375, 812)
                        else:
                            driver.set_window_size(1024, 768)

                    logging.info(f"本次作答方式: {'微信来源' if use_weixin else '普通渠道'} (UA已切换)")

                    driver.get(self.config["url"])
                    time.sleep(self.config["page_load_delay"])

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
                    import traceback
                    traceback.print_exc()
                finally:
                    try:
                        driver.quit()
                    except:
                        pass

                submit_count += 1

                # 智能提交间隔/批量休息机制（按原逻辑）
                if self.config.get("enable_smart_gap", True):
                    if self.running and self.cur_num < self.config["target_num"]:
                        batch_size = self.config.get("batch_size", 0)
                        batch_pause = self.config.get("batch_pause", 0)
                        if batch_size > 0 and self.cur_num % batch_size == 0:
                            logging.info(f"已完成{self.cur_num}份，批量休息{batch_pause}分钟...")
                            for i in range(int(batch_pause * 60)):
                                if not self.running:
                                    break
                                time.sleep(1)
                        else:
                            min_gap = self.config.get("min_submit_gap", 10)
                            max_gap = self.config.get("max_submit_gap", 20)
                            if min_gap > max_gap:
                                min_gap, max_gap = max_gap, min_gap
                            submit_interval = random.uniform(min_gap, max_gap) * 60
                            logging.info(f"本次提交后等待{submit_interval / 60:.2f}分钟...")
                            for i in range(int(submit_interval)):
                                if not self.running:
                                    break
                                time.sleep(1)
                # 若不开启，直接跳过间隔与批量休息

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
            # 填写完所有题目后
            self.repair_required_questions(driver)
            # 提交问卷
            return self.submit_survey(driver)

        except Exception as e:
            logging.error(f"填写问卷过程中出错: {str(e)}")
            return False

    def repair_required_questions(self, driver):
        """
        检查所有必答项，自动补全未填写项，包括“其他”多选题下的必答填空。
        """
        try:
            questions = driver.find_elements(By.CSS_SELECTOR, ".div_question, .field, .question")
            for q in questions:
                is_required = False
                # 判断必答标记
                try:
                    if q.find_element(By.CSS_SELECTOR, ".required, .star, .necessary, .wjxnecessary"):
                        is_required = True
                except:
                    if "必答" in q.text or q.get_attribute("data-required") == "1":
                        is_required = True
                if not is_required:
                    continue

                all_inputs = q.find_elements(By.CSS_SELECTOR, "input, textarea, select")
                any_filled = False
                for inp in all_inputs:
                    typ = inp.get_attribute("type")
                    if typ in ("checkbox", "radio"):
                        if inp.is_selected():
                            any_filled = True
                            # 检查“其他”选项的填空
                            if "其他" in inp.get_attribute("value") or "other" in (inp.get_attribute("id") or ""):
                                try:
                                    other_text = q.find_element(By.CSS_SELECTOR, "input[type='text'], textarea")
                                    if not other_text.get_attribute("value"):
                                        other_text.send_keys("自动补全内容")
                                except:
                                    pass
                    elif typ in ("text", None):
                        if inp.get_attribute("value"):
                            any_filled = True
                    elif typ == "select-one":
                        if inp.get_attribute("value"):
                            any_filled = True
                # 未填写自动补全
                if not any_filled:
                    self.auto_fill_question(driver, q)
        except Exception as e:
            logging.warning(f"自动修复必答题时出错: {e}")

    def auto_fill_question(self, driver, question):
        """
        自动补全问题 - 修复版，确保多选题中的'其他'文本必填
        """
        import random
        from selenium.webdriver.common.by import By
        from selenium.common.exceptions import StaleElementReferenceException

        try:
            # 1. 单选题
            try:
                radios = question.find_elements(By.CSS_SELECTOR, "input[type='radio']")
                if radios:
                    random.choice(radios).click()
                    return
            except StaleElementReferenceException:
                pass

            # 2. 多选题
            try:
                checks = question.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
                if checks:
                    # 随机勾选一个
                    chosen = random.choice(checks)
                    try:
                        chosen.click()
                    except:
                        driver.execute_script("arguments[0].click();", chosen)

                    # 获取选项文本
                    option_labels = []
                    label_elems = question.find_elements(By.CSS_SELECTOR, "label")
                    for el in label_elems:
                        try:
                            txt = el.text.strip()
                            if not txt:
                                spans = el.find_elements(By.CSS_SELECTOR, "span")
                                if spans:
                                    txt = spans[0].text.strip()
                            option_labels.append(txt)
                        except StaleElementReferenceException:
                            option_labels.append("")

                    # 检查是否有"其他"选项被选中
                    chose_other = False
                    for idx, chk in enumerate(checks):
                        try:
                            if chk.is_selected() and idx < len(option_labels):
                                label_text = option_labels[idx]
                                if "其他" in label_text or "other" in label_text.lower():
                                    chose_other = True
                                    break
                        except:
                            continue

                    # 如果选中了"其他"选项，填写文本框
                    if chose_other:
                        # 增强定位策略
                        locator_strategies = [
                            (By.XPATH, f".//input[preceding-sibling::label[contains(., '其他')]]"),
                            (By.CSS_SELECTOR, "input[placeholder*='其他'], input[placeholder*='请填写']"),
                            (By.CLASS_NAME, "OtherText"),
                            (By.XPATH, ".//div[contains(@class, 'other')]//input"),
                            (By.CSS_SELECTOR, "input[type='text'], textarea")
                        ]

                        other_inputs = []
                        for strategy in locator_strategies:
                            try:
                                found_inputs = question.find_elements(strategy[0], strategy[1])
                                if found_inputs:
                                    other_inputs = found_inputs
                                    break
                            except:
                                continue

                        # 全局查找
                        if not other_inputs:
                            for strategy in locator_strategies:
                                try:
                                    found_inputs = driver.find_elements(strategy[0], strategy[1])
                                    if found_inputs:
                                        other_inputs = found_inputs
                                        break
                                except:
                                    continue

                        # 填写找到的第一个可见文本框
                        for inp in other_inputs:
                            try:
                                if inp.is_displayed() and not inp.get_attribute("value"):
                                    try:
                                        inp.send_keys("自动补全内容")
                                        logging.info("成功补全'其他'文本框")
                                        break
                                    except:
                                        try:
                                            driver.execute_script("arguments[0].value = '自动补全内容';", inp)
                                            logging.info("通过JS补全'其他'文本框")
                                            break
                                        except:
                                            pass
                            except:
                                continue
                    return
            except StaleElementReferenceException:
                pass

            # 3. 填空题
            try:
                texts = question.find_elements(By.CSS_SELECTOR, "input[type='text'], textarea")
                if texts:
                    for t in texts:
                        if not t.get_attribute("value") and t.is_displayed():
                            try:
                                t.send_keys("自动补全内容")
                            except:
                                try:
                                    driver.execute_script("arguments[0].value = '自动补全内容';", t)
                                except:
                                    pass
                    return
            except StaleElementReferenceException:
                pass

            # 4. 下拉框
            try:
                selects = question.find_elements(By.CSS_SELECTOR, "select")
                if selects:
                    for sel in selects:
                        options = sel.find_elements(By.TAG_NAME, "option")
                        for op in options:
                            try:
                                if op.get_attribute("value") and not op.get_attribute("disabled"):
                                    sel.send_keys(op.get_attribute("value"))
                                    break
                            except:
                                continue
                    return
            except StaleElementReferenceException:
                pass

            # 5. 最后尝试：点击任何可点击元素
            try:
                clickable_elements = question.find_elements(By.CSS_SELECTOR,
                                                            "li, label, div[onclick], span[onclick], .option")
                if clickable_elements:
                    element = random.choice(clickable_elements)
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                                          element)
                    time.sleep(0.2)
                    element.click()
                    return
            except StaleElementReferenceException:
                pass

            logging.warning("无法自动补全问题")
        except Exception as e:
            logging.error(f"自动补全题目时出错: {str(e)}")

    def submit_survey(self, driver):
        """
        增强的问卷提交逻辑，自动适配多种提交按钮和结果检测，自动修复常见异常和验证码
        """
        import time
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        max_retries = 3
        for attempt in range(max_retries):
            try:
                # 1. 多种方式查找提交按钮
                submit_btn = None
                selectors = [
                    "#submit_button", ".submit-btn", ".submitbutton", "a[id*='submit']", "button[type='submit']",
                    "input[type='submit']", "div.submit", ".btn-submit", ".btn-success", "#ctlNext", "#submit_btn",
                    "#next_button"
                ]
                for sel in selectors:
                    try:
                        btn = driver.find_element(By.CSS_SELECTOR, sel)
                        if btn.is_displayed() and btn.is_enabled():
                            submit_btn = btn
                            break
                    except Exception:
                        continue
                # 2. 若还找不到，用文本查找
                if not submit_btn:
                    try:
                        submit_btn = driver.find_element(By.XPATH, "//*[contains(text(),'提交')]")
                    except Exception:
                        pass

                # 3. 若找到按钮，尝试click
                if submit_btn:
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                                          submit_btn)
                    try:
                        submit_btn.click()
                    except Exception:
                        try:
                            driver.execute_script("arguments[0].click();", submit_btn)
                        except Exception:
                            pass
                    time.sleep(2)
                else:
                    # 4. 若没有按钮，尝试直接form.submit
                    try:
                        form = driver.find_element(By.TAG_NAME, "form")
                        driver.execute_script("arguments[0].submit();", form)
                    except Exception:
                        print("找不到可用的提交按钮和form，提交失败！")
                        return False

                # 5. 检查提交结果
                time.sleep(2)
                page_text = driver.page_source
                url = driver.current_url
                if any(keyword in url for keyword in ["complete", "success", "finish", "thank"]):
                    print("问卷提交成功！")
                    return True
                if any(word in page_text for word in ["提交成功", "感谢", "问卷已完成", "谢谢您的参与"]):
                    print("问卷提交成功！")
                    return True
                # 6. 检查是否有错误提示/验证码
                if any(word in page_text for word in ["验证码", "请完成验证"]):
                    print("页面出现验证码，请手动处理后继续。")
                    time.sleep(10)
                    continue
                if any(word in page_text for word in ["还有必答题", "请填写", "错误", "失败"]):
                    print("页面提示填写不全或有错误，尝试自动补全。")
                    self.repair_required_questions(driver)
                    time.sleep(1)
                    continue
                print("提交后页面未变化，重试中...")
            except Exception as e:
                print(f"提交问卷异常: {e}")
            time.sleep(2)
        print("多次重试后提交仍未成功！")
        return False


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
        """
        多选题自动填写，支持概率/必选/随机，并对“其他”文本框强力写入（兼容问卷星所有页面）。
        每个checkbox一一对应真实label，兼容大多数问卷星结构。
        去除所有print（调试）输出。
        """
        import random, time
        from selenium.webdriver.common.by import By

        selectors = [
            f"#div{q_num} .ui-checkbox",
            "input[type='checkbox']",
            ".wjx-checkbox",
            ".option input[type='checkbox']"
        ]
        options = []
        for sel in selectors:
            options = question.find_elements(By.CSS_SELECTOR, sel)
            if options:
                break
        if not options:
            logging.warning(f"多选题{q_num}未找到选项，跳过")
            return

        q_key = str(q_num)
        config = self.config.get("multiple_prob", {}).get(q_key, {
            "prob": [50] * len(options),
            "min_selection": 1,
            "max_selection": len(options)
        })
        probs = config.get("prob", [50] * len(options))
        min_selection = config.get("min_selection", 1)
        max_selection = config.get("max_selection", len(options))
        if max_selection > len(options): max_selection = len(options)
        if min_selection > max_selection: min_selection = max_selection
        probs = probs[:len(options)] if len(probs) > len(options) else probs + [50] * (len(options) - len(probs))

        must_indices = [i for i, prob in enumerate(probs) if prob >= 100]
        selected_indices = list(must_indices)
        for i, prob in enumerate(probs):
            if i not in selected_indices and random.random() * 100 < prob:
                selected_indices.append(i)
        while len(selected_indices) < min_selection:
            left = [i for i in range(len(options)) if i not in selected_indices]
            if not left: break
            selected_indices.append(random.choice(left))
        while len(selected_indices) > max_selection:
            removable = [i for i in selected_indices if i not in must_indices]
            if not removable: break
            selected_indices.remove(random.choice(removable))

        # 一一对应label
        option_labels = []
        for c in options:
            label_text = ""
            label_id = c.get_attribute('id')
            if label_id:
                label = question.find_elements(By.CSS_SELECTOR, f"label[for='{label_id}']")
                if label and label[0].text.strip():
                    label_text = label[0].text.strip()
            if not label_text:
                try:
                    sib = c.find_element(By.XPATH, "following-sibling::*[1]")
                    label_text = sib.text.strip()
                except:
                    label_text = ""
            if not label_text:
                try:
                    parent = c.find_element(By.XPATH, "..")
                    label_text = parent.text.strip()
                except:
                    label_text = ""
            if not label_text:
                try:
                    label_text = c.find_element(By.XPATH, "../..").text.strip()
                except:
                    label_text = ""
            if not label_text:
                label_text = f"未知{len(option_labels) + 1}"
            option_labels.append(label_text)

        chose_other = False
        for idx in selected_indices:
            try:
                if idx >= len(options):
                    continue
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                                      options[idx])
                options[idx].click()
                text = (option_labels[idx] or "").strip().lower()
                if "其他" in text or "other" in text:
                    chose_other = True
                    time.sleep(1.2)
            except Exception as e:
                logging.warning(f"选择选项时出错: {str(e)}")
                continue

        if chose_other:
            other_list = self.config.get("other_texts", {}).get(q_key, ["自动填写内容"])
            other_content = random.choice(other_list) if other_list else "自动填写内容"
            other_inputs = []
            for _ in range(10):
                other_inputs = question.find_elements(By.CSS_SELECTOR, "input.OtherText")
                if other_inputs:
                    break
                time.sleep(0.2)
            else:
                other_inputs = driver.find_elements(By.CSS_SELECTOR, "input.OtherText")
            filled = False
            for inp in other_inputs:
                try:
                    if not inp.is_displayed():
                        continue
                    inp.clear()
                    for c in other_content:
                        inp.send_keys(c)
                        time.sleep(0.08)
                    time.sleep(0.3)
                    if inp.get_attribute("value") == other_content:
                        filled = True
                        break
                    driver.execute_script("""
                        arguments[0].value = arguments[1];
                        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    """, inp, other_content)
                    time.sleep(0.3)
                    if inp.get_attribute("value") == other_content:
                        filled = True
                        break
                except Exception as e:
                    continue
            if not filled:
                logging.warning(f"题目{q_num}：'其他'文本框未能自动填写，建议手动检查。")

        self.random_delay(*self.config.get("per_question_delay", (1.0, 3.0)))

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
            # 全局设置
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
            self.ip_entry.delete(0, tk.END)
            self.ip_entry.insert(0, self.config["ip_api"])
            self.ip_change_mode.set(self.config.get("ip_change_mode", "per_submit"))
            self.ip_change_batch.set(self.config.get("ip_change_batch", 5))
            self.headless_var.set(self.config["headless"])
            # 智能提交间隔/批量休息
            self.enable_smart_gap_var.set(self.config.get("enable_smart_gap", True))
            self.min_submit_gap.set(self.config.get("min_submit_gap", 10))
            self.max_submit_gap.set(self.config.get("max_submit_gap", 20))
            self.batch_size.set(self.config.get("batch_size", 5))
            self.batch_pause.set(self.config.get("batch_pause", 15))
            # 重新加载题型设置
            self.reload_question_settings()
            logging.info("已重置为默认配置")

    def save_config(self):
        """
        保存当前界面配置到self.config，包括多选题“其他”答案配置
        """
        try:
            # 全局设置
            self.config["url"] = self.url_entry.get().strip()
            try:
                self.config["target_num"] = int(self.target_entry.get())
            except ValueError:
                self.config["target_num"] = 100
            self.config["weixin_ratio"] = self.ratio_scale.get()
            try:
                self.config["min_duration"] = int(self.min_duration.get())
                self.config["max_duration"] = int(self.max_duration.get())
                self.config["min_delay"] = float(self.min_delay.get())
                self.config["max_delay"] = float(self.max_delay.get())
                self.config["submit_delay"] = int(self.submit_delay.get())
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
                # 可按需回退到默认值
                pass
            self.config["use_ip"] = self.use_ip_var.get()
            self.config["ip_api"] = self.ip_entry.get().strip()
            self.config["ip_change_mode"] = self.ip_change_mode.get()
            try:
                self.config["ip_change_batch"] = int(self.ip_change_batch.get())
            except Exception:
                self.config["ip_change_batch"] = 5
            self.config["headless"] = self.headless_var.get()
            self.config["enable_smart_gap"] = self.enable_smart_gap_var.get()
            try:
                self.config["min_submit_gap"] = float(self.min_submit_gap.get())
                self.config["max_submit_gap"] = float(self.max_submit_gap.get())
                self.config["batch_size"] = int(self.batch_size.get())
                self.config["batch_pause"] = float(self.batch_pause.get())
            except Exception:
                pass

            # 保存多选题“其他”答案配置
            if hasattr(self, "other_entries"):
                for q_num, entry in self.other_entries.items():
                    val = entry.get().strip()
                    if val:
                        self.config["other_texts"][q_num] = [v.strip() for v in val.split(",") if v.strip()]

            logging.info("配置已保存")
            return True
        except Exception as e:
            logging.error(f"保存配置时出错: {str(e)}")
            messagebox.showerror("错误", f"保存配置时出错: {str(e)}")
            return False

    def get_new_proxy(self):
        """拉取代理IP，返回如 http://ip:port 或 http://user:pwd@ip:port"""
        try:
            url = self.config["ip_api"]
            resp = requests.get(url, timeout=8)
            ip = resp.text.strip()
            if ip and "://" not in ip:
                ip = "http://" + ip
            return ip
        except Exception as e:
            logging.error(f"拉取代理失败: {e}")
            return None

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
