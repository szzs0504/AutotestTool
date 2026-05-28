import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyautogui
import keyboard as kb  # 修改为 kb，避免与pynput.keyboard冲突
import json
import time
from threading import Thread, Lock
import os
from pynput import mouse  # 添加鼠标监听器
from pynput import keyboard  # 添加键盘监听器
import sys
import ctypes
import win32gui
import win32con
import win32com.client
from ctypes import windll
import random  # 添加到文件开头的导入部分
from datetime import datetime
from PIL import ImageGrab

# 截屏存储目录
screenshot_dir = os.path.join(os.path.dirname(__file__), 'screenshots')
os.makedirs(screenshot_dir, exist_ok=True)

def take_screenshot(use_timestamp=True):
    """截屏并保存到指定目录"""
    if use_timestamp:
        filename = datetime.now().strftime('%Y%m%d_%H%M%S') + '.png'
    else:
        existing_files = [f for f in os.listdir(screenshot_dir) if f.endswith('.png')]
        next_index = len(existing_files) + 1
        filename = f'screenshot_{next_index}.png'

    filepath = os.path.join(screenshot_dir, filename)
    try:
        screenshot = ImageGrab.grab()
        screenshot.save(filepath)
        print(f"Screenshot saved to {filepath}")
    except Exception as e:
        print(f"Screenshot failed: {e}")

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except OSError:
        return False

class AutoTestGUI:
    def __init__(self):
        # 检查管理员权限
        if not is_admin():
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()
            
        # 启用UI自动化权限
        try:
            # 设置UI访问权限
            windll.user32.SetProcessDPIAware()
            # 允许模拟输入
            windll.user32.AllowSetForegroundWindow(-1)  # -1 表示允许所有窗口
        except Exception as e:
            print(f"权限设置失败: {e}")
        
        self.root = tk.Tk()
        self.root.title("自动测试工具")
        self.root.geometry("1000x600")
        
        # 配置主窗口网格布局
        self.root.grid_rowconfigure(0, weight=1)  # 让主框架占据所有可用空间
        self.root.grid_columnconfigure(0, weight=1)
        
        # 添加窗口关闭事件处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.actions = []
        self.actions_lock = Lock()  # 线程锁保护actions
        self.is_recording = False
        self.is_playing = False
        self.in_dialog_operation = False  # 在对话框操作时不记录
        self.test_index = 0
        self.mouse_listener = None
        self.keyboard_listener = None
        self.last_click_time = 0
        self.last_click_position = (0, 0)
        self.double_click_threshold = 0.3  # 双击检测阈值（秒）
        self.double_click_distance_threshold = 10  # 双击距离阈值（像素）
        self.last_move_time = 0
        self.move_threshold = 0.5  # 移动记录阈值（秒）
        self.loop_count = tk.IntVar(value=1)  # 默认循环1次
        self.loop_interval = tk.DoubleVar(value=1.0)  # 默认循环间隔1秒
        self.interval_random = tk.BooleanVar(value=False)  # 是否启用随机间隔
        self.interval_min = tk.DoubleVar(value=0.5)  # 最小间隔
        self.interval_max = tk.DoubleVar(value=1.5)  # 最大间隔
        self.keyboard_combinations = []  # 存储当前按下的键组合
        self.last_scroll_time = 0
        self.scroll_threshold = 0.1  # 滚轮事件记录阈值（秒）
        self.pressed_keys = set()  # 存储当前按下的键
        self.currently_pressed_keys = set()  # 初始化，避免未定义
        self.record_start_time = 0  # 初始化，避免未定义
        self.ui_elements = []  # 存储UI元素的区域
        self.save_dir = os.path.join(os.path.dirname(__file__), 'actions')  # 默认保存目录
        try:
            if not os.path.exists(self.save_dir):
                os.makedirs(self.save_dir)
        except PermissionError:
            # 如果权限不足，使用用户目录
            self.save_dir = os.path.join(os.path.expanduser('~'), 'auto_test_actions')
            os.makedirs(self.save_dir, exist_ok=True)
        except Exception as e:
            print(f"创建保存目录失败: {e}")
        self.recording_thread = None
        self.mouse_queue = []  # 用于缓存鼠标事件
        self.record_mouse_move = tk.BooleanVar(value=False)  # 默认不记录鼠标移动动作
        self.window_topmost = tk.BooleanVar(value=True)  # 默认窗口置顶
        self.default_delay_min = tk.DoubleVar(value=0.5)  # 默认最小延时
        self.default_delay_max = tk.DoubleVar(value=2.0)  # 默认最大延时
        self.default_delay_base = tk.DoubleVar(value=1.0)  # 默认指数延时基数
        self.playback_info = []  # 存储回放信息
        
        self.status_var = tk.StringVar(value="就绪")

        # 快捷键说明文本
        self.shortcut_help = (
            "快捷键: Ctrl+Shift+R 开始/停止录制 | "
            "Ctrl+Shift+P 开始回放 | "
            "Ctrl+Shift+S 停止回放 | "
            "Ctrl+Q 强制停止回放"
        )
        self.setup_ui()

        # 注册快捷键
        self._register_hotkeys()
        
        # 全局监听 Ctrl+Q 停止回放
        kb.add_hotkey('ctrl+q', self._hotkey_stop_playback)
        # 移除原有的 bind_all
        # self.root.bind_all('<Control-q>', self._hotkey_stop_playback)

    def setup_ui(self):
        # 设置窗口置顶（默认开启）
        self.root.attributes('-topmost', self.window_topmost.get())

        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)
        
        # 配置主框架的网格布局
        main_frame.grid_rowconfigure(1, weight=1)  # 让列表框架可以扩展
        main_frame.grid_columnconfigure(0, weight=1)
        
        # 控制按钮框架
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        
        # 录制和回放按钮框架
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(side=tk.TOP, padx=5)
        
        # 添加移动记录控制开关
        self.move_record_cb = ttk.Checkbutton(
            button_frame,
            text="记录移动",
            variable=self.record_mouse_move,
            command=self.toggle_move_record
        )
        self.move_record_cb.pack(side=tk.LEFT, padx=2)

        # 添加窗口置顶开关
        self.topmost_cb = ttk.Checkbutton(
            button_frame,
            text="置顶",
            variable=self.window_topmost,
            command=self.toggle_topmost
        )
        self.topmost_cb.pack(side=tk.LEFT, padx=2)
        
        self.record_btn = ttk.Button(button_frame, text="开始录制", command=self.toggle_recording)
        self.record_btn.pack(side=tk.LEFT, padx=2)
        self.ui_elements.append(self.record_btn)
        
        self.play_btn = ttk.Button(button_frame, text="开始回放", command=self.play_actions)
        self.play_btn.pack(side=tk.LEFT, padx=2)
        self.ui_elements.append(self.play_btn)
        
        self.stop_btn = ttk.Button(button_frame, text="停止回放", command=self.stop_playback, state='disabled')
        self.stop_btn.pack(side=tk.LEFT, padx=2)
        self.ui_elements.append(self.stop_btn)
        
        # 循环控制框架（更紧凑）
        loop_frame = ttk.LabelFrame(control_frame, text="循环设置")
        loop_frame.pack(side=tk.TOP, padx=5, fill=tk.X, pady=2)

        # 次数
        ttk.Label(loop_frame, text="次数:").pack(side=tk.LEFT, padx=2)
        self.loop_entry = ttk.Entry(loop_frame, width=4, textvariable=self.loop_count)
        self.loop_entry.pack(side=tk.LEFT, padx=2)

        # 间隔
        ttk.Label(loop_frame, text="间隔:").pack(side=tk.LEFT, padx=2)
        self.interval_entry = ttk.Entry(loop_frame, width=4, textvariable=self.loop_interval)
        self.interval_entry.pack(side=tk.LEFT, padx=2)
        ttk.Label(loop_frame, text="秒").pack(side=tk.LEFT, padx=2)

        # 随机间隔选项
        ttk.Checkbutton(loop_frame, text="随机", variable=self.interval_random,
                       command=self.toggle_interval_mode).pack(side=tk.LEFT, padx=5)
        ttk.Entry(loop_frame, width=4, textvariable=self.interval_min).pack(side=tk.LEFT, padx=2)
        ttk.Label(loop_frame, text="-").pack(side=tk.LEFT)
        ttk.Entry(loop_frame, width=4, textvariable=self.interval_max).pack(side=tk.LEFT, padx=2)
        
        # 分割动作列表和回放信息
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.grid(row=1, column=0, sticky="nsew", pady=(0, 5))

        # 动作列表框架
        list_frame = ttk.LabelFrame(paned, text="动作列表")
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.grid(row=0, column=1, sticky="ns")

        # 锁定拖动功能
        self.drag_locked = tk.BooleanVar(value=False)
        lock_check = ttk.Checkbutton(list_frame, text="锁定顺序", variable=self.drag_locked)
        lock_check.grid(row=1, column=0, columnspan=2, pady=2)

        columns = ("action", "delay", "remark")
        self.action_tree = ttk.Treeview(
            list_frame,
            columns=columns,
            show="headings",
            selectmode=tk.EXTENDED
        )
        self.action_tree.heading("action", text="动作描述")
        self.action_tree.heading("delay", text="延时(秒)")
        self.action_tree.heading("remark", text="备注")
        self.action_tree.column("action", anchor="w", width=250, stretch=True)
        self.action_tree.column("delay", anchor="center", width=70, stretch=False)
        self.action_tree.column("remark", anchor="w", width=120, stretch=True)
        self.action_tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.action_tree.bind("<Double-1>", self.on_tree_double_click)
        self.action_tree.bind("<Button-1>", self._handle_tree_click)
        self.action_tree.bind("<MouseWheel>", self._hide_delay_spinbox)

        # 拖动排序功能
        self._drag_start_item = None
        self._drag_start_pos = None
        self._drag_start_idx = None
        self._is_dragging = False
        self.action_tree.bind("<Button-1>", self._on_drag_start, add=True)
        self.action_tree.bind("<B1-Motion>", self._on_drag_motion, add=True)
        self.action_tree.bind("<ButtonRelease-1>", self._on_drag_end, add=True)

        scrollbar.config(command=self.action_tree.yview)
        self.action_scrollbar = scrollbar
        self.action_tree.configure(yscrollcommand=self._on_tree_scroll)
        self.ui_elements.append(self.action_tree)

        self.delay_spinbox = None
        self.delay_spinbox_var = tk.DoubleVar(value=0.0)
        self.active_spinbox_item = None
        self._spinbox_updating = False

        # 备注编辑
        self.remark_entry = None
        self.remark_var = tk.StringVar()
        
        # 回放信息框架
        playback_frame = ttk.LabelFrame(paned, text="回放信息")
        playback_frame.grid_rowconfigure(0, weight=1)
        playback_frame.grid_columnconfigure(0, weight=1)
        
        playback_scrollbar = ttk.Scrollbar(playback_frame)
        playback_scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.playback_list = tk.Listbox(
            playback_frame,
            yscrollcommand=playback_scrollbar.set
        )
        self.playback_list.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        playback_scrollbar.config(command=self.playback_list.yview)
        
        # 将两个框架添加到分割窗口
        paned.add(list_frame, weight=1)
        paned.add(playback_frame, weight=1)
        
        # 底部按钮框架
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=2, column=0, sticky="ew", pady=(0, 5))

        # 左侧按钮组（更紧凑）
        left_buttons = ttk.Frame(bottom_frame)
        left_buttons.pack(side=tk.LEFT, padx=3)

        ttk.Button(left_buttons, text="保存动作", command=self.save_actions).pack(side=tk.LEFT, padx=1)
        ttk.Button(left_buttons, text="加载动作", command=self.load_actions).pack(side=tk.LEFT, padx=1)
        ttk.Button(left_buttons, text="插入点击", command=self.insert_click_action).pack(side=tk.LEFT, padx=1)
        ttk.Button(left_buttons, text="删除动作", command=self.delete_selected).pack(side=tk.LEFT, padx=1)
        ttk.Button(left_buttons, text="插入延时", command=self.insert_random_delay).pack(side=tk.LEFT, padx=1)
        ttk.Button(left_buttons, text="循环组", command=self.insert_loop_group).pack(side=tk.LEFT, padx=1)
        ttk.Button(left_buttons, text="输入", command=self.insert_variable).pack(side=tk.LEFT, padx=1)
        
        # 右侧按钮组
        right_buttons = ttk.Frame(bottom_frame)
        right_buttons.pack(side=tk.RIGHT, padx=5)

        # 添加状态栏
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=2, column=0, sticky="ew", padx=5, pady=2)

        # 快捷键提示栏
        self.shortcut_label = ttk.Label(self.root, text=self.shortcut_help, anchor="w", relief=tk.GROOVE)
        self.shortcut_label.grid(row=3, column=0, sticky="ew", padx=5, pady=2)

        # 添加截屏设置框架（更紧凑的布局）
        screenshot_frame = ttk.LabelFrame(main_frame, text="截屏设置")
        screenshot_frame.grid(row=3, column=0, sticky="ew", pady=(0, 5))

        # 第一行：目录
        ttk.Label(screenshot_frame, text="目录:").grid(row=0, column=0, padx=3, pady=3, sticky="e")
        self.screenshot_dir_var = tk.StringVar(value=screenshot_dir)
        screenshot_dir_entry = ttk.Entry(screenshot_frame, textvariable=self.screenshot_dir_var, width=35)
        screenshot_dir_entry.grid(row=0, column=1, padx=3, pady=3, sticky="w")

        ttk.Button(screenshot_frame, text="选择", command=self.select_screenshot_dir).grid(row=0, column=2, padx=3, pady=3)

        # 第二行：文件名和延时
        ttk.Label(screenshot_frame, text="名称:").grid(row=1, column=0, padx=3, pady=3, sticky="e")
        self.screenshot_name_var = tk.StringVar(value="screenshot")
        screenshot_name_entry = ttk.Entry(screenshot_frame, textvariable=self.screenshot_name_var, width=15)
        screenshot_name_entry.grid(row=1, column=1, padx=3, pady=3, sticky="w")

        ttk.Label(screenshot_frame, text="延时:").grid(row=1, column=2, padx=3, pady=3, sticky="e")
        self.screenshot_delay_var = tk.DoubleVar(value=0.0)
        screenshot_delay_entry = ttk.Entry(screenshot_frame, textvariable=self.screenshot_delay_var, width=8)
        screenshot_delay_entry.grid(row=1, column=3, padx=3, pady=3, sticky="w")

        self.screenshot_naming_var = tk.StringVar(value="timestamp")
        ttk.Radiobutton(screenshot_frame, text="时间戳", variable=self.screenshot_naming_var, value="timestamp").grid(row=2, column=0, padx=3, pady=3, sticky="w")
        ttk.Radiobutton(screenshot_frame, text="序号", variable=self.screenshot_naming_var, value="increment").grid(row=2, column=1, padx=3, pady=3, sticky="w")

        ttk.Button(screenshot_frame, text="添加截屏", command=self.add_screenshot_action).grid(row=2, column=2, padx=3, pady=3)

    def select_screenshot_dir(self):
        """选择截屏保存目录"""
        dir_path = filedialog.askdirectory(initialdir=self.screenshot_dir_var.get(), title="选择截屏保存目录")
        if dir_path:
            self.screenshot_dir_var.set(dir_path)

    def is_click_in_ui(self, x, y):
        """检查点击是否在UI元素上"""
        try:
            # 获取程序窗口位置
            window_x = self.root.winfo_x()
            window_y = self.root.winfo_y()
            
            # 检查每个UI元素
            for element in self.ui_elements:
                if not element.winfo_viewable():  # 如果元素不可见则跳过
                    continue
                    
                # 获取元素相对于窗口的位置和大小
                elem_x = window_x + element.winfo_rootx()
                elem_y = window_y + element.winfo_rooty()
                elem_width = element.winfo_width()
                elem_height = element.winfo_height()
                
                # 检查点击是否在元素区域内
                if (elem_x <= x <= elem_x + elem_width and
                    elem_y <= y <= elem_y + elem_height):
                    return True
            return False
        except Exception:
            return False

    def is_in_window(self, x, y):
        """检查坐标是否在工具窗口内"""
        try:
            window_x = self.root.winfo_x()
            window_y = self.root.winfo_y()
            window_width = self.root.winfo_width()
            window_height = self.root.winfo_height()

            return (window_x <= x <= window_x + window_width and
                   window_y <= y <= window_y + window_height)
        except Exception:
            return False

    def is_in_ui_element(self, x, y):
        """检查坐标是否在UI元素（按钮等）内，用于判断是否需要录制按钮点击"""
        try:
            # 检查每个UI元素
            for element in self.ui_elements:
                if not element.winfo_viewable():  # 如果元素不可见则跳过
                    continue

                # 获取元素相对于屏幕的位置和大小
                elem_x = element.winfo_rootx()
                elem_y = element.winfo_rooty()
                elem_width = element.winfo_width()
                elem_height = element.winfo_height()

                # 检查点击是否在元素区域内
                if (elem_x <= x <= elem_x + elem_width and
                    elem_y <= y <= elem_y + elem_height):
                    return True
            return False
        except Exception:
            return False

    def on_click(self, x, y, button, pressed):
        if not pressed:  # 只在释放时记录
            # 如果正在对话框操作中，不记录
            if self.in_dialog_operation:
                return

            # 检查是否在工具窗口内
            if self.is_in_window(x, y):
                # 检查是否在UI元素（按钮）内
                if self.is_in_ui_element(x, y):
                    # 在按钮内点击，录制这个动作（允许录制按钮点击）
                    # 继续执行下面的录制逻辑
                    pass
                else:
                    # 在窗口内但不在按钮上，不录制
                    return

            current_time = time.time() - self.record_start_time  # 使用相对时间
            # 检测双击（使用距离阈值，避免精确坐标匹配问题）
            last_x, last_y = self.last_click_position
            distance = abs(x - last_x) + abs(y - last_y)  # 曼哈顿距离
            is_double_click = (current_time - self.last_click_time < self.double_click_threshold and
                              distance < self.double_click_distance_threshold)

            if is_double_click:
                # 移除上一次的单击记录
                if self.actions and self.actions[-1]['type'] == 'click':
                    self.actions.pop()
                    self.update_action_list_display()
                # 记录双击
                self.actions.append({
                    'type': 'doubleclick',
                    'x': x,
                    'y': y,
                    'button': str(button),
                    'time': current_time
                })
                self.update_action_list_display()
            else:
                # 记录单击
                self.actions.append({
                    'type': 'click',
                    'x': x,
                    'y': y,
                    'button': str(button),
                    'time': current_time
                })
                self.update_action_list_display()
            
            self.last_click_time = current_time
            self.last_click_position = (x, y)

    def on_scroll(self, x, y, dx, dy):
        if self.is_recording and not self.in_dialog_operation:
            # 检查是否在工具窗口内
            if self.is_in_window(x, y):
                # 检查是否在UI元素（按钮）内
                if not self.is_in_ui_element(x, y):
                    # 在窗口内但不在按钮上，不录制
                    return
                
            current_time = time.time() - self.record_start_time
            # 防止滚轮事件记录过于频繁
            if current_time - self.last_scroll_time >= self.scroll_threshold:
                self.actions.append({
                    'type': 'scroll',
                    'x': x,
                    'y': y,
                    'dx': dx,
                    'dy': dy,
                    'time': current_time
                })
                self.update_action_list_display()
                self.last_scroll_time = current_time

    def on_key_down(self, key):
        try:
            if self.is_recording and not self.in_dialog_operation:
                key_str = self.convert_key_name(key)
                self.pressed_keys.add(key_str)
        except AttributeError:
            pass

    def on_key_up(self, key):
        try:
            if self.is_recording and not self.in_dialog_operation:
                current_time = time.time() - self.record_start_time
                key_str = self.convert_key_name(key)
                
                if key_str in self.pressed_keys:
                    self.pressed_keys.remove(key_str)
                    
                    # 获取当前按下的所有键
                    current_keys = list(self.pressed_keys)
                    
                    # 如果是组合键
                    if current_keys:
                        all_keys = current_keys + [key_str]
                        self.actions.append({
                            'type': 'keyboard',
                            'keys': all_keys,
                            'time': current_time
                        })
                        self.update_action_list_display()
                    else:
                        # 单个键
                        self.actions.append({
                            'type': 'keyboard',
                            'keys': [key_str],
                            'time': current_time
                        })
                        self.update_action_list_display()
        except AttributeError:
            pass

    def convert_key_name(self, key):
        """转换键名为更易读和一致的格式"""
        try:
            # 处理组合键特殊情况
            if isinstance(key, keyboard.KeyCode):
                # 检查虚拟键码
                if hasattr(key, 'vk'):
                    # 字母键 (A-Z)
                    if 0x41 <= key.vk <= 0x5A:  # A=0x41, Z=0x5A
                        return chr(key.vk).lower()
                    # 数字键 (0-9)
                    if 0x30 <= key.vk <= 0x39:  # 0=0x30, 9=0x39
                        return chr(key.vk)
                    # 小键盘数字 (0-9)
                    if 0x60 <= key.vk <= 0x69:  # numpad 0-9
                        return f"numpad{key.vk - 0x60}"
                    
                    # 其他特殊键映射
                    vk_mapping = {
                        0x08: 'backspace',  # Backspace
                        0x09: 'tab',        # Tab
                        0x0D: 'enter',      # Enter
                        0x13: 'pause',      # Pause
                        0x14: 'capslock',   # Caps Lock
                        0x1B: 'esc',        # Escape
                        0x20: 'space',      # Spacebar
                        0x21: 'pageup',     # Page Up
                        0x22: 'pagedown',   # Page Down
                        0x23: 'end',        # End
                        0x24: 'home',       # Home
                        0x25: 'left',       # Left Arrow
                        0x26: 'up',         # Up Arrow
                        0x27: 'right',       # Right Arrow
                        0x28: 'down',       # Down Arrow
                        0x2D: 'insert',     # Insert
                        0x2E: 'delete',     # Delete
                    }
                    return vk_mapping.get(key.vk, f'key_{hex(key.vk)}')
                
                # 如果有字符属性则使用字符
                if hasattr(key, 'char') and key.char:
                    return key.char.lower()
                    
            # 处理特殊键
            key_str = str(key).replace('Key.', '')
            key_mapping = {
                # 基础按键
                'space': 'space',
                'enter': 'enter',
                'tab': 'tab',
                'esc': 'escape',
                'escape': 'escape',
                'backspace': 'backspace',
                'delete': 'delete',
                'insert': 'insert',
                
                # 修饰键
                'ctrl_l': 'ctrl',
                'ctrl_r': 'ctrl',
                'shift_l': 'shift',
                'shift_r': 'shift',
                'alt_l': 'alt',
                'alt_r': 'alt',
                'cmd': 'win',
                'cmd_r': 'win',
                
                # 功能键
                'f1': 'f1',
                'f2': 'f2',
                'f3': 'f3',
                'f4': 'f4',
                'f5': 'f5',
                'f6': 'f6',
                'f7': 'f7',
                'f8': 'f8',
                'f9': 'f9',
                'f10': 'f10',
                'f11': 'f11',
                'f12': 'f12',
                
                # 导航键
                'up': 'up',
                'down': 'down',
                'left': 'left',
                'right': 'right',
                'page_up': 'pageup',
                'page_down': 'pagedown',
                'home': 'home',
                'end': 'end',
                
                # 常用组合键的虚拟键码映射
                '0x03': 'c',  # Ctrl+C
                '0x16': 'v',  # Ctrl+V
                '0x18': 'x',  # Ctrl+X
                '0x1A': 'z',  # Ctrl+Z
                '0x01': 'a',  # Ctrl+A
                '0x13': 's',  # Ctrl+S
                '0x04': 'd',  # Ctrl+D
                '0x06': 'f',  # Ctrl+F
                '0x0E': 'n',  # Ctrl+N
                '0x0F': 'o',  # Ctrl+O
                '0x10': 'p',  # Ctrl+P
                '0x19': 'y',  # Ctrl+Y
            }
            return key_mapping.get(key_str, key_str).lower()
        except Exception:
            return str(key).lower()

    def execute_keyboard_action(self, keys):
        """执行键盘动作"""
        try:
            # 映射按键名称为pyautogui可识别的格式
            key_mapping = {
                # 修饰键映射
                'ctrl': 'ctrlleft',
                'shift': 'shiftleft',
                'alt': 'altleft',
                'win': 'winleft',
                'control': 'ctrlleft',
                
                # 特殊键映射
                'escape': 'esc',
                'return': 'enter',
                'delete': 'del',
                'page_up': 'pageup',
                'page_down': 'pagedown',
                
                # 方向键映射
                'up_arrow': 'up',
                'down_arrow': 'down',
                'left_arrow': 'left',
                'right_arrow': 'right',
                
                # 其他常用键映射
                'caps_lock': 'capslock',
                'print_screen': 'printscreen',
                'scroll_lock': 'scrolllock',
                'num_lock': 'numlock',
                'menu': 'apps'
            }
            
            mapped_keys = [key_mapping.get(k, k) for k in keys]
            
            if len(mapped_keys) > 1:
                # 对于组合键，使用keyDown和keyUp
                try:
                    # 按下所有键
                    for k in mapped_keys[:-1]:
                        pyautogui.keyDown(k)
                    # 按下并释放最后一个键
                    pyautogui.press(mapped_keys[-1])
                    # 释放所有键，反序释放
                    for k in reversed(mapped_keys[:-1]):
                        pyautogui.keyUp(k)
                except Exception as e:
                    print(f"组合键执行错误: {e}")
            else:
                # 单个按键直接使用press
                pyautogui.press(mapped_keys[0])
            
            time.sleep(0.1)  # 确保按键执行完成
        except Exception as e:
            print(f"键盘事件执行错误: {e}")

    def on_key_press(self, key):
        try:
            # 检测 Ctrl+Esc
            if key == keyboard.Key.esc and keyboard.Key.ctrl in self.currently_pressed_keys:
                if self.is_recording:
                    self.stop_recording()
        except AttributeError:
            pass

    def stop_recording(self):
        self.is_recording = False

        # 停止并等待监听器
        if self.mouse_listener:
            try:
                self.mouse_listener.stop()
                # 给监听器一点时间来停止
                time.sleep(0.1)
            except Exception as e:
                print(f"停止鼠标监听器失败: {e}")
        if self.keyboard_listener:
            try:
                self.keyboard_listener.stop()
                time.sleep(0.1)
            except Exception as e:
                print(f"停止键盘监听器失败: {e}")

        if self.recording_thread and self.recording_thread.is_alive():
            self.recording_thread.join(timeout=1.0)

        # 更新UI状态
        self.record_btn.config(text="开始录制")
        self.play_btn.config(state='normal')
        if self.actions:  # 如果有录制的动作
            self.status_var.set("录制完成")
        else:
            self.status_var.set("就绪")

    def stop_playback(self):
        """立即停止所有播放动作并关闭播放线程"""
        self.is_playing = False

        # 等待播放线程结束
        if hasattr(self, 'playback_thread') and self.playback_thread.is_alive():
            self.playback_thread.join(timeout=1.0)

        # 更新UI状态
        self.play_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.record_btn.config(state='normal')
        self.status_var.set("回放已停止")

        # 启用其他按钮
        for button in self.ui_elements:
            if isinstance(button, ttk.Button):
                button.config(state='normal')

        self.update_playback_info("回放已停止")

    def play_actions(self):
        if not self.actions:
            messagebox.showwarning("警告", "没有可回放的动作！")
            return

        self.is_playing = True
        self.play_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.test_index+=1
        self.playback_thread = Thread(target=self.play_recorded_actions,args=(self.test_index,))
        self.playback_thread.start()

    def toggle_recording(self):
        if not self.is_recording:
            self.actions = []
            self._update_action_list()
            self.is_recording = True
            self.record_btn.config(text="停止录制")
            
            # 初始化当前按下的键集合
            self.currently_pressed_keys = set()
            self.keyboard_combinations = []
            self.pressed_keys = set()  # 重置按键状态
            
            # 使用新的线程管理方式
            self.recording_thread = Thread(target=self.record_actions)
            self.recording_thread.daemon = True  # 设置为守护线程
            self.recording_thread.start()
            
            # 初始化鼠标监听器，使用队列处理事件
            self.mouse_listener = mouse.Listener(
                on_click=self.on_click,
                on_scroll=self.on_scroll,
                suppress=False  # 不阻止事件传播
            )
            self.mouse_listener.start()

            # 键盘监听器配置
            self.keyboard_listener = keyboard.Listener(
                on_press=self.on_key_down,
                on_release=self.on_key_up,
                suppress=False  # 不阻止事件传播
            )
            self.keyboard_listener.start()
        else:
            self.stop_recording()

    def toggle_move_record(self):
        """切换鼠标移动记录状态"""
        state = "开启" if self.record_mouse_move.get() else "关闭"
        self.status_var.set(f"鼠标移动记录已{state}")

    def toggle_topmost(self):
        """切换窗口置顶状态"""
        topmost = self.window_topmost.get()
        self.root.attributes('-topmost', topmost)
        state = "开启" if topmost else "关闭"
        self.status_var.set(f"窗口置顶已{state}")

    def toggle_interval_mode(self):
        """切换固定/随机间隔模式"""
        is_random = self.interval_random.get()
        self.interval_entry.configure(state='disabled' if is_random else 'normal')
        self.status_var.set(f"使用{'随机' if is_random else '固定'}时间间隔")

    def record_actions(self):
        """优化后的录制函数，确保弹窗出现时不停止录制"""
        self.record_start_time = time.time()
        last_position = None
        corner_threshold = 10

        # 禁用 pyautogui 的故障保护
        pyautogui.FAILSAFE = False

        try:
            while self.is_recording:
                try:
                    current_position = pyautogui.position()
                    current_time = time.time() - self.record_start_time

                    # 使用更高效的左上角检测
                    if current_position[0] <= corner_threshold and current_position[1] <= corner_threshold:
                        self.root.after(0, self.stop_recording)
                        break

                    # 只在启用移动记录时记录移动
                    if self.record_mouse_move.get():
                        if not self.is_in_window(current_position[0], current_position[1]):
                            # 优化移动事件记录
                            if (last_position != current_position and 
                                current_time - self.last_move_time >= self.move_threshold):
                                self.actions.append({
                                    'type': 'move',
                                    'x': current_position[0],
                                    'y': current_position[1],
                                    'time': current_time
                                })
                                # 使用 root.after 确保在主线程中更新 UI
                                self.root.after(0, lambda pos=current_position: 
                                    self.update_action_list_display())
                                last_position = current_position
                                self.last_move_time = current_time

                    time.sleep(0.005)  # 减少CPU使用率但保持响应性
                except Exception as e:
                    print(f"记录动作时出错: {e}")
                    continue
        finally:
            pyautogui.FAILSAFE = True  # 恢复故障保护

    def _execute_single_action(self, action):
        """执行单个动作（用于循环组）"""
        if not action:
            return

        action_type = action.get('type')
        if action_type == 'click':
            pyautogui.moveTo(action['x'], action['y'], duration=0.1)
            pyautogui.click(x=action['x'], y=action['y'])
        elif action_type == 'doubleclick':
            pyautogui.moveTo(action['x'], action['y'], duration=0.1)
            pyautogui.doubleClick(x=action['x'], y=action['y'])
        elif action_type == 'keyboard':
            keys = [k.lower() for k in action.get('keys', [])]
            self.execute_keyboard_action(keys)
        elif action_type == 'scroll':
            pyautogui.moveTo(action['x'], action['y'], duration=0.1)
            scroll_amount = int(action.get('dy', 0) * 100)
            if scroll_amount != 0:
                pyautogui.scroll(scroll_amount)
        elif action_type == 'move':
            pyautogui.moveTo(action['x'], action['y'], duration=0.1)
        elif action_type == 'screenshot':
            self.take_screenshot(delay=action.get('delay', 0))
        elif action_type == 'random_delay':
            delay = random.uniform(action.get('min_delay', 0), action.get('max_delay', 0))
            time.sleep(delay)
        elif action_type == 'multiply_delay':
            # 倍数延时：这里在循环组内不使用倍数，直接使用基数
            delay = action.get('base_delay', 1.0)
            time.sleep(delay)
        elif action_type == 'arithmetic_delay':
            # 等差延时：这里在循环组内使用起始延时
            delay = action.get('start_delay', 0.5)
            time.sleep(delay)
        elif action_type == 'loop_group':
            # 嵌套循环组：执行内部动作
            nested_loop_count = action.get('loop_count', 1)
            nested_loop_actions = action.get('loop_actions', [])
            interval_type = action.get('loop_interval_type', 'fixed')
            interval_params = action.get('loop_interval', {})

            if not nested_loop_actions:
                return

            for _ in range(nested_loop_count):
                if not self.is_playing:
                    return
                # 执行嵌套循环动作
                for nested_action in nested_loop_actions:
                    if not self.is_playing:
                        return
                    self._execute_single_action(nested_action)

                # 嵌套循环间隔
                if _ < nested_loop_count - 1:
                    if interval_type == 'fixed':
                        wait_time = interval_params.get('value', 0)
                    elif interval_type == 'range':
                        wait_time = random.uniform(interval_params.get('min', 0), interval_params.get('max', 0))
                    elif interval_type == 'list':
                        values = interval_params.get('values', [0])
                        wait_time = random.choice(values)
                    else:  # random
                        wait_time = random.uniform(interval_params.get('min', 0), interval_params.get('max', 0))
                    if wait_time > 0:
                        time.sleep(wait_time)
        elif action_type == 'variable_input':
            # 变量输入
            target_x = action.get('x')
            target_y = action.get('y')
            value = action.get('value', '')
            clear_text = action.get('clear_text', True)

            if target_x is not None and target_y is not None:
                pyautogui.moveTo(target_x, target_y, duration=0.1)
                pyautogui.click(x=target_x, y=target_y)

            if clear_text:
                # 全选并删除
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.05)
                pyautogui.press('backspace')

            # 输入文本
            pyautogui.write(str(value), interval=0.05)

    def play_recorded_actions(self, index):
        # 清空回放信息列表（使用root.after确保线程安全）
        self.root.after(0, lambda: self.playback_list.delete(0, tk.END))
        self.root.after(0, lambda: self.status_var.set("正在执行..."))
        if not self.actions:
            return

        try:
            # 获取前台窗口信息（添加异常处理）
            try:
                hwnd = win32gui.GetForegroundWindow()
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                print(f"窗口切换失败: {e}")
            time.sleep(0.5)

            loops = max(1, self.loop_count.get())
            if self.interval_random.get():
                min_interval = max(0, self.interval_min.get())
                max_interval = max(min_interval, self.interval_max.get())
            else:
                interval = max(0, self.loop_interval.get())

            for loop in range(loops):
                if not self.is_playing:
                    self.root.after(0, lambda: self.status_var.set("回放已中止"))
                    break
                self.update_playback_info(f"开始执行第 {loop + 1}/{loops} 轮")

                last_action_time = 0  # 初始化上一个动作的时间

                for i, action in enumerate(self.actions):
                    # 检查是否按下Ctrl+Q（兼容性处理，防止遗漏）
                    if not self.is_playing:
                        return
                    if self.test_index != index:
                        return
                    # 检查鼠标是否在屏幕左上角
                    current_position = pyautogui.position()
                    if current_position[0] <= 10 and current_position[1] <= 10:
                        self.stop_playback()
                        return

                    # 计算需要等待的时间
                    if i > 0:
                        wait_time = action.get('time', 0) - last_action_time
                        if wait_time > 0:
                            time.sleep(wait_time)

                    if action['type'] == 'move':
                        next_is_click = (i < len(self.actions) - 1 and 
                                       self.actions[i + 1]['type'] in ['click', 'doubleclick'])
                        if next_is_click:
                            pyautogui.moveTo(action['x'], action['y'], duration=0.2)
                    elif action['type'] == 'click':
                        pyautogui.moveTo(action['x'], action['y'], duration=0.2)
                        pyautogui.click(x=action['x'], y=action['y'])
                    elif action['type'] == 'doubleclick':
                        pyautogui.moveTo(action['x'], action['y'], duration=0.2)
                        pyautogui.doubleClick(x=action['x'], y=action['y'])
                    elif action['type'] == 'keyboard':
                        keys = [k.lower() for k in action['keys']]
                        self.execute_keyboard_action(keys)
                    elif action['type'] == 'scroll':
                        pyautogui.moveTo(action['x'], action['y'], duration=0.2)
                        scroll_amount = int(action['dy'] * 100)
                        if scroll_amount != 0:
                            pyautogui.scroll(scroll_amount)
                    elif action['type'] == 'random_delay':
                        delay = random.uniform(action['min_delay'], action['max_delay'])
                        self.update_playback_info(f"等待随机延时: {delay:.1f}秒")
                        time.sleep(delay)
                    elif action['type'] == 'multiply_delay':
                        # 计算倍数延时: base * (loop_index + 1)
                        base_delay = action.get('base_delay', 1.0)
                        delay = base_delay * (loop + 1)
                        self.update_playback_info(f"等待倍数延时(第{loop+1}轮): {delay:.1f}秒")
                        time.sleep(delay)
                    elif action['type'] == 'arithmetic_delay':
                        # 计算等差延时: start + step * loop_index
                        start_delay = action.get('start_delay', 0.5)
                        step_delay = action.get('step_delay', 0.1)
                        delay = start_delay + step_delay * loop
                        self.update_playback_info(f"等待等差延时(第{loop+1}轮): {delay:.1f}秒")
                        time.sleep(delay)
                    elif action['type'] == 'screenshot':
                        self.take_screenshot(delay=action.get('delay', 0))
                    elif action['type'] == 'loop_group':
                        # 循环组：执行选中动作N次
                        loop_count = action.get('loop_count', 1)
                        loop_actions = action.get('loop_actions', [])
                        interval_type = action.get('loop_interval_type', 'fixed')
                        interval_params = action.get('loop_interval', {})

                        if not loop_actions:
                            self.update_playback_info("循环组: 无循环动作")
                            continue

                        for _ in range(loop_count):
                            if not self.is_playing:
                                return
                            # 执行循环动作
                            for loop_action in loop_actions:
                                if not self.is_playing:
                                    return
                                self._execute_single_action(loop_action)

                            # 循环间隔
                            if _ < loop_count - 1:
                                if interval_type == 'fixed':
                                    wait_time = interval_params.get('value', 0)
                                elif interval_type == 'range':
                                    wait_time = random.uniform(interval_params.get('min', 0), interval_params.get('max', 0))
                                elif interval_type == 'list':
                                    values = interval_params.get('values', [0])
                                    wait_time = random.choice(values)
                                else:  # random
                                    wait_time = random.uniform(interval_params.get('min', 0), interval_params.get('max', 0))
                                if wait_time > 0:
                                    time.sleep(wait_time)

                        self.update_playback_info(f"循环组完成: 共执行{loop_count}次")
                    elif action['type'] == 'numeric_loop':
                        # 数值循环：通过{n}和步距执行输入动作
                        start = action.get('start', 0)
                        step = action.get('step', 1)
                        count = action.get('count', 10)
                        input_text = action.get('input_text', '')
                        target_x = action.get('x')
                        target_y = action.get('y')
                        interval = action.get('interval', 0.5)

                        if step == 0:
                            self.update_playback_info("数值循环: 步距不能为0")
                            continue

                        for i in range(count):
                            if not self.is_playing:
                                return

                            # 计算当前数值：起始值 + 步距 * 索引
                            value = start + step * i

                            # 如果有目标位置，先移动到目标位置
                            if target_x is not None and target_y is not None:
                                pyautogui.moveTo(target_x, target_y, duration=0.1)

                            # 替换文本中的占位符并输入
                            text_to_input = input_text.replace('{n}', str(value)).replace('{i}', str(value))
                            pyautogui.write(text_to_input, interval=0.05)
                            self.update_playback_info(f"数值循环: 输入 {text_to_input}")

                            # 等待间隔
                            if interval > 0:
                                time.sleep(interval)

                        self.update_playback_info(f"数值循环完成: 共{count}次")
                    elif action['type'] == 'variable_input':
                        # 变量输入：输入带变量的文本
                        start = action.get('start', 0)
                        step = action.get('step', 1)
                        input_text = action.get('input_text', '')
                        target_x = action.get('x')
                        target_y = action.get('y')
                        clear_text = action.get('clear_text', False)

                        # 计算当前变量值
                        if not hasattr(self, '_variable_counter'):
                            self._variable_counter = {}
                        key = id(action)
                        if key not in self._variable_counter:
                            self._variable_counter[key] = start
                        current_value = self._variable_counter[key]

                        # 如果有目标位置，先移动到目标位置
                        if target_x is not None and target_y is not None:
                            pyautogui.moveTo(target_x, target_y, duration=0.1)
                            # 点击文本框获取焦点
                            pyautogui.click()

                            # 如果需要清空文本框
                            if clear_text:
                                # 全选并删除
                                pyautogui.hotkey('ctrl', 'a')
                                time.sleep(0.05)
                                pyautogui.press('backspace')
                                time.sleep(0.05)

                        # 替换文本中的占位符并输入
                        text_to_input = input_text.replace('{' + str(start) + '}', str(current_value)).replace('{n}', str(current_value)).replace('{i}', str(current_value))
                        pyautogui.write(text_to_input, interval=0.05)
                        self.update_playback_info(f"变量输入: {text_to_input}")

                        # 更新变量值
                        self._variable_counter[key] = current_value + step

                    last_action_time = action.get('time', 0)

                if loop < loops - 1:
                    if self.interval_random.get():
                        wait_time = random.uniform(min_interval, max_interval)
                        self.update_playback_info(f"等待 {wait_time:.1f} 秒后开始下一轮...")
                    else:
                        wait_time = interval
                        self.update_playback_info(f"等待 {interval} 秒后开始下一轮...")
                    time.sleep(wait_time)

            self.is_playing = False
            self.play_btn.config(state='normal')
            self.stop_btn.config(state='disabled')
            if self.status_var.get() != "回放已中止":
                self.status_var.set("执行完成")
            self.update_playback_info("播放完成")

        except Exception as e:
            print(f"回放错误: {e}")
            self.root.after(0, lambda: self.status_var.set("回放出错"))
            return
    
    def update_playback_info(self, text):
        """更新回放信息列表（线程安全）"""
        # 使用root.after确保在主线程中执行UI操作
        self.root.after(0, lambda: self._do_update_playback_info(text))

    def _do_update_playback_info(self, text):
        """实际更新UI的方法（必须在主线程中调用）"""
        try:
            self.playback_list.insert(tk.END, text)
            self.playback_list.yview_moveto(1)
            self.playback_list.update()
        except tk.TclError:
            # 窗口可能已关闭
            pass

    def save_actions(self):
        if not self.actions:
            messagebox.showwarning("警告", "没有动作可保存！")
            return
        
        # 创建保存文件对话框
        filepath = filedialog.asksaveasfilename(
            initialdir=self.save_dir,
            title="保存动作文件",
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
            initialfile="新动作.json"
        )
        
        if filepath:
            try:
                # 确保目录存在（检查dirname是否为空）
                dir_path = os.path.dirname(filepath)
                if dir_path:
                    os.makedirs(dir_path, exist_ok=True)
                # 保存文件
                with open(filepath, 'w', encoding='utf-8') as f:
                    json.dump(self.actions, f, ensure_ascii=False, indent=2)
                self.status_var.set(f"动作已保存至: {filepath}")
                messagebox.showinfo("成功", "动作已保存！")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {str(e)}")

    def load_actions(self):
        """加载动作文件"""
        filepath = filedialog.askopenfilename(
            initialdir=self.save_dir,
            title="选择要加载的动作文件",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        
        if not filepath:
            return
            
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                self.actions = json.load(f)
            self._update_action_list()  # 更新动作列表显示
            messagebox.showinfo("成功", "动作已加载！")
        except Exception as e:
            messagebox.showerror("错误", f"加载失败: {str(e)}")

    def _update_action_list(self, scroll_to_end=False, select_index=None):
        """刷新动作列表为Treeview格式，附带延时信息"""
        if not hasattr(self, 'action_tree'):
            return

        self._hide_delay_spinbox()
        previous_selection = [int(self.action_tree.index(item)) for item in self.action_tree.selection()] if self.action_tree.selection() else []

        for item in self.action_tree.get_children():
            self.action_tree.delete(item)

        idx = 0
        while idx < len(self.actions):
            action = self.actions[idx]
            action_text = self._describe_action(action)
            delay_text = self._format_delay_text(idx)
            remark_text = action.get('remark', '')

            if action['type'] == 'loop_group':
                # 循环组：使用树形结构显示，支持展开
                parent_id = f"loop_{idx}"
                self.action_tree.insert("", tk.END, iid=parent_id, values=(action_text, delay_text, remark_text), open=False)
                # 递归添加内部动作
                loop_actions = action.get('loop_actions', [])
                self._add_loop_actions_to_tree(parent_id, loop_actions, idx, scroll_to_end, select_index, previous_selection)
                idx += 1
            elif action['type'] == 'numeric_loop':
                # 数值循环显示
                self.action_tree.insert("", tk.END, iid=str(idx), values=(action_text, delay_text, remark_text))
                idx += 1
            else:
                self.action_tree.insert("", tk.END, iid=str(idx), values=(action_text, delay_text, remark_text))
                idx += 1

        # 如果没有循环组执行递归调用，在这里执行选择逻辑
        has_loop_group = any(a.get('type') == 'loop_group' for a in self.actions)
        if not has_loop_group:
            self._restore_selection(scroll_to_end, select_index, previous_selection)

    def _restore_selection(self, scroll_to_end, select_index, previous_selection):
        """恢复选择状态"""
        target_selection = None
        if select_index is not None and 0 <= select_index < len(self.actions):
            target_selection = (select_index,)
        elif previous_selection:
            target_selection = tuple(i for i in previous_selection if i < len(self.actions))

        if self.action_tree.selection():
            self.action_tree.selection_remove(self.action_tree.selection())
        if target_selection:
            new_selection = [str(idx) for idx in target_selection if self.action_tree.exists(str(idx))]
            if not new_selection and target_selection:
                fallback = str(target_selection[0])
                if self.action_tree.exists(fallback):
                    new_selection = [fallback]
            if new_selection:
                self.action_tree.selection_set(new_selection)
                self.action_tree.see(new_selection[0])
        elif scroll_to_end and self.actions:
            last_item = self.action_tree.get_children()[-1]
            self.action_tree.see(last_item)

    def _add_loop_actions_to_tree(self, parent_id, loop_actions, parent_idx, scroll_to_end=False, select_index=None, previous_selection=None, depth=0):
        """递归添加循环组内的动作到树形结构"""
        indent = "  " * depth  # 根据深度添加缩进

        for j, loop_action in enumerate(loop_actions):
            action_type = loop_action.get('type', '')
            if action_type == 'loop_group':
                # 嵌套循环组
                loop_text = self._describe_action(loop_action)
                loop_remark = loop_action.get('remark', '')
                nested_parent_id = f"{parent_id}_{j}"
                self.action_tree.insert(parent_id, tk.END, iid=nested_parent_id, values=(indent + loop_text, "", loop_remark), open=False)
                # 递归添加嵌套循环组的动作
                nested_loop_actions = loop_action.get('loop_actions', [])
                self._add_loop_actions_to_tree(nested_parent_id, nested_loop_actions, parent_idx, None, None, None, depth + 1)
            else:
                # 普通动作
                action_text = self._describe_action(loop_action)
                loop_remark = loop_action.get('remark', '')
                self.action_tree.insert(parent_id, tk.END, iid=f"{parent_id}_{j}", values=(indent + action_text, "", loop_remark))

        # 恢复选择状态（只在顶层调用）
        if previous_selection is not None:
            self._restore_selection(scroll_to_end, select_index, previous_selection)

    def _describe_action(self, action):
        action_type = action['type']
        if action_type == 'random_delay':
            return f"随机延时 ({action.get('min_delay', 0):.3f}-{action.get('max_delay', 0):.3f}s)"
        if action_type == 'multiply_delay':
            return f"倍数延时 (基数: {action.get('base_delay', 0):.3f}s)"
        if action_type == 'arithmetic_delay':
            return f"等差延时 (起始: {action.get('start_delay', 0):.3f}s, 步距: {action.get('step_delay', 0):.3f}s)"
        if action_type == 'loop_group':
            loop_count = action.get('loop_count', 1)
            interval_params = action.get('loop_interval', {})
            interval_type = action.get('loop_interval_type', 'fixed')
            if interval_type == 'fixed':
                interval_str = f"固定{interval_params.get('value', 0)}秒"
            elif interval_type == 'range':
                interval_str = f"范围{interval_params.get('min', 0)}-{interval_params.get('max', 0)}秒"
            elif interval_type == 'list':
                values = interval_params.get('values', [])
                interval_str = f"列表{values}"
            else:  # random
                interval_str = f"随机{interval_params.get('min', 0)}-{interval_params.get('max', 0)}秒"
            return f"循环组 (执行{loop_count}次,{interval_str})"
        if action_type == 'numeric_loop':
            start = action.get('start', 0)
            step = action.get('step', 1)
            count = action.get('count', 10)
            text = action.get('input_text', '')
            return f"数值循环 (起始{start},步距{step},次数{count}): 输入'{text}'"
        if action_type == 'variable_input':
            start = action.get('start', 0)
            step = action.get('step', 1)
            text = action.get('input_text', '')
            return f"变量输入 (起始{start},步距{step}): {text}"
        if action_type == 'move':
            return f"移动到 ({action['x']}, {action['y']})"
        if action_type == 'doubleclick':
            button = action.get('button', '')
            return f"双击 {button} at ({action['x']}, {action['y']})"
        if action_type == 'click':
            button = action.get('button', '')
            return f"单击 {button} at ({action['x']}, {action['y']})"
        if action_type == 'keyboard':
            keys = action.get('keys', [])
            if len(keys) > 1:
                return f"组合键: {'+'.join(keys)}"
            if keys:
                return f"按键: {keys[0]}"
            return "按键"
        if action_type == 'scroll':
            direction = "上" if action.get('dy', 0) > 0 else "下"
            return f"滚轮{direction}滚 at ({action['x']}, {action['y']})"
        if action_type == 'screenshot':
            naming = "时间戳" if action.get('naming', 'timestamp') == 'timestamp' else "递增序号"
            return f"截屏: {action.get('filename', 'screenshot')} ({naming})"
        return action_type

    def _format_delay_text(self, index):
        action = self.actions[index]
        if action['type'] == 'random_delay':
            return f"{action.get('min_delay', 0):.3f}-{action.get('max_delay', 0):.3f}s"
        if action['type'] == 'multiply_delay':
            return f"Base: {action.get('base_delay', 0):.3f}s"
        if action['type'] == 'arithmetic_delay':
            return f"Start: {action.get('start_delay', 0):.3f}s Step: {action.get('step_delay', 0):.3f}s"
        if action['type'] == 'loop_group':
            return f"循环{action.get('loop_count', 1)}次"
        if action['type'] == 'numeric_loop':
            start = action.get('start', 0)
            end = action.get('end', 0)
            step = action.get('step', 1)
            count = abs(int((end - start) / step)) + 1 if step != 0 else 0
            return f"{count}次循环"
        if action['type'] == 'screenshot':
            return f"{action.get('delay', 0):.3f}s"
        relative_delay = self._get_relative_delay(index)
        return f"{relative_delay:.3f}s" if relative_delay is not None else "-"

    def _get_relative_delay(self, index):
        """获取相对延时，从前向后遍历一次即可"""
        if not self.actions or index < 0 or index >= len(self.actions):
            return None

        action = self.actions[index]
        if 'time' not in action:
            return None

        # 从前往后计算，更高效
        prev_time = 0
        for j in range(index):
            if 'time' in self.actions[j]:
                prev_time = self.actions[j]['time']
        return max(0.0, action['time'] - prev_time)

    def _get_selected_indices(self):
        """获取选中的动作索引，支持顶层动作和嵌套循环组内的动作"""
        if not hasattr(self, 'action_tree'):
            return tuple()
        selection = self.action_tree.selection()
        if not selection:
            return tuple()

        # 分离顶层动作和循环组内动作
        top_indices = []
        loop_selections = {}  # {parent_idx: [child_indices]}

        for item in selection:
            if item.startswith('loop_'):
                # 格式: loop_顶层索引_第一层子索引_第二层子索引...
                parts = item.split('_')[1:]  # 去掉 'loop' 前缀
                if len(parts) == 1:
                    # 选择的是循环组本身 (loop_0, loop_1 等)
                    parent_idx = int(parts[0])
                    top_indices.append(parent_idx)
                elif len(parts) >= 2:
                    # 选择的是循环组内的动作
                    parent_idx = int(parts[0])
                    child_idx = int(parts[1])
                    if parent_idx not in loop_selections:
                        loop_selections[parent_idx] = []
                    loop_selections[parent_idx].append(child_idx)
            else:
                try:
                    top_indices.append(int(item))
                except ValueError:
                    pass

        # 合并结果
        result = []
        # 添加顶层动作
        result.extend(sorted(top_indices))
        # 添加循环组内动作
        for parent_idx in sorted(loop_selections.keys()):
            child_indices = sorted(loop_selections[parent_idx])
            result.append(('loop', parent_idx, child_indices[0], child_indices[-1], child_indices))

        return tuple(result)

    def _handle_tree_click(self, event):
        column = self.action_tree.identify_column(event.x)
        row_id = self.action_tree.identify_row(event.y)

        # 点击空白区域或表头时隐藏编辑器
        if not row_id or column == '#0':
            self._hide_delay_spinbox()
            return

        # 检测是否是循环组项，用于展开/折叠
        if row_id.startswith('loop_'):
            action = self._get_action_by_row_id(row_id)
            if action and action.get('type') == 'loop_group':
                # 单击循环组：展开或折叠，并选中
                if self.action_tree.item(row_id, 'open'):
                    self.action_tree.item(row_id, open=False)
                else:
                    self.action_tree.item(row_id, open=True)
                # 选中该行
                self.action_tree.selection_set(row_id)
                return

        # 检测Ctrl键 - 只在普通点击时处理选择
        ctrl_pressed = (event.state & 0x4) != 0

        if not ctrl_pressed:
            # 普通点击：选中该行
            self.action_tree.selection_set(row_id)

        # 处理列点击
        if column == "#2":
            self._show_delay_spinbox(row_id, column)
        elif column == "#3":
            self._edit_remark(row_id)
        else:
            self._hide_delay_spinbox()

        # 不阻止默认行为，让Treeview处理Ctrl+多选

    def _on_drag_start(self, event):
        """开始拖动"""
        if self.drag_locked.get():
            return
        row_id = self.action_tree.identify_row(event.y)
        if not row_id:
            return

        # 检测Ctrl键：使用event.state
        ctrl_pressed = (event.state & 0x4) != 0 or (event.state & 0x80) != 0 or (event.state & 0x100) != 0

        if ctrl_pressed:
            # Ctrl+拖动：多选模式，记录起始位置
            self._drag_start_item = row_id
            self._drag_start_pos = (event.x, event.y)
            self._is_dragging = False
            self._drag_start_idx = None
        elif row_id.startswith('loop_'):
            # 循环组内的动作：可以多选但不能拖动排序
            self._drag_start_item = row_id
            self._drag_start_pos = (event.x, event.y)
            self._is_dragging = False
            self._drag_start_idx = None
        elif '_' not in row_id:
            # 普通拖动：排序模式
            self._drag_start_item = row_id
            self._drag_start_pos = (event.x, event.y)
            self._is_dragging = False
            try:
                self._drag_start_idx = int(row_id)
            except ValueError:
                self._drag_start_idx = None

    def _on_drag_motion(self, event):
        """拖动过程中"""
        if self.drag_locked.get():
            return
        if not self._drag_start_pos:
            return

        # 检测Ctrl键
        ctrl_pressed = (event.state & 0x4) != 0 or (event.state & 0x80) != 0 or (event.state & 0x100) != 0

        dx = abs(event.x - self._drag_start_pos[0])
        dy = abs(event.y - self._drag_start_pos[1])

        if dx > 3 or dy > 3:
            self._is_dragging = True

            if ctrl_pressed and self._drag_start_item:
                # Ctrl+拖动：多选模式
                row_id = self.action_tree.identify_row(event.y)
                if row_id and row_id != self._drag_start_item:
                    current_selection = self.action_tree.selection()
                    if row_id not in current_selection:
                        self.action_tree.selection_add(row_id)
                # 滚动视图
                row_id = self.action_tree.identify_row(event.y)
                if row_id:
                    self.action_tree.see(row_id)
            else:
                # 普通拖动：滚动视图
                row_id = self.action_tree.identify_row(event.y)
                if row_id:
                    self.action_tree.see(row_id)

    def _on_drag_end(self, event):
        """结束拖动，重新排序或多选"""
        # 如果拖动被锁定，则不处理
        if self.drag_locked.get():
            self._drag_start_item = None
            self._drag_start_pos = None
            self._is_dragging = False
            self._drag_start_idx = None
            return

        # 如果不是真正拖动（只是点击），则不处理
        if not self._is_dragging:
            self._drag_start_item = None
            self._drag_start_pos = None
            self._is_dragging = False
            self._drag_start_idx = None
            return

        # Ctrl+拖动（多选模式）不需要重新排序
        if self._drag_start_idx is None:
            self._drag_start_item = None
            self._drag_start_pos = None
            self._is_dragging = False
            return

        if not self._drag_start_item:
            self._drag_start_item = None
            self._drag_start_pos = None
            self._is_dragging = False
            self._drag_start_idx = None
            return

        row_id = self.action_tree.identify_row(event.y)
        if row_id and row_id != self._drag_start_item:
            # 循环组内的项不能单独拖动
            if '_' in self._drag_start_item and self._drag_start_item.startswith('loop_'):
                self._drag_start_item = None
                self._drag_start_pos = None
                self._is_dragging = False
                self._drag_start_idx = None
                return

            # 目标索引
            try:
                end_idx = int(row_id)
            except ValueError:
                if '_' in row_id and row_id.startswith('loop_'):
                    self._drag_start_item = None
                    self._drag_start_pos = None
                    self._is_dragging = False
                    self._drag_start_idx = None
                    return
                self._drag_start_item = None
                self._drag_start_pos = None
                self._is_dragging = False
                self._drag_start_idx = None
                return

            start_idx = self._drag_start_idx

            # 如果索引有效，进行重排
            if 0 <= start_idx < len(self.actions) and 0 <= end_idx < len(self.actions) and start_idx != end_idx:
                # 保存动作
                action = self.actions.pop(start_idx)
                # 调整目标索引
                if end_idx > start_idx:
                    end_idx -= 1
                # 插入到新位置
                self.actions.insert(end_idx, action)
                # 刷新列表
                self._update_action_list(select_index=end_idx)

        self._drag_start_item = None
        self._drag_start_pos = None
        self._is_dragging = False
        self._drag_start_idx = None

    def on_tree_double_click(self, event):
        column = self.action_tree.identify_column(event.x)
        row_id = self.action_tree.identify_row(event.y)
        if not row_id:
            return
        self.action_tree.selection_set(row_id)

        # 获取动作类型
        action = None
        try:
            if row_id.startswith('loop_'):
                action = self._get_action_by_row_id(row_id)
            else:
                index = int(row_id)
                if 0 <= index < len(self.actions):
                    action = self.actions[index]
        except (ValueError, IndexError):
            pass

        # 如果无法获取动作，直接返回
        if not action:
            return

        action_type = action.get('type', '')

        # 嵌套的循环组也可以编辑
        if action_type == 'loop_group':
            self._hide_delay_spinbox()
            self.edit_action_time(event)
        elif column == "#2":
            self._show_delay_spinbox(row_id)
        elif column == "#3":
            # 点击备注列：编辑备注
            self._edit_remark(row_id)
        elif action_type in ['click', 'doubleclick', 'move', 'variable_input']:
            # 点击/双击/移动/变量输入动作 - 编辑坐标和内容
            self._hide_delay_spinbox()
            self.edit_action(event)
        else:
            self._hide_delay_spinbox()
            self.edit_action_time(event)

    def _edit_remark(self, row_id):
        """编辑动作备注 - 内联编辑"""
        # 获取动作
        try:
            index = int(row_id)
            if 0 <= index < len(self.actions):
                self._show_inline_remark_editor(row_id, self.actions[index])
        except ValueError:
            if row_id.startswith('loop_') and '_' in row_id:
                action = self._get_action_by_row_id(row_id)
                if action:
                    self._show_inline_remark_editor(row_id, action)

    def _get_action_by_row_id(self, row_id):
        """根据row_id获取对应的动作"""
        if row_id.startswith('loop_'):
            parts = row_id.split('_')[1:]  # 去掉 'loop' 前缀
            # parts[0] = 顶层循环组索引, parts[1] = 第一层子动作索引, ...
            if not parts:
                return None

            # 获取顶层循环组
            idx = int(parts[0])
            if idx >= len(self.actions):
                return None

            action = self.actions[idx]
            if action.get('type') != 'loop_group':
                return None

            # 逐层往下找，递归处理嵌套循环组
            for i in range(1, len(parts)):
                child_idx = int(parts[i])
                # 如果当前action是循环组，获取其loop_actions
                if action.get('type') == 'loop_group':
                    loop_actions = action.get('loop_actions', [])
                else:
                    return None

                if child_idx >= len(loop_actions):
                    return None
                action = loop_actions[child_idx]
                # 如果当前action还是循环组，继续循环

            return action
        else:
            # 顶层动作
            idx = int(row_id)
            if 0 <= idx < len(self.actions):
                return self.actions[idx]
        return None

    def _show_inline_remark_editor(self, row_id, action):
        """显示内联备注编辑器"""
        self._hide_delay_spinbox()

        self.action_tree.update_idletasks()
        bbox = self.action_tree.bbox(row_id, column='#3')
        if not bbox:
            self.action_tree.see(row_id)
            bbox = self.action_tree.bbox(row_id, column='#3')
            if not bbox:
                return

        x, y, width, height = bbox
        self.remark_var.set(action.get('remark', ''))

        self.remark_entry = ttk.Entry(
            self.action_tree,
            textvariable=self.remark_var,
            width=20
        )
        self.remark_entry.place(x=x, y=y, width=width, height=height)
        self.remark_entry.focus_set()
        self.remark_entry.select_range(0, tk.END)

        def save_remark(event=None):
            action['remark'] = self.remark_var.get()
            self._hide_delay_spinbox()
            self._update_action_list()

        def cancel_edit(event=None):
            self._hide_delay_spinbox()

        self.remark_entry.bind('<Return>', save_remark)
        self.remark_entry.bind('<Escape>', cancel_edit)
        self.remark_entry.bind('<FocusOut>', save_remark)

    def _show_delay_spinbox(self, item_id, column="#2"):
        if not self.actions:
            return

        # 检查是否是循环组内的动作
        action = None
        if item_id.startswith('loop_'):
            # 获取循环组内的动作
            action = self._get_action_by_row_id(item_id)
            if not action:
                return
        else:
            try:
                index = int(item_id)
            except ValueError:
                index = int(self.action_tree.index(item_id))

            if index < 0 or index >= len(self.actions):
                return
            action = self.actions[index]

        action_type = action.get('type', '')

        # 根据动作类型获取延时值
        if action_type == 'screenshot':
            current_value = float(action.get('delay', 0.0))
        elif action_type == 'random_delay':
            current_value = float(action.get('min_delay', 0.0))
        elif action_type == 'multiply_delay':
            current_value = float(action.get('base_delay', 0.0))
        elif action_type == 'arithmetic_delay':
            current_value = float(action.get('start_delay', 0.0))
        elif action_type == 'loop_group':
            interval_type = action.get('loop_interval_type', 'fixed')
            interval_params = action.get('loop_interval', {})
            if interval_type == 'fixed':
                current_value = float(interval_params.get('value', 0.0))
            elif interval_type in ('range', 'random'):
                current_value = float(interval_params.get('min', 0.0))
            else:  # list
                values = interval_params.get('values', [0])
                current_value = float(values[0] if values else 0.0)
        elif action_type == 'numeric_loop':
            current_value = float(action.get('interval', 0.5))
        elif action_type == 'variable_input':
            current_value = float(action.get('time', 0.0))
        elif action_type in ['move', 'click', 'doubleclick']:
            relative = self._get_relative_delay(index)
            if relative is None:
                return
            current_value = float(relative)
        else:
            return

        self.action_tree.update_idletasks()
        bbox = self.action_tree.bbox(item_id, column=column)
        if not bbox:
            self.action_tree.see(item_id)
            bbox = self.action_tree.bbox(item_id, column=column)
            if not bbox:
                return

        self._hide_delay_spinbox()
        x, y, width, height = bbox
        self.delay_spinbox_var.set(round(current_value, 3))
        self.delay_spinbox = ttk.Spinbox(
            self.action_tree,
            from_=0.0,
            to=9999.0,
            increment=0.1,
            textvariable=self.delay_spinbox_var
        )
        self.delay_spinbox.place(x=x, y=y, width=width, height=height)
        self.delay_spinbox.focus_set()
        self.delay_spinbox.selection_range(0, tk.END)
        self.delay_spinbox.bind("<Return>", self._apply_spinbox_value)
        self.delay_spinbox.bind("<FocusOut>", self._apply_spinbox_value)
        self.delay_spinbox.bind("<Escape>", self._hide_delay_spinbox)
        self.active_spinbox_item = item_id

    def _apply_spinbox_value(self, *_args):
        if not self.delay_spinbox or self.active_spinbox_item is None:
            return
        if self._spinbox_updating:
            return
        self._spinbox_updating = True
        try:
            new_value = float(self.delay_spinbox_var.get())
            if new_value < 0:
                raise ValueError("时间不能为负数")

            # 获取动作
            action = None
            item_id = self.active_spinbox_item

            if item_id.startswith('loop_'):
                # 循环组内的动作
                action = self._get_action_by_row_id(item_id)
            else:
                try:
                    index = int(item_id)
                except ValueError:
                    index = int(self.action_tree.index(item_id))

                if 0 <= index < len(self.actions):
                    action = self.actions[index]

            if not action:
                return

            action_type = action.get('type', '')

            if action_type == 'screenshot':
                action['delay'] = new_value
            elif action_type == 'random_delay':
                action['min_delay'] = new_value
                # 保持 max_delay >= min_delay
                if action.get('max_delay', 0) < new_value:
                    action['max_delay'] = new_value
            elif action_type == 'multiply_delay':
                action['base_delay'] = new_value
            elif action_type == 'arithmetic_delay':
                action['start_delay'] = new_value
            elif action_type == 'loop_group':
                interval_type = action.get('loop_interval_type', 'fixed')
                interval_params = action.get('loop_interval', {})
                if interval_type == 'fixed':
                    interval_params['value'] = new_value
                elif interval_type in ('range', 'random'):
                    interval_params['min'] = new_value
                elif interval_type == 'list':
                    values = interval_params.get('values', [0])
                    if values:
                        values[0] = new_value
                        interval_params['values'] = values
                action['loop_interval'] = interval_params
            elif action_type == 'numeric_loop':
                action['interval'] = new_value
            elif action_type == 'variable_input':
                action['time'] = new_value
            elif action_type in ['move', 'click', 'doubleclick']:
                # 对于嵌套在循环组内的动作，直接设置时间
                if item_id.startswith('loop_'):
                    action['time'] = new_value
                else:
                    prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
                    old_time = action.get('time', 0)
                    new_absolute_time = prev_time + new_value
                    delta = new_absolute_time - old_time
                    action['time'] = new_absolute_time
                    self._shift_action_times(index + 1, delta)
            else:
                return

            # 对于嵌套动作，找到其父循环组的索引并刷新
            if item_id.startswith('loop_'):
                parts = item_id.split('_')
                if len(parts) >= 2:
                    parent_idx = int(parts[1])
                    self._update_action_list(select_index=parent_idx)
            else:
                self._update_action_list(select_index=index)
        except ValueError as exc:
            messagebox.showerror("错误", str(exc))
        finally:
            self._spinbox_updating = False
            self._hide_delay_spinbox()

    def _hide_delay_spinbox(self, *_args):
        if self.delay_spinbox:
            self.delay_spinbox.destroy()
            self.delay_spinbox = None
        if hasattr(self, 'remark_entry') and self.remark_entry:
            self.remark_entry.destroy()
            self.remark_entry = None
        self.active_spinbox_item = None

    def _on_tree_scroll(self, *args):
        if hasattr(self, 'action_scrollbar') and self.action_scrollbar:
            self.action_scrollbar.set(*args)
        self._hide_delay_spinbox()
    
    def delete_selected(self):
        """删除选中的动作"""
        try:
            # 获取选中项的索引
            selection = self._get_selected_indices()
            if not selection:
                messagebox.showinfo("提示", "请先选择要删除的动作")
                return

            # 分离顶层动作和循环组内动作
            top_indices = []
            loop_deletes = {}  # {parent_idx: [child_indices]}

            for item in selection:
                if isinstance(item, tuple) and len(item) >= 5 and item[0] == 'loop':
                    # 循环组内动作: ('loop', parent_idx, start, end, child_indices_list)
                    parent_idx = item[1]
                    child_indices = item[4]  # 包含所有选中的子索引
                    if parent_idx not in loop_deletes:
                        loop_deletes[parent_idx] = []
                    loop_deletes[parent_idx].extend(child_indices)
                else:
                    # 顶层动作（整数索引）
                    top_indices.append(item)

            deleted_count = 0

            # 从后往前删除顶层动作（添加边界检查）
            for index in reversed(sorted(top_indices)):
                if 0 <= index < len(self.actions):
                    del self.actions[index]
                    deleted_count += 1

            # 删除循环组内动作（从后往前，添加边界检查）
            for parent_idx in sorted(loop_deletes.keys(), reverse=True):
                if 0 <= parent_idx < len(self.actions):
                    action = self.actions[parent_idx]
                    if action.get('type') in ['loop_group', 'numeric_loop'] and 'loop_actions' in action:
                        loop_actions = action['loop_actions']
                        child_indices = sorted(loop_deletes[parent_idx], reverse=True)
                        for child_idx in child_indices:
                            if 0 <= child_idx < len(loop_actions):
                                del loop_actions[child_idx]
                                deleted_count += 1

            next_index = None
            if self.actions:
                if top_indices:
                    next_index = min(top_indices[0], len(self.actions) - 1)
                else:
                    next_index = 0

            self._update_action_list(select_index=next_index)
            self.status_var.set(f"已删除 {deleted_count} 个动作")
        except Exception as e:
            messagebox.showerror("错误", f"删除失败: {str(e)}")
    
    def update_action_list_display(self, *_args, **_kwargs):
        """线程安全地刷新动作列表并滚动到底部"""
        self.root.after(0, lambda: self._update_action_list(scroll_to_end=True))
    
    def insert_click_action(self):
        """插入点击动作"""
        self.in_dialog_operation = True

        # 获取当前选中项
        selection = self._get_selected_indices()

        if selection:
            insert_pos = selection[-1] + 1
        else:
            insert_pos = len(self.actions)

        # 计算相对延时
        relative_delay = 0
        if insert_pos > 0 and 'time' in self.actions[insert_pos - 1]:
            relative_delay = self.actions[insert_pos - 1].get('time', 0)

        dialog = tk.Toplevel(self.root)
        dialog.title("插入点击")
        dialog.geometry("500x180")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 目标位置
        ttk.Label(frame, text="点击位置:").grid(row=0, column=0, padx=5, pady=8, sticky="e")
        x_var = tk.StringVar(value="")
        y_var = tk.StringVar(value="")

        pos_frame = ttk.Frame(frame)
        pos_frame.grid(row=0, column=1, padx=5, pady=8, sticky="w")

        # 录制位置按钮（放在前面）
        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        x_entry = ttk.Entry(pos_frame, width=8, textvariable=x_var)
        x_entry.pack(side=tk.LEFT, padx=2)

        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        y_entry = ttk.Entry(pos_frame, width=8, textvariable=y_var)
        y_entry.pack(side=tk.LEFT, padx=2)

        # 相对延时
        ttk.Label(frame, text="相对延时(秒):").grid(row=1, column=0, padx=5, pady=8, sticky="e")
        delay_var = tk.DoubleVar(value=relative_delay)
        ttk.Entry(frame, width=15, textvariable=delay_var).grid(row=1, column=1, padx=5, pady=8, sticky="w")

        # 状态标签
        status_label = ttk.Label(frame, text="", foreground="blue")
        status_label.grid(row=2, column=0, columnspan=2, pady=2)

        # 录制位置函数
        def start_record_position():
            dialog.withdraw()
            dialog.update()
            time.sleep(0.5)
            status_label.config(text="请在2秒后点击目标位置...")
            dialog.update()
            time.sleep(2)
            status_label.config(text="请现在点击目标位置...")
            dialog.update()

            # 等待点击
            def on_click(_x, _y, _button, pressed):
                if pressed:
                    return False

            listener = mouse.Listener(on_click=on_click)
            listener.start()
            listener.join()

            # 使用win32api获取最终鼠标位置
            try:
                import win32api
                x, y = win32api.GetCursorPos()
                x_var.set(str(x))
                y_var.set(str(y))
                status_label.config(text=f"已录制: ({x}, {y})", foreground="green")
            except Exception:
                status_label.config(text="录制失败，请手动输入", foreground="red")

            dialog.deiconify()
            dialog.update()

        record_btn.config(command=start_record_position)

        def confirm():
            try:
                target_x = int(x_var.get()) if x_var.get().strip() else None
                target_y = int(y_var.get()) if y_var.get().strip() else None

                if target_x is None or target_y is None:
                    raise ValueError("请输入有效的坐标")

                action = {
                    'type': 'click',
                    'x': target_x,
                    'y': target_y,
                    'button': 'left',
                    'time': delay_var.get()
                }

                self.actions.insert(insert_pos, action)
                self._update_action_list(select_index=insert_pos)
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=3, column=0, columnspan=2, pady=5)

    def insert_random_delay(self):
        """插入延时动作（支持随机延时和指数延时）"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("插入延时")
        dialog.geometry("500x200")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        dialog.grab_set()

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

        # 延时类型选择
        type_var = tk.StringVar(value="random")
        
        # 延时设置框架
        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 类型选择
        ttk.Label(frame, text="延时类型:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        type_frame = ttk.Frame(frame)
        type_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(type_frame, text="随机范围", variable=type_var, value="random",
                       command=lambda: update_ui("random")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(type_frame, text="倍数增长", variable=type_var, value="multiply",
                       command=lambda: update_ui("multiply")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(type_frame, text="等差递增", variable=type_var, value="arithmetic",
                       command=lambda: update_ui("arithmetic")).pack(side=tk.LEFT, padx=2)

        # 随机延时参数
        random_frame = ttk.Frame(frame)
        random_frame.grid(row=1, column=0, columnspan=2, sticky="ew")

        ttk.Label(random_frame, text="最小延时(秒):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        min_entry = ttk.Entry(random_frame, width=10, textvariable=self.default_delay_min)
        min_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(random_frame, text="最大延时(秒):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        max_entry = ttk.Entry(random_frame, width=10, textvariable=self.default_delay_max)
        max_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # 倍数延时参数
        exp_frame = ttk.Frame(frame)
        exp_frame.grid(row=1, column=0, columnspan=2, sticky="ew")

        ttk.Label(exp_frame, text="基数延时(秒):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        base_entry = ttk.Entry(exp_frame, width=10, textvariable=self.default_delay_base)
        base_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(exp_frame, text="(每轮循环增加一倍)").grid(row=0, column=2, padx=5, pady=5, sticky="w")

        # 等差递增参数
        arithmetic_frame = ttk.Frame(frame)
        arithmetic_frame.grid(row=1, column=0, columnspan=2, sticky="ew")

        ttk.Label(arithmetic_frame, text="起始延时(秒):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        arithmetic_start_var = tk.DoubleVar(value=0.5)
        ttk.Entry(arithmetic_frame, width=10, textvariable=arithmetic_start_var).grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(arithmetic_frame, text="步距(秒):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        arithmetic_step_var = tk.DoubleVar(value=0.1)
        ttk.Entry(arithmetic_frame, width=10, textvariable=arithmetic_step_var).grid(row=1, column=1, padx=5, pady=5, sticky="w")

        def update_ui(delay_type):
            if delay_type == "random":
                random_frame.grid()
                exp_frame.grid_remove()
                arithmetic_frame.grid_remove()
            elif delay_type == "multiply":
                random_frame.grid_remove()
                exp_frame.grid()
                arithmetic_frame.grid_remove()
            else:  # arithmetic
                random_frame.grid_remove()
                exp_frame.grid_remove()
                arithmetic_frame.grid()

        # 初始化UI状态
        update_ui("random")

        def confirm():
            try:
                delay_type = type_var.get()

                if delay_type == "random":
                    min_delay = float(min_entry.get())
                    max_delay = float(max_entry.get())
                    if (min_delay < 0 or max_delay < min_delay):
                        raise ValueError("无效的延时范围")

                    action = {
                        'type': 'random_delay',
                        'min_delay': min_delay,
                        'max_delay': max_delay,
                        'time': 0
                    }
                elif delay_type == "multiply":
                    base_delay = float(base_entry.get())
                    if base_delay < 0:
                        raise ValueError("基数不能为负数")

                    action = {
                        'type': 'multiply_delay',
                        'base_delay': base_delay,
                        'time': 0
                    }
                else:  # arithmetic
                    start_delay = float(arithmetic_start_var.get())
                    step_delay = float(arithmetic_step_var.get())
                    if start_delay < 0 or step_delay < 0:
                        raise ValueError("延时和步距不能为负数")

                    action = {
                        'type': 'arithmetic_delay',
                        'start_delay': start_delay,
                        'step_delay': step_delay,
                        'time': 0
                    }

                # 获取当前选中项
                selection = self._get_selected_indices()

                if selection:
                    # 插入到选中项之后
                    insert_pos = selection[-1] + 1
                else:
                    insert_pos = len(self.actions)

                if insert_pos > 0 and self.actions:
                    action['time'] = self.actions[insert_pos - 1].get('time', 0)

                self.actions.insert(insert_pos, action)
                self._update_action_list(select_index=insert_pos)

                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=2, column=0, columnspan=2, pady=10)

    def insert_loop_group(self):
        """插入循环组：执行选中动作N次"""
        self.in_dialog_operation = True

        selection = self._get_selected_indices()

        if not selection:
            messagebox.showinfo("提示", "请先选择要作为循环组的动作")
            self.in_dialog_operation = False
            return

        # 检查是否是循环组内的选择
        is_loop_inner = False
        parent_loop_idx = None
        start_idx = 0
        end_idx = 0

        if isinstance(selection[0], tuple) and selection[0][0] == 'loop':
            # 选择的是循环组内的动作
            # 新格式: ('loop', parent_idx, start_child, end_child, [child_list])
            is_loop_inner = True
            parent_loop_idx = selection[0][1]
            start_idx = selection[0][2]  # start child index
            end_idx = selection[0][3]    # end child index

            parent_action = self.actions[parent_loop_idx]
            loop_actions = parent_action.get('loop_actions', [])
            loop_actions_selected = loop_actions[start_idx:end_idx+1]
            info_text = f"循环组{parent_loop_idx+1}内: 第{start_idx+1}-{end_idx+1}个动作 (共{end_idx-start_idx+1}个)"
        else:
            # 普通选择
            start_idx = selection[0]
            end_idx = selection[-1]
            loop_actions_selected = self.actions[start_idx:end_idx+1]
            info_text = f"选中的动作: {start_idx+1} - {end_idx+1} (共{end_idx-start_idx+1}个)"

        dialog = tk.Toplevel(self.root)
        dialog.title("插入循环组")
        dialog.geometry("450x200")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 循环次数
        ttk.Label(frame, text="循环次数:").grid(row=0, column=0, padx=5, pady=10, sticky="e")
        loop_count_var = tk.IntVar(value=3)
        ttk.Entry(frame, width=15, textvariable=loop_count_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")

        # 循环间隔类型选择
        interval_type_var = tk.StringVar(value="fixed")

        ttk.Label(frame, text="间隔模式:").grid(row=1, column=0, padx=5, pady=10, sticky="e")
        interval_type_frame = ttk.Frame(frame)
        interval_type_frame.grid(row=1, column=1, padx=5, pady=10, sticky="w")

        ttk.Radiobutton(interval_type_frame, text="固定", variable=interval_type_var, value="fixed",
                        command=lambda: update_interval_ui("fixed")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(interval_type_frame, text="范围", variable=interval_type_var, value="range",
                        command=lambda: update_interval_ui("range")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(interval_type_frame, text="列表", variable=interval_type_var, value="list",
                        command=lambda: update_interval_ui("list")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(interval_type_frame, text="随机", variable=interval_type_var, value="random",
                        command=lambda: update_interval_ui("random")).pack(side=tk.LEFT, padx=2)

        # 间隔参数框架
        interval_frame = ttk.Frame(frame)
        interval_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        # 固定间隔
        fixed_frame = ttk.Frame(interval_frame)
        ttk.Label(fixed_frame, text="间隔(秒):").pack(side=tk.LEFT, padx=2)
        loop_interval_var = tk.DoubleVar(value=0.5)
        ttk.Entry(fixed_frame, width=10, textvariable=loop_interval_var).pack(side=tk.LEFT, padx=2)

        # 范围间隔
        range_frame = ttk.Frame(interval_frame)
        ttk.Label(range_frame, text="最小:").pack(side=tk.LEFT, padx=2)
        range_min_var = tk.DoubleVar(value=0.5)
        ttk.Entry(range_frame, width=8, textvariable=range_min_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(range_frame, text="最大:").pack(side=tk.LEFT, padx=2)
        range_max_var = tk.DoubleVar(value=2.0)
        ttk.Entry(range_frame, width=8, textvariable=range_max_var).pack(side=tk.LEFT, padx=2)

        # 列表间隔
        list_frame = ttk.Frame(interval_frame)
        ttk.Label(list_frame, text="列表(逗号分隔):").pack(side=tk.LEFT, padx=2)
        list_values_var = tk.StringVar(value="0.5,1.0,1.5,2.0")
        ttk.Entry(list_frame, width=20, textvariable=list_values_var).pack(side=tk.LEFT, padx=2)

        # 随机间隔
        random_frame = ttk.Frame(interval_frame)
        ttk.Label(random_frame, text="最小:").pack(side=tk.LEFT, padx=2)
        random_min_var = tk.DoubleVar(value=0.5)
        ttk.Entry(random_frame, width=8, textvariable=random_min_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(random_frame, text="最大:").pack(side=tk.LEFT, padx=2)
        random_max_var = tk.DoubleVar(value=2.0)
        ttk.Entry(random_frame, width=8, textvariable=random_max_var).pack(side=tk.LEFT, padx=2)

        def update_interval_ui(interval_type):
            fixed_frame.grid_remove()
            range_frame.grid_remove()
            list_frame.grid_remove()
            random_frame.grid_remove()
            if interval_type == "fixed":
                fixed_frame.grid()
            elif interval_type == "range":
                range_frame.grid()
            elif interval_type == "list":
                list_frame.grid()
            else:  # random
                random_frame.grid()

        # 初始化UI状态
        update_interval_ui("fixed")

        # 信息显示
        info_label = ttk.Label(frame, text=info_text)
        info_label.grid(row=3, column=0, columnspan=2, pady=10)

        def confirm():
            try:
                loop_count = loop_count_var.get()

                if loop_count <= 0:
                    raise ValueError("循环次数必须大于0")

                interval_type = interval_type_var.get()
                if interval_type == "fixed":
                    loop_interval = loop_interval_var.get()
                    if loop_interval < 0:
                        raise ValueError("间隔不能为负数")
                    interval_params = {'type': 'fixed', 'value': loop_interval}
                elif interval_type == "range":
                    min_val = range_min_var.get()
                    max_val = range_max_var.get()
                    if min_val < 0 or max_val < min_val:
                        raise ValueError("无效的范围")
                    interval_params = {'type': 'range', 'min': min_val, 'max': max_val}
                elif interval_type == "list":
                    values_str = list_values_var.get().strip()
                    if not values_str:
                        raise ValueError("列表不能为空")
                    try:
                        values = [float(v.strip()) for v in values_str.split(',')]
                    except ValueError:
                        raise ValueError("列表格式无效，请使用逗号分隔的数字")
                    interval_params = {'type': 'list', 'values': values}
                else:  # random
                    min_val = random_min_var.get()
                    max_val = random_max_var.get()
                    if min_val < 0 or max_val < min_val:
                        raise ValueError("无效的范围")
                    interval_params = {'type': 'random', 'min': min_val, 'max': max_val}

                if is_loop_inner:
                    # 在循环组内插入循环组
                    parent_action = self.actions[parent_loop_idx]
                    loop_actions = parent_action.get('loop_actions', [])

                    # 获取选中的动作
                    selected_actions = loop_actions[start_idx:end_idx+1]

                    # 创建新的嵌套循环组
                    nested_loop = {
                        'type': 'loop_group',
                        'loop_count': loop_count,
                        'loop_interval_type': interval_params['type'],
                        'loop_interval': interval_params,
                        'loop_actions': selected_actions,
                        'time': 0
                    }

                    # 删除选中的动作并插入嵌套循环组
                    del loop_actions[start_idx:end_idx+1]
                    loop_actions.insert(start_idx, nested_loop)

                    # 更新父循环组
                    parent_action['loop_actions'] = loop_actions
                    self._update_action_list(select_index=parent_loop_idx)
                else:
                    # 创建循环组动作
                    action = {
                        'type': 'loop_group',
                        'loop_count': loop_count,
                        'loop_interval_type': interval_params['type'],
                        'loop_interval': interval_params,
                        'loop_actions': loop_actions_selected,
                        'time': self.actions[start_idx].get('time', 0) if self.actions else 0
                    }

                    # 删除选中的动作，并在原位置插入循环组
                    del self.actions[start_idx:end_idx+1]
                    self.actions.insert(start_idx, action)
                    self._update_action_list(select_index=start_idx)

                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=4, column=0, columnspan=2, pady=10)

    def insert_variable(self):
        """插入变量：在选中动作上添加一个带变量的输入动作"""
        self.in_dialog_operation = True

        # 获取当前选中项
        selection = self._get_selected_indices()

        if selection:
            # 如果有选中项，在选中项之后插入
            insert_pos = selection[-1] + 1
        else:
            # 否则添加到末尾
            insert_pos = len(self.actions)

        # 计算相对延时（与上一个动作的时间差）
        relative_delay = 0
        if insert_pos > 0 and 'time' in self.actions[insert_pos - 1]:
            relative_delay = self.actions[insert_pos - 1].get('time', 0)

        dialog = tk.Toplevel(self.root)
        dialog.title("插入变量")
        dialog.geometry("550x320")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 输入文本（带变量初始值）
        ttk.Label(frame, text="输入文本:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        input_text_var = tk.StringVar(value="text{0}")
        ttk.Entry(frame, width=25, textvariable=input_text_var).grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # 提示
        ttk.Label(frame, text="(如: text{0} 表示从0开始)").grid(row=1, column=1, padx=5, pady=2, sticky="w")

        # 步距
        ttk.Label(frame, text="步距:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        step_var = tk.IntVar(value=1)
        ttk.Entry(frame, width=15, textvariable=step_var).grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # 目标位置
        ttk.Label(frame, text="点击位置:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        x_var = tk.StringVar(value="")
        y_var = tk.StringVar(value="")

        pos_frame = ttk.Frame(frame)
        pos_frame.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        x_entry = ttk.Entry(pos_frame, width=8, textvariable=x_var)
        x_entry.pack(side=tk.LEFT, padx=2)

        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        y_entry = ttk.Entry(pos_frame, width=8, textvariable=y_var)
        y_entry.pack(side=tk.LEFT, padx=2)

        # 录制位置按钮
        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        # 相对延时
        ttk.Label(frame, text="相对延时(秒):").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        delay_var = tk.DoubleVar(value=relative_delay)
        ttk.Entry(frame, width=15, textvariable=delay_var).grid(row=4, column=1, padx=5, pady=5, sticky="w")

        # 清空文本选项
        clear_text_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="输入前清空文本框", variable=clear_text_var).grid(row=5, column=0, columnspan=2, pady=2, sticky="w")
        clear_text_var.set(True)  # 默认清空文本框

        # 状态标签
        status_label = ttk.Label(frame, text="", foreground="blue")
        status_label.grid(row=6, column=0, columnspan=2, pady=2)

        # 录制位置函数
        def start_record_position():
            dialog.withdraw()
            dialog.update()
            time.sleep(0.5)  # 等待对话框完全隐藏
            status_label.config(text="请在2秒后点击目标位置...")
            dialog.update()
            time.sleep(2)
            status_label.config(text="请现在点击目标位置...")
            dialog.update()

            # 等待点击，位置由win32api获取
            def on_click(___x, ___y, __button, pressed):
                if pressed:
                    return False

            listener = mouse.Listener(on_click=on_click)
            listener.start()
            listener.join()

            # 使用win32api获取最终鼠标位置
            try:
                import win32api
                x, y = win32api.GetCursorPos()
                x_var.set(str(x))
                y_var.set(str(y))
                status_label.config(text=f"已录制: ({x}, {y})", foreground="green")
            except Exception:
                status_label.config(text="录制失败，请手动输入", foreground="red")

            dialog.deiconify()
            dialog.update()

        record_btn.config(command=start_record_position)

        def confirm():
            try:
                input_text = input_text_var.get()
                if not input_text:
                    raise ValueError("输入文本不能为空")

                # 解析变量初始值
                import re
                match = re.search(r'\{(\d+)\}', input_text)
                if match:
                    start_value = int(match.group(1))
                else:
                    start_value = 0

                step = step_var.get()
                if step == 0:
                    raise ValueError("步距不能为0")

                # 解析目标位置
                target_x = None
                target_y = None
                if x_var.get().strip():
                    target_x = int(x_var.get())
                if y_var.get().strip():
                    target_y = int(y_var.get())

                action = {
                    'type': 'variable_input',
                    'start': start_value,
                    'step': step,
                    'input_text': input_text,
                    'x': target_x,
                    'y': target_y,
                    'clear_text': clear_text_var.get(),
                    'time': delay_var.get()
                }

                self.actions.insert(insert_pos, action)
                self._update_action_list(select_index=insert_pos)
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=7, column=0, columnspan=2, pady=8)

    def _shift_action_times(self, start_index, delta):
        """整体平移从指定索引开始的动作时间，保持相对间隔不变"""
        if delta == 0 or start_index >= len(self.actions):
            return

        for i in range(start_index, len(self.actions)):
            if 'time' in self.actions[i]:
                self.actions[i]['time'] = max(0, self.actions[i]['time'] + delta)

    def edit_action(self, event):
        """编辑动作 - 根据动作类型显示不同的编辑对话框"""
        try:
            # 获取选中的动作
            selection = self._get_selected_indices()
            if not selection:
                return
            row_id = self.action_tree.selection()[0] if self.action_tree.selection() else None
            if not row_id:
                return

            self._hide_delay_spinbox()

            # 判断是嵌套动作还是顶层动作
            is_nested = row_id.startswith('loop_')
            if is_nested:
                action = self._get_action_by_row_id(row_id)
                if not action:
                    return
                action_type = action.get('type', '')
                # 对于嵌套的点击/移动动作，可以编辑坐标和时间
                if action_type in ['click', 'doubleclick', 'move']:
                    self._edit_nested_click_action(action, row_id)
                    return
                elif action_type == 'variable_input':
                    self._edit_nested_variable_input_action(action, row_id)
                    return
                else:
                    self.edit_action_time(event)
                    return
            else:
                index = selection[0]
                action = self.actions[index]
                action_type = action.get('type', '')

            # 点击动作 - 编辑坐标
            if action_type in ['click', 'doubleclick', 'move']:
                self._edit_click_action(action, index)
            # 延时动作 - 编辑延时参数
            elif action_type in ['random_delay', 'multiply_delay', 'arithmetic_delay']:
                self.edit_action_time(event)
            # 输入动作 - 编辑输入
            elif action_type == 'variable_input':
                self._edit_variable_input_action(action, index)
            # 循环组 - 显示编辑
            elif action_type == 'loop_group':
                self.edit_action_time(event)
            # 其他动作
            else:
                messagebox.showinfo("提示", "该动作类型不支持编辑")
        except Exception as e:
            messagebox.showerror("错误", f"编辑失败: {str(e)}")

    def _edit_click_action(self, action, index):
        """编辑点击动作的坐标"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑点击动作")
        dialog.geometry("500x220")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 动作类型
        action_type = action.get('type', 'click')
        type_text = {"click": "点击", "doubleclick": "双击", "move": "移动"}
        ttk.Label(frame, text=f"动作类型: {type_text.get(action_type, action_type)}").grid(row=0, column=0, columnspan=2, pady=10)

        # 目标位置
        ttk.Label(frame, text="点击位置:").grid(row=1, column=0, padx=5, pady=10, sticky="e")
        x_var = tk.StringVar(value=str(action.get('x', '')))
        y_var = tk.StringVar(value=str(action.get('y', '')))

        pos_frame = ttk.Frame(frame)
        pos_frame.grid(row=1, column=1, padx=5, pady=10, sticky="w")

        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        x_entry = ttk.Entry(pos_frame, width=8, textvariable=x_var)
        x_entry.pack(side=tk.LEFT, padx=2)

        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        y_entry = ttk.Entry(pos_frame, width=8, textvariable=y_var)
        y_entry.pack(side=tk.LEFT, padx=2)

        # 录制位置按钮
        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        # 相对延时
        prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
        relative_delay = action.get('time', 0) - prev_time

        ttk.Label(frame, text="相对延时(秒):").grid(row=2, column=0, padx=5, pady=10, sticky="e")
        delay_var = tk.DoubleVar(value=relative_delay)
        ttk.Entry(frame, width=15, textvariable=delay_var).grid(row=2, column=1, padx=5, pady=10, sticky="w")

        # 状态标签
        status_label = ttk.Label(frame, text="", foreground="blue")
        status_label.grid(row=3, column=0, columnspan=2, pady=5)

        # 录制位置函数
        def start_record_position():
            dialog.withdraw()
            dialog.update()
            time.sleep(0.5)
            status_label.config(text="请在2秒后点击目标位置...")
            dialog.update()
            time.sleep(2)
            status_label.config(text="请现在点击目标位置...")
            dialog.update()

            def on_click(_x, _y, _button, pressed):
                if pressed:
                    return False

            listener = mouse.Listener(on_click=on_click)
            listener.start()
            listener.join()

            # 获取鼠标位置
            try:
                x, y = pyautogui.position()
                x_var.set(str(x))
                y_var.set(str(y))
                status_label.config(text=f"已录制: ({x}, {y})")
            except Exception as e:
                status_label.config(text=f"录制失败: {str(e)}")

            dialog.deiconify()
            dialog.lift()

        record_btn.config(command=start_record_position)

        def confirm():
            try:
                target_x = x_var.get().strip()
                target_y = y_var.get().strip()

                if not target_x or not target_y:
                    raise ValueError("请输入有效的坐标")

                target_x = int(target_x)
                target_y = int(target_y)

                # 更新动作
                action['x'] = target_x
                action['y'] = target_y

                # 计算新的绝对时间
                prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
                new_absolute_time = prev_time + delay_var.get()
                old_time = action.get('time', 0)
                delta = new_absolute_time - old_time
                action['time'] = new_absolute_time
                self._shift_action_times(index + 1, delta)

                self._update_action_list(select_index=index)
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=4, column=0, columnspan=2, pady=10)

    def _edit_variable_input_action(self, action, index):
        """编辑变量输入动作"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑变量输入")
        dialog.geometry("550x320")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 输入文本（带变量初始值）
        ttk.Label(frame, text="输入文本:").grid(row=0, column=0, padx=5, pady=10, sticky="e")
        input_text_var = tk.StringVar(value=action.get('value', 'text{0}'))
        ttk.Entry(frame, width=25, textvariable=input_text_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")

        # 提示
        ttk.Label(frame, text="(如: text{0} 表示从0开始)").grid(row=1, column=1, padx=5, pady=2, sticky="w")

        # 步距
        ttk.Label(frame, text="步距:").grid(row=2, column=0, padx=5, pady=10, sticky="e")
        step_var = tk.IntVar(value=action.get('step', 1))
        ttk.Entry(frame, width=15, textvariable=step_var).grid(row=2, column=1, padx=5, pady=10, sticky="w")

        # 目标位置
        ttk.Label(frame, text="点击位置:").grid(row=3, column=0, padx=5, pady=10, sticky="e")
        x_var = tk.StringVar(value=str(action.get('x', '')))
        y_var = tk.StringVar(value=str(action.get('y', '')))

        pos_frame = ttk.Frame(frame)
        pos_frame.grid(row=3, column=1, padx=5, pady=10, sticky="w")

        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        x_entry = ttk.Entry(pos_frame, width=8, textvariable=x_var)
        x_entry.pack(side=tk.LEFT, padx=2)

        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        y_entry = ttk.Entry(pos_frame, width=8, textvariable=y_var)
        y_entry.pack(side=tk.LEFT, padx=2)

        # 录制按钮
        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        # 相对延时
        prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
        relative_delay = action.get('time', 0) - prev_time

        ttk.Label(frame, text="相对延时(秒):").grid(row=4, column=0, padx=5, pady=10, sticky="e")
        delay_var = tk.DoubleVar(value=relative_delay)
        ttk.Entry(frame, width=15, textvariable=delay_var).grid(row=4, column=1, padx=5, pady=10, sticky="w")

        # 清空文本选项
        clear_text_var = tk.BooleanVar(value=action.get('clear_text', True))
        ttk.Checkbutton(frame, text="输入前清空文本框", variable=clear_text_var).grid(row=5, column=0, columnspan=2, pady=5, sticky="w")

        # 状态标签
        status_label = ttk.Label(frame, text="", foreground="blue")
        status_label.grid(row=6, column=0, columnspan=2, pady=5)

        # 录制函数
        def start_record_position():
            dialog.withdraw()
            dialog.update()
            time.sleep(0.5)
            status_label.config(text="请在2秒后点击目标位置...")
            dialog.update()
            time.sleep(2)
            status_label.config(text="请现在点击目标位置...")
            dialog.update()

            def on_click(_x, _y, _button, pressed):
                if pressed:
                    return False

            listener = mouse.Listener(on_click=on_click)
            listener.start()
            listener.join()

            try:
                import win32api
                x, y = win32api.GetCursorPos()
                x_var.set(str(x))
                y_var.set(str(y))
                status_label.config(text=f"已录制: ({x}, {y})", foreground="green")
            except Exception:
                status_label.config(text="录制失败，请手动输入", foreground="red")

            dialog.deiconify()
            dialog.update()

        record_btn.config(command=start_record_position)

        def confirm():
            try:
                input_text = input_text_var.get()
                if not input_text:
                    raise ValueError("输入文本不能为空")

                # 解析变量初始值
                import re
                match = re.search(r'\{(\d+)\}', input_text)
                if match:
                    start_value = int(match.group(1))
                else:
                    start_value = 0

                step = step_var.get()
                if step < 0:
                    raise ValueError("步距不能为负数")

                target_x = x_var.get().strip()
                target_y = y_var.get().strip()

                if not target_x or not target_y:
                    raise ValueError("请输入有效的坐标")

                target_x = int(target_x)
                target_y = int(target_y)

                # 更新动作
                action['value'] = input_text
                action['start'] = start_value
                action['step'] = step
                action['x'] = target_x
                action['y'] = target_y
                action['clear_text'] = clear_text_var.get()

                # 计算新的绝对时间
                prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
                new_absolute_time = prev_time + delay_var.get()
                old_time = action.get('time', 0)
                delta = new_absolute_time - old_time
                action['time'] = new_absolute_time
                self._shift_action_times(index + 1, delta)

                self._update_action_list(select_index=index)
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=4, column=0, columnspan=2, pady=10)

    def _edit_nested_click_action(self, action, row_id):
        """编辑嵌套的点击动作坐标"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑点击动作")
        dialog.geometry("500x220")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 动作类型
        action_type = action.get('type', 'click')
        type_text = {"click": "点击", "doubleclick": "双击", "move": "移动"}
        ttk.Label(frame, text=f"动作类型: {type_text.get(action_type, action_type)}").grid(row=0, column=0, columnspan=2, pady=10)

        # 目标位置
        ttk.Label(frame, text="点击位置:").grid(row=1, column=0, padx=5, pady=10, sticky="e")
        x_var = tk.StringVar(value=str(action.get('x', '')))
        y_var = tk.StringVar(value=str(action.get('y', '')))

        pos_frame = ttk.Frame(frame)
        pos_frame.grid(row=1, column=1, padx=5, pady=10, sticky="w")

        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        x_entry = ttk.Entry(pos_frame, width=8, textvariable=x_var)
        x_entry.pack(side=tk.LEFT, padx=2)

        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        y_entry = ttk.Entry(pos_frame, width=8, textvariable=y_var)
        y_entry.pack(side=tk.LEFT, padx=2)

        # 录制位置按钮
        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        # 相对延时
        ttk.Label(frame, text="相对延时(秒):").grid(row=2, column=0, padx=5, pady=10, sticky="e")
        delay_var = tk.DoubleVar(value=action.get('time', 0))
        ttk.Entry(frame, width=15, textvariable=delay_var).grid(row=2, column=1, padx=5, pady=10, sticky="w")

        # 状态标签
        status_label = ttk.Label(frame, text="", foreground="blue")
        status_label.grid(row=3, column=0, columnspan=2, pady=5)

        # 录制位置函数
        def start_record_position():
            dialog.withdraw()
            dialog.update()
            time.sleep(0.5)
            status_label.config(text="请在2秒后点击目标位置...")
            dialog.update()
            time.sleep(2)
            status_label.config(text="请现在点击目标位置...")
            dialog.update()

            def on_click(_x, _y, _button, pressed):
                if pressed:
                    return False

            listener = mouse.Listener(on_click=on_click)
            listener.start()
            listener.join()

            try:
                import win32api
                x, y = win32api.GetCursorPos()
                x_var.set(str(x))
                y_var.set(str(y))
                status_label.config(text=f"已录制: ({x}, {y})", foreground="green")
            except Exception:
                status_label.config(text="录制失败，请手动输入", foreground="red")

            dialog.deiconify()
            dialog.update()

        record_btn.config(command=start_record_position)

        def confirm():
            try:
                target_x = x_var.get().strip()
                target_y = y_var.get().strip()

                if not target_x or not target_y:
                    raise ValueError("请输入有效的坐标")

                target_x = int(target_x)
                target_y = int(target_y)

                # 更新动作
                action['x'] = target_x
                action['y'] = target_y
                action['time'] = delay_var.get()

                # 刷新列表
                self._update_action_list()
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=4, column=0, columnspan=2, pady=10)

    def _edit_nested_variable_input_action(self, action, row_id):
        """编辑嵌套的变量输入动作"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑变量输入")
        dialog.geometry("550x320")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 输入文本（带变量初始值）
        ttk.Label(frame, text="输入文本:").grid(row=0, column=0, padx=5, pady=10, sticky="e")
        input_text_var = tk.StringVar(value=action.get('value', 'text{0}'))
        ttk.Entry(frame, width=25, textvariable=input_text_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")

        # 提示
        ttk.Label(frame, text="(如: text{0} 表示从0开始)").grid(row=1, column=1, padx=5, pady=2, sticky="w")

        # 步距
        ttk.Label(frame, text="步距:").grid(row=2, column=0, padx=5, pady=10, sticky="e")
        step_var = tk.IntVar(value=action.get('step', 1))
        ttk.Entry(frame, width=15, textvariable=step_var).grid(row=2, column=1, padx=5, pady=10, sticky="w")

        # 目标位置
        ttk.Label(frame, text="点击位置:").grid(row=3, column=0, padx=5, pady=10, sticky="e")
        x_var = tk.StringVar(value=str(action.get('x', '')))
        y_var = tk.StringVar(value=str(action.get('y', '')))

        pos_frame = ttk.Frame(frame)
        pos_frame.grid(row=3, column=1, padx=5, pady=10, sticky="w")

        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        x_entry = ttk.Entry(pos_frame, width=8, textvariable=x_var)
        x_entry.pack(side=tk.LEFT, padx=2)

        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        y_entry = ttk.Entry(pos_frame, width=8, textvariable=y_var)
        y_entry.pack(side=tk.LEFT, padx=2)

        # 录制按钮
        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        # 相对延时
        ttk.Label(frame, text="相对延时(秒):").grid(row=4, column=0, padx=5, pady=10, sticky="e")
        delay_var = tk.DoubleVar(value=action.get('time', 0))
        ttk.Entry(frame, width=15, textvariable=delay_var).grid(row=4, column=1, padx=5, pady=10, sticky="w")

        # 清空文本选项
        clear_text_var = tk.BooleanVar(value=action.get('clear_text', True))
        ttk.Checkbutton(frame, text="输入前清空文本框", variable=clear_text_var).grid(row=5, column=0, columnspan=2, pady=5, sticky="w")

        # 状态标签
        status_label = ttk.Label(frame, text="", foreground="blue")
        status_label.grid(row=6, column=0, columnspan=2, pady=5)

        # 录制函数
        def start_record_position():
            dialog.withdraw()
            dialog.update()
            time.sleep(0.5)
            status_label.config(text="请在2秒后点击目标位置...")
            dialog.update()
            time.sleep(2)
            status_label.config(text="请现在点击目标位置...")
            dialog.update()

            def on_click(_x, _y, _button, pressed):
                if pressed:
                    return False

            listener = mouse.Listener(on_click=on_click)
            listener.start()
            listener.join()

            try:
                import win32api
                x, y = win32api.GetCursorPos()
                x_var.set(str(x))
                y_var.set(str(y))
                status_label.config(text=f"已录制: ({x}, {y})", foreground="green")
            except Exception:
                status_label.config(text="录制失败，请手动输入", foreground="red")

            dialog.deiconify()
            dialog.update()

        record_btn.config(command=start_record_position)

        def confirm():
            try:
                input_text = input_text_var.get()
                if not input_text:
                    raise ValueError("输入文本不能为空")

                # 解析变量初始值
                import re
                match = re.search(r'\{(\d+)\}', input_text)
                if match:
                    start_value = int(match.group(1))
                else:
                    start_value = 0

                step = step_var.get()
                if step < 0:
                    raise ValueError("步距不能为负数")

                target_x = x_var.get().strip()
                target_y = y_var.get().strip()

                if not target_x or not target_y:
                    raise ValueError("请输入有效的坐标")

                target_x = int(target_x)
                target_y = int(target_y)

                # 更新动作
                action['value'] = input_text
                action['start'] = start_value
                action['step'] = step
                action['x'] = target_x
                action['y'] = target_y
                action['clear_text'] = clear_text_var.get()
                action['time'] = delay_var.get()

                self._update_action_list()
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=7, column=0, columnspan=2, pady=15)

    def _edit_nested_loop_group(self, action, row_id):
        """编辑嵌套的循环组"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑循环组")
        dialog.geometry("550x320")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 循环次数
        ttk.Label(frame, text="循环次数:").grid(row=0, column=0, padx=5, pady=10, sticky="e")
        loop_count_var = tk.IntVar(value=action.get('loop_count', 1))
        ttk.Entry(frame, width=15, textvariable=loop_count_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")

        # 循环间隔类型选择
        interval_type_var = tk.StringVar(value=action.get('loop_interval_type', 'fixed'))
        interval_params = action.get('loop_interval', {})

        ttk.Label(frame, text="间隔模式:").grid(row=1, column=0, padx=5, pady=10, sticky="e")
        interval_type_frame = ttk.Frame(frame)
        interval_type_frame.grid(row=1, column=1, padx=5, pady=10, sticky="w")

        ttk.Radiobutton(interval_type_frame, text="固定", variable=interval_type_var, value="fixed",
                        command=lambda: update_interval_ui("fixed")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(interval_type_frame, text="范围", variable=interval_type_var, value="range",
                        command=lambda: update_interval_ui("range")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(interval_type_frame, text="列表", variable=interval_type_var, value="list",
                        command=lambda: update_interval_ui("list")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(interval_type_frame, text="随机", variable=interval_type_var, value="random",
                        command=lambda: update_interval_ui("random")).pack(side=tk.LEFT, padx=2)

        # 间隔参数框架
        interval_frame = ttk.Frame(frame)
        interval_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        # 固定间隔
        fixed_frame = ttk.Frame(interval_frame)
        ttk.Label(fixed_frame, text="间隔(秒):").pack(side=tk.LEFT, padx=2)
        loop_interval_var = tk.DoubleVar(value=interval_params.get('value', 0.5))
        ttk.Entry(fixed_frame, width=10, textvariable=loop_interval_var).pack(side=tk.LEFT, padx=2)

        # 范围间隔
        range_frame = ttk.Frame(interval_frame)
        ttk.Label(range_frame, text="最小:").pack(side=tk.LEFT, padx=2)
        range_min_var = tk.DoubleVar(value=interval_params.get('min', 0.5))
        ttk.Entry(range_frame, width=8, textvariable=range_min_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(range_frame, text="最大:").pack(side=tk.LEFT, padx=2)
        range_max_var = tk.DoubleVar(value=interval_params.get('max', 2.0))
        ttk.Entry(range_frame, width=8, textvariable=range_max_var).pack(side=tk.LEFT, padx=2)

        # 列表间隔
        list_frame = ttk.Frame(interval_frame)
        ttk.Label(list_frame, text="列表(逗号分隔):").pack(side=tk.LEFT, padx=2)
        list_values = interval_params.get('values', [0.5, 1.0, 1.5, 2.0])
        list_values_var = tk.StringVar(value=','.join(str(v) for v in list_values))
        ttk.Entry(list_frame, width=20, textvariable=list_values_var).pack(side=tk.LEFT, padx=2)

        # 随机间隔
        random_frame = ttk.Frame(interval_frame)
        ttk.Label(random_frame, text="最小:").pack(side=tk.LEFT, padx=2)
        random_min_var = tk.DoubleVar(value=interval_params.get('min', 0.5))
        ttk.Entry(random_frame, width=8, textvariable=random_min_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(random_frame, text="最大:").pack(side=tk.LEFT, padx=2)
        random_max_var = tk.DoubleVar(value=interval_params.get('max', 2.0))
        ttk.Entry(random_frame, width=8, textvariable=random_max_var).pack(side=tk.LEFT, padx=2)

        def update_interval_ui(interval_type):
            fixed_frame.grid_remove()
            range_frame.grid_remove()
            list_frame.grid_remove()
            random_frame.grid_remove()
            if interval_type == "fixed":
                fixed_frame.grid()
            elif interval_type == "range":
                range_frame.grid()
            elif interval_type == "list":
                list_frame.grid()
            else:  # random
                random_frame.grid()

        # 初始化UI状态
        update_interval_ui(interval_type_var.get())

        def confirm():
            try:
                loop_count = loop_count_var.get()
                if loop_count < 1:
                    raise ValueError("循环次数至少为1")

                interval_type = interval_type_var.get()
                if interval_type == "fixed":
                    loop_interval = loop_interval_var.get()
                    if loop_interval < 0:
                        raise ValueError("间隔不能为负数")
                    interval_params = {'type': 'fixed', 'value': loop_interval}
                elif interval_type == "range":
                    min_val = range_min_var.get()
                    max_val = range_max_var.get()
                    if min_val < 0 or max_val < min_val:
                        raise ValueError("无效的范围")
                    interval_params = {'type': 'range', 'min': min_val, 'max': max_val}
                elif interval_type == "list":
                    values_str = list_values_var.get().strip()
                    if not values_str:
                        raise ValueError("列表不能为空")
                    try:
                        values = [float(v.strip()) for v in values_str.split(',')]
                    except ValueError:
                        raise ValueError("列表格式无效，请使用逗号分隔的数字")
                    interval_params = {'type': 'list', 'values': values}
                else:  # random
                    min_val = random_min_var.get()
                    max_val = random_max_var.get()
                    if min_val < 0 or max_val < min_val:
                        raise ValueError("无效的范围")
                    interval_params = {'type': 'random', 'min': min_val, 'max': max_val}

                action['loop_count'] = loop_count
                action['loop_interval_type'] = interval_type
                action['loop_interval'] = interval_params

                self._update_action_list()
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=3, column=0, columnspan=2, pady=10)

    def _edit_nested_action_time(self, action, row_id):
        """编辑嵌套动作的时间（延时等）"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("修改延时")
        dialog.geometry("500x250")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        action_type = action.get('type', '')

        if action_type == 'random_delay':
            ttk.Label(frame, text="最小延时(秒):").grid(row=0, column=0, padx=5, pady=10, sticky="e")
            min_var = tk.DoubleVar(value=action.get('min_delay', 0.5))
            ttk.Entry(frame, width=15, textvariable=min_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")

            ttk.Label(frame, text="最大延时(秒):").grid(row=1, column=0, padx=5, pady=10, sticky="e")
            max_var = tk.DoubleVar(value=action.get('max_delay', 1.0))
            ttk.Entry(frame, width=15, textvariable=max_var).grid(row=1, column=1, padx=5, pady=10, sticky="w")

            def confirm():
                try:
                    min_delay = min_var.get()
                    max_delay = max_var.get()
                    if min_delay < 0 or max_delay < min_delay:
                        raise ValueError("无效的延时范围")
                    action['min_delay'] = min_delay
                    action['max_delay'] = max_delay
                    self._update_action_list()
                    self.in_dialog_operation = False
                    dialog.destroy()
                except Exception as e:
                    messagebox.showerror("错误", f"输入无效: {str(e)}")
        elif action_type == 'multiply_delay':
            ttk.Label(frame, text="基数延时(秒):").grid(row=0, column=0, padx=5, pady=10, sticky="e")
            base_var = tk.DoubleVar(value=action.get('base_delay', 1.0))
            ttk.Entry(frame, width=15, textvariable=base_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")

            def confirm():
                try:
                    base_delay = base_var.get()
                    if base_delay < 0:
                        raise ValueError("基数不能为负数")
                    action['base_delay'] = base_delay
                    self._update_action_list()
                    self.in_dialog_operation = False
                    dialog.destroy()
                except Exception as e:
                    messagebox.showerror("错误", f"输入无效: {str(e)}")
        elif action_type == 'arithmetic_delay':
            ttk.Label(frame, text="起始延时(秒):").grid(row=0, column=0, padx=5, pady=10, sticky="e")
            start_var = tk.DoubleVar(value=action.get('start_delay', 0.5))
            ttk.Entry(frame, width=15, textvariable=start_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")

            ttk.Label(frame, text="步距(秒):").grid(row=1, column=0, padx=5, pady=10, sticky="e")
            step_var = tk.DoubleVar(value=action.get('step_delay', 0.1))
            ttk.Entry(frame, width=15, textvariable=step_var).grid(row=1, column=1, padx=5, pady=10, sticky="w")

            def confirm():
                try:
                    start_delay = start_var.get()
                    step_delay = step_var.get()
                    if start_delay < 0 or step_delay < 0:
                        raise ValueError("延时和步距不能为负数")
                    action['start_delay'] = start_delay
                    action['step_delay'] = step_delay
                    self._update_action_list()
                    self.in_dialog_operation = False
                    dialog.destroy()
                except Exception as e:
                    messagebox.showerror("错误", f"输入无效: {str(e)}")
        else:
            ttk.Label(frame, text="相对延时(秒):").grid(row=0, column=0, padx=5, pady=10, sticky="e")
            time_var = tk.DoubleVar(value=action.get('time', 0))
            ttk.Entry(frame, width=15, textvariable=time_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")

            def confirm():
                try:
                    new_time = time_var.get()
                    if new_time < 0:
                        raise ValueError("时间不能为负数")
                    action['time'] = new_time
                    self._update_action_list()
                    self.in_dialog_operation = False
                    dialog.destroy()
                except Exception as e:
                    messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=2, column=0, columnspan=2, pady=10)

    def edit_action_time(self, event):
        """双击动作列表项以修改等待时间（相对于上一个动作或截屏延时）"""
        self.in_dialog_operation = True
        try:
            # 获取选中的动作
            row_id = self.action_tree.selection()[0] if self.action_tree.selection() else None
            if not row_id:
                self.in_dialog_operation = False
                return

            self._hide_delay_spinbox()

            # 判断是嵌套动作还是顶层动作
            is_nested = row_id.startswith('loop_')
            if is_nested:
                action = self._get_action_by_row_id(row_id)
                if not action:
                    self.in_dialog_operation = False
                    return
                # 嵌套动作使用不同的编辑方式
                if action.get('type') == 'loop_group':
                    self._edit_nested_loop_group(action, row_id)
                    return
                elif action.get('type') in ['click', 'doubleclick', 'move']:
                    self._edit_nested_click_action(action, row_id)
                    return
                elif action.get('type') == 'variable_input':
                    self._edit_nested_variable_input_action(action, row_id)
                    return
                else:
                    # 嵌套的其他动作类型
                    self._edit_nested_action_time(action, row_id)
                    return
            else:
                # 顶层动作
                selection = self._get_selected_indices()
                if not selection:
                    self.in_dialog_operation = False
                    return
                index = selection[0]
                action = self.actions[index]

            # 检查动作是否支持时间修改
            if action['type'] not in ['move', 'click', 'doubleclick', 'random_delay', 'multiply_delay', 'arithmetic_delay', 'screenshot', 'loop_group', 'numeric_loop', 'variable_input']:
                messagebox.showinfo("提示", "该动作不支持修改时间")
                self.in_dialog_operation = False
                return

            # 弹出对话框修改时间
            dialog = tk.Toplevel(self.root)
            dialog.title("修改延时")
            dialog.geometry("500x250")
            dialog.transient(self.root)

            def on_dialog_close():
                self.in_dialog_operation = False
                dialog.destroy()

            dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
            dialog.grab_set()

            frame = ttk.Frame(dialog, padding="10")
            frame.pack(fill=tk.BOTH, expand=True)

            # 判断当前动作类型
            action_type = action['type']

            # 延时动作使用完整的延时类型选择界面
            if action_type in ['random_delay', 'multiply_delay', 'arithmetic_delay']:
                # 延时类型选择
                type_var = tk.StringVar(value=action_type)

                # 类型选择
                ttk.Label(frame, text="延时类型:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
                type_frame = ttk.Frame(frame)
                type_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
                ttk.Radiobutton(type_frame, text="随机范围", variable=type_var, value="random",
                               command=lambda: update_ui("random")).pack(side=tk.LEFT, padx=2)
                ttk.Radiobutton(type_frame, text="倍数增长", variable=type_var, value="multiply",
                               command=lambda: update_ui("multiply")).pack(side=tk.LEFT, padx=2)
                ttk.Radiobutton(type_frame, text="等差递增", variable=type_var, value="arithmetic",
                               command=lambda: update_ui("arithmetic")).pack(side=tk.LEFT, padx=2)

                # 随机延时参数
                random_frame = ttk.Frame(frame)
                random_frame.grid(row=1, column=0, columnspan=2, sticky="ew")

                ttk.Label(random_frame, text="最小延时(秒):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
                min_var = tk.DoubleVar(value=action.get('min_delay', 0.5))
                min_entry = ttk.Entry(random_frame, width=10, textvariable=min_var)
                min_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

                ttk.Label(random_frame, text="最大延时(秒):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
                max_var = tk.DoubleVar(value=action.get('max_delay', 1.0))
                max_entry = ttk.Entry(random_frame, width=10, textvariable=max_var)
                max_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

                # 倍数延时参数
                exp_frame = ttk.Frame(frame)
                exp_frame.grid(row=1, column=0, columnspan=2, sticky="ew")

                ttk.Label(exp_frame, text="基数延时(秒):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
                base_var = tk.DoubleVar(value=action.get('base_delay', 1.0))
                base_entry = ttk.Entry(exp_frame, width=10, textvariable=base_var)
                base_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
                ttk.Label(exp_frame, text="(每轮循环增加一倍)").grid(row=0, column=2, padx=5, pady=5, sticky="w")

                # 等差递增参数
                arithmetic_frame = ttk.Frame(frame)
                arithmetic_frame.grid(row=1, column=0, columnspan=2, sticky="ew")

                ttk.Label(arithmetic_frame, text="起始延时(秒):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
                arithmetic_start_var = tk.DoubleVar(value=action.get('start_delay', 0.5))
                ttk.Entry(arithmetic_frame, width=10, textvariable=arithmetic_start_var).grid(row=0, column=1, padx=5, pady=5, sticky="w")

                ttk.Label(arithmetic_frame, text="步距(秒):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
                arithmetic_step_var = tk.DoubleVar(value=action.get('step_delay', 0.1))
                ttk.Entry(arithmetic_frame, width=10, textvariable=arithmetic_step_var).grid(row=1, column=1, padx=5, pady=5, sticky="w")

                def update_ui(delay_type):
                    if delay_type == "random":
                        random_frame.grid()
                        exp_frame.grid_remove()
                        arithmetic_frame.grid_remove()
                    elif delay_type == "multiply":
                        random_frame.grid_remove()
                        exp_frame.grid()
                        arithmetic_frame.grid_remove()
                    else:  # arithmetic
                        random_frame.grid_remove()
                        exp_frame.grid_remove()
                        arithmetic_frame.grid()

                # 初始化UI状态
                update_ui(action_type)

                def confirm():
                    try:
                        delay_type = type_var.get()

                        if delay_type == "random":
                            min_delay = float(min_var.get())
                            max_delay = float(max_var.get())
                            if min_delay < 0 or max_delay < min_delay:
                                raise ValueError("无效的延时范围")
                            action['type'] = 'random_delay'
                            action['min_delay'] = min_delay
                            action['max_delay'] = max_delay
                        elif delay_type == "multiply":
                            base_delay = float(base_var.get())
                            if base_delay < 0:
                                raise ValueError("基数不能为负数")
                            action['type'] = 'multiply_delay'
                            action['base_delay'] = base_delay
                        else:  # arithmetic
                            start_delay = float(arithmetic_start_var.get())
                            step_delay = float(arithmetic_step_var.get())
                            if start_delay < 0 or step_delay < 0:
                                raise ValueError("延时和步距不能为负数")
                            action['type'] = 'arithmetic_delay'
                            action['start_delay'] = start_delay
                            action['step_delay'] = step_delay

                        self._update_action_list(select_index=index)
                        self.in_dialog_operation = False
                        dialog.destroy()
                    except Exception as e:
                        messagebox.showerror("错误", f"输入无效: {str(e)}")

                ttk.Button(frame, text="确定", command=confirm).grid(row=2, column=0, columnspan=2, pady=10)
            else:
                # 其他动作类型（click, move, doubleclick, screenshot, loop_group, numeric_loop, variable_input）
                if action_type == 'screenshot':
                    ttk.Label(frame, text="截屏延时(秒):").grid(row=0, column=0, padx=5, pady=10, sticky="e")
                    time_var = tk.DoubleVar(value=action.get('delay', 0))
                    ttk.Entry(frame, width=15, textvariable=time_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")
                elif action_type == 'loop_group':
                    # 获取间隔参数
                    interval_type = action.get('loop_interval_type', 'fixed')
                    interval_params = action.get('loop_interval', {})
                    if interval_type == 'fixed':
                        default_val = interval_params.get('value', 0)
                    elif interval_type in ('range', 'random'):
                        default_val = interval_params.get('min', 0)
                    else:  # list
                        values = interval_params.get('values', [0])
                        default_val = values[0] if values else 0
                    ttk.Label(frame, text="循环间隔(秒):").grid(row=0, column=0, padx=5, pady=10, sticky="e")
                    time_var = tk.DoubleVar(value=default_val)
                    ttk.Entry(frame, width=15, textvariable=time_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")
                elif action_type == 'numeric_loop':
                    ttk.Label(frame, text="循环间隔(秒):").grid(row=0, column=0, padx=5, pady=10, sticky="e")
                    time_var = tk.DoubleVar(value=action.get('interval', 0.5))
                    ttk.Entry(frame, width=15, textvariable=time_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")
                elif action_type == 'variable_input':
                    ttk.Label(frame, text="相对延时(秒):").grid(row=0, column=0, padx=5, pady=10, sticky="e")
                    prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
                    relative_time = action.get('time', 0) - prev_time
                    time_var = tk.DoubleVar(value=relative_time)
                    ttk.Entry(frame, width=15, textvariable=time_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")
                else:
                    ttk.Label(frame, text="相对延时(秒):").grid(row=0, column=0, padx=5, pady=10, sticky="e")
                    prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
                    relative_time = action.get('time', 0) - prev_time
                    time_var = tk.DoubleVar(value=relative_time)
                    ttk.Entry(frame, width=15, textvariable=time_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")

                def confirm():
                    try:
                        new_value = float(time_var.get())
                        if new_value < 0:
                            raise ValueError("时间不能为负数")

                        if action_type == 'screenshot':
                            action['delay'] = new_value
                        elif action_type == 'loop_group':
                            interval_type = action.get('loop_interval_type', 'fixed')
                            interval_params = action.get('loop_interval', {})
                            if interval_type == 'fixed':
                                interval_params['value'] = new_value
                            elif interval_type in ('range', 'random'):
                                interval_params['min'] = new_value
                            elif interval_type == 'list':
                                values = interval_params.get('values', [0])
                                if values:
                                    values[0] = new_value
                                    interval_params['values'] = values
                            action['loop_interval'] = interval_params
                        elif action_type == 'numeric_loop':
                            action['interval'] = new_value
                        elif action_type == 'variable_input':
                            prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
                            new_absolute_time = prev_time + new_value
                            old_time = action.get('time', 0)
                            delta = new_absolute_time - old_time
                            action['time'] = new_absolute_time
                            self._shift_action_times(index + 1, delta)
                        else:
                            # click, move, doubleclick
                            prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
                            new_absolute_time = prev_time + new_value
                            old_time = action.get('time', 0)
                            delta = new_absolute_time - old_time
                            action['time'] = new_absolute_time
                            self._shift_action_times(index + 1, delta)

                        self._update_action_list(select_index=index)
                        self.in_dialog_operation = False
                        dialog.destroy()
                    except Exception as e:
                        messagebox.showerror("错误", f"输入无效: {str(e)}")

                ttk.Button(frame, text="确定", command=confirm).grid(row=1, column=0, columnspan=2, pady=10)

        except Exception as e:
            self.in_dialog_operation = False
            messagebox.showerror("错误", f"修改时间失败: {str(e)}")

    def add_screenshot_action(self):
        """添加截屏动作到动作列表，支持时间戳和序号递增选择"""
        action = {
            'type': 'screenshot',
            'naming': self.screenshot_naming_var.get(),
            'directory': self.screenshot_dir_var.get(),
            'filename': self.screenshot_name_var.get(),
            'delay': self.screenshot_delay_var.get()
        }
        self.actions.append(action)
        self._update_action_list(scroll_to_end=True)

    def take_screenshot(self, delay=0):
        """截屏并保存到指定目录，支持延时和命名方式选择"""
        time.sleep(delay)  # 添加延时
        directory = self.screenshot_dir_var.get()
        if not os.path.exists(directory):
            os.makedirs(directory)

        if self.screenshot_naming_var.get() == "timestamp":
            filename = f"{self.screenshot_name_var.get()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        else:
            existing_files = [f for f in os.listdir(directory) if f.startswith(self.screenshot_name_var.get()) and f.endswith('.png')]
            next_index = len(existing_files) + 1
            filename = f"{self.screenshot_name_var.get()}_{next_index}.png"

        filepath = os.path.join(directory, filename)
        try:
            screenshot = ImageGrab.grab()
            screenshot.save(filepath)
            self.update_playback_info(f"截屏已保存: {filepath}")
        except Exception as e:
            self.update_playback_info(f"截屏失败: {e}")

    def on_closing(self):
        """处理窗口关闭事件"""
        # 停止录制和播放
        if self.is_recording:
            self.stop_recording()
        if self.is_playing:
            self.stop_playback()

        # 清理快捷键
        try:
            kb.remove_hotkey('ctrl+q')
            kb.unhook_all_hotkeys()
        except Exception:
            pass

        # 清理监听器
        if self.mouse_listener:
            try:
                self.mouse_listener.stop()
            except Exception:
                pass
        if self.keyboard_listener:
            try:
                self.keyboard_listener.stop()
            except Exception:
                pass

        self.root.destroy()
    
    def run(self):
        self.root.mainloop()

    def _hotkey_stop_playback(self):
        """Ctrl+Q 快捷键停止回放"""
        if self.is_playing:
            self.stop_playback()

    def _register_hotkeys(self):
        """注册全局快捷键"""
        try:
            kb.add_hotkey('ctrl+shift+r', self.toggle_recording)
            kb.add_hotkey('ctrl+shift+p', self.play_actions)
            kb.add_hotkey('ctrl+shift+s', self.stop_playback)
        except Exception as e:
            print(f"快捷键注册失败: {e}")

if __name__ == "__main__":
    app = AutoTestGUI()
    app.run()
