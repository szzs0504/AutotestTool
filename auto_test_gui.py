import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyautogui
import keyboard as kb  # 修改为 kb，避免与pynput.keyboard冲突
import json
import time
from threading import Thread, Lock, Event
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
        self.root.geometry("900x600")
        
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
        self.test_index_lock = Lock()  # 保护test_index的锁
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
        self._relative_delay_cache = {}  # 缓存相对延时计算结果
        self._relative_delay_valid = False  # 缓存是否有效
        self._listener_stop_scheduled = False  # 标记是否已计划停止监听器
        self.screenshot_dir_var = tk.StringVar(value=screenshot_dir)
        self.screenshot_name_var = tk.StringVar(value="screenshot")
        self.screenshot_delay_var = tk.DoubleVar(value=0.0)
        self.screenshot_naming_var = tk.StringVar(value="timestamp")
        self._keyboard_recording_paused = False  # 标记键盘录制是否暂停
        self.status_var = tk.StringVar(value="就绪")

        # 快捷键说明文本
        self.shortcut_help = (
            "快捷键: Ctrl+Shift+R 开始/停止录制 | "
            "Ctrl+Shift+P 开始回放 | "
            "Ctrl+Shift+S 停止回放 | "
            "Ctrl+Q 强制停止回放"
        )

        self.setup_ui()

    def _record_position_with_dialog(self, dialog, x_var, y_var, status_label):
        """录制位置 - 共享方法，支持不同对话框

        Args:
            dialog: 父对话框（会隐藏和恢复）
            x_var: X坐标的StringVar
            y_var: Y坐标的StringVar
            status_label: 状态标签Label

        Returns:
            bool: 是否成功录制
        """
        dialog.withdraw()
        dialog.update()

        captured_pos = [None, None]

        def on_click(x, y, _button, pressed):
            if pressed:
                captured_pos[0] = x
                captured_pos[1] = y
                return False  # 停止监听

        listener = mouse.Listener(on_click=on_click)
        listener.start()
        listener.join()

        if captured_pos[0] is not None:
            x_var.set(str(captured_pos[0]))
            y_var.set(str(captured_pos[1]))
            status_label.config(text=f"已录制: ({captured_pos[0]}, {captured_pos[1]})", foreground="green")
        else:
            status_label.config(text="录制失败，请重试", foreground="red")

        dialog.deiconify()
        return True

    def setup_ui(self):
        # 设置窗口置顶（默认开启）
        self.root.attributes('-topmost', self.window_topmost.get())

        # 绑定焦点事件：当焦点在程序内时不录制键盘输入
        self._last_focus_owner = None
        self.root.bind('<FocusIn>', self._on_app_focus_in)
        self.root.bind('<FocusOut>', self._on_app_focus_out)

        # 创建菜单栏
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="保存动作", command=self.save_actions)
        file_menu.add_command(label="加载动作", command=self.load_actions)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)

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
        self.loop_count = tk.IntVar(value=1)  # 默认循环1次
        self.loop_spinbox = ttk.Spinbox(loop_frame, from_=1, to=9999, width=4, textvariable=self.loop_count)
        self.loop_spinbox.pack(side=tk.LEFT, padx=2)
        self._update_loop_count_max()

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
        self.action_tree.column("delay", anchor="center", width=85, stretch=False)
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
        # 拖动视觉效果相关
        self._drag_indicator = None  # 拖动指示器（抓手当蓝色效果）
        self._drag_placeholder = None  # 拖动占位符
        self._drag_over_item = None   # 当前悬停在哪一行
        # 定义拖动抓手当宽度
        self._GRAB_HANDLE_WIDTH = 50
        self._GRAB_HANDLE_ICON = "☰"  # 抓手当符号（手型）
        self.action_tree.tag_configure('drag_handle', background='#0078D7')  # 蓝色抓手当效果
        self.action_tree.tag_configure('drag_over', background='#E5F3FF')    # 悬停时蓝色背景
        self.action_tree.tag_configure('placeholder', background='#F0F0F0')  # 占位符灰色
        self.action_tree.tag_configure('grab_handle', font='')  # 抓手当标签
        self.action_tree.bind("<Button-1>", self._on_drag_start, add=True)
        self.action_tree.bind("<B1-Motion>", self._on_drag_motion, add=True)
        self.action_tree.bind("<ButtonRelease-1>", self._on_drag_end, add=True)
        self.action_tree.bind("<Motion>", self._on_tree_motion)  # 鼠标移动事件用于改变光标

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

        ttk.Button(left_buttons, text="插入点击", command=self.insert_click_action).pack(side=tk.LEFT, padx=1)
        ttk.Button(left_buttons, text="删除动作", command=self.delete_selected).pack(side=tk.LEFT, padx=1)
        ttk.Button(left_buttons, text="插入延时", command=self.insert_random_delay).pack(side=tk.LEFT, padx=1)
        ttk.Button(left_buttons, text="添加截屏", command=self.add_screenshot_action).pack(side=tk.LEFT, padx=1)
        ttk.Button(left_buttons, text="输入", command=self.insert_traverse_input).pack(side=tk.LEFT, padx=1)
        ttk.Button(left_buttons, text="设为循环组", command=self.insert_loop_group).pack(side=tk.LEFT, padx=1)

        # 右侧按钮组
        right_buttons = ttk.Frame(bottom_frame)
        right_buttons.pack(side=tk.RIGHT, padx=5)

        # 添加状态栏
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=2, column=0, sticky="ew", padx=5, pady=2)

        # 快捷键提示栏
        self.shortcut_label = ttk.Label(self.root, text=self.shortcut_help, anchor="w", relief=tk.GROOVE)
        self.shortcut_label.grid(row=3, column=0, sticky="ew", padx=5, pady=2)

        
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

    def _on_app_focus_in(self, event):
        """当应用获得焦点时，暂停键盘录制"""
        self._keyboard_recording_paused = True
        self.pressed_keys.clear()  # 清除按下的键，避免恢复时误录

    def _on_app_focus_out(self, event):
        """当应用失去焦点时，恢复键盘录制"""
        self._keyboard_recording_paused = False

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

    def is_in_any_dialog(self, x, y):
        """检查坐标是否在任何对话框内"""
        try:
            # 遍历所有Toplevel窗口
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Toplevel):
                    if widget.winfo_viewable():  # 如果窗口可见
                        dialog_x = widget.winfo_rootx()
                        dialog_y = widget.winfo_rooty()
                        dialog_width = widget.winfo_width()
                        dialog_height = widget.winfo_height()

                        if (dialog_x <= x <= dialog_x + dialog_width and
                            dialog_y <= y <= dialog_y + dialog_height):
                            return True
            return False
        except Exception:
            return False

    def on_click(self, x, y, button, pressed):
        if not pressed:  # 只在释放时记录
            # 如果正在对话框操作中，不记录
            if self.in_dialog_operation:
                return

            # 检查是否在工具窗口内或任何对话框内（不记录）
            if self.is_in_window(x, y) or self.is_in_any_dialog(x, y):
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
                self._invalidate_relative_delay_cache()
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
                self._invalidate_relative_delay_cache()
            
            self.last_click_time = current_time
            self.last_click_position = (x, y)

    def on_scroll(self, x, y, dx, dy):
        if self.is_recording and not self.in_dialog_operation:
            # 检查是否在工具窗口内或任何对话框内（不记录）
            if self.is_in_window(x, y) or self.is_in_any_dialog(x, y):
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
                self._invalidate_relative_delay_cache()
                self.last_scroll_time = current_time

    def on_key_down(self, key):
        try:
            if self.is_recording and not self.in_dialog_operation:
                # 如果延时spinbox正在编辑，不记录键盘事件
                if self.active_spinbox_item is not None:
                    return
                key_str = self.convert_key_name(key)
                self.pressed_keys.add(key_str)
        except AttributeError:
            pass

    def on_key_up(self, key):
        try:
            if self.is_recording and not self.in_dialog_operation:
                # 如果延时spinbox正在编辑，不记录键盘事件
                if self.active_spinbox_item is not None:
                    return
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
                        self._invalidate_relative_delay_cache()
                    else:
                        # 单个键
                        self.actions.append({
                            'type': 'keyboard',
                            'keys': [key_str],
                            'time': current_time
                        })
                        self.update_action_list_display()
                        self._invalidate_relative_delay_cache()
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

        # 如果从非主线程调用，使用root.after在主线程执行停止操作
        if not self._listener_stop_scheduled:
            self._listener_stop_scheduled = True
            self.root.after(0, self._stop_listeners_internal)
            self.root.after(0, self._finalize_recording_ui)

    def _stop_listeners_internal(self):
        """在主线程中安全地停止监听器"""
        try:
            if self.mouse_listener:
                self.mouse_listener.stop()
        except Exception as e:
            print(f"停止鼠标监听器失败: {e}")
        try:
            if self.keyboard_listener:
                self.keyboard_listener.stop()
        except Exception as e:
            print(f"停止键盘监听器失败: {e}")
        self._listener_stop_scheduled = False

    def _finalize_recording_ui(self):
        """更新UI状态（在主线程执行）"""
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
        with self.test_index_lock:
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

    def _update_loop_count_max(self):
        """更新主界面循环次数的最大值，如果动作列表中有列表遍历类型的遍历输入"""
        max_count = 9999
        for action in self.actions:
            # 只检查顶层动作，不检查循环组内的嵌套动作
            if action.get('type') == 'traverse_input' and action.get('traverse_type') == 'list':
                values = action.get('traverse_values', [])
                if values:
                    max_count = len(values)
                    break
        self.loop_spinbox.configure(to=max_count)
        if self.loop_count.get() > max_count:
            self.loop_count.set(max_count)

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

    def _execute_single_action(self, action, loop_index=0):
        """执行单个动作（用于循环组）

        Args:
            action: 动作字典
            loop_index: 当前循环组的循环索引（从0开始），用于遍历输入等动作
        """
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
            self.take_screenshot(
                delay=action.get('delay', 0),
                directory=action.get('directory'),
                filename=action.get('filename'),
                naming=action.get('naming')
            )
        elif action_type == 'random_delay':
            delay = random.uniform(action.get('min_delay', 0), action.get('max_delay', 0))
            time.sleep(delay)
        elif action_type == 'multiply_delay':
            # 倍数延时：使用基数 * (loop_index + 1)
            base_delay = action.get('base_delay', 1.0)
            delay = base_delay * (loop_index + 1)
            time.sleep(delay)
        elif action_type == 'arithmetic_delay':
            # 等差延时：这里在循环组内使用起始延时
            delay = action.get('start_delay', 0.5)
            time.sleep(delay)
        elif action_type == 'delay':
            # 固定延时
            delay = action.get('delay', 0)
            if delay > 0:
                time.sleep(delay)
        elif action_type == 'loop_group':
            # 嵌套循环组：执行内部动作
            nested_loop_count = action.get('loop_count', 1)
            nested_loop_actions = action.get('loop_actions', [])
            interval_type = action.get('loop_interval_type', 'fixed')
            interval_params = action.get('loop_interval', {})

            if not nested_loop_actions:
                return

            for nested_loop_index in range(nested_loop_count):
                if not self.is_playing:
                    return
                # 执行嵌套循环动作
                for nested_action in nested_loop_actions:
                    if not self.is_playing:
                        return
                    self._execute_single_action(nested_action, nested_loop_index)

                # 嵌套循环间隔
                if nested_loop_index < nested_loop_count - 1:
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
                        self._safe_sleep_with_failsafe(wait_time)
        elif action_type == 'variable_input':
            # 变量输入
            target_x = action.get('x')
            target_y = action.get('y')
            value = action.get('value', '')
            clear_text = action.get('clear_text', True)

            # 处理相对延时
            action_time = action.get('time', 0)
            if action_time > 0:
                time.sleep(action_time)

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
        elif action_type == 'input':
            # 输入动作：点击坐标并输入文本
            target_x = action.get('x')
            target_y = action.get('y')
            input_text = action.get('text', '')

            if target_x is not None and target_y is not None:
                pyautogui.moveTo(target_x, target_y, duration=0.1)
                pyautogui.click(x=target_x, y=target_y)
                time.sleep(0.1)

            # 清空输入框
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.05)
            pyautogui.press('delete')

            if input_text:
                pyautogui.write(input_text, interval=0.05)

        elif action_type == 'traverse_input':
            # 遍历输入：根据循环索引使用不同的值
            traverse_type = action.get('traverse_type', 'list')
            target_x = action.get('x')
            target_y = action.get('y')

            # 计算当前值
            if traverse_type == 'sequence':
                start = action.get('start', 0)
                step = action.get('step', 1)
                value = start + step * loop_index
                template = action.get('template', 'text{n}')
                text_to_input = template.replace('{n}', str(value))
            elif traverse_type == 'fixed_text':
                text_to_input = action.get('fixed_text', '')
            else:  # list
                traverse_values = action.get('traverse_values', [])
                if traverse_values:
                    # 如果索引超出列表长度，使用最后一个值
                    if loop_index < len(traverse_values):
                        text_to_input = traverse_values[loop_index]
                    else:
                        text_to_input = traverse_values[-1]
                else:
                    text_to_input = ''

            # 执行输入
            if target_x is not None and target_y is not None:
                pyautogui.moveTo(target_x, target_y, duration=0.1)
                pyautogui.click()
                time.sleep(0.05)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.05)
                pyautogui.press('backspace')

            if text_to_input:
                if traverse_type == 'sequence':
                    pyautogui.write(text_to_input, interval=0.05)
                else:
                    pyautogui.write(str(text_to_input), interval=0.05)

        elif action_type == 'input_loop':
            # 输入循环：执行输入并执行循环内动作
            loop_type = action.get('loop_type', 'fixed')
            template = action.get('input_template', '')
            target_x = action.get('target_x')
            target_y = action.get('target_y')
            clear_text = action.get('clear_text', True)
            interval_mode = action.get('interval_mode', 'fixed')
            loop_interval = action.get('loop_interval', 1.0)
            loop_actions = action.get('loop_actions', [])

            # 根据类型生成输入值列表
            if loop_type == 'fixed':
                count = action.get('loop_count', 3)
                values = list(range(count))
            elif loop_type == 'range':
                start = action.get('start', 0)
                end = action.get('end', 10)
                step = action.get('step', 1)
                values = list(range(start, end + 1, step))
            else:  # list
                values = action.get('text_list', [])

            for i, value in enumerate(values):
                if not self.is_playing:
                    return

                # 点击目标位置
                if target_x is not None and target_y is not None:
                    pyautogui.moveTo(target_x, target_y, duration=0.1)
                    pyautogui.click(x=target_x, y=target_y)
                    time.sleep(0.1)

                # 清空文本框
                if clear_text:
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(0.05)
                    pyautogui.press('backspace')
                    time.sleep(0.05)

                # 生成并输入文本
                if loop_type == 'list':
                    text = str(value)
                else:
                    text = template.replace('{n}', str(value))
                pyautogui.write(text, interval=0.05)

                # 执行循环内动作
                for act in loop_actions:
                    if not self.is_playing:
                        return
                    self._execute_single_action(act)

                # 计算实际间隔
                if i < len(values) - 1:
                    if interval_mode == 'fixed':
                        actual_interval = loop_interval if isinstance(loop_interval, (int, float)) else 1.0
                    elif interval_mode == 'range':
                        # 范围间隔：等差或等比
                        if isinstance(loop_interval, dict) and loop_interval.get('sub_type') == 'geometric':
                            # 等比：start * ratio^i
                            interval_geo_start = loop_interval.get('start', 1.0)
                            interval_ratio = loop_interval.get('ratio', 1.5)
                            interval_steps = loop_interval.get('steps', 10)
                            actual_interval = interval_geo_start * (interval_ratio ** i)
                            if i >= interval_steps:
                                actual_interval = interval_geo_start * (interval_ratio ** interval_steps)
                        else:
                            # 等差：start + i * step
                            interval_start = loop_interval.get('start', 1.0) if isinstance(loop_interval, dict) else 1.0
                            interval_step = loop_interval.get('step', 0.1) if isinstance(loop_interval, dict) else 0.1
                            interval_end = loop_interval.get('end', 2.0) if isinstance(loop_interval, dict) else 2.0
                            actual_interval = interval_start + i * interval_step
                            if actual_interval > interval_end:
                                actual_interval = interval_end
                    elif interval_mode == 'list':
                        vals = loop_interval.get('values', [1.0]) if isinstance(loop_interval, dict) else [1.0]
                        actual_interval = random.choice(vals)
                    else:  # random
                        min_val = loop_interval.get('min', 0.5) if isinstance(loop_interval, dict) else 0.5
                        max_val = loop_interval.get('max', 2.0) if isinstance(loop_interval, dict) else 2.0
                        actual_interval = random.uniform(min_val, max_val)
                    self._safe_sleep_with_failsafe(actual_interval)

    def play_recorded_actions(self, index):
        # 禁用 pyautogui 的故障保护
        pyautogui.FAILSAFE = False

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
                # 确保max不小于min
                if max_interval < min_interval:
                    max_interval = min_interval
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
                        self.play_btn.config(state='normal')
                        self.stop_btn.config(state='disabled')
                        return
                    with self.test_index_lock:
                        if self.test_index != index:
                            self.play_btn.config(state='normal')
                            self.stop_btn.config(state='disabled')
                            return
                    # 检查鼠标是否在屏幕左上角
                    current_position = pyautogui.position()
                    if current_position[0] <= 10 and current_position[1] <= 10:
                        pyautogui.FAILSAFE = True
                        self.is_playing = False
                        self.play_btn.config(state='normal')
                        self.stop_btn.config(state='disabled')
                        self.root.after(0, lambda: self.status_var.set("回放已中止（FAILSAFE触发）"))
                        return

                    # 计算需要等待的时间
                    if i > 0:
                        wait_time = action.get('time', 0) - last_action_time
                        if wait_time > 0:
                            if not self._safe_sleep_with_failsafe(wait_time, index):
                                # 被故障保护中断
                                self.play_btn.config(state='normal')
                                self.stop_btn.config(state='disabled')
                                return

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
                    elif action['type'] == 'input':
                        pyautogui.moveTo(action['x'], action['y'], duration=0.2)
                        pyautogui.click(x=action['x'], y=action['y'])
                        time.sleep(0.1)
                        # 清空输入框
                        pyautogui.hotkey('ctrl', 'a')
                        time.sleep(0.05)
                        pyautogui.press('delete')
                        # 输入文本
                        pyautogui.write(action.get('text', ''), interval=0.05)
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
                    elif action['type'] == 'delay':
                        # 固定延时
                        delay = action.get('delay', 0)
                        if delay > 0:
                            self.update_playback_info(f"等待固定延时: {delay:.1f}秒")
                            time.sleep(delay)
                    elif action['type'] == 'screenshot':
                        self.take_screenshot(
                            delay=action.get('delay', 0),
                            directory=action.get('directory'),
                            filename=action.get('filename'),
                            naming=action.get('naming')
                        )
                    elif action['type'] == 'loop_group':
                        # 循环组：执行选中动作N次
                        loop_count = action.get('loop_count', 1)
                        loop_actions = action.get('loop_actions', [])
                        interval_type = action.get('loop_interval_type', 'fixed')
                        interval_params = action.get('loop_interval', {})

                        if not loop_actions:
                            self.update_playback_info("循环组: 无循环动作")
                            continue

                        for loop_idx in range(loop_count):
                            if not self.is_playing:
                                return
                            # 执行循环动作
                            for loop_action in loop_actions:
                                if not self.is_playing:
                                    return
                                self._execute_single_action(loop_action, loop_idx)

                            # 循环间隔
                            if loop_idx < loop_count - 1:
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
                                    if not self._safe_sleep_with_failsafe(wait_time, index):
                                        # 被故障保护中断，停止播放
                                        self.play_btn.config(state='normal')
                                        self.stop_btn.config(state='disabled')
                                        return

                        self.update_playback_info(f"循环组完成: 共执行{loop_count}次")
                    elif action['type'] == 'input_loop':
                        # 输入循环：执行输入并执行循环内动作
                        self._execute_single_action(action)
                        self.update_playback_info(f"输入循环完成")
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
                    elif action['type'] == 'traverse_input':
                        # 遍历输入：根据循环轮次使用不同的值
                        traverse_type = action.get('traverse_type', 'list')
                        target_x = action.get('x')
                        target_y = action.get('y')

                        # 计算当前值
                        if traverse_type == 'sequence':
                            start = action.get('start', 0)
                            step = action.get('step', 1)
                            value = start + step * loop
                            template = action.get('template', 'text{n}')
                            text_to_input = template.replace('{n}', str(value))
                        elif traverse_type == 'fixed_text':
                            value = action.get('fixed_text', '')
                        else:  # list
                            traverse_values = action.get('traverse_values', [])
                            if traverse_values:
                                # 如果索引超出列表长度，使用最后一个值
                                if loop < len(traverse_values):
                                    value = traverse_values[loop]
                                else:
                                    value = traverse_values[-1]  # 使用最后一个值
                            else:
                                value = None

                        if value is not None:
                            if target_x is not None and target_y is not None:
                                pyautogui.moveTo(target_x, target_y, duration=0.1)
                                pyautogui.click()
                                time.sleep(0.05)
                                pyautogui.hotkey('ctrl', 'a')
                                time.sleep(0.05)
                                pyautogui.press('backspace')
                            if traverse_type == 'sequence':
                                pyautogui.write(text_to_input, interval=0.05)
                                self.update_playback_info(f"遍历输入: {text_to_input}")
                            else:
                                pyautogui.write(str(value), interval=0.05)
                                self.update_playback_info(f"遍历输入: {value}")
                        else:
                            self.update_playback_info(f"遍历输入: 无可用值(循环{loop+1}超过列表长度)")

                    last_action_time = action.get('time', 0)

                if loop < loops - 1:
                    if self.interval_random.get():
                        wait_time = random.uniform(min_interval, max_interval)
                        self.update_playback_info(f"等待 {wait_time:.1f} 秒后开始下一轮...")
                    else:
                        wait_time = interval
                        self.update_playback_info(f"等待 {interval} 秒后开始下一轮...")
                    if not self._safe_sleep_with_failsafe(wait_time, index):
                        # 被故障保护中断
                        self.play_btn.config(state='normal')
                        self.stop_btn.config(state='disabled')
                        return

            self.is_playing = False
            self.play_btn.config(state='normal')
            self.stop_btn.config(state='disabled')
            if self.status_var.get() != "回放已中止":
                self.status_var.set("执行完成")
            self.update_playback_info("播放完成")
            pyautogui.FAILSAFE = True  # 恢复故障保护

        except Exception as e:
            print(f"回放错误: {e}")
            self.root.after(0, lambda: self.status_var.set("回放出错"))
            pyautogui.FAILSAFE = True  # 恢复故障保护
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

    def _safe_sleep_with_failsafe(self, duration, index=None):
        """带故障保护检查的睡眠方法

        Args:
            duration: 需要睡眠的总时长（秒）
            index: 当前测试索引，用于检查是否过期（可选）

        Returns:
            True: 睡眠正常完成
            False: 被故障保护中断
        """
        check_interval = 0.1  # 每0.1秒检查一次
        elapsed = 0.0
        while elapsed < duration:
            remaining = duration - elapsed
            sleep_time = min(check_interval, remaining)
            time.sleep(sleep_time)
            elapsed += sleep_time

            # 检查是否已停止
            if not self.is_playing:
                return False

            # 检查测试索引是否过期（仅当提供了index参数时）
            if index is not None:
                with self.test_index_lock:
                    if self.test_index != index:
                        return False

            # 检查鼠标是否在屏幕左上角（故障保护）
            try:
                current_position = pyautogui.position()
                if current_position[0] <= 10 and current_position[1] <= 10:
                    pyautogui.FAILSAFE = True
                    self.is_playing = False
                    self.root.after(0, lambda: self.status_var.set("回放已中止（故障保护触发）"))
                    return False
            except Exception:
                pass

        return True

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
            # 为每个动作设置 display_delay
            for i, action in enumerate(self.actions):
                if 'display_delay' not in action:
                    if i == 0:
                        action['display_delay'] = action.get('time', 0)
                    else:
                        prev_time = self.actions[i - 1].get('time', 0)
                        curr_time = action.get('time', 0)
                        action['display_delay'] = curr_time - prev_time
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

        # 保存当前展开状态
        expanded_items = set()
        for item in self.action_tree.get_children():
            if self.action_tree.item(item, 'open'):
                expanded_items.add(item)

        for item in self.action_tree.get_children():
            self.action_tree.delete(item)

        idx = 0
        while idx < len(self.actions):
            action = self.actions[idx]
            action_text = f"{self._GRAB_HANDLE_ICON} {self._describe_action(action)}"
            delay_text = self._format_delay_text(idx)
            remark_text = action.get('remark', '')

            if action['type'] == 'loop_group':
                # 循环组：使用树形结构显示，支持展开
                parent_id = f"loop_{idx}"
                # 恢复展开状态
                is_expanded = parent_id in expanded_items
                self.action_tree.insert("", tk.END, iid=parent_id, values=(action_text, delay_text, remark_text), open=is_expanded)
                # 递归添加内部动作
                loop_actions = action.get('loop_actions', [])
                self._add_loop_actions_to_tree(parent_id, loop_actions, idx, scroll_to_end, select_index, previous_selection)
                idx += 1
            elif action['type'] == 'input_loop':
                # 输入循环：使用树形结构显示，支持展开
                parent_id = f"input_loop_{idx}"
                # 恢复展开状态
                is_expanded = parent_id in expanded_items
                self.action_tree.insert("", tk.END, iid=parent_id, values=(action_text, delay_text, remark_text), open=is_expanded)
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
        has_loop_group = any(a.get('type') in ['loop_group', 'input_loop'] for a in self.actions)
        if not has_loop_group:
            self._restore_selection(scroll_to_end, select_index, previous_selection)

        # 更新主界面循环次数的最大值
        self._update_loop_count_max()

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
        # 循环组内的动作缩进4个字符
        indent = "    " if depth == 0 else "        "

        for j, loop_action in enumerate(loop_actions):
            action_type = loop_action.get('type', '')
            if action_type == 'loop_group':
                # 嵌套循环组
                loop_text = self._describe_action(loop_action)
                loop_remark = loop_action.get('remark', '')
                nested_parent_id = f"{parent_id}_{j}"
                self.action_tree.insert(parent_id, tk.END, iid=nested_parent_id, values=(indent + f"{self._GRAB_HANDLE_ICON} {loop_text}", "", loop_remark), open=False)
                # 递归添加嵌套循环组的动作
                nested_loop_actions = loop_action.get('loop_actions', [])
                self._add_loop_actions_to_tree(nested_parent_id, nested_loop_actions, parent_idx, None, None, None, depth + 1)
            else:
                # 普通动作 - 计算延时
                action_text = self._describe_action(loop_action)
                loop_remark = loop_action.get('remark', '')

                # 计算相对延时
                if j == 0:
                    delay_text = f"{loop_action.get('time', 0):.1f}s"
                else:
                    prev_time = loop_actions[j - 1].get('time', 0)
                    curr_time = loop_action.get('time', 0)
                    delay_text = f"{curr_time - prev_time:.1f}s"

                self.action_tree.insert(parent_id, tk.END, iid=f"{parent_id}_{j}", values=(indent + f"{self._GRAB_HANDLE_ICON} {action_text}", delay_text, loop_remark))

        # 恢复选择状态（只在顶层调用）
        if previous_selection is not None:
            self._restore_selection(scroll_to_end, select_index, previous_selection)

    def _describe_action(self, action):
        action_type = action['type']
        if action_type == 'random_delay':
            return f"随机延时 ({action.get('min_delay', 0):.1f}-{action.get('max_delay', 0):.1f}s)"
        if action_type == 'multiply_delay':
            return f"倍数延时 (基数: {action.get('base_delay', 0):.1f}s)"
        if action_type == 'arithmetic_delay':
            return f"等差延时 (起始: {action.get('start_delay', 0):.1f}s, 步距: {action.get('step_delay', 0):.1f}s)"
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
        if action_type == 'input_loop':
            loop_type = action.get('loop_type', 'fixed')
            template = action.get('input_template', '')
            interval_mode = action.get('interval_mode', 'fixed')
            loop_interval = action.get('loop_interval', 1.0)
            inner_count = len(action.get('loop_actions', []))

            # 构建间隔描述
            if interval_mode == 'fixed':
                interval_desc = f"固定{loop_interval}s" if isinstance(loop_interval, (int, float)) else f"固定{loop_interval.get('value', 1.0)}s"
            elif interval_mode == 'range':
                if isinstance(loop_interval, dict) and loop_interval.get('sub_type') == 'geometric':
                    interval_desc = f"等比起始{loop_interval.get('start', 1.0)}系数{loop_interval.get('ratio', 1.5)}步数{loop_interval.get('steps', 10)}"
                elif isinstance(loop_interval, dict):
                    interval_desc = f"等差起始{loop_interval.get('start', 1.0)}步距{loop_interval.get('step', 0.1)}上限{loop_interval.get('end', 2.0)}"
                else:
                    interval_desc = "等差起始1.0步距0.1上限2.0"
            elif interval_mode == 'list':
                values = loop_interval.get('values', []) if isinstance(loop_interval, dict) else []
                interval_desc = f"列表{values}"
            else:  # random
                interval_desc = f"随机{loop_interval.get('min', 0.5)}-{loop_interval.get('max', 2.0)}秒" if isinstance(loop_interval, dict) else "随机0.5-2.0秒"

            if loop_type == 'range':
                start = action.get('start', 0)
                end = action.get('end', 10)
                step = action.get('step', 1)
                return f"输入循环[范围] ({start}-{end}步{step},{interval_desc}): {template}"
            else:  # list
                items = action.get('text_list', [])
                return f"输入循环[列表] ({len(items)}项,{interval_desc}): {items[0] if items else ''}..."
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
        if action_type == 'traverse_input':
            traverse_type = action.get('traverse_type', 'list')
            values = action.get('traverse_values', [])
            x = action.get('x', 0)
            y = action.get('y', 0)
            if traverse_type == 'list':
                return f"遍历输入[列表] ({len(values)}项) at ({x}, {y})"
            elif traverse_type == 'fixed_text':
                fixed_text = action.get('fixed_text', '')
                return f"遍历输入[固定文本] ({fixed_text}) at ({x}, {y})"
            else:
                start = action.get('start', 0)
                step = action.get('step', 1)
                template = action.get('template', 'text{n}')
                return f"遍历输入[序列] ({template}) at ({x}, {y})"
        if action_type == 'move':
            return f"移动到 ({action['x']}, {action['y']})"
        if action_type == 'doubleclick':
            button = action.get('button', '')
            return f"双击 {button} at ({action['x']}, {action['y']})"
        if action_type == 'click':
            button = action.get('button', '')
            return f"单击 {button} at ({action['x']}, {action['y']})"
        if action_type == 'input':
            return f"输入: {action.get('text', '')} at ({action.get('x', 0)}, {action.get('y', 0)})"
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
        if action_type == 'delay':
            return f"延时 ({action.get('delay', 1.0):.1f}s)"
        return action_type

    def _format_delay_text(self, index):
        action = self.actions[index]
        # 如果动作有 display_delay 字段，直接使用它（跟着动作本身）
        if 'display_delay' in action:
            return f"{action['display_delay']:.1f}s"
        if action['type'] == 'random_delay':
            return f"{action.get('min_delay', 0):.1f}-{action.get('max_delay', 0):.1f}s"
        if action['type'] == 'multiply_delay':
            return f"Base: {action.get('base_delay', 0):.1f}s"
        if action['type'] == 'arithmetic_delay':
            return f"Start: {action.get('start_delay', 0):.1f}s Step: {action.get('step_delay', 0):.1f}s"
        if action['type'] == 'loop_group':
            return f"循环{action.get('loop_count', 1)}次"
        if action['type'] == 'input_loop':
            interval = action.get('loop_interval', 0.5)
            loop_type = action.get('loop_type', 'fixed')
            if loop_type == 'fixed':
                return f"{action.get('loop_count', 3)}次"
            elif loop_type == 'range':
                return f"{action.get('start', 0)}-{action.get('end', 10)}"
            else:
                return f"{len(action.get('text_list', []))}项"
        if action['type'] == 'numeric_loop':
            start = action.get('start', 0)
            end = action.get('end', 0)
            step = action.get('step', 1)
            count = abs(int((end - start) / step)) + 1 if step != 0 else 0
            return f"{count}次循环"
        if action['type'] == 'screenshot':
            return f"{action.get('delay', 0):.1f}s"
        if action['type'] == 'delay':
            return f"{action.get('delay', 1.0):.1f}s"
        relative_delay = self._get_relative_delay(index)
        return f"{relative_delay:.1f}s" if relative_delay is not None else "-"

    def _get_relative_delay(self, index):
        """获取相对延时，使用缓存优化性能"""
        if not self.actions or index < 0 or index >= len(self.actions):
            return None

        action = self.actions[index]
        if 'time' not in action:
            return None

        # 检查缓存是否有效
        if self._relative_delay_valid and index in self._relative_delay_cache:
            return self._relative_delay_cache[index]

        # 缓存无效或不存在，构建缓存
        if not self._relative_delay_valid:
            self._build_relative_delay_cache()

        # 如果缓存中存在，直接返回
        if index in self._relative_delay_cache:
            return self._relative_delay_cache[index]

        return None

    def _build_relative_delay_cache(self):
        """构建相对延时缓存，从前往后遍历一次"""
        self._relative_delay_cache = {}
        if not self.actions:
            self._relative_delay_valid = True
            return

        prev_time = 0
        for j, action in enumerate(self.actions):
            if 'time' in action:
                current_time = action['time']
                self._relative_delay_cache[j] = max(0.0, current_time - prev_time)
                prev_time = current_time
            else:
                self._relative_delay_cache[j] = 0.0

        self._relative_delay_valid = True

    def _invalidate_relative_delay_cache(self):
        """使相对延时缓存失效"""
        self._relative_delay_cache = {}
        self._relative_delay_valid = False

    def _rebuild_action_times(self, action_time_map=None, actions_list_before=None):
        """重建所有动作的 time 字段，设置 display_delay 保持显示延时不变

        display_delay 跟着动作本身移动，排序后显示的延时保持不变。
        """
        if not self.actions:
            return

        if action_time_map is None:
            return

        # 使用排序前的列表来计算显示延时
        if actions_list_before is None:
            actions_list_before = list(self.actions)

        # 记录排序前每个动作的显示延时
        # 显示延时 = 当前动作的 time - 前一个动作的 time
        action_display_delays = {}
        for i, action in enumerate(actions_list_before):
            action_id = id(action)
            if i == 0:
                # 第一个动作的显示延时就是它的 time
                action_display_delays[action_id] = action_time_map.get(action_id, 0)
            else:
                prev_action = actions_list_before[i - 1]
                prev_action_id = id(prev_action)
                curr_time = action_time_map.get(action_id, 0)
                prev_time = action_time_map.get(prev_action_id, 0)
                action_display_delays[action_id] = curr_time - prev_time

        # 排序后，设置每个动作的 display_delay（跟着动作本身）
        # 同时计算新的 time 值（绝对时间）
        cumulative_time = 0
        for action in self.actions:
            action_id = id(action)
            display_delay = action_display_delays.get(action_id, 0)
            action['display_delay'] = display_delay
            cumulative_time += display_delay
            action['time'] = cumulative_time

        # 使缓存失效
        self._invalidate_relative_delay_cache()

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
            if item.startswith('input_loop_'):
                # 输入循环: input_loop_顶层索引_子索引...
                parts = item.split('_')[2:]  # 去掉 'input_loop' 前缀
                if len(parts) == 1:
                    parent_idx = int(parts[0])
                    top_indices.append(parent_idx)
                elif len(parts) >= 2:
                    parent_idx = int(parts[0])
                    child_idx = int(parts[1])
                    if parent_idx not in loop_selections:
                        loop_selections[parent_idx] = []
                    loop_selections[parent_idx].append(child_idx)
            elif item.startswith('loop_'):
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

        # 检测是否是循环组项或输入循环项，用于展开/折叠
        if row_id.startswith('input_loop_'):
            action = self._get_action_by_row_id(row_id)
            if action and action.get('type') == 'input_loop':
                # 单击输入循环：展开或折叠，并选中
                if self.action_tree.item(row_id, 'open'):
                    self.action_tree.item(row_id, open=False)
                else:
                    self.action_tree.item(row_id, open=True)
                self.action_tree.selection_set(row_id)
                return
        elif row_id.startswith('loop_'):
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

        # 首先检测是否点击了抓手当区域（动作列的前50像素）
        column = self.action_tree.identify_column(event.x)
        if column == "#1" and event.x <= self._GRAB_HANDLE_WIDTH and '_' not in row_id:
            # 点击了抓手当区域，开始拖动
            self._drag_start_item = row_id
            self._drag_start_pos = (event.x, event.y)
            self._is_dragging = True
            try:
                self._drag_start_idx = int(row_id)
            except ValueError:
                self._drag_start_idx = None
            # 清除之前的拖动效果
            self._clear_drag_visual()
            # 添加抓手当蓝色效果
            current_tags = self.action_tree.item(row_id, 'tags')
            if current_tags:
                self.action_tree.item(row_id, tags=current_tags + ('drag_handle',))
            else:
                self.action_tree.item(row_id, tags=('drag_handle',))
            self._drag_indicator = row_id
            # 阻止Treeview处理这个点击
            return 'break'

        if ctrl_pressed:
            # Ctrl+拖动：多选模式，记录起始位置
            self._drag_start_item = row_id
            self._drag_start_pos = (event.x, event.y)
            self._is_dragging = False
            self._drag_start_idx = None
        elif row_id.startswith('loop_') or row_id.startswith('input_loop_'):
            # 循环组内的动作：阻止Treeview的展开/合并行为，但标记为已处理
            self._drag_start_item = row_id
            self._drag_start_pos = (event.x, event.y)
            self._is_dragging = False
            # 设置拖动索引，用于循环组内排序
            try:
                parts = row_id.split('_')
                if len(parts) >= 3:
                    self._drag_start_idx = int(parts[2])  # 子动作索引
                else:
                    self._drag_start_idx = None
            except (ValueError, IndexError):
                self._drag_start_idx = None
            # 阻止Treeview处理这个点击（防止触发展开/合并）
            return 'break'
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

        # 更新拖动视觉效果
        if self._is_dragging and self._drag_indicator:
            current_row = self.action_tree.identify_row(event.y)
            if current_row and current_row != self._drag_indicator:
                # 清除之前的悬停效果
                if self._drag_over_item and self._drag_over_item != self._drag_indicator:
                    current_tags = self.action_tree.item(self._drag_over_item, 'tags')
                    if current_tags:
                        new_tags = tuple(t for t in current_tags if t not in ('drag_over',))
                        self.action_tree.item(self._drag_over_item, tags=new_tags)
                # 添加新的悬停效果
                if current_row != self._drag_indicator:
                    current_tags = self.action_tree.item(current_row, 'tags')
                    if current_tags:
                        new_tags = current_tags + ('drag_over',)
                    else:
                        new_tags = ('drag_over',)
                    self.action_tree.item(current_row, tags=new_tags)
                self._drag_over_item = current_row

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
            # 检查是否是循环组内的子动作
            start_parts = self._drag_start_item.split('_')
            is_start_inner = len(start_parts) >= 3 and (start_parts[0] in ('loop', 'input_loop'))

            # 获取目标动作的父循环组信息
            row_parts = row_id.split('_')
            is_end_inner = len(row_parts) >= 3 and (row_parts[0] in ('loop', 'input_loop'))

            # 如果开始和结束都在同一个循环组内，允许排序
            if is_start_inner and is_end_inner and start_parts[0] == row_parts[0] and start_parts[1] == row_parts[1]:
                # 同一个循环组内的排序
                try:
                    parent_idx = int(start_parts[1])
                    start_child_idx = int(start_parts[2])
                    end_child_idx = int(row_parts[2])

                    parent_action = self.actions[parent_idx]
                    loop_actions = parent_action.get('loop_actions', [])

                    if 0 <= start_child_idx < len(loop_actions) and 0 <= end_child_idx < len(loop_actions) and start_child_idx != end_child_idx:
                        # 在 pop/insert 之前，记录每个动作的 time 和 display_delay
                        action_time_map = {}
                        action_display_map = {}
                        for i, loop_action in enumerate(loop_actions):
                            action_time_map[id(loop_action)] = loop_action.get('time', 0)
                            # 计算原始 display_delay
                            if i == 0:
                                action_display_map[id(loop_action)] = loop_action.get('time', 0)
                            else:
                                prev_time = loop_actions[i - 1].get('time', 0)
                                curr_time = loop_action.get('time', 0)
                                action_display_map[id(loop_action)] = curr_time - prev_time

                        # 执行排序
                        action = loop_actions.pop(start_child_idx)
                        loop_actions.insert(end_child_idx, action)

                        # 重新计算 time 和 display_delay
                        cumulative = 0
                        for loop_action in loop_actions:
                            action_id = id(loop_action)
                            display_delay = action_display_map.get(action_id, 0)
                            cumulative += display_delay
                            loop_action['time'] = cumulative
                            loop_action['display_delay'] = display_delay

                        self._update_action_list(select_index=parent_idx)
                except (ValueError, IndexError):
                    pass
            elif '_' in self._drag_start_item and (self._drag_start_item.startswith('loop_') or self._drag_start_item.startswith('input_loop_')):
                # 循环组内的项不能拖动到循环组外
                self._drag_start_item = None
                self._drag_start_pos = None
                self._is_dragging = False
                self._drag_start_idx = None
                return
            else:
                # 顶层动作的排序
                try:
                    start_idx = int(self._drag_start_item)
                    # 如果 row_id 为 None（拖动到空白区域），则移动到最后
                    if row_id is None:
                        end_idx = len(self.actions) - 1  # 目标位置是最后
                    else:
                        end_idx = int(row_id)

                    if 0 <= start_idx < len(self.actions) and start_idx != end_idx:
                        # 限制 end_idx 在有效范围内
                        end_idx = max(0, min(end_idx, len(self.actions) - 1))

                        # 在 pop/insert 之前，记录每个动作的 time 值
                        # 用 id(action) 作为键，这样排序后还能找到每个动作原本的 time 值
                        action_time_map = {}
                        for action in self.actions:
                            action_time_map[id(action)] = action.get('time', 0)

                        # 在 pop/insert 之前，记录排序前的列表副本
                        actions_list_before = list(self.actions)

                        # 记录移动的动作
                        action = self.actions[start_idx]

                        # 执行 pop 和 insert
                        self.actions.pop(start_idx)
                        insert_pos = end_idx if end_idx <= start_idx else end_idx
                        self.actions.insert(insert_pos, action)

                        # 重建 time 字段：根据动作原本的 time 值，重新计算列表中的 time
                        self._rebuild_action_times(action_time_map, actions_list_before)
                        self._update_action_list(select_index=insert_pos)
                except (ValueError, IndexError):
                    pass

        self._drag_start_item = None
        self._drag_start_pos = None
        self._is_dragging = False
        self._drag_start_idx = None
        # 清除拖动视觉效果
        self._clear_drag_visual()

    def _clear_drag_visual(self):
        """清除拖动视觉效果"""
        # 清除抓手当蓝色效果
        if self._drag_indicator:
            current_tags = self.action_tree.item(self._drag_indicator, 'tags')
            if current_tags:
                new_tags = tuple(t for t in current_tags if t not in ('drag_handle', 'drag_over'))
                self.action_tree.item(self._drag_indicator, tags=new_tags)
            self._drag_indicator = None
        # 清除悬停蓝色背景效果
        if self._drag_over_item:
            current_tags = self.action_tree.item(self._drag_over_item, 'tags')
            if current_tags:
                new_tags = tuple(t for t in current_tags if t not in ('drag_handle', 'drag_over'))
                self.action_tree.item(self._drag_over_item, tags=new_tags)
            self._drag_over_item = None

    def _on_tree_motion(self, event):
        """鼠标移动事件，用于在抓手当区域显示手型光标"""
        if self.drag_locked.get():
            return
        column = self.action_tree.identify_column(event.x)
        row_id = self.action_tree.identify_row(event.y)
        # 只有在动作列的前50像素区域内且是顶层动作时才显示手型
        if column == "#1" and event.x <= self._GRAB_HANDLE_WIDTH and row_id and '_' not in row_id:
            self.action_tree.config(cursor="hand1")
        else:
            self.action_tree.config(cursor="")

    def on_tree_double_click(self, event):
        column = self.action_tree.identify_column(event.x)
        row_id = self.action_tree.identify_row(event.y)
        if not row_id:
            return
        self.action_tree.selection_set(row_id)

        # 获取动作类型
        action = None
        try:
            if row_id.startswith('loop_') or row_id.startswith('input_loop_'):
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
        elif action_type == 'input_loop':
            self._hide_delay_spinbox()
            # 正确解析input_loop的索引
            if row_id.startswith('input_loop_'):
                index = int(row_id.split('_')[2])
            else:
                index = int(row_id)
            self._edit_input_loop_action(action, index)
        elif action_type == 'screenshot':
            # 截屏动作 - 双击打开编辑对话框
            self._hide_delay_spinbox()
            index = int(row_id) if row_id.isdigit() else None
            self._edit_screenshot_action(action, index)
        elif column == "#2":
            self._show_delay_spinbox(row_id)
        elif column == "#3":
            # 点击备注列：编辑备注
            self._edit_remark(row_id)
        elif action_type in ['click', 'doubleclick', 'move', 'variable_input', 'input', 'traverse_input']:
            # 点击/双击/移动/变量输入/遍历输入动作 - 编辑坐标和内容
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
            if (row_id.startswith('loop_') or row_id.startswith('input_loop_')) and '_' in row_id:
                action = self._get_action_by_row_id(row_id)
                if action:
                    self._show_inline_remark_editor(row_id, action)

    def _get_action_by_row_id(self, row_id):
        """根据row_id获取对应的动作"""
        if row_id.startswith('input_loop_'):
            parts = row_id.split('_')[2:]  # 去掉 'input_loop' 前缀
            if not parts:
                return None

            idx = int(parts[0])
            if idx >= len(self.actions):
                return None

            action = self.actions[idx]
            if action.get('type') != 'input_loop':
                return None

            # 逐层往下找，递归处理嵌套
            for i in range(1, len(parts)):
                child_idx = int(parts[i])
                if action.get('type') in ['loop_group', 'input_loop']:
                    loop_actions = action.get('loop_actions', [])
                else:
                    return None

                if child_idx >= len(loop_actions):
                    return None
                action = loop_actions[child_idx]

            return action
        elif row_id.startswith('loop_'):
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

    def _get_parent_loop_info(self, row_id):
        """获取循环组内动作的父循环组信息

        返回 {'loop_actions': [...], 'child_idx': int} 或 None
        """
        if row_id.startswith('input_loop_'):
            parts = row_id.split('_')[2:]
            if not parts:
                return None
            idx = int(parts[0])
            if idx >= len(self.actions):
                return None
            action = self.actions[idx]
            if action.get('type') != 'input_loop':
                return None
            # 逐层往下找
            for i in range(1, len(parts)):
                child_idx = int(parts[i])
                loop_actions = action.get('loop_actions', [])
                if child_idx >= len(loop_actions):
                    return None
                action = loop_actions[child_idx]
            # 找到动作后，获取父级loop_actions
            if len(parts) >= 2:
                # 有父级
                parent_parts = parts[:-1]
                parent_action = self.actions[int(parent_parts[0])]
                for j in range(1, len(parent_parts)):
                    parent_action = parent_action.get('loop_actions', [])[int(parent_parts[j])]
                loop_actions = parent_action.get('loop_actions', [])
                child_idx = int(parts[-1])
                return {'loop_actions': loop_actions, 'child_idx': child_idx}
            else:
                # 顶层input_loop的直接子动作
                return {'loop_actions': action.get('loop_actions', []), 'child_idx': int(parts[-1])}
        elif row_id.startswith('loop_'):
            parts = row_id.split('_')[1:]
            if not parts:
                return None
            idx = int(parts[0])
            if idx >= len(self.actions):
                return None
            action = self.actions[idx]
            if action.get('type') != 'loop_group':
                return None
            # 逐层往下找
            for i in range(1, len(parts)):
                child_idx = int(parts[i])
                loop_actions = action.get('loop_actions', [])
                if child_idx >= len(loop_actions):
                    return None
                action = loop_actions[child_idx]
            # 找到动作后，获取父级loop_actions
            if len(parts) >= 2:
                # 有父级
                parent_parts = parts[:-1]
                parent_action = self.actions[int(parent_parts[0])]
                for j in range(1, len(parent_parts)):
                    parent_action = parent_action.get('loop_actions', [])[int(parent_parts[j])]
                loop_actions = parent_action.get('loop_actions', [])
                child_idx = int(parts[-1])
                return {'loop_actions': loop_actions, 'child_idx': child_idx}
            else:
                # 顶层loop_group的直接子动作
                return {'loop_actions': action.get('loop_actions', []), 'child_idx': int(parts[-1])}
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
        """显示延时编辑"""
        if not self.actions:
            return

        # 检查是否是循环组内的动作
        action = None
        index = None
        is_loop_action = False
        loop_actions = None
        child_idx = None

        if item_id.startswith('loop_') or item_id.startswith('input_loop_'):
            # 获取循环组内的动作
            action = self._get_action_by_row_id(item_id)
            if not action:
                return
            is_loop_action = True
            # 获取同组内的动作列表和索引，用于计算相对延时
            parent_info = self._get_parent_loop_info(item_id)
            if parent_info:
                loop_actions = parent_info['loop_actions']
                child_idx = parent_info['child_idx']
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
        if action_type == 'delay':
            current_value = float(action.get('delay', 1.0))
        elif action_type == 'screenshot':
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
        elif action_type == 'input_loop':
            interval_mode = action.get('interval_mode', 'fixed')
            loop_interval = action.get('loop_interval', 1.0)
            if interval_mode == 'fixed':
                current_value = float(loop_interval) if isinstance(loop_interval, (int, float)) else 1.0
            elif interval_mode == 'range':
                current_value = float(loop_interval.get('start', 1.0)) if isinstance(loop_interval, dict) else 1.0
            elif interval_mode == 'list':
                values = loop_interval.get('values', [1.0]) if isinstance(loop_interval, dict) else [1.0]
                current_value = float(values[0] if values else 1.0)
            else:  # random
                current_value = float(loop_interval.get('min', 0.5)) if isinstance(loop_interval, dict) else 0.5
        elif action_type == 'numeric_loop':
            current_value = float(action.get('interval', 0.5))
        elif action_type == 'variable_input':
            current_value = float(action.get('time', 0.0))
        elif action_type in ['move', 'click', 'doubleclick', 'input', 'traverse_input', 'keyboard', 'scroll']:
            if is_loop_action:
                # 循环组内动作：根据同组内前一个动作计算相对延时
                if loop_actions is None or child_idx is None:
                    return
                if child_idx == 0:
                    # 组内第一个动作，相对延时就是自己的 time
                    current_value = float(action.get('time', 0.0))
                else:
                    prev_action = loop_actions[child_idx - 1]
                    curr_time = action.get('time', 0.0)
                    prev_time = prev_action.get('time', 0.0)
                    current_value = max(0.0, curr_time - prev_time)
            else:
                if index is None:
                    return
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
        self.delay_spinbox_var.set(round(current_value, 1))
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
        self.delay_spinbox.bind("<FocusOut>", self._on_spinbox_focus_out)
        self.delay_spinbox.bind("<Escape>", self._hide_delay_spinbox)
        self.active_spinbox_item = item_id

    def _on_spinbox_focus_out(self, event=None):
        """处理spinbox失去焦点的事件"""
        if self._spinbox_updating:
            return
        # 延迟执行，让Enter键有机会先处理
        self.root.after(100, self._apply_spinbox_value)

    def _apply_spinbox_value(self, *_args):
        """应用延时值"""
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
            index = None

            if item_id.startswith('loop_') or item_id.startswith('input_loop_'):
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

            if action_type == 'delay':
                action['delay'] = new_value
            elif action_type == 'screenshot':
                action['delay'] = new_value
            elif action_type == 'random_delay':
                action['min_delay'] = new_value
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
            elif action_type == 'input_loop':
                interval_mode = action.get('interval_mode', 'fixed')
                loop_interval = action.get('loop_interval', 1.0)
                if interval_mode == 'fixed':
                    action['loop_interval'] = new_value
                elif interval_mode == 'range':
                    if isinstance(loop_interval, dict):
                        loop_interval['start'] = new_value
                    else:
                        loop_interval = {'type': 'range', 'start': new_value, 'step': 0.1, 'end': 2.0}
                    action['loop_interval'] = loop_interval
                elif interval_mode == 'list':
                    if isinstance(loop_interval, dict):
                        values = loop_interval.get('values', [1.0])
                        if values:
                            values[0] = new_value
                            loop_interval['values'] = values
                        action['loop_interval'] = loop_interval
                    else:
                        action['loop_interval'] = {'type': 'list', 'values': [new_value]}
                else:  # random
                    if isinstance(loop_interval, dict):
                        loop_interval['min'] = new_value
                    else:
                        loop_interval = {'type': 'random', 'min': new_value, 'max': new_value + 1}
                    action['loop_interval'] = loop_interval
            elif action_type == 'numeric_loop':
                action['interval'] = new_value
            elif action_type == 'variable_input':
                action['time'] = new_value
            elif action_type in ['move', 'click', 'doubleclick', 'input', 'traverse_input', 'keyboard', 'scroll']:
                if item_id.startswith('loop_') or item_id.startswith('input_loop_'):
                    # 循环组内的动作：需要计算绝对时间
                    parent_info = self._get_parent_loop_info(item_id)
                    if parent_info:
                        loop_actions = parent_info['loop_actions']
                        child_idx = parent_info['child_idx']

                        # 计算前一个动作的绝对时间
                        if child_idx > 0:
                            prev_time = loop_actions[child_idx - 1].get('time', 0)
                        else:
                            prev_time = 0

                        old_time = action.get('time', 0)
                        new_absolute_time = prev_time + new_value
                        delta = new_absolute_time - old_time
                        action['time'] = new_absolute_time

                        # 调整后续动作的时间
                        for i in range(child_idx + 1, len(loop_actions)):
                            loop_actions[i]['time'] = loop_actions[i].get('time', 0) + delta
                    else:
                        return
                elif index is not None:
                    prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
                    old_time = action.get('time', 0)
                    new_absolute_time = prev_time + new_value
                    delta = new_absolute_time - old_time
                    action['time'] = new_absolute_time
                    self._shift_action_times(index + 1, delta)
                    self._invalidate_relative_delay_cache()
            else:
                return

            # 对于嵌套动作，找到其父循环组的索引并刷新
            if item_id.startswith('input_loop_'):
                parts = item_id.split('_')[2:]  # 跳过 'input' 和 'loop'
                if parts:
                    parent_idx = int(parts[0])
                    self._update_action_list(select_index=parent_idx)
            elif item_id.startswith('loop_'):
                parts = item_id.split('_')[1:]  # 跳过 'loop'
                if parts:
                    parent_idx = int(parts[0])
                    self._update_action_list(select_index=parent_idx)
            elif index is not None:
                self._update_action_list(select_index=index)
        except ValueError as exc:
            messagebox.showerror("错误", str(exc))
        finally:
            self._spinbox_updating = False
            self._hide_delay_spinbox()

    def _hide_delay_spinbox(self, *_args):
        """隐藏延时编辑器"""
        if self.delay_spinbox:
            self.delay_spinbox.destroy()
            self.delay_spinbox = None
        if hasattr(self, 'remark_entry') and self.remark_entry:
            self.remark_entry.destroy()
            self.remark_entry = None
        self.active_spinbox_item = None

    def _on_tree_scroll(self, *args):
        """树形列表滚动"""
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
                    child_indices = item[4]
                    if parent_idx not in loop_deletes:
                        loop_deletes[parent_idx] = []
                    loop_deletes[parent_idx].extend(child_indices)
                else:
                    # 顶层动作（整数索引）
                    top_indices.append(item)

            deleted_count = 0

            # 从后往前删除顶层动作
            for index in reversed(sorted(top_indices)):
                if 0 <= index < len(self.actions):
                    del self.actions[index]
                    deleted_count += 1

            # 删除循环组内动作（从后往前）
            for parent_idx in sorted(loop_deletes.keys(), reverse=True):
                if 0 <= parent_idx < len(self.actions):
                    action = self.actions[parent_idx]
                    if action.get('type') in ['loop_group', 'numeric_loop', 'input_loop'] and 'loop_actions' in action:
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

        # 计算相对延时
        relative_delay = 0
        insert_pos = len(self.actions)

        # 检查是否选择了循环组内的动作
        has_loop_inner = any(isinstance(s, tuple) and s[0] == 'loop' for s in selection)
        if not has_loop_inner:
            int_indices = [s for s in selection if isinstance(s, int)]
            if int_indices:
                insert_pos = int_indices[-1] + 1
                if insert_pos > 0 and 'time' in self.actions[insert_pos - 1]:
                    relative_delay = self.actions[insert_pos - 1].get('time', 0)

        dialog = tk.Toplevel(self.root)
        dialog.title("插入点击")
        dialog.geometry("600x200")
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
        ttk.Label(frame, text="点击位置:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        x_var = tk.StringVar(value="")
        y_var = tk.StringVar(value="")

        pos_frame = ttk.Frame(frame)
        pos_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        x_entry = ttk.Entry(pos_frame, width=8, textvariable=x_var)
        x_entry.pack(side=tk.LEFT, padx=2)

        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        y_entry = ttk.Entry(pos_frame, width=8, textvariable=y_var)
        y_entry.pack(side=tk.LEFT, padx=2)

        # 录制位置按钮（放在最后）
        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        # 相对延时
        ttk.Label(frame, text="相对延时(秒):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        delay_var = tk.DoubleVar(value=round(relative_delay, 1))
        ttk.Spinbox(frame, width=13, textvariable=delay_var, from_=0, to=999, increment=0.1, state='normal').grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # 状态标签
        status_label = ttk.Label(frame, text="", foreground="blue")
        status_label.grid(row=2, column=0, columnspan=2, pady=2)

        record_btn.config(command=lambda: self._record_position_with_dialog(dialog, x_var, y_var, status_label))

        def confirm():
            try:
                target_x = int(x_var.get()) if x_var.get().strip() else None
                target_y = int(y_var.get()) if y_var.get().strip() else None

                if target_x is None or target_y is None:
                    raise ValueError("请输入有效的坐标")

                # 重新获取选择并检查是否循环组内
                selection = self._get_selected_indices()
                has_loop_inner = any(isinstance(s, tuple) and s[0] == 'loop' for s in selection)

                if has_loop_inner:
                    # 计算新动作的绝对时间 = 前一个动作的time + 相对延时（循环组内）
                    new_time = delay_var.get()
                    if child_idx > 0 and 'time' in loop_actions[child_idx - 1]:
                        new_time = loop_actions[child_idx - 1].get('time', 0) + delay_var.get()
                else:
                    # 计算新动作的绝对时间 = 前一个动作的time + 相对延时（顶层）
                    new_time = delay_var.get()
                    if insert_pos > 0 and 'time' in self.actions[insert_pos - 1]:
                        new_time = self.actions[insert_pos - 1].get('time', 0) + delay_var.get()

                action = {
                    'type': 'click',
                    'x': target_x,
                    'y': target_y,
                    'button': 'left',
                    'time': new_time
                }

                if has_loop_inner:
                    # 选择的是循环组内的动作
                    for s in selection:
                        if isinstance(s, tuple) and s[0] == 'loop':
                            parent_loop_idx = s[1]
                            child_idx = s[2]
                            break
                    parent_action = self.actions[parent_loop_idx]
                    loop_actions = parent_action.setdefault('loop_actions', [])
                    insert_child_pos = child_idx + 1
                    loop_actions.insert(insert_child_pos, action)
                    # 调整循环组内后续动作的time，保持相对延时不变
                    self._shift_loop_action_times(loop_actions, insert_child_pos + 1, delay_var.get())
                    self._update_action_list(select_index=parent_loop_idx)
                else:
                    # 普通选择
                    int_indices = [s for s in selection if isinstance(s, int)]
                    if int_indices:
                        insert_pos = int_indices[-1] + 1
                    else:
                        insert_pos = len(self.actions)
                    self.actions.insert(insert_pos, action)
                    # 调整后续动作的time，保持相对延时不变
                    self._shift_action_times(insert_pos + 1, delay_var.get())
                    self._invalidate_relative_delay_cache()
                    self._update_action_list(select_index=insert_pos)

                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=3, column=0, columnspan=2, pady=3)

    def insert_input_action(self):
        """插入输入动作"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("插入输入")
        dialog.geometry("500x250")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        dialog.grab_set()

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 输入内容
        ttk.Label(frame, text="输入内容：").grid(row=0, column=0, padx=5, pady=10, sticky="e")
        input_text_var = tk.StringVar()
        ttk.Entry(frame, width=30, textvariable=input_text_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")

        # 输入坐标
        ttk.Label(frame, text="输入坐标：").grid(row=1, column=0, padx=5, pady=10, sticky="e")
        pos_frame = ttk.Frame(frame)
        pos_frame.grid(row=1, column=1, padx=5, pady=10, sticky="w")

        x_var = tk.StringVar()
        y_var = tk.StringVar()
        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        x_entry = ttk.Entry(pos_frame, width=8, textvariable=x_var)
        x_entry.pack(side=tk.LEFT, padx=2)
        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        y_entry = ttk.Entry(pos_frame, width=8, textvariable=y_var)
        y_entry.pack(side=tk.LEFT, padx=2)

        record_btn = ttk.Button(pos_frame, text="录制",
                               command=lambda: self._record_position_simple(dialog, x_var, y_var))
        record_btn.pack(side=tk.LEFT, padx=10)

        # 延时设置
        ttk.Label(frame, text="延时(秒)：").grid(row=2, column=0, padx=5, pady=10, sticky="e")
        delay_var = tk.DoubleVar(value=0.5)
        ttk.Spinbox(frame, width=13, textvariable=delay_var, from_=0, to=999, increment=0.1, state='normal').grid(row=2, column=1, padx=5, pady=10, sticky="w")

        def confirm():
            try:
                input_text = input_text_var.get()
                x = int(x_var.get())
                y = int(y_var.get())
                delay = delay_var.get()

                if not input_text:
                    raise ValueError("输入内容不能为空")

                # 计算time（基于最后一个动作的时间）
                time = 0
                if self.actions:
                    time = self.actions[-1].get('time', 0) + delay

                action = {
                    'type': 'input',
                    'text': input_text,
                    'x': x,
                    'y': y,
                    'time': time
                }

                self.actions.append(action)
                self._invalidate_relative_delay_cache()
                self._update_action_list(scroll_to_end=True)
                self.in_dialog_operation = False
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=3, column=0, columnspan=2, pady=10)

    def _record_position_simple(self, dialog, x_var, y_var):
        """录制位置（简单版）- 等待用户点击目标位置"""
        dialog.withdraw()
        dialog.update()

        captured_pos = [None, None]

        def on_click(x, y, button, pressed):
            if pressed:
                captured_pos[0] = x
                captured_pos[1] = y
                return False  # 停止监听

        listener = mouse.Listener(on_click=on_click)
        listener.start()
        listener.join()

        if captured_pos[0] is not None:
            x_var.set(str(captured_pos[0]))
            y_var.set(str(captured_pos[1]))

        dialog.deiconify()

    def insert_traverse_input(self):
        """插入遍历输入：根据主页循环次数遍历参数"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("遍历输入")
        dialog.geometry("500x340")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        dialog.grab_set()

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 遍历类型选择
        ttk.Label(main_frame, text="遍历类型:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        type_var = tk.StringVar(value="fixed_text")
        type_frame = ttk.Frame(main_frame)
        type_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(type_frame, text="固定文本", variable=type_var, value="fixed_text",
                       command=lambda: update_type_ui("fixed_text")).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="列表遍历", variable=type_var, value="list",
                       command=lambda: update_type_ui("list")).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="序列", variable=type_var, value="sequence",
                       command=lambda: update_type_ui("sequence")).pack(side=tk.LEFT, padx=5)

        # 点击位置
        ttk.Label(main_frame, text="点击位置:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        pos_frame = ttk.Frame(main_frame)
        pos_frame.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        x_var = tk.StringVar(value="")
        y_var = tk.StringVar(value="")
        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        ttk.Entry(pos_frame, width=8, textvariable=x_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        ttk.Entry(pos_frame, width=8, textvariable=y_var).pack(side=tk.LEFT, padx=2)

        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        # 参数框架（动态切换）
        params_frame = ttk.Frame(main_frame)
        params_frame.grid(row=2, column=0, columnspan=2, sticky="w", pady=10)

        # 列表遍历参数
        list_frame = ttk.Frame(params_frame)
        list_row1 = ttk.Frame(list_frame)
        list_row1.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(list_row1, text="遍历列表:").pack(side=tk.LEFT, padx=2)
        list_text = tk.Text(list_row1, width=25, height=6)
        list_text.pack(side=tk.LEFT, padx=2)
        list_text.insert(tk.END, "value1\nvalue2\nvalue3\nvalue4\nvalue5")
        ttk.Label(list_frame, text="(每行一个值)").pack(side=tk.TOP, padx=2, pady=(2, 0))

        # 序列参数
        seq_frame = ttk.Frame(params_frame)
        seq_row1 = ttk.Frame(seq_frame)
        seq_row1.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(seq_row1, text="起始值:").pack(side=tk.LEFT, padx=2)
        start_var = tk.DoubleVar(value=0)
        ttk.Entry(seq_row1, width=8, textvariable=start_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(seq_row1, text="步距:").pack(side=tk.LEFT, padx=2)
        step_var = tk.DoubleVar(value=1)
        ttk.Entry(seq_row1, width=8, textvariable=step_var).pack(side=tk.LEFT, padx=2)
        seq_row2 = ttk.Frame(seq_frame)
        seq_row2.pack(side=tk.TOP, fill=tk.X, pady=(5, 0))
        ttk.Label(seq_row2, text="模板:").pack(side=tk.LEFT, padx=2)
        seq_template_var = tk.StringVar(value="text{n}text")
        ttk.Entry(seq_row2, width=15, textvariable=seq_template_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(seq_row2, text="(用{n}替换)").pack(side=tk.LEFT, padx=2)

        # 固定文本参数
        fixed_text_frame = ttk.Frame(params_frame)
        ttk.Label(fixed_text_frame, text="输入文本:").pack(side=tk.LEFT, padx=2)
        fixed_text_var = tk.StringVar(value="text1")
        ttk.Entry(fixed_text_frame, width=25, textvariable=fixed_text_var).pack(side=tk.LEFT, padx=2)

        def update_type_ui(mode):
            list_frame.grid_forget()
            seq_frame.grid_forget()
            fixed_text_frame.grid_forget()
            if mode == "list":
                list_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
            elif mode == "sequence":
                seq_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
            else:  # fixed_text
                fixed_text_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)

        # 初始化UI
        update_type_ui("fixed_text")

        # 录制按钮功能
        def start_record():
            dialog.withdraw()
            dialog.update()
            captured_pos = [None, None]

            def on_click(x, y, button, pressed):
                if pressed:
                    captured_pos[0] = x
                    captured_pos[1] = y
                    return False

            listener = mouse.Listener(on_click=on_click)
            listener.start()
            listener.join()

            if captured_pos[0] is not None:
                x_var.set(str(captured_pos[0]))
                y_var.set(str(captured_pos[1]))
            dialog.deiconify()

        record_btn.config(command=start_record)

        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=15)

        def confirm():
            try:
                x = int(x_var.get()) if x_var.get() else 0
                y = int(y_var.get()) if y_var.get() else 0

                # 获取遍历值列表
                traverse_type = type_var.get()
                values = []
                start = 0
                step = 1
                fixed_text = ""
                if traverse_type == "list":
                    text_content = list_text.get("1.0", tk.END).strip()
                    if not text_content:
                        raise ValueError("遍历列表不能为空")
                    values = [line.strip() for line in text_content.split('\n') if line.strip()]
                elif traverse_type == "sequence":
                    start = start_var.get()
                    step = step_var.get()
                    seq_template = seq_template_var.get()
                else:  # fixed_text
                    fixed_text = fixed_text_var.get()
                    if not fixed_text:
                        raise ValueError("输入文本不能为空")

                if not values and traverse_type == "list":
                    raise ValueError("遍历列表不能为空")

                # 构建遍历输入动作
                action = {
                    'type': 'traverse_input',
                    'traverse_type': traverse_type,
                    'x': x,
                    'y': y,
                    'traverse_values': values if traverse_type == "list" else [],
                    'start': start if traverse_type == "sequence" else None,
                    'step': step if traverse_type == "sequence" else None,
                    'template': seq_template if traverse_type == "sequence" else None,
                    'fixed_text': fixed_text if traverse_type == "fixed_text" else None,
                    'remark': ''
                }

                # 插入到动作列表
                selection = self._get_selected_indices()

                # 检查是否选择了循环组内的动作
                has_loop_inner = any(isinstance(s, tuple) and s[0] == 'loop' for s in selection)

                if has_loop_inner:
                    # 选择的是循环组内的动作，在循环组内插入
                    for s in selection:
                        if isinstance(s, tuple) and s[0] == 'loop':
                            parent_loop_idx = s[1]
                            child_idx = s[2]
                            break

                    parent_action = self.actions[parent_loop_idx]
                    loop_actions = parent_action.setdefault('loop_actions', [])

                    # 在循环组内选中动作之后插入
                    insert_child_pos = child_idx + 1
                    loop_actions.insert(insert_child_pos, action)

                    # 调整循环组内后续动作的time，保持相对延时不变
                    self._shift_loop_action_times(loop_actions, insert_child_pos + 1, 0.5)
                    # 更新循环组内的延时
                    self._update_action_list(select_index=parent_loop_idx)
                else:
                    # 普通选择 - 在顶层动作列表中插入
                    int_indices = [s for s in selection if isinstance(s, int)]

                    if int_indices:
                        insert_pos = int_indices[-1] + 1
                    else:
                        insert_pos = len(self.actions)

                    # 计算新动作的绝对时间 = 前一个动作的time + 相对延时(默认0.5)
                    if insert_pos > 0 and 'time' in self.actions[insert_pos - 1]:
                        action['time'] = self.actions[insert_pos - 1].get('time', 0) + 0.5
                    elif self.actions:
                        action['time'] = self.actions[-1].get('time', 0) + 0.5

                    self.actions.insert(insert_pos, action)
                    # 调整后续动作的time，保持相对延时不变
                    self._shift_action_times(insert_pos + 1, 0.5)
                    self._invalidate_relative_delay_cache()
                    self._update_action_list(scroll_to_end=True)

                self.in_dialog_operation = False
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(btn_frame, text="确定", command=confirm).pack(side=tk.LEFT, padx=20)
        ttk.Button(btn_frame, text="取消", command=on_dialog_close).pack(side=tk.LEFT, padx=5)

    def insert_random_delay(self):
        """插入延时动作（支持随机延时和指数延时）"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("插入延时")
        dialog.geometry("540x160")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        dialog.grab_set()

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

        # 延时类型选择
        type_var = tk.StringVar(value="fixed")

        # 延时设置框架
        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 类型选择
        ttk.Label(frame, text="延时类型:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        type_frame = ttk.Frame(frame)
        type_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(type_frame, text="固定延时", variable=type_var, value="fixed",
                       command=lambda: update_ui("fixed")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(type_frame, text="随机范围", variable=type_var, value="random",
                       command=lambda: update_ui("random")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(type_frame, text="倍数增长", variable=type_var, value="multiply",
                       command=lambda: update_ui("multiply")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(type_frame, text="等差递增", variable=type_var, value="arithmetic",
                       command=lambda: update_ui("arithmetic")).pack(side=tk.LEFT, padx=2)

        # 固定延时参数
        fixed_frame = ttk.Frame(frame)
        fixed_frame.grid(row=1, column=0, columnspan=2, sticky="ew")

        ttk.Label(fixed_frame, text="延时(秒):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        fixed_delay_var = tk.DoubleVar(value=1.0)
        ttk.Entry(fixed_frame, width=10, textvariable=fixed_delay_var).grid(row=0, column=1, padx=5, pady=5, sticky="w")

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
            if delay_type == "fixed":
                fixed_frame.grid()
                random_frame.grid_remove()
                exp_frame.grid_remove()
                arithmetic_frame.grid_remove()
            elif delay_type == "random":
                fixed_frame.grid_remove()
                random_frame.grid()
                exp_frame.grid_remove()
                arithmetic_frame.grid_remove()
            elif delay_type == "multiply":
                fixed_frame.grid_remove()
                random_frame.grid_remove()
                exp_frame.grid()
                arithmetic_frame.grid_remove()
            else:  # arithmetic
                fixed_frame.grid_remove()
                random_frame.grid_remove()
                exp_frame.grid_remove()
                arithmetic_frame.grid()

        # 初始化UI状态
        update_ui("fixed")

        def confirm():
            try:
                delay_type = type_var.get()

                if delay_type == "fixed":
                    delay = float(fixed_delay_var.get())
                    if delay < 0:
                        raise ValueError("延时不能为负数")

                    action = {
                        'type': 'delay',
                        'delay': delay,
                        'time': 0
                    }
                elif delay_type == "random":
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

                # 检查是否选择了循环组内的动作
                has_loop_inner = any(isinstance(s, tuple) and s[0] == 'loop' for s in selection)

                if has_loop_inner:
                    # 选择的是循环组内的动作，在循环组内插入延时
                    for s in selection:
                        if isinstance(s, tuple) and s[0] == 'loop':
                            parent_loop_idx = s[1]
                            child_idx = s[2]  # 选中的子动作索引
                            break

                    parent_action = self.actions[parent_loop_idx]
                    loop_actions = parent_action.setdefault('loop_actions', [])

                    # 在循环组内选中动作之后插入的位置
                    insert_child_pos = child_idx + 1

                    # 计算新动作的time（基于前一个动作）
                    if insert_child_pos > 0 and insert_child_pos <= len(loop_actions) and 'time' in loop_actions[insert_child_pos - 1]:
                        prev_time = loop_actions[insert_child_pos - 1].get('time', 0)
                        if delay_type == "fixed":
                            action['time'] = prev_time + action.get('delay', 0)
                        elif delay_type == "random":
                            action['time'] = prev_time + action.get('min_delay', 0)
                        elif delay_type == "multiply":
                            action['time'] = prev_time + action.get('base_delay', 0)
                        else:  # arithmetic
                            action['time'] = prev_time + action.get('start_delay', 0)

                    loop_actions.insert(insert_child_pos, action)

                    # 调整循环组内后续动作的time，保持相对延时不变
                    if delay_type == "fixed":
                        shift = action.get('delay', 0)
                    elif delay_type == "random":
                        shift = action.get('min_delay', 0)
                    elif delay_type == "multiply":
                        shift = action.get('base_delay', 0)
                    else:  # arithmetic
                        shift = action.get('start_delay', 0)
                    self._shift_loop_action_times(loop_actions, insert_child_pos + 1, shift)
                    # 更新循环组内的延时
                    self._update_action_list(select_index=parent_loop_idx)
                else:
                    # 普通选择 - 在顶层动作列表中插入
                    int_indices = [s for s in selection if isinstance(s, int)]

                    if int_indices:
                        # 插入到选中项之后
                        insert_pos = int_indices[-1] + 1
                    else:
                        insert_pos = len(self.actions)

                    if insert_pos > 0 and self.actions:
                        prev_time = self.actions[insert_pos - 1].get('time', 0)
                        # 根据延时类型计算新动作的time
                        if delay_type == "fixed":
                            action['time'] = prev_time + action.get('delay', 0)
                        elif delay_type == "random":
                            action['time'] = prev_time + action.get('min_delay', 0)
                        elif delay_type == "multiply":
                            action['time'] = prev_time + action.get('base_delay', 0)
                        else:  # arithmetic
                            action['time'] = prev_time + action.get('start_delay', 0)

                    self.actions.insert(insert_pos, action)
                    # 调整后续动作的time
                    if delay_type == "fixed":
                        shift = action.get('delay', 0)
                    elif delay_type == "random":
                        shift = action.get('min_delay', 0)
                    elif delay_type == "multiply":
                        shift = action.get('base_delay', 0)
                    else:  # arithmetic
                        shift = action.get('start_delay', 0)
                    self._shift_action_times(insert_pos + 1, shift)
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

        # 检查是否选择了循环组内的动作
        has_loop_inner = any(isinstance(s, tuple) and s[0] == 'loop' for s in selection)

        if has_loop_inner:
            # 选择的是循环组内的动作
            for s in selection:
                if isinstance(s, tuple) and s[0] == 'loop':
                    is_loop_inner = True
                    parent_loop_idx = s[1]
                    start_idx = s[2]  # start child index
                    end_idx = s[3]    # end child index
                    break

            parent_action = self.actions[parent_loop_idx]
            loop_actions = parent_action.setdefault('loop_actions', [])
            loop_actions_selected = loop_actions[start_idx:end_idx+1]
            info_text = f"循环组{parent_loop_idx+1}内: 第{start_idx+1}-{end_idx+1}个动作 (共{end_idx-start_idx+1}个)"
        else:
            # 普通选择 - 只取整数索引
            int_indices = [s for s in selection if isinstance(s, int)]
            if int_indices:
                start_idx = int_indices[0]
                end_idx = int_indices[-1]
                loop_actions_selected = self.actions[start_idx:end_idx+1]
                info_text = f"选中的动作: {start_idx+1} - {end_idx+1} (共{end_idx-start_idx+1}个)"
            else:
                messagebox.showinfo("提示", "请选择有效的动作")
                self.in_dialog_operation = False
                return

        dialog = tk.Toplevel(self.root)
        dialog.title("设为循环组")
        dialog.geometry("350x190")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        dialog.grab_set()

        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 计算循环次数上限（如果包含列表遍历类型的遍历输入，最大值设为列表长度）
        max_loop_count = 9999
        for act in loop_actions_selected:
            if act.get('type') == 'traverse_input' and act.get('traverse_type') == 'list':
                values = act.get('traverse_values', [])
                if values:
                    max_loop_count = len(values)
                    break

        # 循环次数
        loop_count_frame = ttk.Frame(main_frame)
        loop_count_frame.pack(fill=tk.X, pady=10)
        ttk.Label(loop_count_frame, text="循环次数：").pack(side=tk.LEFT)
        loop_count_var = tk.IntVar(value=min(3, max_loop_count))
        loop_count_spinbox = ttk.Spinbox(loop_count_frame, from_=1, to=max_loop_count, width=15, textvariable=loop_count_var)
        loop_count_spinbox.pack(side=tk.LEFT, padx=8)

        # 选中的动作提示
        info_label = ttk.Label(main_frame, text=info_text)
        info_label.pack(pady=10)

        def confirm():
            try:
                loop_count = loop_count_var.get()

                if loop_count <= 0:
                    raise ValueError("循环次数必须大于0")

                # 检查循环组内的遍历输入动作，自动调整循环次数为列表长度
                for act in loop_actions_selected:
                    if act.get('type') == 'traverse_input' and act.get('traverse_type') == 'list':
                        values = act.get('traverse_values', [])
                        if values and len(values) < loop_count:
                            loop_count = len(values)
                            loop_count_var.set(loop_count)
                            break

                # 使用固定间隔，默认0.5秒
                interval_type = "fixed"
                interval_params = {'type': 'fixed', 'value': 0.5}

                if is_loop_inner:
                    parent_action = self.actions[parent_loop_idx]
                    loop_actions = parent_action.setdefault('loop_actions', [])
                    selected_actions = loop_actions[start_idx:end_idx+1]

                    nested_loop = {
                        'type': 'loop_group',
                        'loop_count': loop_count,
                        'loop_interval_type': interval_type,
                        'loop_interval': interval_params,
                        'loop_actions': selected_actions,
                        'time': 0
                    }

                    del loop_actions[start_idx:end_idx+1]
                    loop_actions.insert(start_idx, nested_loop)
                    parent_action['loop_actions'] = loop_actions
                    self._update_action_list(select_index=parent_loop_idx)
                else:
                    action = {
                        'type': 'loop_group',
                        'loop_count': loop_count,
                        'loop_interval_type': interval_type,
                        'loop_interval': interval_params,
                        'loop_actions': loop_actions_selected,
                        'time': self.actions[start_idx].get('time', 0) if self.actions else 0
                    }

                    del self.actions[start_idx:end_idx+1]
                    self.actions.insert(start_idx, action)
                    self._update_action_list(select_index=start_idx)

                self.in_dialog_operation = False
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(main_frame, text="确定", command=confirm).pack(pady=10)

    def insert_input_loop(self):
        """插入输入循环：三种类型的输入循环组"""
        self.in_dialog_operation = True

        # 获取当前选中项
        selection = self._get_selected_indices()

        # 检查是否选择了循环组内的动作
        has_loop_inner = any(isinstance(s, tuple) and s[0] == 'loop' for s in selection)

        if selection and not has_loop_inner:
            int_indices = [s for s in selection if isinstance(s, int)]
            if int_indices:
                insert_pos = int_indices[-1] + 1
            else:
                insert_pos = len(self.actions)
        else:
            insert_pos = len(self.actions)

        dialog = tk.Toplevel(self.root)
        dialog.title("输入循环")
        dialog.geometry("800x700")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        dialog.grab_set()

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

        # 主框架
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 循环类型选择
        ttk.Label(main_frame, text="循环类型:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        type_var = tk.StringVar(value="range")
        type_frame = ttk.Frame(main_frame)
        type_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(type_frame, text="数值范围", variable=type_var, value="range",
                       command=lambda: update_type_ui("range")).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="列表遍历", variable=type_var, value="list",
                       command=lambda: update_type_ui("list")).pack(side=tk.LEFT, padx=5)

        # 输入模板（固定重复和数值范围使用）
        ttk.Label(main_frame, text="输入模板:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        template_var = tk.StringVar(value="text{n}")
        template_entry = ttk.Entry(main_frame, width=30, textvariable=template_var)
        template_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(main_frame, text="(使用{n}作为循环变量)").grid(row=2, column=1, padx=5, pady=2, sticky="w")

        # 点击位置
        ttk.Label(main_frame, text="点击位置:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        pos_frame = ttk.Frame(main_frame)
        pos_frame.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        x_var = tk.StringVar(value="")
        y_var = tk.StringVar(value="")
        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        ttk.Entry(pos_frame, width=8, textvariable=x_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        ttk.Entry(pos_frame, width=8, textvariable=y_var).pack(side=tk.LEFT, padx=2)

        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        # 循环参数框架（动态切换）
        params_frame = ttk.Frame(main_frame)
        params_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=5)

        # 固定重复参数
        fixed_frame = ttk.Frame(params_frame)
        ttk.Label(fixed_frame, text="重复次数:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        loop_count_var = tk.IntVar(value=3)
        ttk.Entry(fixed_frame, width=10, textvariable=loop_count_var).grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # 数值范围参数
        range_frame = ttk.Frame(params_frame)
        ttk.Label(range_frame, text="起始值:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        start_var = tk.IntVar(value=0)
        ttk.Entry(range_frame, width=8, textvariable=start_var).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(range_frame, text="步距:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        step_var = tk.IntVar(value=1)
        ttk.Entry(range_frame, width=8, textvariable=step_var).grid(row=0, column=3, padx=5, pady=5, sticky="w")
        ttk.Label(range_frame, text="结束值:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        end_var = tk.IntVar(value=10)
        ttk.Entry(range_frame, width=8, textvariable=end_var).grid(row=0, column=5, padx=5, pady=5, sticky="w")

        # 列表遍历参数
        list_frame = ttk.Frame(params_frame)
        ttk.Label(list_frame, text="文本列表:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        list_text = tk.Text(list_frame, width=30, height=5)
        list_text.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        list_text.insert(tk.END, "apple\nbanana\ncherry")

        ttk.Label(list_frame, text="(每行一个文本)").grid(row=1, column=1, padx=5, pady=2, sticky="w")
        ttk.Button(list_frame, text="从文件加载", command=lambda: load_list_from_file()).grid(row=0, column=2, padx=5, pady=5)

        list_file_var = tk.StringVar(value="")
        list_file_label = ttk.Label(list_frame, textvariable=list_file_var, foreground="blue")
        list_file_label.grid(row=1, column=2, padx=5, pady=2)

        # 循环间隔和清空选项
        options_frame = ttk.Frame(main_frame)
        options_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=5)

        # 间隔模式选择
        interval_mode_var = tk.StringVar(value="fixed")

        ttk.Label(options_frame, text="间隔模式:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        mode_frame = ttk.Frame(options_frame)
        mode_frame.grid(row=0, column=1, columnspan=2, sticky="w")
        ttk.Radiobutton(mode_frame, text="固定", variable=interval_mode_var, value="fixed",
                        command=lambda: update_interval_mode_ui("fixed")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(mode_frame, text="范围", variable=interval_mode_var, value="range",
                        command=lambda: update_interval_mode_ui("range")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(mode_frame, text="列表", variable=interval_mode_var, value="list",
                        command=lambda: update_interval_mode_ui("list")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(mode_frame, text="随机", variable=interval_mode_var, value="random",
                        command=lambda: update_interval_mode_ui("random")).pack(side=tk.LEFT, padx=2)

        # 间隔参数框架
        interval_params_frame = ttk.Frame(options_frame)
        interval_params_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=5)

        # 固定间隔
        fixed_interval_frame = ttk.Frame(interval_params_frame)
        ttk.Label(fixed_interval_frame, text="间隔(秒):").pack(side=tk.LEFT, padx=2)
        interval_var = tk.DoubleVar(value=1.0)
        ttk.Entry(fixed_interval_frame, width=10, textvariable=interval_var).pack(side=tk.LEFT, padx=2)

        # 范围间隔（等差/等比）
        range_interval_frame = ttk.Frame(interval_params_frame)
        range_interval_frame.pack_propagate(False)  # 防止框架扩展
        range_sub_type_var = tk.StringVar(value="arithmetic")  # arithmetic / geometric

        # Radiobutton 放在外层，始终可见
        ttk.Radiobutton(range_interval_frame, text="等差", variable=range_sub_type_var, value="arithmetic",
                        command=lambda: update_range_sub_ui()).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(range_interval_frame, text="等比", variable=range_sub_type_var, value="geometric",
                        command=lambda: update_range_sub_ui()).pack(side=tk.LEFT, padx=2)

        # 等差子模式 - 参数区域
        arithmetic_frame = ttk.Frame(range_interval_frame)
        ttk.Label(arithmetic_frame, text="起始:").pack(side=tk.LEFT, padx=2)
        interval_start_var = tk.DoubleVar(value=1.0)
        ttk.Entry(arithmetic_frame, width=8, textvariable=interval_start_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(arithmetic_frame, text="步距:").pack(side=tk.LEFT, padx=2)
        interval_step_var = tk.DoubleVar(value=0.1)
        ttk.Entry(arithmetic_frame, width=8, textvariable=interval_step_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(arithmetic_frame, text="上限:").pack(side=tk.LEFT, padx=2)
        interval_end_var = tk.DoubleVar(value=2.0)
        ttk.Entry(arithmetic_frame, width=8, textvariable=interval_end_var).pack(side=tk.LEFT, padx=2)

        # 等比子模式 - 参数区域
        geometric_frame = ttk.Frame(range_interval_frame)
        ttk.Label(geometric_frame, text="起始:").pack(side=tk.LEFT, padx=2)
        interval_geo_start_var = tk.DoubleVar(value=1.0)
        ttk.Entry(geometric_frame, width=8, textvariable=interval_geo_start_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(geometric_frame, text="等比系数:").pack(side=tk.LEFT, padx=2)
        interval_ratio_var = tk.DoubleVar(value=1.5)
        ttk.Entry(geometric_frame, width=8, textvariable=interval_ratio_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(geometric_frame, text="步数:").pack(side=tk.LEFT, padx=2)
        interval_steps_var = tk.IntVar(value=10)
        ttk.Entry(geometric_frame, width=8, textvariable=interval_steps_var).pack(side=tk.LEFT, padx=2)

        def update_range_sub_ui():
            arithmetic_frame.pack_forget()
            geometric_frame.pack_forget()
            if range_sub_type_var.get() == "arithmetic":
                arithmetic_frame.pack(side=tk.LEFT, padx=2, anchor="w")
            else:
                geometric_frame.pack(side=tk.LEFT, padx=2, anchor="w")

        update_range_sub_ui()

        # 列表间隔
        list_interval_frame = ttk.Frame(interval_params_frame)
        ttk.Label(list_interval_frame, text="列表(逗号分隔):").pack(side=tk.LEFT, padx=2)
        interval_list_var = tk.StringVar(value="0.5,1.0,1.5,2.0")
        ttk.Entry(list_interval_frame, width=20, textvariable=interval_list_var).pack(side=tk.LEFT, padx=2)

        # 随机间隔
        random_interval_frame = ttk.Frame(interval_params_frame)
        ttk.Label(random_interval_frame, text="最小:").pack(side=tk.LEFT, padx=2)
        interval_random_min_var = tk.DoubleVar(value=0.5)
        ttk.Entry(random_interval_frame, width=8, textvariable=interval_random_min_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(random_interval_frame, text="最大:").pack(side=tk.LEFT, padx=2)
        interval_random_max_var = tk.DoubleVar(value=2.0)
        ttk.Entry(random_interval_frame, width=8, textvariable=interval_random_max_var).pack(side=tk.LEFT, padx=2)

        def update_interval_mode_ui(mode):
            fixed_interval_frame.pack_forget()
            range_interval_frame.pack_forget()
            list_interval_frame.pack_forget()
            random_interval_frame.pack_forget()
            if mode == "fixed":
                fixed_interval_frame.pack(side=tk.LEFT, padx=2, anchor="w")
            elif mode == "range":
                range_interval_frame.pack(side=tk.LEFT, padx=2, anchor="w")
                update_range_sub_ui()
            elif mode == "list":
                list_interval_frame.pack(side=tk.LEFT, padx=2, anchor="w")
            else:  # random
                random_interval_frame.pack(side=tk.LEFT, padx=2, anchor="w")

        # 初始化间隔UI
        update_interval_mode_ui("fixed")

        clear_text_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="输入前清空文本框", variable=clear_text_var).grid(row=2, column=0, columnspan=3, pady=5)

        # 循环内动作列表 - 使用Treeview支持拖动排序
        ttk.Label(main_frame, text="循环内动作:").grid(row=6, column=0, padx=5, pady=5, sticky="ne")
        actions_frame = ttk.Frame(main_frame)
        actions_frame.grid(row=6, column=1, padx=5, pady=5, sticky="nsew")

        columns = ("action", "delay", "remark")
        loop_action_tree = ttk.Treeview(actions_frame, columns=columns, show="headings", height=8)
        loop_action_tree.heading("action", text="动作描述")
        loop_action_tree.heading("delay", text="延时")
        loop_action_tree.heading("remark", text="备注")
        loop_action_tree.column("action", anchor="w", width=220, stretch=True)
        loop_action_tree.column("delay", anchor="center", width=60, stretch=False)
        loop_action_tree.column("remark", anchor="w", width=140, stretch=True)
        loop_action_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        loop_scrollbar = ttk.Scrollbar(actions_frame, orient=tk.VERTICAL, command=loop_action_tree.yview)
        loop_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        loop_action_tree.configure(yscrollcommand=loop_scrollbar.set)

        loop_actions = []  # 存储循环内动作

        # 拖动排序相关变量
        drag_start_item = [None]
        drag_start_idx = [None]
        is_dragging = [False]

        def refresh_actions_list():
            """刷新循环内动作列表"""
            for item in loop_action_tree.get_children():
                loop_action_tree.delete(item)
            for idx, act in enumerate(loop_actions):
                desc = self._describe_action(act)
                act_type = act.get('type')
                if act_type == 'delay':
                    delay_text = f"{act.get('delay', 1.0):.1f}s"
                else:
                    delay_text = f"{act.get('time', 0):.1f}s"
                remark_text = act.get('remark', '')
                loop_action_tree.insert("", tk.END, iid=str(idx), values=(desc, delay_text, remark_text))

        # 合并的点击处理函数
        def on_loop_action_click(event):
            """处理点击事件：先处理编辑器销毁，再处理延时/备注编辑"""
            # 先销毁编辑器
            if remark_entry_in_loop[0]:
                remark_entry_in_loop[0].destroy()
                remark_entry_in_loop[0] = None
            if delay_entry_in_loop[0]:
                delay_entry_in_loop[0].destroy()
                delay_entry_in_loop[0] = None
            # 延时列和备注列的编辑
            row_id = loop_action_tree.identify_row(event.y)
            if not row_id:
                return
            column = loop_action_tree.identify_column(event.x)
            try:
                action_idx = int(row_id)
            except ValueError:
                return
            if action_idx >= len(loop_actions):
                return
            action = loop_actions[action_idx]

            # 延时列编辑
            if column == "#2":
                loop_action_tree.update_idletasks()
                bbox = loop_action_tree.bbox(row_id, column='#2')
                if not bbox:
                    return
                x, y, width, height = bbox
                act_type = action.get('type')
                if act_type == 'delay':
                    delay_var_in_loop[0].set(f"{action.get('delay', 1.0):.1f}")
                else:
                    delay_var_in_loop[0].set(f"{action.get('time', 0):.1f}")
                delay_entry_in_loop[0] = ttk.Entry(loop_action_tree, textvariable=delay_var_in_loop[0], width=10)
                delay_entry_in_loop[0].place(x=x, y=y, width=width, height=height)
                delay_entry_in_loop[0].focus_set()
                delay_entry_in_loop[0].select_range(0, tk.END)

                def save_loop_delay(event=None):
                    try:
                        new_delay = float(delay_var_in_loop[0].get())
                        act_type = loop_actions[action_idx].get('type')
                        if act_type == 'delay':
                            loop_actions[action_idx]['delay'] = new_delay
                        else:
                            loop_actions[action_idx]['time'] = new_delay
                    except ValueError:
                        pass
                    delay_entry_in_loop[0].destroy()
                    delay_entry_in_loop[0] = None
                    refresh_actions_list()

                delay_entry_in_loop[0].bind('<Return>', save_loop_delay)
                delay_entry_in_loop[0].bind('<Escape>', lambda e: delay_entry_in_loop[0].destroy() if delay_entry_in_loop[0] else None)
                delay_entry_in_loop[0].bind('<FocusOut>', save_loop_delay)
                return

            # 备注列编辑
            if column == "#3":
                loop_action_tree.update_idletasks()
                bbox = loop_action_tree.bbox(row_id, column='#3')
                if not bbox:
                    return
                x, y, width, height = bbox
                remark_var_in_loop[0].set(action.get('remark', ''))
                remark_entry_in_loop[0] = ttk.Entry(loop_action_tree, textvariable=remark_var_in_loop[0], width=20)
                remark_entry_in_loop[0].place(x=x, y=y, width=width, height=height)
                remark_entry_in_loop[0].focus_set()
                remark_entry_in_loop[0].select_range(0, tk.END)

                def save_loop_remark(event=None):
                    loop_actions[action_idx]['remark'] = remark_var_in_loop[0].get()
                    remark_entry_in_loop[0].destroy()
                    remark_entry_in_loop[0] = None
                    refresh_actions_list()

                remark_entry_in_loop[0].bind('<Return>', save_loop_remark)
                remark_entry_in_loop[0].bind('<Escape>', lambda e: remark_entry_in_loop[0].destroy() if remark_entry_in_loop[0] else None)
                remark_entry_in_loop[0].bind('<FocusOut>', save_loop_remark)
                return

            # 动作列: 允许正常拖动排序
            drag_start_item[0] = row_id
            drag_start_idx[0] = action_idx
            is_dragging[0] = False

        def on_loop_action_drag_motion(event):
            """拖动中"""
            dx = abs(event.x - loop_action_tree.winfo_rootx())
            dy = abs(event.y - loop_action_tree.winfo_rooty())
            if dx > 3 or dy > 3:
                is_dragging[0] = True
                row_id = loop_action_tree.identify_row(event.y)
                if row_id:
                    loop_action_tree.see(row_id)

        def on_loop_action_drag_end(event):
            """结束拖动"""
            if drag_start_idx[0] is None:
                return
            if not is_dragging[0]:
                drag_start_item[0] = None
                drag_start_idx[0] = None
                return

            row_id = loop_action_tree.identify_row(event.y)
            if row_id and row_id != drag_start_item[0]:
                try:
                    end_idx = int(row_id)
                    start_idx = drag_start_idx[0]
                    if 0 <= start_idx < len(loop_actions) and 0 <= end_idx < len(loop_actions):
                        action = loop_actions.pop(start_idx)
                        if end_idx > start_idx:
                            end_idx -= 1
                        loop_actions.insert(end_idx, action)
                        refresh_actions_list()
                except ValueError:
                    pass

            drag_start_item[0] = None
            drag_start_idx[0] = None
            is_dragging[0] = False

        loop_action_tree.bind("<Button-1>", on_loop_action_click)
        loop_action_tree.bind("<B1-Motion>", on_loop_action_drag_motion)
        loop_action_tree.bind("<ButtonRelease-1>", on_loop_action_drag_end)

        # 备注和延时编辑（内联）
        remark_entry_in_loop = [None]
        remark_var_in_loop = [tk.StringVar()]
        delay_entry_in_loop = [None]
        delay_var_in_loop = [tk.StringVar()]

        def show_loop_remark_editor(event):
            """显示循环内动作的备注或延时编辑器"""
            row_id = loop_action_tree.identify_row(event.y)
            if not row_id:
                return
            column = loop_action_tree.identify_column(event.x)
            try:
                action_idx = int(row_id)
            except ValueError:
                return
            if action_idx >= len(loop_actions):
                return
            action = loop_actions[action_idx]

            # 延时列编辑 - 所有动作的time都可以修改
            if column == "#2":
                # 隐藏备注编辑器
                if remark_entry_in_loop[0]:
                    remark_entry_in_loop[0].destroy()
                    remark_entry_in_loop[0] = None
                # 所有动作的延时都可以修改
                loop_action_tree.update_idletasks()
                bbox = loop_action_tree.bbox(row_id, column='#2')
                if not bbox:
                    return
                x, y, width, height = bbox
                act_type = action.get('type')
                if act_type == 'delay':
                    delay_var_in_loop[0].set(f"{action.get('delay', 1.0):.1f}")
                else:
                    delay_var_in_loop[0].set(f"{action.get('time', 0):.1f}")
                delay_entry_in_loop[0] = ttk.Entry(loop_action_tree, textvariable=delay_var_in_loop[0], width=10)
                delay_entry_in_loop[0].place(x=x, y=y, width=width, height=height)
                delay_entry_in_loop[0].focus_set()
                delay_entry_in_loop[0].select_range(0, tk.END)

                def save_loop_delay(event=None):
                    try:
                        new_delay = float(delay_var_in_loop[0].get())
                        act_type = loop_actions[action_idx].get('type')
                        if act_type == 'delay':
                            loop_actions[action_idx]['delay'] = new_delay
                        else:
                            loop_actions[action_idx]['time'] = new_delay
                    except ValueError:
                        pass
                    delay_entry_in_loop[0].destroy()
                    delay_entry_in_loop[0] = None
                    refresh_actions_list()

                delay_entry_in_loop[0].bind('<Return>', save_loop_delay)
                delay_entry_in_loop[0].bind('<Escape>', lambda e: delay_entry_in_loop[0].destroy() if delay_entry_in_loop[0] else None)
                delay_entry_in_loop[0].bind('<FocusOut>', save_loop_delay)
                return

            # 备注列编辑
            if column == "#3":
                # 隐藏延时编辑器
                if delay_entry_in_loop[0]:
                    delay_entry_in_loop[0].destroy()
                    delay_entry_in_loop[0] = None
                if remark_entry_in_loop[0]:
                    remark_entry_in_loop[0].destroy()
                loop_action_tree.update_idletasks()
                bbox = loop_action_tree.bbox(row_id, column='#3')
                if not bbox:
                    return
                x, y, width, height = bbox
                remark_var_in_loop[0].set(action.get('remark', ''))
                remark_entry_in_loop[0] = ttk.Entry(loop_action_tree, textvariable=remark_var_in_loop[0], width=20)
                remark_entry_in_loop[0].place(x=x, y=y, width=width, height=height)
                remark_entry_in_loop[0].focus_set()
                remark_entry_in_loop[0].select_range(0, tk.END)

                def save_loop_remark(event=None):
                    loop_actions[action_idx]['remark'] = remark_var_in_loop[0].get()
                    remark_entry_in_loop[0].destroy()
                    remark_entry_in_loop[0] = None
                    refresh_actions_list()

                remark_entry_in_loop[0].bind('<Return>', save_loop_remark)
                remark_entry_in_loop[0].bind('<Escape>', lambda e: remark_entry_in_loop[0].destroy() if remark_entry_in_loop[0] else None)
                remark_entry_in_loop[0].bind('<FocusOut>', save_loop_remark)

        def on_loop_double_click(event):
            """双击动作列：如果是可录制坐标的动作，重新录制坐标；截屏则编辑参数"""
            row_id = loop_action_tree.identify_row(event.y)
            if not row_id:
                return
            column = loop_action_tree.identify_column(event.x)
            if column != "#1":  # 非动作列不处理
                return
            try:
                action_idx = int(row_id)
            except ValueError:
                return
            if action_idx >= len(loop_actions):
                return
            action = loop_actions[action_idx]
            action_type = action.get('type')

            # 截屏动作 - 编辑参数
            if action_type == 'screenshot':
                # 隐藏所有编辑器
                if remark_entry_in_loop[0]:
                    remark_entry_in_loop[0].destroy()
                    remark_entry_in_loop[0] = None
                if delay_entry_in_loop[0]:
                    delay_entry_in_loop[0].destroy()
                    delay_entry_in_loop[0] = None
                # 打开编辑对话框
                edit_screenshot_in_loop(action)
                return

            # 只有有坐标的动作才能重新录制
            if action_type not in ['click', 'doubleclick', 'move', 'scroll']:
                return

            # 隐藏所有编辑器
            if remark_entry_in_loop[0]:
                remark_entry_in_loop[0].destroy()
                remark_entry_in_loop[0] = None
            if delay_entry_in_loop[0]:
                delay_entry_in_loop[0].destroy()
                delay_entry_in_loop[0] = None

            # 隐藏对话框，开始录制新坐标
            dialog.grab_release()
            dialog.withdraw()
            dialog.update()

            captured_pos = [None, None]

            def on_capture_click(x, y, button, pressed):
                if pressed:
                    captured_pos[0] = x
                    captured_pos[1] = y
                    return False

            listener = mouse.Listener(on_click=on_capture_click)
            listener.start()
            listener.join()

            if captured_pos[0] is not None:
                loop_actions[action_idx]['x'] = captured_pos[0]
                loop_actions[action_idx]['y'] = captured_pos[1]
                refresh_actions_list()

            dialog.deiconify()
            dialog.grab_set()
            dialog.update()

        def edit_screenshot_in_loop(action):
            """编辑循环内截屏动作"""
            edit_dialog = tk.Toplevel(dialog)
            edit_dialog.title("编辑截屏")
            edit_dialog.geometry("450x200")
            edit_dialog.transient(dialog)
            edit_dialog.attributes('-topmost', True)
            edit_dialog.grab_set()

            def on_edit_close():
                edit_dialog.destroy()

            edit_dialog.protocol("WM_DELETE_WINDOW", on_edit_close)

            main_frame_ss = ttk.Frame(edit_dialog, padding="10")
            main_frame_ss.pack(fill=tk.BOTH, expand=True)

            # 目录
            ttk.Label(main_frame_ss, text="目录:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
            ss_dir_var = tk.StringVar(value=action.get('directory', ''))
            ss_dir_entry = ttk.Entry(main_frame_ss, textvariable=ss_dir_var, width=30)
            ss_dir_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

            def select_ss_dir():
                dir_path = filedialog.askdirectory(initialdir=ss_dir_var.get(), title="选择截屏保存目录")
                if dir_path:
                    ss_dir_var.set(dir_path)

            ttk.Button(main_frame_ss, text="选择", command=select_ss_dir).grid(row=0, column=2, padx=5, pady=5)

            # 名称
            ttk.Label(main_frame_ss, text="名称:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
            ss_name_var = tk.StringVar(value=action.get('filename', 'screenshot'))
            ttk.Entry(main_frame_ss, textvariable=ss_name_var, width=15).grid(row=1, column=1, padx=5, pady=5, sticky="w")

            # 延时
            ttk.Label(main_frame_ss, text="延时:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
            ss_delay_var = tk.DoubleVar(value=action.get('delay', 0))
            ttk.Entry(main_frame_ss, textvariable=ss_delay_var, width=8).grid(row=1, column=3, padx=5, pady=5, sticky="w")

            # 命名方式
            ttk.Label(main_frame_ss, text="命名:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
            ss_naming_var = tk.StringVar(value=action.get('naming', 'timestamp'))
            ttk.Radiobutton(main_frame_ss, text="时间戳", variable=ss_naming_var, value="timestamp").grid(row=2, column=1, padx=5, pady=5, sticky="w")
            ttk.Radiobutton(main_frame_ss, text="递增序号", variable=ss_naming_var, value="increment").grid(row=2, column=2, padx=5, pady=5, sticky="w")

            def confirm_edit():
                action['directory'] = ss_dir_var.get()
                action['filename'] = ss_name_var.get()
                action['delay'] = ss_delay_var.get()
                action['naming'] = ss_naming_var.get()
                refresh_actions_list()
                edit_dialog.destroy()

            btn_frame = ttk.Frame(main_frame_ss)
            btn_frame.grid(row=3, column=0, columnspan=4, pady=10)
            ttk.Button(btn_frame, text="确定", command=confirm_edit).pack(side=tk.LEFT, padx=10)
            ttk.Button(btn_frame, text="取消", command=on_edit_close).pack(side=tk.LEFT, padx=10)

        loop_action_tree.bind("<Double-1>", on_loop_double_click)
        loop_action_tree.bind("<Button-1>", on_loop_action_click)

        actions_btn_frame = ttk.Frame(main_frame)
        actions_btn_frame.grid(row=7, column=1, padx=5, pady=5, sticky="w")

        # 用于重选动作的临时窗口引用
        select_hint_window = None

        def add_selected_actions():
            """添加当前选中的动作（直接添加）"""
            sel = self._get_selected_indices()
            import copy
            added_count = 0
            for item in sel:
                if isinstance(item, int) and 0 <= item < len(self.actions):
                    loop_actions.append(copy.deepcopy(self.actions[item]))
                    added_count += 1
            refresh_actions_list()
            if added_count > 0:
                edit_status_label.config(text=f"已添加 {added_count} 个动作", foreground="green")

        def start_reselect_actions():
            """开始重选动作模式 - 隐藏对话框，释放grab，显示提示"""
            # 释放grab让主界面可以交互
            dialog.grab_release()
            dialog.withdraw()
            dialog.update()

            # 创建提示窗口（不设置grab，让主界面可交互）
            select_hint_window = tk.Toplevel(self.root)
            select_hint_window.title("选择动作")
            select_hint_window.geometry("300x150")
            select_hint_window.attributes('-topmost', True)
            # 不设置 transient 和 grab，让用户可以在主界面操作

            ttk.Label(select_hint_window, text="请在主界面选择动作\n（支持Ctrl多选）\n完成后点击下方按钮确认",
                     font=('Arial', 10)).pack(pady=10)

            hint_status = ttk.Label(select_hint_window, text="当前选中: 0 个动作", foreground="blue")
            hint_status.pack(pady=5)

            def confirm_and_add():
                """确认添加并恢复对话框"""
                sel = self._get_selected_indices()
                import copy
                added_count = 0
                for item in sel:
                    if isinstance(item, int) and 0 <= item < len(self.actions):
                        loop_actions.append(copy.deepcopy(self.actions[item]))
                        added_count += 1
                refresh_actions_list()
                select_hint_window.destroy()
                dialog.deiconify()
                dialog.lift()
                dialog.attributes('-topmost', True)
                dialog.grab_set()  # 重新设置grab
                if added_count > 0:
                    edit_status_label.config(text=f"已添加 {added_count} 个动作", foreground="green")

            def cancel_reselect():
                """取消重选，恢复对话框"""
                select_hint_window.destroy()
                dialog.deiconify()
                dialog.lift()
                dialog.attributes('-topmost', True)
                dialog.grab_set()  # 重新设置grab

            btn_frame = ttk.Frame(select_hint_window)
            btn_frame.pack(pady=10)
            ttk.Button(btn_frame, text="确认添加", command=confirm_and_add).pack(side=tk.LEFT, padx=10)
            ttk.Button(btn_frame, text="取消", command=cancel_reselect).pack(side=tk.LEFT, padx=10)

            select_hint_window.protocol("WM_DELETE_WINDOW", cancel_reselect)

            # 更新选中计数
            def update_selection_count():
                try:
                    sel = self._get_selected_indices()
                    count = len([i for i in sel if isinstance(i, int)])
                    hint_status.config(text=f"当前选中: {count} 个动作")
                    select_hint_window.after(300, update_selection_count)
                except:
                    pass
            update_selection_count()

        def delete_action_from_loop():
            selection = loop_action_tree.selection()
            if selection:
                # 按索引倒序删除（避免删除后索引变化）
                indices = sorted([int(item) for item in selection], reverse=True)
                for idx in indices:
                    if 0 <= idx < len(loop_actions):
                        del loop_actions[idx]
                refresh_actions_list()

        def clear_loop_actions():
            loop_actions.clear()
            refresh_actions_list()

        def start_recording():
            """开始录制动作模式 - 保持对话框显示，录制动作直接显示在列表中"""
            # 保持对话框显示，只释放grab
            dialog.grab_release()
            dialog.update()

            # 更新动作列表显示
            def update_temp_actions_list():
                """更新循环内动作列表显示"""
                for item in loop_action_tree.get_children():
                    loop_action_tree.delete(item)
                for idx, act in enumerate(loop_actions):
                    desc = self._describe_action(act)
                    act_type = act.get('type')
                    if act_type == 'delay':
                        delay_text = f"{act.get('delay', 1.0):.1f}s"
                    else:
                        delay_text = f"{act.get('time', 0):.1f}s"
                    remark_text = act.get('remark', '')
                    loop_action_tree.insert("", tk.END, iid=str(idx), values=(desc, delay_text, remark_text))
                if loop_actions:
                    last_item = loop_action_tree.get_children()[-1]
                    loop_action_tree.see(last_item)

            # 录制状态
            temp_recording = {'is_recording': True}

            # 录制回调
            def on_recording_click(x, y, button, pressed):
                if not pressed:
                    return
                if not temp_recording['is_recording']:
                    return
                # 检查是否在主窗口内
                if self.is_in_window(x, y):
                    if self.is_in_ui_element(x, y):
                        pass  # 允许录制按钮点击
                    else:
                        return
                # 检查是否在输入循环对话框内
                try:
                    dialog_x = dialog.winfo_x()
                    dialog_y = dialog.winfo_y()
                    dialog_width = dialog.winfo_width()
                    dialog_height = dialog.winfo_height()
                    if (dialog_x <= x <= dialog_x + dialog_width and
                        dialog_y <= y <= dialog_y + dialog_height):
                        return  # 在对话框内不记录
                except Exception:
                    pass

                current_time = time.time() - self.record_start_time
                last_x, last_y = self.last_click_position
                distance = abs(x - last_x) + abs(y - last_y)
                is_double_click = (current_time - self.last_click_time < self.double_click_threshold and
                                  distance < self.double_click_distance_threshold)

                if is_double_click:
                    if loop_actions and loop_actions[-1]['type'] == 'click':
                        loop_actions.pop()
                    loop_actions.append({
                        'type': 'doubleclick', 'x': x, 'y': y,
                        'button': str(button), 'time': current_time
                    })
                else:
                    loop_actions.append({
                        'type': 'click', 'x': x, 'y': y,
                        'button': str(button), 'time': current_time
                    })
                self.last_click_time = current_time
                self.last_click_position = (x, y)
                loop_action_tree.after(0, update_temp_actions_list)

            def on_recording_scroll(x, y, dx, dy):
                if not temp_recording['is_recording']:
                    return
                if self.is_in_window(x, y) and self.is_in_ui_element(x, y):
                    return
                # 检查是否在输入循环对话框内
                try:
                    dialog_x = dialog.winfo_x()
                    dialog_y = dialog.winfo_y()
                    dialog_width = dialog.winfo_width()
                    dialog_height = dialog.winfo_height()
                    if (dialog_x <= x <= dialog_x + dialog_width and
                        dialog_y <= y <= dialog_y + dialog_height):
                        return  # 在对话框内不记录
                except Exception:
                    pass
                current_time = time.time() - self.record_start_time
                if current_time - self.last_scroll_time >= self.scroll_threshold:
                    loop_actions.append({
                        'type': 'scroll', 'x': x, 'y': y,
                        'dx': dx, 'dy': dy, 'time': current_time
                    })
                    self.last_scroll_time = current_time
                    loop_action_tree.after(0, update_temp_actions_list)

            def on_recording_key_down(key):
                try:
                    if temp_recording['is_recording']:
                        key_str = self.convert_key_name(key)
                        self.pressed_keys.add(key_str)
                except AttributeError:
                    pass

            def on_recording_key_up(key):
                try:
                    if temp_recording['is_recording']:
                        # 当焦点在程序内时，不录制键盘输入
                        if self._keyboard_recording_paused:
                            return
                        current_time = time.time() - self.record_start_time
                        key_str = self.convert_key_name(key)
                        if key_str in self.pressed_keys:
                            self.pressed_keys.remove(key_str)
                        current_keys = list(self.pressed_keys)
                        if current_keys:
                            loop_actions.append({
                                'type': 'keyboard', 'keys': current_keys + [key_str], 'time': current_time
                            })
                        else:
                            loop_actions.append({
                                'type': 'keyboard', 'keys': [key_str], 'time': current_time
                            })
                        loop_action_tree.after(0, update_temp_actions_list)
                except AttributeError:
                    pass

            # 开始录制
            self.record_start_time = time.time()
            self.currently_pressed_keys = set()
            self.pressed_keys = set()
            self.last_scroll_time = 0

            rec_mouse_listener = mouse.Listener(on_click=on_recording_click, on_scroll=on_recording_scroll, suppress=False)
            rec_mouse_listener.start()
            rec_keyboard_listener = keyboard.Listener(on_press=on_recording_key_down, on_release=on_recording_key_up, suppress=False)
            rec_keyboard_listener.start()

            # 更新按钮文字
            record_btn_widget = actions_btn_frame.winfo_children()[0]
            record_btn_widget.config(text="停止录制")

            def stop_recording_and_close():
                temp_recording['is_recording'] = False
                try:
                    rec_mouse_listener.stop()
                    rec_keyboard_listener.stop()
                    time.sleep(0.1)
                except Exception as e:
                    print(f"停止监听器失败: {e}")
                record_btn_widget.config(text="录制")
                dialog.grab_set()
                dialog.update()
                edit_status_label.config(text=f"录制完成 ({len(loop_actions)} 个动作)", foreground="green")

            # 绑定ESC键和录制按钮
            dialog.bind('<Escape>', lambda e: stop_recording_and_close())
            record_btn_widget.config(command=stop_recording_and_close)

        def insert_delay():
            """插入固定延时 - 弹出配置对话框"""
            delay_dialog = tk.Toplevel(dialog)
            delay_dialog.title("延时设置")
            delay_dialog.geometry("250x120")
            delay_dialog.transient(dialog)
            delay_dialog.attributes('-topmost', True)
            delay_dialog.grab_set()

            def on_dialog_close():
                delay_dialog.destroy()

            delay_dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

            main_frame_d = ttk.Frame(delay_dialog, padding="10")
            main_frame_d.pack(fill=tk.BOTH, expand=True)

            ttk.Label(main_frame_d, text="延时时间(秒):").pack(pady=5)
            delay_var = tk.DoubleVar(value=1.0)
            delay_entry = ttk.Entry(main_frame_d, textvariable=delay_var, width=10)
            delay_entry.pack(pady=5)
            delay_entry.focus_set()

            def confirm_delay():
                delay_action = {
                    'type': 'delay',
                    'delay': delay_var.get()
                }
                loop_actions.append(delay_action)
                refresh_actions_list()
                edit_status_label.config(text=f"已插入延时 ({delay_var.get()}秒)", foreground="green")
                delay_dialog.destroy()

            btn_frame = ttk.Frame(main_frame_d)
            btn_frame.pack(pady=10)
            ttk.Button(btn_frame, text="确定", command=confirm_delay).pack(side=tk.LEFT, padx=10)
            ttk.Button(btn_frame, text="取消", command=on_dialog_close).pack(side=tk.LEFT, padx=10)

            delay_entry.bind('<Return>', lambda e: confirm_delay())

        def insert_screenshot_action():
            """插入截屏动作 - 弹出配置对话框"""
            screenshot_dialog = tk.Toplevel(dialog)
            screenshot_dialog.title("截屏设置")
            screenshot_dialog.geometry("450x200")
            screenshot_dialog.transient(dialog)
            screenshot_dialog.attributes('-topmost', True)
            screenshot_dialog.grab_set()

            def on_dialog_close():
                screenshot_dialog.destroy()

            screenshot_dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

            main_frame_ss = ttk.Frame(screenshot_dialog, padding="10")
            main_frame_ss.pack(fill=tk.BOTH, expand=True)

            # 目录
            ttk.Label(main_frame_ss, text="目录:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
            ss_dir_var = tk.StringVar(value=self.screenshot_dir_var.get())
            ss_dir_entry = ttk.Entry(main_frame_ss, textvariable=ss_dir_var, width=30)
            ss_dir_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

            def select_ss_dir():
                dir_path = filedialog.askdirectory(initialdir=ss_dir_var.get(), title="选择截屏保存目录")
                if dir_path:
                    ss_dir_var.set(dir_path)

            ttk.Button(main_frame_ss, text="选择", command=select_ss_dir).grid(row=0, column=2, padx=5, pady=5)

            # 名称
            ttk.Label(main_frame_ss, text="名称:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
            ss_name_var = tk.StringVar(value=self.screenshot_name_var.get())
            ss_name_entry = ttk.Entry(main_frame_ss, textvariable=ss_name_var, width=15)
            ss_name_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

            # 延时
            ttk.Label(main_frame_ss, text="延时:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
            ss_delay_var = tk.DoubleVar(value=self.screenshot_delay_var.get())
            ttk.Entry(main_frame_ss, textvariable=ss_delay_var, width=8).grid(row=1, column=3, padx=5, pady=5, sticky="w")

            # 命名方式
            ttk.Label(main_frame_ss, text="命名:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
            ss_naming_var = tk.StringVar(value="timestamp")
            ttk.Radiobutton(main_frame_ss, text="时间戳", variable=ss_naming_var, value="timestamp").grid(row=2, column=1, padx=5, pady=5, sticky="w")
            ttk.Radiobutton(main_frame_ss, text="递增序号", variable=ss_naming_var, value="increment").grid(row=2, column=2, padx=5, pady=5, sticky="w")

            def confirm_screenshot():
                screenshot_action = {
                    'type': 'screenshot',
                    'naming': ss_naming_var.get(),
                    'directory': ss_dir_var.get(),
                    'filename': ss_name_var.get(),
                    'delay': ss_delay_var.get()
                }
                loop_actions.append(screenshot_action)
                refresh_actions_list()
                edit_status_label.config(text="已插入截屏动作", foreground="green")
                screenshot_dialog.destroy()

            btn_frame = ttk.Frame(main_frame_ss)
            btn_frame.grid(row=3, column=0, columnspan=4, pady=10)
            ttk.Button(btn_frame, text="确定", command=confirm_screenshot).pack(side=tk.LEFT, padx=10)
            ttk.Button(btn_frame, text="取消", command=on_dialog_close).pack(side=tk.LEFT, padx=10)

        ttk.Button(actions_btn_frame, text="录制", command=start_recording).pack(side=tk.LEFT, padx=2)
        ttk.Button(actions_btn_frame, text="插入延时", command=insert_delay).pack(side=tk.LEFT, padx=2)
        ttk.Button(actions_btn_frame, text="插入截屏", command=insert_screenshot_action).pack(side=tk.LEFT, padx=2)
        ttk.Button(actions_btn_frame, text="删除选中", command=delete_action_from_loop).pack(side=tk.LEFT, padx=2)
        ttk.Button(actions_btn_frame, text="清空", command=clear_loop_actions).pack(side=tk.LEFT, padx=2)

        # 状态标签
        edit_status_label = ttk.Label(main_frame, text="", foreground="green")
        edit_status_label.grid(row=8, column=0, columnspan=2, pady=2)

        def update_type_ui(type_name):
            fixed_frame.grid_remove()
            range_frame.grid_remove()
            list_frame.grid_remove()
            template_entry.config(state='normal' if type_name != 'list' else 'disabled')
            if type_name == "range":
                range_frame.grid()
            else:
                list_frame.grid()

        # 初始化显示
        update_type_ui("range")

        # 从文件加载列表
        def load_list_from_file():
            filepath = filedialog.askopenfilename(
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
            )
            if filepath:
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        content = f.read()
                        list_text.delete(tk.END, '1.0')
                        list_text.insert(tk.END, content)
                    list_file_var.set(filepath)
                except Exception as e:
                    messagebox.showerror("错误", f"加载失败: {str(e)}")

        record_btn.config(command=lambda: self._record_position_with_dialog(dialog, x_var, y_var, edit_status_label))

        def confirm():
            try:
                loop_type = type_var.get()
                template = template_var.get()

                target_x = None
                target_y = None
                if x_var.get().strip():
                    target_x = int(x_var.get())
                if y_var.get().strip():
                    target_y = int(y_var.get())

                clear_text = clear_text_var.get()

                # 调试输出
                print(f"[保存input_loop] loop_actions数量={len(loop_actions)}")
                if loop_actions:
                    print(f"[保存input_loop] 循环内动作: {[a.get('type') for a in loop_actions]}")

                action = {
                    'type': 'input_loop',
                    'loop_type': loop_type,
                    'input_template': template,
                    'target_x': target_x,
                    'target_y': target_y,
                    'clear_text': clear_text,
                    'interval_mode': interval_mode_var.get(),
                    'loop_actions': loop_actions.copy(),
                    'time': 0
                }

                # 保存间隔参数
                interval_mode = interval_mode_var.get()
                if interval_mode == 'fixed':
                    action['loop_interval'] = interval_var.get()
                elif interval_mode == 'range':
                    if range_sub_type_var.get() == "arithmetic":
                        action['loop_interval'] = {
                            'type': 'range',
                            'sub_type': 'arithmetic',
                            'start': interval_start_var.get(),
                            'step': interval_step_var.get(),
                            'end': interval_end_var.get()
                        }
                    else:
                        action['loop_interval'] = {
                            'type': 'range',
                            'sub_type': 'geometric',
                            'start': interval_geo_start_var.get(),
                            'ratio': interval_ratio_var.get(),
                            'steps': interval_steps_var.get()
                        }
                elif interval_mode == 'list':
                    try:
                        values = [float(v.strip()) for v in interval_list_var.get().split(',')]
                        action['loop_interval'] = {'type': 'list', 'values': values}
                    except ValueError:
                        raise ValueError("列表格式无效，请使用逗号分隔的数字")
                else:  # random
                    action['loop_interval'] = {'type': 'random', 'min': interval_random_min_var.get(), 'max': interval_random_max_var.get()}

                if loop_type == 'range':
                    start = start_var.get()
                    step = step_var.get()
                    end = end_var.get()
                    if step == 0:
                        raise ValueError("步距不能为0")
                    action['start'] = start
                    action['step'] = step
                    action['end'] = end

                else:  # list
                    text_content = list_text.get('1.0', tk.END).strip()
                    if not text_content:
                        raise ValueError("文本列表不能为空")
                    action['text_list'] = [line.strip() for line in text_content.split('\n') if line.strip()]
                    action['list_file'] = list_file_var.get()

                self.actions.insert(insert_pos, action)
                self._update_action_list(select_index=insert_pos)
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(main_frame, text="确定", command=confirm).grid(row=9, column=0, columnspan=2, pady=10)

    def _shift_action_times(self, start_index, delta):
        """整体平移从指定索引开始的动作时间，保持相对间隔不变"""
        if delta == 0 or start_index >= len(self.actions):
            return

        for i in range(start_index, len(self.actions)):
            if 'time' in self.actions[i]:
                self.actions[i]['time'] = max(0, self.actions[i]['time'] + delta)

    def _shift_loop_action_times(self, loop_actions, start_index, delta):
        """整体平移循环组内从指定索引开始的动作时间，保持相对间隔不变"""
        if delta == 0 or start_index >= len(loop_actions):
            return

        for i in range(start_index, len(loop_actions)):
            if 'time' in loop_actions[i]:
                loop_actions[i]['time'] = max(0, loop_actions[i]['time'] + delta)

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
            is_nested = row_id.startswith('loop_') or row_id.startswith('input_loop_')
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
                elif action_type == 'traverse_input':
                    self._edit_nested_traverse_input_action(action, row_id)
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
            # 输入动作 - 编辑输入内容
            elif action_type == 'input':
                self._edit_input_action(action, index)
            # 延时动作 - 编辑延时参数
            elif action_type in ['random_delay', 'multiply_delay', 'arithmetic_delay']:
                self.edit_action_time(event)
            # 输入动作 - 编辑输入
            elif action_type == 'variable_input':
                self._edit_variable_input_action(action, index)
            # 遍历输入 - 编辑遍历参数
            elif action_type == 'traverse_input':
                self._edit_traverse_input_action(action, index)
            # 输入循环 - 编辑参数
            elif action_type == 'input_loop':
                self._edit_input_loop_action(action, index)
            # 循环组 - 显示编辑
            elif action_type == 'loop_group':
                self._edit_top_loop_group(action, index)
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
        if 'display_delay' in action:
            relative_delay = action['display_delay']
        else:
            prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
            relative_delay = round(action.get('time', 0) - prev_time, 1)

        ttk.Label(frame, text="相对延时(秒):").grid(row=2, column=0, padx=5, pady=10, sticky="e")
        delay_var = tk.DoubleVar(value=round(relative_delay, 1))
        ttk.Spinbox(frame, width=13, textvariable=delay_var, from_=0, to=999, increment=0.1, state='normal').grid(row=2, column=1, padx=5, pady=10, sticky="w")

        # 状态标签
        status_label = ttk.Label(frame, text="", foreground="blue")
        status_label.grid(row=3, column=0, columnspan=2, pady=5)

        record_btn.config(command=lambda: self._record_position_with_pyautogui(dialog, x_var, y_var, status_label))

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
                action['display_delay'] = delay_var.get()
                self._shift_action_times(index + 1, delta)

                self._update_action_list(select_index=index)
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=4, column=0, columnspan=2, pady=10)

    def _edit_input_action(self, action, index):
        """编辑输入动作"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑输入动作")
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

        # 输入内容
        ttk.Label(frame, text="输入内容：").grid(row=0, column=0, padx=5, pady=10, sticky="e")
        input_text_var = tk.StringVar(value=action.get('text', ''))
        ttk.Entry(frame, width=30, textvariable=input_text_var).grid(row=0, column=1, padx=5, pady=10, sticky="w")

        # 输入坐标
        ttk.Label(frame, text="输入坐标：").grid(row=1, column=0, padx=5, pady=10, sticky="e")
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

        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        # 相对延时
        if 'display_delay' in action:
            relative_delay = action['display_delay']
        else:
            prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
            relative_delay = round(action.get('time', 0) - prev_time, 1)

        ttk.Label(frame, text="相对延时(秒):").grid(row=2, column=0, padx=5, pady=10, sticky="e")
        delay_var = tk.DoubleVar(value=round(relative_delay, 1))
        ttk.Spinbox(frame, width=13, textvariable=delay_var, from_=0, to=999, increment=0.1, state='normal').grid(row=2, column=1, padx=5, pady=10, sticky="w")

        # 状态标签
        status_label = ttk.Label(frame, text="", foreground="blue")
        status_label.grid(row=3, column=0, columnspan=2, pady=5)

        record_btn.config(command=lambda: self._record_position_with_pyautogui(dialog, x_var, y_var, status_label))

        def confirm():
            try:
                target_x = x_var.get().strip()
                target_y = y_var.get().strip()
                input_text = input_text_var.get().strip()

                if not input_text:
                    raise ValueError("输入内容不能为空")
                if not target_x or not target_y:
                    raise ValueError("请输入有效的坐标")

                target_x = int(target_x)
                target_y = int(target_y)

                # 更新动作
                action['text'] = input_text
                action['x'] = target_x
                action['y'] = target_y

                # 计算新的绝对时间
                prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
                new_absolute_time = prev_time + delay_var.get()
                old_time = action.get('time', 0)
                delta = new_absolute_time - old_time
                action['time'] = new_absolute_time
                action['display_delay'] = delay_var.get()
                self._shift_action_times(index + 1, delta)

                self._update_action_list(select_index=index)
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=4, column=0, columnspan=2, pady=10)

    def _edit_traverse_input_action(self, action, index):
        """编辑遍历输入动作"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑遍历输入")
        dialog.geometry("500x340")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        dialog.grab_set()

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 遍历类型选择
        ttk.Label(main_frame, text="遍历类型:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        type_var = tk.StringVar(value=action.get('traverse_type', 'list'))
        type_frame = ttk.Frame(main_frame)
        type_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(type_frame, text="固定文本", variable=type_var, value="fixed_text",
                       command=lambda: update_type_ui("fixed_text")).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="列表遍历", variable=type_var, value="list",
                       command=lambda: update_type_ui("list")).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="序列", variable=type_var, value="sequence",
                       command=lambda: update_type_ui("sequence")).pack(side=tk.LEFT, padx=5)

        # 点击位置
        ttk.Label(main_frame, text="点击位置:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        pos_frame = ttk.Frame(main_frame)
        pos_frame.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        x_var = tk.StringVar(value=str(action.get('x', '')))
        y_var = tk.StringVar(value=str(action.get('y', '')))
        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        ttk.Entry(pos_frame, width=8, textvariable=x_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        ttk.Entry(pos_frame, width=8, textvariable=y_var).pack(side=tk.LEFT, padx=2)

        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        # 参数框架（动态切换）
        params_frame = ttk.Frame(main_frame)
        params_frame.grid(row=2, column=0, columnspan=2, sticky="w", pady=10)

        # 列表遍历参数
        list_frame = ttk.Frame(params_frame)
        list_row1 = ttk.Frame(list_frame)
        list_row1.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(list_row1, text="遍历列表:").pack(side=tk.LEFT, padx=2)
        list_text = tk.Text(list_row1, width=25, height=6)
        list_text.pack(side=tk.LEFT, padx=2)
        # 填充现有值
        traverse_values = action.get('traverse_values', [])
        if traverse_values:
            list_text.insert(tk.END, '\n'.join(str(v) for v in traverse_values))
        ttk.Label(list_frame, text="(每行一个值)").pack(side=tk.TOP, padx=2, pady=(2, 0))

        # 序列参数
        seq_frame = ttk.Frame(params_frame)
        seq_row1 = ttk.Frame(seq_frame)
        seq_row1.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(seq_row1, text="起始值:").pack(side=tk.LEFT, padx=2)
        start_var = tk.DoubleVar(value=action.get('start', 0))
        ttk.Entry(seq_row1, width=8, textvariable=start_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(seq_row1, text="步距:").pack(side=tk.LEFT, padx=2)
        step_var = tk.DoubleVar(value=action.get('step', 1))
        ttk.Entry(seq_row1, width=8, textvariable=step_var).pack(side=tk.LEFT, padx=2)
        seq_row2 = ttk.Frame(seq_frame)
        seq_row2.pack(side=tk.TOP, fill=tk.X, pady=(5, 0))
        ttk.Label(seq_row2, text="模板:").pack(side=tk.LEFT, padx=2)
        seq_template_var = tk.StringVar(value=action.get('template', 'text{n}'))
        ttk.Entry(seq_row2, width=15, textvariable=seq_template_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(seq_row2, text="(用{n}替换)").pack(side=tk.LEFT, padx=2)

        # 固定文本参数
        fixed_text_frame = ttk.Frame(params_frame)
        ttk.Label(fixed_text_frame, text="输入文本:").pack(side=tk.LEFT, padx=2)
        fixed_text_var = tk.StringVar(value=action.get('fixed_text', ''))
        ttk.Entry(fixed_text_frame, width=25, textvariable=fixed_text_var).pack(side=tk.LEFT, padx=2)

        def update_type_ui(mode):
            list_frame.grid_forget()
            seq_frame.grid_forget()
            fixed_text_frame.grid_forget()
            if mode == "list":
                list_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
            elif mode == "sequence":
                seq_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
            else:  # fixed_text
                fixed_text_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)

        # 初始化UI
        current_type = action.get('traverse_type', 'list')
        update_type_ui(current_type)

        # 录制按钮功能
        def start_record():
            dialog.withdraw()
            dialog.update()
            captured_pos = [None, None]

            def on_click(x, y, button, pressed):
                if pressed:
                    captured_pos[0] = x
                    captured_pos[1] = y
                    return False

            listener = mouse.Listener(on_click=on_click)
            listener.start()
            listener.join()

            if captured_pos[0] is not None:
                x_var.set(str(captured_pos[0]))
                y_var.set(str(captured_pos[1]))
            dialog.deiconify()

        record_btn.config(command=start_record)

        # 相对延时
        if 'display_delay' in action:
            relative_delay = action['display_delay']
        else:
            prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
            relative_delay = round(action.get('time', 0) - prev_time, 1)

        ttk.Label(main_frame, text="相对延时(秒):").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        delay_var = tk.DoubleVar(value=round(relative_delay, 1))
        ttk.Spinbox(main_frame, width=13, textvariable=delay_var, from_=0, to=999, increment=0.1, state='normal').grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=15)

        def confirm():
            try:
                x = int(x_var.get()) if x_var.get() else 0
                y = int(y_var.get()) if y_var.get() else 0

                # 获取遍历值列表
                traverse_type = type_var.get()
                values = []
                start = 0
                step = 1
                fixed_text = ""
                if traverse_type == "list":
                    text_content = list_text.get("1.0", tk.END).strip()
                    if not text_content:
                        raise ValueError("遍历列表不能为空")
                    values = [line.strip() for line in text_content.split('\n') if line.strip()]
                elif traverse_type == "sequence":
                    start = start_var.get()
                    step = step_var.get()
                    seq_template = seq_template_var.get()
                else:  # fixed_text
                    fixed_text = fixed_text_var.get()
                    if not fixed_text:
                        raise ValueError("输入文本不能为空")

                # 更新动作
                action['traverse_type'] = traverse_type
                action['x'] = x
                action['y'] = y
                action['traverse_values'] = values if traverse_type == "list" else []
                action['start'] = start if traverse_type == "sequence" else None
                action['step'] = step if traverse_type == "sequence" else None
                action['template'] = seq_template if traverse_type == "sequence" else None
                action['fixed_text'] = fixed_text if traverse_type == "fixed_text" else None

                # 计算新的绝对时间
                prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
                new_absolute_time = prev_time + delay_var.get()
                old_time = action.get('time', 0)
                delta = new_absolute_time - old_time
                action['time'] = new_absolute_time
                action['display_delay'] = delay_var.get()
                self._shift_action_times(index + 1, delta)

                self._update_action_list(select_index=index)
                self.in_dialog_operation = False
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(btn_frame, text="确定", command=confirm).pack(side=tk.LEFT, padx=20)
        ttk.Button(btn_frame, text="取消", command=on_dialog_close).pack(side=tk.LEFT, padx=5)

    def _edit_screenshot_action(self, action, index):
        """编辑截屏动作"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑截屏")
        dialog.geometry("560x220")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        dialog.grab_set()

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 目录
        ttk.Label(main_frame, text="目录:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        dir_var = tk.StringVar(value=action.get('directory', ''))
        dir_entry = ttk.Entry(main_frame, textvariable=dir_var, width=30)
        dir_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        def select_dir():
            dir_path = filedialog.askdirectory(initialdir=dir_var.get(), title="选择截屏保存目录")
            if dir_path:
                dir_var.set(dir_path)

        ttk.Button(main_frame, text="选择", command=select_dir).grid(row=0, column=2, padx=5, pady=5)

        # 名称
        ttk.Label(main_frame, text="名称:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        name_var = tk.StringVar(value=action.get('filename', 'screenshot'))
        ttk.Entry(main_frame, textvariable=name_var, width=15).grid(row=1, column=1, padx=0, pady=5, sticky="w")

        # 延时
        delay_frame = ttk.Frame(main_frame)
        delay_frame.grid(row=1, column=2, padx=0, pady=5, sticky="w")
        ttk.Label(delay_frame, text="延时:").pack(side=tk.LEFT, padx=0)
        delay_var = tk.DoubleVar(value=action.get('delay', 0))
        ttk.Entry(delay_frame, textvariable=delay_var, width=8).pack(side=tk.LEFT, padx=0)

        # 命名方式
        ttk.Label(main_frame, text="命名:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        naming_var = tk.StringVar(value=action.get('naming', 'timestamp'))
        ttk.Radiobutton(main_frame, text="时间戳", variable=naming_var, value="timestamp").grid(row=2, column=1, padx=2, pady=5, sticky="w")
        ttk.Radiobutton(main_frame, text="递增序号", variable=naming_var, value="increment").grid(row=2, column=2, padx=2, pady=5, sticky="w")

        def confirm():
            action['directory'] = dir_var.get()
            action['filename'] = name_var.get()
            action['delay'] = delay_var.get()
            action['naming'] = naming_var.get()
            self._update_action_list(select_index=index)
            self.in_dialog_operation = False
            dialog.destroy()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, columnspan=4, pady=10)
        ttk.Button(btn_frame, text="确定", command=confirm).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_dialog_close).pack(side=tk.LEFT, padx=10)

    def _edit_input_loop_action(self, action, index):
        """编辑输入循环动作"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑输入循环")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        dialog.grab_set()

        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        loop_type = action.get('loop_type', 'fixed')

        # 循环类型显示
        type_names = {'fixed': '固定重复', 'range': '数值范围', 'list': '列表遍历'}
        ttk.Label(main_frame, text=f"循环类型: {type_names.get(loop_type, loop_type)}").grid(row=0, column=0, columnspan=2, pady=5, sticky="w")

        # 输入模板
        ttk.Label(main_frame, text="输入模板:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        template_var = tk.StringVar(value=action.get('input_template', 'text{n}'))
        template_entry = ttk.Entry(main_frame, width=30, textvariable=template_var)
        template_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        if loop_type == 'list':
            template_entry.config(state='disabled')

        # 点击位置
        ttk.Label(main_frame, text="点击位置:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        pos_frame = ttk.Frame(main_frame)
        pos_frame.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        x_var = tk.StringVar(value=str(action.get('target_x', '')) if action.get('target_x') else '')
        y_var = tk.StringVar(value=str(action.get('target_y', '')) if action.get('target_y') else '')
        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        ttk.Entry(pos_frame, width=8, textvariable=x_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        ttk.Entry(pos_frame, width=8, textvariable=y_var).pack(side=tk.LEFT, padx=2)

        # 循环参数（根据类型）
        if loop_type == 'fixed':
            ttk.Label(main_frame, text="重复次数:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
            loop_count_var = tk.IntVar(value=action.get('loop_count', 3))
            ttk.Entry(main_frame, width=10, textvariable=loop_count_var).grid(row=3, column=1, padx=5, pady=5, sticky="w")
        elif loop_type == 'range':
            range_frame = ttk.Frame(main_frame)
            range_frame.grid(row=3, column=0, columnspan=2, pady=5, sticky="w")
            ttk.Label(range_frame, text="起始值:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
            start_var = tk.IntVar(value=action.get('start', 0))
            ttk.Entry(range_frame, width=8, textvariable=start_var).grid(row=0, column=1, padx=5, pady=5, sticky="w")
            ttk.Label(range_frame, text="步距:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
            step_var = tk.IntVar(value=action.get('step', 1))
            ttk.Entry(range_frame, width=8, textvariable=step_var).grid(row=0, column=3, padx=5, pady=5, sticky="w")
            ttk.Label(range_frame, text="结束值:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
            end_var = tk.IntVar(value=action.get('end', 10))
            ttk.Entry(range_frame, width=8, textvariable=end_var).grid(row=0, column=5, padx=5, pady=5, sticky="w")
        else:  # list
            ttk.Label(main_frame, text="文本列表:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
            list_frame = ttk.Frame(main_frame)
            list_frame.grid(row=3, column=1, padx=5, pady=5, sticky="w")
            list_text = tk.Text(list_frame, width=30, height=5)
            list_text.pack(side=tk.LEFT)
            text_list = action.get('text_list', [])
            list_text.insert(tk.END, '\n'.join(text_list))

        # 间隔模式选择
        interval_mode_var = tk.StringVar(value=action.get('interval_mode', 'fixed'))
        interval_params = action.get('loop_interval', 1.0)

        interval_mode_frame = ttk.Frame(main_frame)
        interval_mode_frame.grid(row=4, column=0, columnspan=2, pady=5, sticky="w")
        ttk.Label(interval_mode_frame, text="间隔模式:").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(interval_mode_frame, text="固定", variable=interval_mode_var, value="fixed",
                        command=lambda: update_interval_ui("fixed")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(interval_mode_frame, text="范围", variable=interval_mode_var, value="range",
                        command=lambda: update_interval_ui("range")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(interval_mode_frame, text="列表", variable=interval_mode_var, value="list",
                        command=lambda: update_interval_ui("list")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(interval_mode_frame, text="随机", variable=interval_mode_var, value="random",
                        command=lambda: update_interval_ui("random")).pack(side=tk.LEFT, padx=2)

        # 间隔参数框架
        interval_params_frame = ttk.Frame(main_frame)
        interval_params_frame.grid(row=5, column=0, columnspan=2, pady=5, sticky="w")

        # 固定间隔
        fixed_interval_frame = ttk.Frame(interval_params_frame)
        ttk.Label(fixed_interval_frame, text="间隔(秒):").pack(side=tk.LEFT, padx=2)
        if isinstance(interval_params, (int, float)):
            interval_var = tk.DoubleVar(value=interval_params)
        else:
            interval_var = tk.DoubleVar(value=interval_params.get('value', 1.0) if isinstance(interval_params, dict) else 1.0)
        ttk.Entry(fixed_interval_frame, width=10, textvariable=interval_var).pack(side=tk.LEFT, padx=2)

        # 范围间隔（等差/等比）
        range_interval_frame = ttk.Frame(interval_params_frame)
        range_sub_type_var = tk.StringVar(value="geometric" if (isinstance(interval_params, dict) and interval_params.get('sub_type') == 'geometric') else "arithmetic")

        # Radiobutton 放在外层，始终可见
        ttk.Radiobutton(range_interval_frame, text="等差", variable=range_sub_type_var, value="arithmetic",
                        command=lambda: update_range_sub_ui()).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(range_interval_frame, text="等比", variable=range_sub_type_var, value="geometric",
                        command=lambda: update_range_sub_ui()).pack(side=tk.LEFT, padx=2)

        # 等差子模式 - 参数区域
        arithmetic_frame = ttk.Frame(range_interval_frame)
        ttk.Label(arithmetic_frame, text="起始:").pack(side=tk.LEFT, padx=2)
        if isinstance(interval_params, dict) and interval_params.get('sub_type') != 'geometric':
            interval_start_var = tk.DoubleVar(value=interval_params.get('start', 1.0))
            interval_step_var = tk.DoubleVar(value=interval_params.get('step', 0.1))
            interval_end_var = tk.DoubleVar(value=interval_params.get('end', 2.0))
        else:
            interval_start_var = tk.DoubleVar(value=1.0)
            interval_step_var = tk.DoubleVar(value=0.1)
            interval_end_var = tk.DoubleVar(value=2.0)
        ttk.Entry(arithmetic_frame, width=8, textvariable=interval_start_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(arithmetic_frame, text="步距:").pack(side=tk.LEFT, padx=2)
        ttk.Entry(arithmetic_frame, width=8, textvariable=interval_step_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(arithmetic_frame, text="上限:").pack(side=tk.LEFT, padx=2)
        ttk.Entry(arithmetic_frame, width=8, textvariable=interval_end_var).pack(side=tk.LEFT, padx=2)

        # 等比子模式 - 参数区域
        geometric_frame = ttk.Frame(range_interval_frame)
        ttk.Label(geometric_frame, text="起始:").pack(side=tk.LEFT, padx=2)
        if isinstance(interval_params, dict) and interval_params.get('sub_type') == 'geometric':
            interval_geo_start_var = tk.DoubleVar(value=interval_params.get('start', 1.0))
            interval_ratio_var = tk.DoubleVar(value=interval_params.get('ratio', 1.5))
            interval_steps_var = tk.IntVar(value=interval_params.get('steps', 10))
        else:
            interval_geo_start_var = tk.DoubleVar(value=1.0)
            interval_ratio_var = tk.DoubleVar(value=1.5)
            interval_steps_var = tk.IntVar(value=10)
        ttk.Entry(geometric_frame, width=8, textvariable=interval_geo_start_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(geometric_frame, text="等比系数:").pack(side=tk.LEFT, padx=2)
        ttk.Entry(geometric_frame, width=8, textvariable=interval_ratio_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(geometric_frame, text="步数:").pack(side=tk.LEFT, padx=2)
        ttk.Entry(geometric_frame, width=8, textvariable=interval_steps_var).pack(side=tk.LEFT, padx=2)

        def update_range_sub_ui():
            arithmetic_frame.pack_forget()
            geometric_frame.pack_forget()
            if range_sub_type_var.get() == "arithmetic":
                arithmetic_frame.pack(side=tk.LEFT, padx=2)
            else:
                geometric_frame.pack(side=tk.LEFT, padx=2)

        update_range_sub_ui()

        # 列表间隔
        list_interval_frame = ttk.Frame(interval_params_frame)
        ttk.Label(list_interval_frame, text="列表(逗号分隔):").pack(side=tk.LEFT, padx=2)
        if isinstance(interval_params, dict):
            interval_list_values = interval_params.get('values', [0.5, 1.0, 1.5, 2.0])
            interval_list_var = tk.StringVar(value=','.join(str(v) for v in interval_list_values))
        else:
            interval_list_var = tk.StringVar(value="0.5,1.0,1.5,2.0")
        ttk.Entry(list_interval_frame, width=20, textvariable=interval_list_var).pack(side=tk.LEFT, padx=2)

        # 随机间隔
        random_interval_frame = ttk.Frame(interval_params_frame)
        ttk.Label(random_interval_frame, text="最小:").pack(side=tk.LEFT, padx=2)
        if isinstance(interval_params, dict):
            interval_random_min_var = tk.DoubleVar(value=interval_params.get('min', 0.5))
        else:
            interval_random_min_var = tk.DoubleVar(value=0.5)
        ttk.Entry(random_interval_frame, width=8, textvariable=interval_random_min_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(random_interval_frame, text="最大:").pack(side=tk.LEFT, padx=2)
        if isinstance(interval_params, dict):
            interval_random_max_var = tk.DoubleVar(value=interval_params.get('max', 2.0))
        else:
            interval_random_max_var = tk.DoubleVar(value=2.0)
        ttk.Entry(random_interval_frame, width=8, textvariable=interval_random_max_var).pack(side=tk.LEFT, padx=2)

        def update_interval_ui(mode):
            fixed_interval_frame.pack_forget()
            range_interval_frame.pack_forget()
            list_interval_frame.pack_forget()
            random_interval_frame.pack_forget()
            if mode == "fixed":
                fixed_interval_frame.pack(side=tk.LEFT, padx=2)
            elif mode == "range":
                range_interval_frame.pack(side=tk.LEFT, padx=2)
                update_range_sub_ui()
            elif mode == "list":
                list_interval_frame.pack(side=tk.LEFT, padx=2)
            else:  # random
                random_interval_frame.pack(side=tk.LEFT, padx=2)

        update_interval_ui(interval_mode_var.get())

        # 清空选项
        clear_text_var = tk.BooleanVar(value=action.get('clear_text', True))
        ttk.Checkbutton(main_frame, text="输入前清空文本框", variable=clear_text_var).grid(row=6, column=0, columnspan=2, pady=5, sticky="w")

        # 循环内动作列表
        ttk.Label(main_frame, text="循环内动作:").grid(row=7, column=0, padx=5, pady=5, sticky="e")
        actions_frame = ttk.Frame(main_frame)
        actions_frame.grid(row=7, column=1, padx=5, pady=5, sticky="ew")

        actions_listbox = tk.Listbox(actions_frame, width=40, height=6)
        actions_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        actions_scrollbar = ttk.Scrollbar(actions_frame, orient=tk.VERTICAL, command=actions_listbox.yview)
        actions_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        actions_listbox.configure(yscrollcommand=actions_scrollbar.set)

        loop_actions = action.get('loop_actions', []).copy()
        for i, act in enumerate(loop_actions):
            desc = self._describe_action(act)
            actions_listbox.insert(tk.END, f"{i+1}. {desc}")

        actions_btn_frame = ttk.Frame(main_frame)
        actions_btn_frame.grid(row=8, column=1, padx=5, pady=5, sticky="w")

        # 状态标签（用于显示添加结果）
        edit_status_label = ttk.Label(main_frame, text="", foreground="green")
        edit_status_label.grid(row=8, column=0, columnspan=2, pady=2)

        def refresh_list():
            actions_listbox.delete(0, tk.END)
            for i, act in enumerate(loop_actions):
                actions_listbox.insert(tk.END, f"{i+1}. {self._describe_action(act)}")

        def add_selected_actions():
            """添加当前选中的动作（直接添加）"""
            sel = self._get_selected_indices()
            import copy
            added_count = 0
            for item in sel:
                if isinstance(item, int) and 0 <= item < len(self.actions):
                    loop_actions.append(copy.deepcopy(self.actions[item]))
                    added_count += 1
            refresh_list()
            if added_count > 0:
                edit_status_label.config(text=f"已添加 {added_count} 个动作")

        def start_reselect_actions():
            """开始重选动作模式 - 隐藏对话框，释放grab，显示提示"""
            # 释放grab让主界面可以交互
            dialog.grab_release()
            dialog.withdraw()
            dialog.update()

            # 创建提示窗口（不设置grab，让主界面可交互）
            hint_win = tk.Toplevel(self.root)
            hint_win.title("选择动作")
            hint_win.geometry("300x150")
            hint_win.attributes('-topmost', True)
            # 不设置 transient 和 grab，让用户可以在主界面操作

            ttk.Label(hint_win, text="请在主界面选择动作后\n点击下方按钮确认添加\n(可按住Ctrl多选)",
                     font=('Arial', 10)).pack(pady=15)

            hint_status = ttk.Label(hint_win, text="", foreground="blue")
            hint_status.pack(pady=5)

            def confirm_and_add():
                """确认添加并恢复对话框"""
                sel = self._get_selected_indices()
                import copy
                added_count = 0
                for item in sel:
                    if isinstance(item, int) and 0 <= item < len(self.actions):
                        loop_actions.append(copy.deepcopy(self.actions[item]))
                        added_count += 1
                refresh_list()
                hint_win.destroy()
                dialog.deiconify()
                dialog.lift()
                dialog.attributes('-topmost', True)
                dialog.grab_set()  # 重新设置grab
                if added_count > 0:
                    edit_status_label.config(text=f"已添加 {added_count} 个动作")

            def cancel_reselect():
                """取消重选，恢复对话框"""
                hint_win.destroy()
                dialog.deiconify()
                dialog.lift()
                dialog.attributes('-topmost', True)
                dialog.grab_set()  # 重新设置grab

            btn_frame = ttk.Frame(hint_win)
            btn_frame.pack(pady=10)
            ttk.Button(btn_frame, text="确认添加", command=confirm_and_add).pack(side=tk.LEFT, padx=10)
            ttk.Button(btn_frame, text="取消", command=cancel_reselect).pack(side=tk.LEFT, padx=10)

            hint_win.protocol("WM_DELETE_WINDOW", cancel_reselect)

            # 更新选中计数
            def update_selection_count():
                try:
                    sel = self._get_selected_indices()
                    count = len([i for i in sel if isinstance(i, int)])
                    hint_status.config(text=f"当前选中: {count} 个动作")
                    hint_win.after(500, update_selection_count)
                except:
                    pass
            update_selection_count()

        def delete_action_from_loop():
            idx = actions_listbox.curselection()
            if idx:
                del loop_actions[idx[0]]
                refresh_list()

        def clear_loop_actions():
            loop_actions.clear()
            refresh_list()

        def start_recording():
            """开始录制动作模式"""
            dialog.grab_release()
            dialog.withdraw()
            dialog.update()

            select_hint_window = tk.Toplevel(self.root)
            select_hint_window.title("录制中...")
            select_hint_window.attributes('-topmost', True)
            tk.Label(select_hint_window, text="开始录制动作...\n按 ESC 结束录制", font=('Arial', 12)).pack(padx=20, pady=15)
            select_hint_window.update()

            def stop_recording_and_close():
                select_hint_window.destroy()
                self.stop_recording()
                dialog.deiconify()
                dialog.grab_set()
                dialog.update()
                refresh_list()

            select_hint_window.bind('<Escape>', lambda e: stop_recording_and_close())
            select_hint_window.protocol('WM_DELETE_WINDOW', stop_recording_and_close)

        def insert_delay():
            """插入固定延时"""
            delay_action = {
                'type': 'delay',
                'delay': 1.0
            }
            loop_actions.append(delay_action)
            refresh_list()
            edit_status_label.config(text="已插入固定延时", foreground="green")

        def insert_screenshot_action():
            """插入截屏动作"""
            screenshot_action = {
                'type': 'screenshot',
                'naming': self.screenshot_naming_var.get(),
                'directory': self.screenshot_dir_var.get(),
                'filename': self.screenshot_name_var.get(),
                'delay': self.screenshot_delay_var.get()
            }
            loop_actions.append(screenshot_action)
            refresh_list()
            edit_status_label.config(text="已插入截屏动作", foreground="green")

        ttk.Button(actions_btn_frame, text="录制", command=start_recording).pack(side=tk.LEFT, padx=2)
        ttk.Button(actions_btn_frame, text="插入延时", command=insert_delay).pack(side=tk.LEFT, padx=2)
        ttk.Button(actions_btn_frame, text="插入截屏", command=insert_screenshot_action).pack(side=tk.LEFT, padx=2)
        ttk.Button(actions_btn_frame, text="删除选中", command=delete_action_from_loop).pack(side=tk.LEFT, padx=2)
        ttk.Button(actions_btn_frame, text="清空", command=clear_loop_actions).pack(side=tk.LEFT, padx=2)

        # 状态标签
        edit_status_label = ttk.Label(main_frame, text="", foreground="green")
        edit_status_label.grid(row=8, column=0, columnspan=2, pady=2)

        def confirm():
            try:
                action['input_template'] = template_var.get()
                target_x = None
                target_y = None
                if x_var.get().strip():
                    target_x = int(x_var.get())
                if y_var.get().strip():
                    target_y = int(y_var.get())
                action['target_x'] = target_x
                action['target_y'] = target_y
                action['interval_mode'] = interval_mode_var.get()
                action['clear_text'] = clear_text_var.get()
                action['loop_actions'] = loop_actions.copy()

                # 保存间隔参数
                interval_mode = interval_mode_var.get()
                if interval_mode == 'fixed':
                    action['loop_interval'] = interval_var.get()
                elif interval_mode == 'range':
                    if range_sub_type_var.get() == "arithmetic":
                        action['loop_interval'] = {
                            'type': 'range',
                            'sub_type': 'arithmetic',
                            'start': interval_start_var.get(),
                            'step': interval_step_var.get(),
                            'end': interval_end_var.get()
                        }
                    else:
                        action['loop_interval'] = {
                            'type': 'range',
                            'sub_type': 'geometric',
                            'start': interval_geo_start_var.get(),
                            'ratio': interval_ratio_var.get(),
                            'steps': interval_steps_var.get()
                        }
                elif interval_mode == 'list':
                    try:
                        values = [float(v.strip()) for v in interval_list_var.get().split(',')]
                        action['loop_interval'] = {'type': 'list', 'values': values}
                    except ValueError:
                        raise ValueError("列表格式无效，请使用逗号分隔的数字")
                else:  # random
                    action['loop_interval'] = {'type': 'random', 'min': interval_random_min_var.get(), 'max': interval_random_max_var.get()}

                if loop_type == 'fixed':
                    count = loop_count_var.get()
                    if count <= 0:
                        raise ValueError("重复次数必须大于0")
                    action['loop_count'] = count
                elif loop_type == 'range':
                    step = step_var.get()
                    if step == 0:
                        raise ValueError("步距不能为0")
                    action['start'] = start_var.get()
                    action['step'] = step
                    action['end'] = end_var.get()
                else:  # list
                    text_content = list_text.get('1.0', tk.END).strip()
                    if not text_content:
                        raise ValueError("文本列表不能为空")
                    action['text_list'] = [line.strip() for line in text_content.split('\n') if line.strip()]

                self._update_action_list(select_index=index)
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(main_frame, text="确定", command=confirm).grid(row=9, column=0, columnspan=2, pady=10)

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
        if 'display_delay' in action:
            relative_delay = action['display_delay']
        else:
            prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
            relative_delay = round(action.get('time', 0) - prev_time, 1)

        ttk.Label(frame, text="相对延时(秒):").grid(row=4, column=0, padx=5, pady=10, sticky="e")
        delay_var = tk.DoubleVar(value=round(relative_delay, 1))
        ttk.Spinbox(frame, width=13, textvariable=delay_var, from_=0, to=999, increment=0.1, state='normal').grid(row=4, column=1, padx=5, pady=10, sticky="w")

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
                action['display_delay'] = delay_var.get()
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

        # 相对延时 - 使用 display_delay（如果存在），否则计算
        if 'display_delay' in action:
            relative_delay = round(action['display_delay'], 1)
        else:
            parts = row_id.split('_')
            if len(parts) >= 3:
                parent_idx = int(parts[1])
                child_idx = int(parts[2])
                parent_action = self.actions[parent_idx]
                loop_actions = parent_action.get('loop_actions', [])
                if child_idx > 0 and loop_actions:
                    prev_time = loop_actions[child_idx - 1].get('time', 0)
                    curr_time = action.get('time', 0)
                    relative_delay = round(curr_time - prev_time, 1)
                else:
                    relative_delay = round(action.get('time', 0), 1)
            else:
                relative_delay = round(action.get('time', 0), 1)
        delay_var = tk.DoubleVar(value=relative_delay)
        ttk.Label(frame, text="相对延时(秒):").grid(row=2, column=0, padx=5, pady=10, sticky="e")
        ttk.Spinbox(frame, width=13, textvariable=delay_var, from_=0, to=999, increment=0.1, state='normal').grid(row=2, column=1, padx=5, pady=10, sticky="w")

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
                # 计算新的绝对时间：prev_time + 相对延时
                parts = row_id.split('_')
                if len(parts) >= 3:
                    parent_idx = int(parts[1])
                    child_idx = int(parts[2])
                    parent_action = self.actions[parent_idx]
                    loop_actions = parent_action.get('loop_actions', [])
                    if child_idx > 0 and loop_actions:
                        prev_time = loop_actions[child_idx - 1].get('time', 0)
                    else:
                        prev_time = 0
                else:
                    prev_time = 0
                action['time'] = prev_time + delay_var.get()

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

        # 相对延时 - 计算相对于前一个动作的延时
        parts = row_id.split('_')
        if len(parts) >= 3:
            parent_idx = int(parts[1])
            child_idx = int(parts[2])
            parent_action = self.actions[parent_idx]
            loop_actions = parent_action.get('loop_actions', [])
            if child_idx > 0 and loop_actions:
                prev_time = loop_actions[child_idx - 1].get('time', 0)
                curr_time = action.get('time', 0)
                relative_delay = round(curr_time - prev_time, 1)
            else:
                relative_delay = round(action.get('time', 0), 1)
        else:
            relative_delay = round(action.get('time', 0), 1)
        delay_var = tk.DoubleVar(value=round(relative_delay, 1))
        ttk.Label(frame, text="相对延时(秒):").grid(row=4, column=0, padx=5, pady=10, sticky="e")
        ttk.Spinbox(frame, width=13, textvariable=delay_var, from_=0, to=999, increment=0.1, state='normal').grid(row=4, column=1, padx=5, pady=10, sticky="w")

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
                # 计算新的绝对时间：prev_time + 相对延时
                parts = row_id.split('_')
                if len(parts) >= 3:
                    parent_idx = int(parts[1])
                    child_idx = int(parts[2])
                    parent_action = self.actions[parent_idx]
                    loop_actions = parent_action.get('loop_actions', [])
                    if child_idx > 0 and loop_actions:
                        prev_time = loop_actions[child_idx - 1].get('time', 0)
                    else:
                        prev_time = 0
                else:
                    prev_time = 0
                action['time'] = prev_time + delay_var.get()

                self._update_action_list()
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).grid(row=7, column=0, columnspan=2, pady=15)

    def _edit_nested_traverse_input_action(self, action, row_id):
        """编辑嵌套的遍历输入动作"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑遍历输入")
        dialog.geometry("500x340")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        dialog.grab_set()

        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 遍历类型选择
        ttk.Label(main_frame, text="遍历类型:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        type_var = tk.StringVar(value=action.get('traverse_type', 'list'))
        type_frame = ttk.Frame(main_frame)
        type_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(type_frame, text="固定文本", variable=type_var, value="fixed_text",
                       command=lambda: update_type_ui("fixed_text")).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="列表遍历", variable=type_var, value="list",
                       command=lambda: update_type_ui("list")).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="序列", variable=type_var, value="sequence",
                       command=lambda: update_type_ui("sequence")).pack(side=tk.LEFT, padx=5)

        # 点击位置
        ttk.Label(main_frame, text="点击位置:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        pos_frame = ttk.Frame(main_frame)
        pos_frame.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        x_var = tk.StringVar(value=str(action.get('x', '')))
        y_var = tk.StringVar(value=str(action.get('y', '')))
        ttk.Label(pos_frame, text="X:").pack(side=tk.LEFT)
        ttk.Entry(pos_frame, width=8, textvariable=x_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(pos_frame, text="Y:").pack(side=tk.LEFT)
        ttk.Entry(pos_frame, width=8, textvariable=y_var).pack(side=tk.LEFT, padx=2)

        record_btn = ttk.Button(pos_frame, text="录制")
        record_btn.pack(side=tk.LEFT, padx=5)

        # 参数框架（动态切换）
        params_frame = ttk.Frame(main_frame)
        params_frame.grid(row=2, column=0, columnspan=2, sticky="w", pady=10)

        # 列表遍历参数
        list_frame = ttk.Frame(params_frame)
        list_row1 = ttk.Frame(list_frame)
        list_row1.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(list_row1, text="遍历列表:").pack(side=tk.LEFT, padx=2)
        list_text = tk.Text(list_row1, width=25, height=6)
        list_text.pack(side=tk.LEFT, padx=2)
        traverse_values = action.get('traverse_values', [])
        if traverse_values:
            list_text.insert(tk.END, '\n'.join(str(v) for v in traverse_values))
        ttk.Label(list_frame, text="(每行一个值)").pack(side=tk.TOP, padx=2, pady=(2, 0))

        # 序列参数
        seq_frame = ttk.Frame(params_frame)
        seq_row1 = ttk.Frame(seq_frame)
        seq_row1.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(seq_row1, text="起始值:").pack(side=tk.LEFT, padx=2)
        start_var = tk.DoubleVar(value=action.get('start', 0))
        ttk.Entry(seq_row1, width=8, textvariable=start_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(seq_row1, text="步距:").pack(side=tk.LEFT, padx=2)
        step_var = tk.DoubleVar(value=action.get('step', 1))
        ttk.Entry(seq_row1, width=8, textvariable=step_var).pack(side=tk.LEFT, padx=2)
        seq_row2 = ttk.Frame(seq_frame)
        seq_row2.pack(side=tk.TOP, fill=tk.X, pady=(5, 0))
        ttk.Label(seq_row2, text="模板:").pack(side=tk.LEFT, padx=2)
        seq_template_var = tk.StringVar(value=action.get('template', 'text{n}'))
        ttk.Entry(seq_row2, width=15, textvariable=seq_template_var).pack(side=tk.LEFT, padx=2)
        ttk.Label(seq_row2, text="(用{n}替换)").pack(side=tk.LEFT, padx=2)

        # 固定文本参数
        fixed_text_frame = ttk.Frame(params_frame)
        ttk.Label(fixed_text_frame, text="输入文本:").pack(side=tk.LEFT, padx=2)
        fixed_text_var = tk.StringVar(value=action.get('fixed_text', ''))
        ttk.Entry(fixed_text_frame, width=25, textvariable=fixed_text_var).pack(side=tk.LEFT, padx=2)

        def update_type_ui(mode):
            list_frame.grid_forget()
            seq_frame.grid_forget()
            fixed_text_frame.grid_forget()
            if mode == "list":
                list_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
            elif mode == "sequence":
                seq_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)
            else:  # fixed_text
                fixed_text_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=5)

        # 初始化UI
        current_type = action.get('traverse_type', 'list')
        update_type_ui(current_type)

        # 相对延时 - 计算相对于前一个动作的延时
        parts = row_id.split('_')
        if len(parts) >= 3:
            parent_idx = int(parts[1])
            child_idx = int(parts[2])
            parent_action = self.actions[parent_idx]
            loop_actions = parent_action.get('loop_actions', [])
            if child_idx > 0 and loop_actions:
                prev_time = loop_actions[child_idx - 1].get('time', 0)
                curr_time = action.get('time', 0)
                relative_delay = round(curr_time - prev_time, 1)
            else:
                relative_delay = round(action.get('time', 0), 1)
        else:
            relative_delay = round(action.get('time', 0), 1)
        delay_var = tk.DoubleVar(value=round(relative_delay, 1))
        ttk.Label(main_frame, text="相对延时(秒):").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        ttk.Spinbox(main_frame, width=13, textvariable=delay_var, from_=0, to=999, increment=0.1, state='normal').grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=15)

        # 录制按钮功能
        def start_record():
            dialog.withdraw()
            dialog.update()
            captured_pos = [None, None]

            def on_click(x, y, button, pressed):
                if pressed:
                    captured_pos[0] = x
                    captured_pos[1] = y
                    return False

            listener = mouse.Listener(on_click=on_click)
            listener.start()
            listener.join()

            if captured_pos[0] is not None:
                x_var.set(str(captured_pos[0]))
                y_var.set(str(captured_pos[1]))
            dialog.deiconify()

        record_btn.config(command=start_record)

        def confirm():
            try:
                x = int(x_var.get()) if x_var.get() else 0
                y = int(y_var.get()) if y_var.get() else 0

                traverse_type = type_var.get()
                values = []
                start = 0
                step = 1
                fixed_text = ""
                if traverse_type == "list":
                    text_content = list_text.get("1.0", tk.END).strip()
                    if not text_content:
                        raise ValueError("遍历列表不能为空")
                    values = [line.strip() for line in text_content.split('\n') if line.strip()]
                elif traverse_type == "sequence":
                    start = start_var.get()
                    step = step_var.get()
                    seq_template = seq_template_var.get()
                else:  # fixed_text
                    fixed_text = fixed_text_var.get()
                    if not fixed_text:
                        raise ValueError("输入文本不能为空")

                # 更新动作
                action['traverse_type'] = traverse_type
                action['x'] = x
                action['y'] = y
                action['traverse_values'] = values if traverse_type == "list" else []
                action['start'] = start if traverse_type == "sequence" else None
                action['step'] = step if traverse_type == "sequence" else None
                action['template'] = seq_template if traverse_type == "sequence" else None
                action['fixed_text'] = fixed_text if traverse_type == "fixed_text" else None
                # 计算新的绝对时间：prev_time + 相对延时
                parts = row_id.split('_')
                if len(parts) >= 3:
                    parent_idx = int(parts[1])
                    child_idx = int(parts[2])
                    parent_action = self.actions[parent_idx]
                    loop_actions = parent_action.get('loop_actions', [])
                    if child_idx > 0 and loop_actions:
                        prev_time = loop_actions[child_idx - 1].get('time', 0)
                    else:
                        prev_time = 0
                else:
                    prev_time = 0
                action['time'] = prev_time + delay_var.get()

                self._update_action_list()
                self.in_dialog_operation = False
                dialog.destroy()
            except ValueError as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(btn_frame, text="确定", command=confirm).pack(side=tk.LEFT, padx=20)
        ttk.Button(btn_frame, text="取消", command=on_dialog_close).pack(side=tk.LEFT, padx=5)

    def _edit_nested_loop_group(self, action, row_id):
        """编辑嵌套的循环组"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑循环组")
        dialog.geometry("385x225")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # 计算循环次数上限（如果包含列表遍历类型的遍历输入，最大值设为列表长度）
        loop_actions = action.get('loop_actions', [])
        max_loop_count = 9999
        for act in loop_actions:
            if act.get('type') == 'traverse_input' and act.get('traverse_type') == 'list':
                values = act.get('traverse_values', [])
                if values:
                    max_loop_count = len(values)
                    break

        # 循环次数
        ttk.Label(frame, text="循环次数:").grid(row=0, column=0, padx=5, pady=10, sticky="e")
        loop_count_var = tk.IntVar(value=min(action.get('loop_count', 1), max_loop_count))
        loop_count_spinbox = ttk.Spinbox(frame, from_=1, to=max_loop_count, width=15, textvariable=loop_count_var)
        loop_count_spinbox.grid(row=0, column=1, padx=5, pady=10, sticky="w")

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

    def _edit_top_loop_group(self, action, index):
        """编辑顶层的循环组"""
        self.in_dialog_operation = True

        dialog = tk.Toplevel(self.root)
        dialog.title("编辑循环组")
        dialog.geometry("170x90")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        def on_dialog_close():
            self.in_dialog_operation = False
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        # 计算循环次数上限（如果包含列表遍历类型的遍历输入，最大值设为列表长度）
        loop_actions = action.get('loop_actions', [])
        max_loop_count = 9999
        for act in loop_actions:
            if act.get('type') == 'traverse_input' and act.get('traverse_type') == 'list':
                values = act.get('traverse_values', [])
                if values:
                    max_loop_count = len(values)
                    break

        # 循环次数
        loop_count_frame = ttk.Frame(frame)
        loop_count_frame.pack(fill=tk.X, pady=10)
        ttk.Label(loop_count_frame, text="循环次数：").pack(side=tk.LEFT)
        loop_count_var = tk.IntVar(value=min(action.get('loop_count', 1), max_loop_count))
        loop_count_spinbox = ttk.Spinbox(loop_count_frame, from_=1, to=max_loop_count, width=15, textvariable=loop_count_var)
        loop_count_spinbox.pack(side=tk.LEFT, padx=8)

        def confirm():
            try:
                loop_count = loop_count_var.get()
                if loop_count < 1:
                    raise ValueError("循环次数至少为1")
                action['loop_count'] = loop_count
                self._update_action_list(select_index=index)
                self.in_dialog_operation = False
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"输入无效: {str(e)}")

        ttk.Button(frame, text="确定", command=confirm).pack(pady=10)

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
            is_nested = row_id.startswith('loop_') or row_id.startswith('input_loop_')
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
            if action['type'] not in ['move', 'click', 'doubleclick', 'delay', 'random_delay', 'multiply_delay', 'arithmetic_delay', 'screenshot', 'loop_group', 'numeric_loop', 'variable_input']:
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
            if action_type in ['delay', 'random_delay', 'multiply_delay', 'arithmetic_delay']:
                # 延时类型选择 - 将动作类型映射到UI类型
                ui_type = "fixed" if action_type == "delay" else action_type
                type_var = tk.StringVar(value=ui_type)

                # 类型选择
                ttk.Label(frame, text="延时类型:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
                type_frame = ttk.Frame(frame)
                type_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
                ttk.Radiobutton(type_frame, text="固定延时", variable=type_var, value="fixed",
                               command=lambda: update_ui("fixed")).pack(side=tk.LEFT, padx=2)
                ttk.Radiobutton(type_frame, text="随机范围", variable=type_var, value="random",
                               command=lambda: update_ui("random")).pack(side=tk.LEFT, padx=2)
                ttk.Radiobutton(type_frame, text="倍数增长", variable=type_var, value="multiply",
                               command=lambda: update_ui("multiply")).pack(side=tk.LEFT, padx=2)
                ttk.Radiobutton(type_frame, text="等差递增", variable=type_var, value="arithmetic",
                               command=lambda: update_ui("arithmetic")).pack(side=tk.LEFT, padx=2)

                # 固定延时参数
                fixed_frame = ttk.Frame(frame)
                fixed_frame.grid(row=1, column=0, columnspan=2, sticky="ew")

                ttk.Label(fixed_frame, text="延时(秒):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
                fixed_var = tk.DoubleVar(value=action.get('delay', 1.0))
                ttk.Entry(fixed_frame, width=10, textvariable=fixed_var).grid(row=0, column=1, padx=5, pady=5, sticky="w")

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
                    if delay_type in ("fixed", "delay"):
                        fixed_frame.grid()
                        random_frame.grid_remove()
                        exp_frame.grid_remove()
                        arithmetic_frame.grid_remove()
                    elif delay_type == "random":
                        fixed_frame.grid_remove()
                        random_frame.grid()
                        exp_frame.grid_remove()
                        arithmetic_frame.grid_remove()
                    elif delay_type == "multiply":
                        fixed_frame.grid_remove()
                        random_frame.grid_remove()
                        exp_frame.grid()
                        arithmetic_frame.grid_remove()
                    else:  # arithmetic
                        fixed_frame.grid_remove()
                        random_frame.grid_remove()
                        exp_frame.grid_remove()
                        arithmetic_frame.grid()

                # 初始化UI状态 - 将动作类型映射到UI类型
                ui_type = "fixed" if action_type == "delay" else action_type
                update_ui(ui_type)

                def confirm():
                    try:
                        delay_type = type_var.get()

                        if delay_type == "fixed":
                            delay = fixed_var.get()
                            if delay < 0:
                                raise ValueError("延时不能为负数")
                            action['type'] = 'delay'
                            action['delay'] = delay
                        elif delay_type == "random":
                            min_delay = min_var.get()
                            max_delay = max_var.get()
                            if min_delay < 0 or max_delay < min_delay:
                                raise ValueError("无效的延时范围")
                            action['type'] = 'random_delay'
                            action['min_delay'] = min_delay
                            action['max_delay'] = max_delay
                        elif delay_type == "multiply":
                            base_delay = base_var.get()
                            if base_delay < 0:
                                raise ValueError("基数不能为负数")
                            action['type'] = 'multiply_delay'
                            action['base_delay'] = base_delay
                        else:  # arithmetic
                            start_delay = arithmetic_start_var.get()
                            step_delay = arithmetic_step_var.get()
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
                            action['display_delay'] = new_value
                            self._shift_action_times(index + 1, delta)
                        else:
                            # click, move, doubleclick
                            prev_time = self.actions[index - 1].get('time', 0) if index > 0 else 0
                            new_absolute_time = prev_time + new_value
                            old_time = action.get('time', 0)
                            delta = new_absolute_time - old_time
                            action['time'] = new_absolute_time
                            action['display_delay'] = new_value
                            self._shift_action_times(index + 1, delta)

                        self._update_action_list(select_index=index)
                        self.in_dialog_operation = False
                        dialog.destroy()
                    except Exception as e:
                        messagebox.showerror("错误", f"输入无效: {str(e)}")

                ttk.Button(frame, text="确定", command=confirm).grid(row=2, column=0, columnspan=2, pady=10)

        except Exception as e:
            self.in_dialog_operation = False
            messagebox.showerror("错误", f"修改时间失败: {str(e)}")

    def add_screenshot_action(self):
        """添加截屏动作 - 弹出配置对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("截屏设置")
        dialog.geometry("560x200")
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)
        dialog.grab_set()

        def on_dialog_close():
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 目录
        ttk.Label(main_frame, text="目录:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ss_dir_var = tk.StringVar(value=self.screenshot_dir_var.get())
        ss_dir_entry = ttk.Entry(main_frame, textvariable=ss_dir_var, width=30)
        ss_dir_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        def select_ss_dir():
            dir_path = filedialog.askdirectory(initialdir=ss_dir_var.get(), title="选择截屏保存目录")
            if dir_path:
                ss_dir_var.set(dir_path)

        ttk.Button(main_frame, text="选择", command=select_ss_dir).grid(row=0, column=2, padx=5, pady=5)

        # 名称
        ttk.Label(main_frame, text="名称:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        ss_name_var = tk.StringVar(value=self.screenshot_name_var.get())
        ss_name_entry = ttk.Entry(main_frame, textvariable=ss_name_var, width=15)
        ss_name_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # 延时
        delay_frame = ttk.Frame(main_frame)
        delay_frame.grid(row=1, column=2, columnspan=2, padx=5, pady=5, sticky="w")
        ttk.Label(delay_frame, text="延时:").pack(side=tk.LEFT)
        ss_delay_var = tk.DoubleVar(value=self.screenshot_delay_var.get())
        ttk.Entry(delay_frame, textvariable=ss_delay_var, width=8).pack(side=tk.LEFT, padx=2)

        # 命名方式
        ttk.Label(main_frame, text="命名:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        ss_naming_var = tk.StringVar(value="timestamp")
        ttk.Radiobutton(main_frame, text="时间戳", variable=ss_naming_var, value="timestamp").grid(row=2, column=1, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(main_frame, text="递增序号", variable=ss_naming_var, value="increment").grid(row=2, column=2, padx=5, pady=5, sticky="w")

        def confirm_screenshot():
            screenshot_action = {
                'type': 'screenshot',
                'naming': ss_naming_var.get(),
                'directory': ss_dir_var.get(),
                'filename': ss_name_var.get(),
                'delay': ss_delay_var.get()
            }
            self.actions.append(screenshot_action)
            self._update_action_list(scroll_to_end=True)
            dialog.destroy()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, columnspan=4, pady=10)
        ttk.Button(btn_frame, text="确定", command=confirm_screenshot).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def take_screenshot(self, delay=0, directory=None, filename=None, naming=None):
        """截屏并保存到指定目录，支持延时和命名方式选择"""
        time.sleep(delay)  # 添加延时
        # 使用传入的参数或全局设置
        directory = directory if directory else self.screenshot_dir_var.get()
        filename = filename if filename else self.screenshot_name_var.get()
        naming = naming if naming else self.screenshot_naming_var.get()

        if not os.path.exists(directory):
            os.makedirs(directory)

        if naming == "timestamp":
            filepath = os.path.join(directory, f"{filename}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        else:
            existing_files = [f for f in os.listdir(directory) if f.startswith(filename) and f.endswith('.png')]
            next_index = len(existing_files) + 1
            filepath = os.path.join(directory, f"{filename}_{next_index}.png")

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