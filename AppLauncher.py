import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import subprocess
import sys
import ctypes
from ctypes import Structure, c_void_p, POINTER
from ctypes.wintypes import DWORD, LONG, WORD, UINT, HBITMAP, HDC, HGDIOBJ, RECT, INT, LPARAM, LPCSTR, HICON
import tempfile
import time

class AppLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("应用启动器")
        self.root.geometry("650x750")
        
        # 图标缓存字典
        self.icon_cache = {}
        
        # 默认图标
        self.default_icon = None
        
        # 存储应用数据
        self.apps = {}
        self.data_file = "apps.json"
        
        # 当前选中的应用
        self.selected_app = None
        
        # 触摸滚动相关变量
        self.touch_start_y = None
        self.touch_scrolling = False
        self.press_item = None
        
        # 手柄控制相关变量（使用 XInput）
        self.xinput_available = False
        self.last_move_time = 0
        self.move_delay = 200
        self.last_button_a_state = False
        
        # 加载保存的应用数据
        self.load_apps()
        
        # 创建UI
        self.create_widgets()
        
        # 初始化触摸滚动监听
        self.setup_touch_scrolling()
        
        # 获取本窗口的 Windows 句柄（用于焦点判断）
        self.root.update()  # 确保窗口已创建
        self.hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
        
        # 初始化手柄支持（XInput）
        self.init_xinput()
    
    # ---------- 触摸滚动支持 ----------
    def setup_touch_scrolling(self):
        """为Treeview绑定触摸滚动事件"""
        self.app_tree.bind('<ButtonPress-1>', self.on_tree_press)
        self.app_tree.bind('<B1-Motion>', self.on_tree_motion)
        self.app_tree.bind('<ButtonRelease-1>', self.on_tree_release)
    
    def on_tree_press(self, event):
        self.touch_start_y = event.y
        self.touch_scrolling = False
        self.press_item = self.app_tree.identify_row(event.y)
    
    def on_tree_motion(self, event):
        if self.touch_start_y is None:
            return
        delta = event.y - self.touch_start_y
        if abs(delta) > 10:
            self.touch_scrolling = True
            scroll_units = int(-delta / 8)
            if scroll_units != 0:
                self.app_tree.yview_scroll(scroll_units, "units")
                self.touch_start_y = event.y
            return "break"
    
    def on_tree_release(self, event):
        if self.touch_start_y is not None and not self.touch_scrolling:
            item = self.app_tree.identify_row(event.y)
            if item:
                self.app_tree.selection_set(item)
                self.app_tree.focus(item)
                self.select_app(item)
        self.touch_start_y = None
        self.touch_scrolling = False
        self.press_item = None
    
    # ---------- 手柄支持（XInput，无需pygame） ----------
    def init_xinput(self):
        """初始化 XInput（Windows 原生手柄 API）"""
        try:
            # 定义结构体（只定义一次，保存为实例属性）
            class XINPUT_GAMEPAD(ctypes.Structure):
                _fields_ = [
                    ("wButtons", ctypes.c_ushort),
                    ("bLeftTrigger", ctypes.c_byte),
                    ("bRightTrigger", ctypes.c_byte),
                    ("sThumbLX", ctypes.c_short),
                    ("sThumbLY", ctypes.c_short),
                    ("sThumbRX", ctypes.c_short),
                    ("sThumbRY", ctypes.c_short)
                ]

            class XINPUT_STATE(ctypes.Structure):
                _fields_ = [
                    ("dwPacketNumber", ctypes.c_uint),
                    ("Gamepad", XINPUT_GAMEPAD)
                ]

            self.XINPUT_GAMEPAD = XINPUT_GAMEPAD
            self.XINPUT_STATE = XINPUT_STATE

            # 常量定义
            self.XINPUT_GAMEPAD_DPAD_UP = 0x0001
            self.XINPUT_GAMEPAD_DPAD_DOWN = 0x0002
            self.XINPUT_GAMEPAD_A = 0x1000
            self.XINPUT_GAMEPAD_LEFT_THUMB_DEADZONE = 7849

            # 加载 xinput DLL
            dll_names = ["xinput1_4.dll", "xinput1_3.dll", "xinput9_1_0.dll"]
            self.xinput_dll = None
            for name in dll_names:
                try:
                    self.xinput_dll = ctypes.windll.LoadLibrary(name)
                    break
                except OSError:
                    continue

            if self.xinput_dll is None:
                print("未找到 XInput DLL，手柄支持禁用")
                return

            # 获取函数并设置参数类型
            self.XInputGetState = self.xinput_dll.XInputGetState
            self.XInputGetState.argtypes = [ctypes.c_uint, ctypes.POINTER(XINPUT_STATE)]
            self.XInputGetState.restype = ctypes.c_uint

            self.xinput_available = True
            print("XInput 初始化成功，手柄支持已启用")

            # 开始轮询
            self.poll_xinput()

        except Exception as e:
            print(f"初始化 XInput 失败: {e}")
            self.xinput_available = False

    def poll_xinput(self):
        """定期轮询 XInput 手柄状态"""
        if not self.xinput_available:
            return

        # 检查本程序是否为前台窗口（焦点检查）
        try:
            foreground_hwnd = ctypes.windll.user32.GetForegroundWindow()
        except:
            foreground_hwnd = None

        # 只有前台窗口是自己时才响应手柄操作
        if foreground_hwnd != self.hwnd:
            # 不是前台窗口，继续轮询但忽略按键
            self.root.after(50, self.poll_xinput)
            return

        # 直接使用已定义的结构体类型
        state = self.XINPUT_STATE()
        result = self.XInputGetState(0, ctypes.byref(state))

        if result == 0:  # ERROR_SUCCESS
            gamepad = state.Gamepad

            # 方向检测（摇杆 + 十字键）
            move_detected = False
            move_direction = None

            # 左摇杆 Y 轴
            ly = gamepad.sThumbLY
            dead_zone = self.XINPUT_GAMEPAD_LEFT_THUMB_DEADZONE
            if ly < -dead_zone:
                move_detected = True
                move_direction = "down"
            elif ly > dead_zone:
                move_detected = True
                move_direction = "up"

            # 十字键（DPad）
            buttons = gamepad.wButtons
            if not move_detected:
                if buttons & self.XINPUT_GAMEPAD_DPAD_UP:
                    move_detected = True
                    move_direction = "up"
                elif buttons & self.XINPUT_GAMEPAD_DPAD_DOWN:
                    move_detected = True
                    move_direction = "down"

            # 执行移动（带延迟）
            current_time = time.time() * 1000
            if move_detected and (current_time - self.last_move_time > self.move_delay):
                self.last_move_time = current_time
                self.move_selection(move_direction)

            # 确认按钮（A 键）
            a_pressed = (buttons & self.XINPUT_GAMEPAD_A) != 0
            if a_pressed and not self.last_button_a_state:
                self.launch_app()
            self.last_button_a_state = a_pressed

        # 继续轮询
        self.root.after(50, self.poll_xinput)

    def move_selection(self, direction):
        """移动选中项（上下）"""
        children = self.app_tree.get_children()
        if not children:
            return

        current = self.app_tree.selection()
        if current:
            current_item = current[0]
        else:
            target = children[0] if direction == "down" else children[-1]
            self.set_selected_app(target)
            return

        try:
            idx = children.index(current_item)
            if direction == "up" and idx > 0:
                target = children[idx - 1]
                self.set_selected_app(target)
            elif direction == "down" and idx < len(children) - 1:
                target = children[idx + 1]
                self.set_selected_app(target)
        except ValueError:
            self.set_selected_app(children[0])

    def set_selected_app(self, app_name):
        """设置选中的应用（供手柄调用）"""
        if app_name not in self.apps:
            return
        self.app_tree.selection_set(app_name)
        self.app_tree.focus(app_name)
        self.app_tree.see(app_name)
        self.select_app(app_name)

    # ---------- 图标提取 ----------
    def create_default_icon(self):
        width = height = 32
        rgb_data = bytearray(width * height * 3)
        for y in range(height):
            for x in range(width):
                idx = (y * width + x) * 3
                if 6 <= x <= 25 and 6 <= y <= 25:
                    rgb_data[idx] = 0
                    rgb_data[idx+1] = 0
                    rgb_data[idx+2] = 200
                else:
                    rgb_data[idx] = 240
                    rgb_data[idx+1] = 240
                    rgb_data[idx+2] = 240
        ppm_data = f"P6\n{width} {height}\n255\n".encode() + rgb_data
        self.default_icon = tk.PhotoImage(data=ppm_data, width=width, height=height)

    def extract_icon_with_ctypes(self, app_path, size=32):
        try:
            class BITMAPINFOHEADER(Structure):
                _fields_ = [
                    ("biSize", DWORD), ("biWidth", LONG), ("biHeight", LONG),
                    ("biPlanes", WORD), ("biBitCount", WORD), ("biCompression", DWORD),
                    ("biSizeImage", DWORD), ("biXPelsPerMeter", LONG),
                    ("biYPelsPerMeter", LONG), ("biClrUsed", DWORD), ("biClrImportant", DWORD)
                ]
            class BITMAPINFO(Structure):
                _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", DWORD * 3)]

            shell32 = ctypes.windll.shell32
            user32 = ctypes.windll.user32
            gdi32 = ctypes.windll.gdi32

            large_icons = ctypes.c_void_p()
            small_icons = ctypes.c_void_p()
            icon_count = shell32.ExtractIconExW(app_path, 0, ctypes.byref(large_icons), ctypes.byref(small_icons), 1)
            if icon_count == 0:
                return None
            hIcon = large_icons
            icon_size = 32
            for test_size in [256, 128, 64, 48]:
                test_icon = shell32.ExtractIconExW(app_path, 0, None, None, 1)
                if test_icon > 0:
                    hTestIcon = ctypes.c_void_p()
                    shell32.ExtractIconExW(app_path, 0, ctypes.byref(hTestIcon), None, 1)
                    if hTestIcon:
                        icon_info = ctypes.create_string_buffer(ctypes.sizeof(ctypes.c_void_p) * 5)
                        if user32.GetIconInfo(hTestIcon, ctypes.byref(icon_info)):
                            icon_size = test_size
                            hIcon = hTestIcon
                            if hIcon != large_icons:
                                user32.DestroyIcon(large_icons)

            hdc = user32.GetDC(0)
            memdc = gdi32.CreateCompatibleDC(hdc)
            width = height = min(icon_size, 32)
            hbitmap = gdi32.CreateCompatibleBitmap(hdc, width, height)
            old_bitmap = gdi32.SelectObject(memdc, hbitmap)
            gdi32.PatBlt(memdc, 0, 0, width, height, 0x00F00021)
            user32.DrawIconEx(memdc, 0, 0, hIcon, width, height, 0, 0, 3)
            bmi = BITMAPINFO()
            bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bmi.bmiHeader.biPlanes = 1
            bmi.bmiHeader.biBitCount = 32
            bmi.bmiHeader.biCompression = 0
            bmi.bmiHeader.biClrUsed = 0
            bmi.bmiHeader.biClrImportant = 0
            bmi.bmiHeader.biWidth = width
            bmi.bmiHeader.biHeight = -height
            data_size = width * height * 4
            data = ctypes.create_string_buffer(data_size)
            result = gdi32.GetDIBits(memdc, hbitmap, 0, height, data, ctypes.byref(bmi), 0)
            gdi32.SelectObject(memdc, old_bitmap)
            gdi32.DeleteObject(hbitmap)
            gdi32.DeleteDC(memdc)
            user32.ReleaseDC(0, hdc)
            user32.DestroyIcon(hIcon)
            if result != height:
                return None
            raw_data = bytearray(data)
            rgb_data = bytearray(width * height * 3)
            for i in range(height * width):
                src_offset = i * 4
                dst_offset = i * 3
                alpha = raw_data[src_offset + 3] / 255.0
                b = raw_data[src_offset]
                g = raw_data[src_offset + 1]
                r = raw_data[src_offset + 2]
                rgb_data[dst_offset] = int(r * alpha + 255 * (1 - alpha))
                rgb_data[dst_offset + 1] = int(g * alpha + 255 * (1 - alpha))
                rgb_data[dst_offset + 2] = int(b * alpha + 255 * (1 - alpha))
            return rgb_data
        except Exception as e:
            print(f"使用ctypes提取图标失败: {e}")
            return None

    def get_app_icon_simple(self, app_path):
        try:
            from ctypes import wintypes
            SHGFI_ICON = 0x000000100
            SHGFI_LARGEICON = 0x000000000
            class SHFILEINFO(ctypes.Structure):
                _fields_ = [
                    ("hIcon", ctypes.c_void_p), ("iIcon", ctypes.c_int),
                    ("dwAttributes", ctypes.c_uint), ("szDisplayName", ctypes.c_wchar * 260),
                    ("szTypeName", ctypes.c_wchar * 80)
                ]
            shell32 = ctypes.windll.shell32
            sfi = SHFILEINFO()
            result = shell32.SHGetFileInfoW(app_path, 0, ctypes.byref(sfi), ctypes.sizeof(sfi),
                                            SHGFI_ICON | SHGFI_LARGEICON)
            if result and sfi.hIcon:
                temp_dir = tempfile.gettempdir()
                temp_ico = os.path.join(temp_dir, f"temp_icon_{os.getpid()}.ico")
                ico_data = self.icon_to_ico(sfi.hIcon)
                if ico_data:
                    with open(temp_ico, "wb") as f:
                        f.write(ico_data)
                    icon = tk.PhotoImage(file=temp_ico)
                    os.remove(temp_ico)
                    ctypes.windll.user32.DestroyIcon(sfi.hIcon)
                    return icon
                ctypes.windll.user32.DestroyIcon(sfi.hIcon)
        except Exception as e:
            print(f"使用SHGetFileInfo获取图标失败: {e}")
        return None

    def icon_to_ico(self, hIcon):
        # 简化处理，实际可改进
        return None

    def get_app_icon(self, app_path):
        if not app_path or not os.path.exists(app_path):
            if self.default_icon is None:
                self.create_default_icon()
            return self.default_icon
        if app_path in self.icon_cache:
            return self.icon_cache[app_path]
        icon = self.get_app_icon_simple(app_path)
        if icon:
            self.icon_cache[app_path] = icon
            return icon
        try:
            rgb_data = self.extract_icon_with_ctypes(app_path, 32)
            if rgb_data:
                ppm_data = f"P6\n32 32\n255\n".encode() + rgb_data
                icon = tk.PhotoImage(data=ppm_data, width=32, height=32)
                self.icon_cache[app_path] = icon
                return icon
        except Exception as e:
            print(f"获取图标失败 {app_path}: {e}")
        if self.default_icon is None:
            self.create_default_icon()
        return self.default_icon

    # ---------- UI 构建 ----------
    def create_widgets(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", font=("Arial", 12), rowheight=40, background="#ffffff", fieldbackground="#ffffff")
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"), background="#f0f0f0")

        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        title_label = ttk.Label(main_frame, text="应用启动器", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 15))

        ttk.Label(main_frame, text="应用名称:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.app_name_var = tk.StringVar()
        self.app_name_entry = ttk.Entry(main_frame, textvariable=self.app_name_var, width=40, font=("Arial", 10))
        self.app_name_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=5)
        self.new_btn = ttk.Button(main_frame, text="新建", command=self.add_app)
        self.new_btn.grid(row=1, column=2, padx=(10, 0), pady=5)

        ttk.Label(main_frame, text="启动环境路径:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.env_path_var = tk.StringVar()
        env_frame = ttk.Frame(main_frame)
        env_frame.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.env_entry = ttk.Entry(env_frame, textvariable=self.env_path_var, width=40, font=("Arial", 10))
        self.env_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.env_browse_btn = ttk.Button(env_frame, text="浏览...", width=10, command=self.browse_env_path)
        self.env_browse_btn.pack(side=tk.LEFT, padx=(5, 0))

        ttk.Label(main_frame, text="应用绝对路径:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.app_path_var = tk.StringVar()
        app_path_frame = ttk.Frame(main_frame)
        app_path_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.app_path_entry = ttk.Entry(app_path_frame, textvariable=self.app_path_var, width=40, font=("Arial", 10))
        self.app_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.app_browse_btn = ttk.Button(app_path_frame, text="浏览...", width=10, command=self.browse_app_path)
        self.app_browse_btn.pack(side=tk.LEFT, padx=(5, 0))

        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15)

        ttk.Label(main_frame, text="搜索应用:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_apps)
        self.search_entry = ttk.Entry(main_frame, textvariable=self.search_var, width=40, font=("Arial", 10))
        self.search_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=5)

        list_frame = ttk.LabelFrame(main_frame, text="应用列表", padding="5")
        list_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.app_tree = ttk.Treeview(list_frame, height=6, selectmode="browse", show='tree')
        self.app_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.app_tree.column("#0", width=400, stretch=True, anchor="w")
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.app_tree.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.app_tree.configure(yscrollcommand=scrollbar.set)
        self.app_tree.bind('<<TreeviewSelect>>', self.on_app_select)
        self.app_tree.bind('<Enter>', self.on_app_hover)
        self.app_tree.bind('<Motion>', self.on_app_hover)

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)
        self.launch_btn = ttk.Button(button_frame, text="启动应用", command=self.launch_app, state=tk.DISABLED)
        self.launch_btn.pack(side=tk.LEFT, padx=5)
        self.delete_btn = ttk.Button(button_frame, text="删除应用", command=self.delete_app, state=tk.DISABLED)
        self.delete_btn.pack(side=tk.LEFT, padx=5)

        detail_frame = ttk.LabelFrame(main_frame, text="应用详情", padding="10")
        detail_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        detail_frame.columnconfigure(1, weight=1)
        ttk.Label(detail_frame, text="应用名称:").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.detail_name = ttk.Label(detail_frame, text="", foreground="blue", font=("Arial", 10))
        self.detail_name.grid(row=0, column=1, sticky=tk.W, pady=3)
        ttk.Label(detail_frame, text="环境路径:").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.detail_env = ttk.Label(detail_frame, text="", font=("Arial", 9))
        self.detail_env.grid(row=1, column=1, sticky=tk.W, pady=3)
        ttk.Label(detail_frame, text="应用路径:").grid(row=2, column=0, sticky=tk.W, pady=3)
        self.detail_path = ttk.Label(detail_frame, text="", font=("Arial", 9))
        self.detail_path.grid(row=2, column=1, sticky=tk.W, pady=3)

        self.update_app_list()

    # ---------- 应用管理 ----------
    def browse_env_path(self):
        folder = filedialog.askdirectory(title="选择环境路径")
        if folder:
            self.env_path_var.set(folder)

    def browse_app_path(self):
        file = filedialog.askopenfilename(title="选择应用程序")
        if file:
            self.app_path_var.set(file)

    def add_app(self):
        app_name = self.app_name_var.get().strip()
        env_path = self.env_path_var.get().strip()
        app_path = self.app_path_var.get().strip()
        if not app_name:
            messagebox.showerror("错误", "应用名称不能为空!")
            return
        if not app_path:
            messagebox.showerror("错误", "应用路径不能为空!")
            return
        if app_name in self.apps:
            messagebox.showerror("错误", f"应用 '{app_name}' 已存在!")
            return
        self.apps[app_name] = {"env_path": env_path, "app_path": app_path}
        self.save_apps()
        self.update_app_list()
        self.app_name_var.set("")
        self.env_path_var.set("")
        self.app_path_var.set("")
        messagebox.showinfo("成功", f"应用 '{app_name}' 已添加!")

    def filter_apps(self, *args):
        search_text = self.search_var.get().lower()
        self.update_app_list(search_text)

    def update_app_list(self, filter_text=""):
        for item in self.app_tree.get_children():
            self.app_tree.delete(item)
        app_names = sorted(self.apps.keys())
        for app_name in app_names:
            if filter_text in app_name.lower():
                icon = self.get_app_icon(self.apps[app_name]["app_path"])
                self.app_tree.insert("", "end", iid=app_name, text="  "+app_name, image=icon)

    def on_app_select(self, event):
        selection = self.app_tree.selection()
        if selection:
            self.select_app(selection[0])

    def on_app_hover(self, event):
        item = self.app_tree.identify_row(event.y)
        if item and item in self.apps:
            app_data = self.apps[item]
            self.detail_name.config(text=item)
            self.detail_env.config(text=app_data["env_path"] or "未设置")
            self.detail_path.config(text=app_data["app_path"])

    def select_app(self, app_name):
        if app_name in self.apps:
            self.selected_app = app_name
            app_data = self.apps[app_name]
            self.detail_name.config(text=app_name)
            self.detail_env.config(text=app_data["env_path"] or "未设置")
            self.detail_path.config(text=app_data["app_path"])
            self.launch_btn.config(state=tk.NORMAL)
            self.delete_btn.config(state=tk.NORMAL)

    def launch_app(self):
        if not self.selected_app or self.selected_app not in self.apps:
            messagebox.showerror("错误", "没有选择应用!")
            return
        app_data = self.apps[self.selected_app]
        app_path = app_data["app_path"]
        env_path = app_data["env_path"]
        if not os.path.exists(app_path):
            messagebox.showerror("错误", f"应用路径不存在: {app_path}")
            return
        try:
            if sys.platform == "win32":
                if env_path:
                    subprocess.Popen([app_path], cwd=env_path, shell=True)
                else:
                    subprocess.Popen([app_path], shell=True)
            else:
                if env_path:
                    subprocess.Popen([app_path], cwd=env_path)
                else:
                    subprocess.Popen([app_path])
        except Exception as e:
            messagebox.showerror("错误", f"启动应用时出错: {str(e)}")

    def delete_app(self):
        if not self.selected_app:
            messagebox.showerror("错误", "没有选择应用!")
            return
        if messagebox.askyesno("确认", f"确定要删除应用 '{self.selected_app}' 吗?"):
            del self.apps[self.selected_app]
            self.app_tree.delete(self.selected_app)
            self.save_apps()
            self.detail_name.config(text="")
            self.detail_env.config(text="")
            self.detail_path.config(text="")
            self.launch_btn.config(state=tk.DISABLED)
            self.delete_btn.config(state=tk.DISABLED)
            self.selected_app = None
            messagebox.showinfo("成功", "应用已删除!")

    def save_apps(self):
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(self.apps, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("错误", f"保存应用数据时出错: {str(e)}")

    def load_apps(self):
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    self.apps = json.load(f)
            except Exception as e:
                messagebox.showerror("错误", f"加载应用数据时出错: {str(e)}")
                self.apps = {}
        else:
            self.apps = {}
            self.save_apps()

def main():
    root = tk.Tk()
    try:
        root.iconbitmap("al.ico")
    except:
        pass
    app = AppLauncher(root)
    root.mainloop()

if __name__ == "__main__":
    main()