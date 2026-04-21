import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import subprocess
import sys
import ctypes
from ctypes import Structure, c_void_p, POINTER, WINFUNCTYPE
from ctypes.wintypes import DWORD, LONG, WORD, UINT, HBITMAP, HDC, HGDIOBJ, RECT, INT, LPARAM, LPCSTR, HICON
import tempfile
from pathlib import Path
from io import BytesIO
import threading

# 尝试导入 inputs 库
try:
    from inputs import get_gamepad
    HAS_INPUTS = True
except ImportError:
    HAS_INPUTS = False

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

        # 触屏滚动相关变量
        self.last_y = 0
        
        # 加载保存的应用数据
        self.load_apps()
        
        # 创建UI
        self.create_widgets()

        # 启动手柄监听线程
        if HAS_INPUTS:
            self.stop_gamepad = False
            self.gamepad_thread = threading.Thread(target=self.handle_gamepad, daemon=True)
            self.gamepad_thread.start()

    # --- 手柄监听逻辑 (修正摇杆反向) ---
    def handle_gamepad(self):
        while not getattr(self, 'stop_gamepad', False):
            try:
                events = get_gamepad()
                for event in events:
                    # 十字键控制
                    if event.code == 'ABS_HAT0Y':
                        if event.state == -1: self.root.after(0, lambda: self.move_selection(-1))
                        elif event.state == 1: self.root.after(0, lambda: self.move_selection(1))
                    
                    # 摇杆控制 (修正：ABS_Y负值为向上推，对应列表-1；正值为向下推，对应列表+1)
                    elif event.code == 'ABS_Y':
                        if event.state < -16000: # 向上推
                            self.root.after(0, lambda: self.move_selection(1))
                        elif event.state > 16000: # 向下推
                            self.root.after(0, lambda: self.move_selection(-1))

                    # A键启动
                    elif event.code == 'BTN_SOUTH' and event.state == 1:
                        self.root.after(0, self.launch_app)
            except: pass

    def move_selection(self, direction):
        items = self.app_tree.get_children()
        if not items: return
        sel = self.app_tree.selection()
        if not sel:
            new_sel = items[0]
        else:
            idx = items.index(sel[0])
            new_idx = max(0, min(len(items) - 1, idx + direction))
            new_sel = items[new_idx]
        self.app_tree.selection_set(new_sel)
        self.app_tree.see(new_sel)
        self.on_app_select(None)

    # --- 触屏滑动逻辑 ---
    def on_touch_start(self, event):
        self.last_y = event.y

    def on_touch_scroll(self, event):
        delta = self.last_y - event.y
        self.app_tree.yview_scroll(int(delta/5), "units")
        self.last_y = event.y

    # --- 原始 UI 功能逻辑 ---
    def create_default_icon(self):
        width = height = 32
        rgb_data = bytearray(width * height * 3)
        for y in range(height):
            for x in range(width):
                idx = (y * width + x) * 3
                if 6 <= x <= 25 and 6 <= y <= 25:
                    rgb_data[idx] = 0; rgb_data[idx+1] = 0; rgb_data[idx+2] = 200
                else:
                    rgb_data[idx] = 240; rgb_data[idx+1] = 240; rgb_data[idx+2] = 240
        ppm_data = f"P6\n{width} {height}\n255\n".encode() + rgb_data
        self.default_icon = tk.PhotoImage(data=ppm_data, width=width, height=height)

    def extract_icon_with_ctypes(self, app_path, size=32):
        try:
            class BITMAPINFOHEADER(Structure):
                _fields_ = [("biSize", DWORD), ("biWidth", LONG), ("biHeight", LONG), ("biPlanes", WORD),
                            ("biBitCount", WORD), ("biCompression", DWORD), ("biSizeImage", DWORD),
                            ("biXPelsPerMeter", LONG), ("biYPelsPerMeter", LONG), ("biClrUsed", DWORD), ("biClrImportant", DWORD)]
            class BITMAPINFO(Structure):
                _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", DWORD * 3)]
            shell32 = ctypes.windll.shell32
            user32 = ctypes.windll.user32
            gdi32 = ctypes.windll.gdi32
            large_icons = ctypes.c_void_p()
            small_icons = ctypes.c_void_p()
            icon_count = shell32.ExtractIconExW(app_path, 0, ctypes.byref(large_icons), ctypes.byref(small_icons), 1)
            if icon_count == 0: return None
            hIcon = large_icons
            hdc = user32.GetDC(0)
            memdc = gdi32.CreateCompatibleDC(hdc)
            hbitmap = gdi32.CreateCompatibleBitmap(hdc, 32, 32)
            old_bitmap = gdi32.SelectObject(memdc, hbitmap)
            gdi32.PatBlt(memdc, 0, 0, 32, 32, 0x00F00021)
            user32.DrawIconEx(memdc, 0, 0, hIcon, 32, 32, 0, 0, 3)
            bmi = BITMAPINFO()
            bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bmi.bmiHeader.biPlanes = 1; bmi.bmiHeader.biBitCount = 32; bmi.bmiHeader.biWidth = 32; bmi.bmiHeader.biHeight = -32
            data = ctypes.create_string_buffer(32 * 32 * 4)
            gdi32.GetDIBits(memdc, hbitmap, 0, 32, data, ctypes.byref(bmi), 0)
            gdi32.SelectObject(memdc, old_bitmap); gdi32.DeleteObject(hbitmap); gdi32.DeleteDC(memdc)
            user32.ReleaseDC(0, hdc); user32.DestroyIcon(hIcon)
            raw_data = bytearray(data)
            rgb_data = bytearray(32 * 32 * 3)
            for i in range(32 * 32):
                alpha = raw_data[i*4 + 3] / 255.0
                rgb_data[i*3] = int(raw_data[i*4+2] * alpha + 255 * (1-alpha))
                rgb_data[i*3+1] = int(raw_data[i*4+1] * alpha + 255 * (1-alpha))
                rgb_data[i*3+2] = int(raw_data[i*4] * alpha + 255 * (1-alpha))
            return rgb_data
        except: return None

    def get_app_icon(self, app_path):
        if not app_path or not os.path.exists(app_path):
            if self.default_icon is None: self.create_default_icon()
            return self.default_icon
        if app_path in self.icon_cache: return self.icon_cache[app_path]
        rgb_data = self.extract_icon_with_ctypes(app_path)
        if rgb_data:
            ppm_data = f"P6\n32 32\n255\n".encode() + rgb_data
            icon = tk.PhotoImage(data=ppm_data, width=32, height=32)
            self.icon_cache[app_path] = icon
            return icon
        if self.default_icon is None: self.create_default_icon()
        return self.default_icon

    def create_widgets(self):
        # 100% 还原原始 UI 布局与配置
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", font=("Arial", 12), rowheight=40, background="#ffffff", fieldbackground="#ffffff")
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"), background="#f0f0f0")
        
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1); self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="应用启动器", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        # 应用名称行
        ttk.Label(main_frame, text="应用名称:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.app_name_var = tk.StringVar()
        self.app_name_entry = ttk.Entry(main_frame, textvariable=self.app_name_var, width=40, font=("Arial", 10))
        self.app_name_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=5)
        self.new_btn = ttk.Button(main_frame, text="新建", command=self.add_app)
        self.new_btn.grid(row=1, column=2, padx=(10, 0), pady=5)
        
        # 环境路径行
        ttk.Label(main_frame, text="启动环境路径:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.env_path_var = tk.StringVar()
        env_frame = ttk.Frame(main_frame)
        env_frame.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.env_entry = ttk.Entry(env_frame, textvariable=self.env_path_var, width=40, font=("Arial", 10))
        self.env_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(env_frame, text="浏览...", width=10, command=self.browse_env_path).pack(side=tk.LEFT, padx=(5, 0))
        
        # 应用路径行
        ttk.Label(main_frame, text="应用绝对路径:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.app_path_var = tk.StringVar()
        app_path_frame = ttk.Frame(main_frame)
        app_path_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.app_path_entry = ttk.Entry(app_path_frame, textvariable=self.app_path_var, width=40, font=("Arial", 10))
        self.app_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(app_path_frame, text="浏览...", width=10, command=self.browse_app_path).pack(side=tk.LEFT, padx=(5, 0))
        
        # 分割线
        ttk.Separator(main_frame, orient='horizontal').grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15)
        
        # 搜索行
        ttk.Label(main_frame, text="搜索应用:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_apps)
        self.search_entry = ttk.Entry(main_frame, textvariable=self.search_var, width=40, font=("Arial", 10))
        self.search_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=5)
        
        # 列表区域
        list_frame = ttk.LabelFrame(main_frame, text="应用列表", padding="5")
        list_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        list_frame.columnconfigure(0, weight=1); list_frame.rowconfigure(0, weight=1)
        
        self.app_tree = ttk.Treeview(list_frame, height=6, selectmode="browse", show='tree')
        self.app_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.app_tree.column("#0", width=400, stretch=True, anchor="w")
        
        # 触屏绑定
        self.app_tree.bind("<Button-1>", self.on_touch_start)
        self.app_tree.bind("<B1-Motion>", self.on_touch_scroll)
        self.app_tree.bind('<<TreeviewSelect>>', self.on_app_select)

        self.app_tree.bind('<Enter>', self.on_app_hover)
        self.app_tree.bind('<Motion>', self.on_app_hover)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.app_tree.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.app_tree.configure(yscrollcommand=scrollbar.set)
        
        # 按钮行
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)
        self.launch_btn = ttk.Button(button_frame, text="启动应用", command=self.launch_app, state=tk.DISABLED)
        self.launch_btn.pack(side=tk.LEFT, padx=5)
        self.delete_btn = ttk.Button(button_frame, text="删除应用", command=self.delete_app, state=tk.DISABLED)
        self.delete_btn.pack(side=tk.LEFT, padx=5)
        
        # 详情区域
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

    def browse_env_path(self):
        folder = filedialog.askdirectory()
        if folder: self.env_path_var.set(os.path.normpath(folder))
    
    def browse_app_path(self):
        file = filedialog.askopenfilename()
        if file: self.app_path_var.set(os.path.normpath(file))

    def add_app(self):
        name = self.app_name_var.get().strip()
        env = self.env_path_var.get().strip()
        path = self.app_path_var.get().strip()
        if name and path:
            self.apps[name] = {"env_path": env, "app_path": path}
            self.save_apps(); self.update_app_list()
            self.app_name_var.set(""); self.env_path_var.set(""); self.app_path_var.set("")

    def filter_apps(self, *args):
        self.update_app_list(self.search_var.get().lower())

    def update_app_list(self, filter_text=""):
        for i in self.app_tree.get_children(): self.app_tree.delete(i)
        for name in sorted(self.apps.keys()):
            if filter_text in name.lower():
                icon = self.get_app_icon(self.apps[name]["app_path"])
                self.app_tree.insert("", "end", iid=name, text="  "+name, image=icon)

    def on_app_hover(self, event):
        item = self.app_tree.identify_row(event.y)
        if item:
            app_name = item
            if app_name in self.apps:
                data = self.apps[app_name]
                self.detail_name.config(text=app_name)
                self.detail_env.config(text=data["env_path"] or "未设置")
                self.detail_path.config(text=data["app_path"])

    def on_app_select(self, event):
        sel = self.app_tree.selection()
        if sel:
            app_name = sel[0]
            self.selected_app = app_name
            data = self.apps[app_name]
            self.detail_name.config(text=app_name)
            self.detail_env.config(text=data["env_path"] or "未设置")
            self.detail_path.config(text=data["app_path"])
            self.launch_btn.config(state=tk.NORMAL)
            self.delete_btn.config(state=tk.NORMAL)

    def launch_app(self):
        if not self.selected_app: return
        data = self.apps[self.selected_app]
        try:
            subprocess.Popen([data["app_path"]], cwd=data["env_path"] or None, shell=True)
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def delete_app(self):
        if self.selected_app and messagebox.askyesno("确认", f"确定删除 '{self.selected_app}'?"):
            del self.apps[self.selected_app]
            self.save_apps(); self.update_app_list()
            self.detail_name.config(text=""); self.detail_env.config(text=""); self.detail_path.config(text="")
            self.launch_btn.config(state=tk.DISABLED); self.delete_btn.config(state=tk.DISABLED)
            self.selected_app = None

    def save_apps(self):
        with open(self.data_file, 'w', encoding='utf-8') as f:
            json.dump(self.apps, f, ensure_ascii=False, indent=2)

    def load_apps(self):
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    self.apps = json.load(f)
            except: self.apps = {}

if __name__ == "__main__":
    root = tk.Tk()
    app = AppLauncher(root)
    root.mainloop()