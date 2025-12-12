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
        
        # 加载保存的应用数据
        self.load_apps()
        
        # 创建UI
        self.create_widgets()
    
    def create_default_icon(self):
        """创建默认图标"""
        # 创建一个简单的默认图标 (32x32)
        width = height = 32
        rgb_data = bytearray(width * height * 3)
        
        # 创建一个简单的蓝色应用图标
        for y in range(height):
            for x in range(width):
                idx = (y * width + x) * 3
                # 创建一个简单的"应用"图标
                if 6 <= x <= 25 and 6 <= y <= 25:
                    # 蓝色方块
                    rgb_data[idx] = 0      # R
                    rgb_data[idx+1] = 0    # G
                    rgb_data[idx+2] = 200  # B
                else:
                    # 浅灰色背景
                    rgb_data[idx] = 240    # R
                    rgb_data[idx+1] = 240  # G
                    rgb_data[idx+2] = 240  # B
        
        # 转换为PhotoImage
        ppm_data = f"P6\n{width} {height}\n255\n".encode() + rgb_data
        self.default_icon = tk.PhotoImage(data=ppm_data, width=width, height=height)
    
    def extract_icon_with_ctypes(self, app_path, size=32):
        """使用ctypes提取exe文件图标 - 改进版本"""
        try:
            # 定义必要的结构
            class BITMAPINFOHEADER(Structure):
                _fields_ = [
                    ("biSize", DWORD),
                    ("biWidth", LONG),
                    ("biHeight", LONG),
                    ("biPlanes", WORD),
                    ("biBitCount", WORD),
                    ("biCompression", DWORD),
                    ("biSizeImage", DWORD),
                    ("biXPelsPerMeter", LONG),
                    ("biYPelsPerMeter", LONG),
                    ("biClrUsed", DWORD),
                    ("biClrImportant", DWORD),
                ]
            
            class BITMAPINFO(Structure):
                _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", DWORD * 3)]
            
            # 获取Windows API函数
            shell32 = ctypes.windll.shell32
            user32 = ctypes.windll.user32
            gdi32 = ctypes.windll.gdi32
            
            # 首先，尝试提取大图标（通常是32x32或更大）
            large_icons = ctypes.c_void_p()
            small_icons = ctypes.c_void_p()
            
            # 提取图标
            icon_count = shell32.ExtractIconExW(app_path, 0, ctypes.byref(large_icons), ctypes.byref(small_icons), 1)
            
            if icon_count == 0:
                return None
            
            # 使用大图标
            hIcon = large_icons
            
            # 获取图标尺寸
            icon_size = 32  # 默认32x32
            
            # 尝试获取更大的图标
            for test_size in [256, 128, 64, 48]:
                # 尝试提取指定大小的图标
                test_icon = shell32.ExtractIconExW(app_path, 0, None, None, 1)
                if test_icon > 0:
                    # 有图标，尝试加载
                    hTestIcon = ctypes.c_void_p()
                    shell32.ExtractIconExW(app_path, 0, ctypes.byref(hTestIcon), None, 1)
                    if hTestIcon:
                        # 检查图标尺寸
                        icon_info = ctypes.create_string_buffer(ctypes.sizeof(ctypes.c_void_p) * 5)
                        if user32.GetIconInfo(hTestIcon, ctypes.byref(icon_info)):
                            # 获取尺寸
                            icon_size = test_size
                            hIcon = hTestIcon
                            # 清理之前的图标
                            if hIcon != large_icons:
                                user32.DestroyIcon(large_icons)
            
            # 准备绘制图标
            hdc = user32.GetDC(0)
            memdc = gdi32.CreateCompatibleDC(hdc)
            
            # 使用提取到的图标尺寸
            width = height = min(icon_size, 32)  # 限制最大为32x32，避免太大
            
            # 创建位图
            hbitmap = gdi32.CreateCompatibleBitmap(hdc, width, height)
            old_bitmap = gdi32.SelectObject(memdc, hbitmap)
            
            # 清空背景为白色
            gdi32.PatBlt(memdc, 0, 0, width, height, 0x00F00021)  # PATCOPY | WHITENESS
            
            # 绘制图标
            user32.DrawIconEx(memdc, 0, 0, hIcon, width, height, 0, 0, 3)
            
            # 获取位图数据
            bmi = BITMAPINFO()
            bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bmi.bmiHeader.biPlanes = 1
            bmi.bmiHeader.biBitCount = 32
            bmi.bmiHeader.biCompression = 0
            bmi.bmiHeader.biClrUsed = 0
            bmi.bmiHeader.biClrImportant = 0
            bmi.bmiHeader.biWidth = width
            bmi.bmiHeader.biHeight = -height  # 负值表示从上到下的位图
            
            data_size = width * height * 4
            data = ctypes.create_string_buffer(data_size)
            
            result = gdi32.GetDIBits(memdc, hbitmap, 0, height, data, ctypes.byref(bmi), 0)
            
            # 清理资源
            gdi32.SelectObject(memdc, old_bitmap)
            gdi32.DeleteObject(hbitmap)
            gdi32.DeleteDC(memdc)
            user32.ReleaseDC(0, hdc)
            user32.DestroyIcon(hIcon)
            
            if result != height:
                return None
            
            # 转换BGRA到RGB
            raw_data = bytearray(data)
            rgb_data = bytearray(width * height * 3)
            
            for i in range(height * width):
                src_offset = i * 4
                dst_offset = i * 3
                
                # 获取alpha值
                alpha = raw_data[src_offset + 3] / 255.0
                
                # BGRA -> RGB，考虑alpha通道
                b = raw_data[src_offset]
                g = raw_data[src_offset + 1]
                r = raw_data[src_offset + 2]
                
                # 简单的alpha混合到白色背景
                rgb_data[dst_offset] = int(r * alpha + 255 * (1 - alpha))     # R
                rgb_data[dst_offset + 1] = int(g * alpha + 255 * (1 - alpha)) # G
                rgb_data[dst_offset + 2] = int(b * alpha + 255 * (1 - alpha)) # B
            
            return rgb_data
            
        except Exception as e:
            print(f"使用ctypes提取图标失败: {e}")
            return None
    
    def get_app_icon_simple(self, app_path):
        """简单的图标获取方法，尝试获取最高质量的图标"""
        try:
            # 使用SHGetFileInfo获取图标
            from ctypes import wintypes
            import struct
            
            # 定义结构体和常量
            SHGFI_ICON = 0x000000100
            SHGFI_LARGEICON = 0x000000000
            SHGFI_SMALLICON = 0x000000001
            
            class SHFILEINFO(ctypes.Structure):
                _fields_ = [
                    ("hIcon", ctypes.c_void_p),
                    ("iIcon", ctypes.c_int),
                    ("dwAttributes", ctypes.c_uint),
                    ("szDisplayName", ctypes.c_wchar * 260),
                    ("szTypeName", ctypes.c_wchar * 80)
                ]
            
            shell32 = ctypes.windll.shell32
            
            # 获取文件信息
            sfi = SHFILEINFO()
            result = shell32.SHGetFileInfoW(
                app_path, 0, ctypes.byref(sfi), ctypes.sizeof(sfi), 
                SHGFI_ICON | SHGFI_LARGEICON
            )
            
            if result and sfi.hIcon:
                # 创建临时ICO文件
                temp_dir = tempfile.gettempdir()
                temp_ico = os.path.join(temp_dir, f"temp_icon_{os.getpid()}.ico")
                
                # 保存图标到文件
                ico_data = self.icon_to_ico(sfi.hIcon)
                if ico_data:
                    with open(temp_ico, "wb") as f:
                        f.write(ico_data)
                    
                    # 加载为PhotoImage
                    icon = tk.PhotoImage(file=temp_ico)
                    
                    # 清理
                    os.remove(temp_ico)
                    ctypes.windll.user32.DestroyIcon(sfi.hIcon)
                    
                    return icon
                
                ctypes.windll.user32.DestroyIcon(sfi.hIcon)
        
        except Exception as e:
            print(f"使用SHGetFileInfo获取图标失败: {e}")
        
        return None
    
    def icon_to_ico(self, hIcon):
        """将图标句柄转换为ICO文件数据"""
        try:
            # 获取图标信息
            from ctypes import wintypes
            
            class ICONINFO(ctypes.Structure):
                _fields_ = [
                    ("fIcon", wintypes.BOOL),
                    ("xHotspot", wintypes.DWORD),
                    ("yHotspot", wintypes.DWORD),
                    ("hbmMask", wintypes.HBITMAP),
                    ("hbmColor", wintypes.HBITMAP)
                ]
            
            icon_info = ICONINFO()
            if not ctypes.windll.user32.GetIconInfo(hIcon, ctypes.byref(icon_info)):
                return None
            
            # 这里简化处理，实际需要将位图数据转换为ICO格式
            # 由于比较复杂，这里返回None，使用备用方法
            return None
            
        except Exception as e:
            print(f"转换图标失败: {e}")
            return None
    
    def get_app_icon(self, app_path):
        """获取应用程序图标 - 主方法"""
        if not app_path or not os.path.exists(app_path):
            if self.default_icon is None:
                self.create_default_icon()
            return self.default_icon
        
        # 检查缓存
        if app_path in self.icon_cache:
            return self.icon_cache[app_path]
        
        # 首先尝试简单方法
        icon = self.get_app_icon_simple(app_path)
        if icon:
            self.icon_cache[app_path] = icon
            return icon
        
        # 如果简单方法失败，使用ctypes方法
        try:
            rgb_data = self.extract_icon_with_ctypes(app_path, 32)
            if rgb_data:
                # 创建PPM格式数据
                ppm_data = f"P6\n32 32\n255\n".encode() + rgb_data
                icon = tk.PhotoImage(data=ppm_data, width=32, height=32)
                self.icon_cache[app_path] = icon
                return icon
        except Exception as e:
            print(f"获取图标失败 {app_path}: {e}")
        
        # 返回默认图标
        if self.default_icon is None:
            self.create_default_icon()
        return self.default_icon

    def create_widgets(self):
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 配置Treeview样式
        style.configure("Treeview", 
                      font=("Arial", 12),  # 增大字体
                      rowheight=40,  # 增加行高
                      background="#ffffff",
                      fieldbackground="#ffffff")
        
        style.configure("Treeview.Heading", 
                       font=("Arial", 12, "bold"),
                       background="#f0f0f0")
        
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="应用启动器", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        # 第一行：应用名称输入框和新建按钮
        ttk.Label(main_frame, text="应用名称:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.app_name_var = tk.StringVar()
        self.app_name_entry = ttk.Entry(main_frame, textvariable=self.app_name_var, width=40, font=("Arial", 10))
        self.app_name_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=5)
        
        self.new_btn = ttk.Button(main_frame, text="新建", command=self.add_app)
        self.new_btn.grid(row=1, column=2, padx=(10, 0), pady=5)
        
        # 第二行：启动路径（环境）
        ttk.Label(main_frame, text="启动环境路径:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.env_path_var = tk.StringVar()
        env_frame = ttk.Frame(main_frame)
        env_frame.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.env_entry = ttk.Entry(env_frame, textvariable=self.env_path_var, width=40, font=("Arial", 10))
        self.env_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.env_browse_btn = ttk.Button(env_frame, text="浏览...", width=10, command=self.browse_env_path)
        self.env_browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # 第三行：应用绝对路径
        ttk.Label(main_frame, text="应用绝对路径:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.app_path_var = tk.StringVar()
        app_path_frame = ttk.Frame(main_frame)
        app_path_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.app_path_entry = ttk.Entry(app_path_frame, textvariable=self.app_path_var, width=40, font=("Arial", 10))
        self.app_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.app_browse_btn = ttk.Button(app_path_frame, text="浏览...", width=10, command=self.browse_app_path)
        self.app_browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # 分隔线
        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15)
        
        # 搜索框
        ttk.Label(main_frame, text="搜索应用:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_apps)  # 绑定搜索事件
        self.search_entry = ttk.Entry(main_frame, textvariable=self.search_var, width=40, font=("Arial", 10))
        self.search_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), padx=(5, 0), pady=5)
        
        # 应用列表
        list_frame = ttk.LabelFrame(main_frame, text="应用列表", padding="5")
        list_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 使用Treeview，设置show='tree'隐藏表头
        self.app_tree = ttk.Treeview(list_frame, height=6, selectmode="browse", show='tree')
        self.app_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置列宽
        self.app_tree.column("#0", width=400, stretch=True, anchor="w")
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.app_tree.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.app_tree.configure(yscrollcommand=scrollbar.set)
        
        # 绑定选择事件
        self.app_tree.bind('<<TreeviewSelect>>', self.on_app_select)
        # 绑定鼠标悬停事件
        self.app_tree.bind('<Enter>', self.on_app_hover)
        self.app_tree.bind('<Motion>', self.on_app_hover)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)
        
        self.launch_btn = ttk.Button(button_frame, text="启动应用", command=self.launch_app, state=tk.DISABLED)
        self.launch_btn.pack(side=tk.LEFT, padx=5)
        
        self.delete_btn = ttk.Button(button_frame, text="删除应用", command=self.delete_app, state=tk.DISABLED)
        self.delete_btn.pack(side=tk.LEFT, padx=5)
        
        # 应用详情显示区域
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
        
        # 初始化应用列表
        self.update_app_list()
    
    def browse_env_path(self):
        """浏览并选择环境路径"""
        folder = filedialog.askdirectory(title="选择环境路径")
        if folder:
            self.env_path_var.set(folder)
    
    def browse_app_path(self):
        """浏览并选择应用路径"""
        file = filedialog.askopenfilename(title="选择应用程序")
        if file:
            self.app_path_var.set(file)
    
    def add_app(self):
        """添加新应用到列表"""
        app_name = self.app_name_var.get().strip()
        env_path = self.env_path_var.get().strip()
        app_path = self.app_path_var.get().strip()
        
        if not app_name:
            messagebox.showerror("错误", "应用名称不能为空!")
            return
        
        if not app_path:
            messagebox.showerror("错误", "应用路径不能为空!")
            return
        
        # 检查应用是否已存在
        if app_name in self.apps:
            messagebox.showerror("错误", f"应用 '{app_name}' 已存在!")
            return
        
        # 添加到应用字典
        self.apps[app_name] = {
            "env_path": env_path,
            "app_path": app_path
        }
        
        # 保存到文件
        self.save_apps()
        
        # 更新列表
        self.update_app_list()
        
        # 清空输入框
        self.app_name_var.set("")
        self.env_path_var.set("")
        self.app_path_var.set("")
        
        messagebox.showinfo("成功", f"应用 '{app_name}' 已添加!")
    
    def filter_apps(self, *args):
        """根据搜索框内容过滤应用列表"""
        search_text = self.search_var.get().lower()
        self.update_app_list(search_text)
    
    def update_app_list(self, filter_text=""):
        """更新应用列表"""
        # 清空Treeview
        for item in self.app_tree.get_children():
            self.app_tree.delete(item)
        
        # 获取排序后的应用名称
        app_names = sorted(self.apps.keys())
        
        # 根据筛选条件添加应用
        for app_name in app_names:
            if filter_text in app_name.lower():
                app_data = self.apps[app_name]
                
                # 获取应用图标
                icon = self.get_app_icon(app_data["app_path"])
                
                # 添加到Treeview
                self.app_tree.insert("", "end", 
                                   iid=app_name,
                                   text="  "+app_name,  # 显示应用名称
                                   image=icon)     # 显示图标
    
    def on_app_select(self, event):
        """当从列表中选择应用时触发"""
        selection = self.app_tree.selection()
        if selection:
            app_name = selection[0]
            self.select_app(app_name)
    
    def on_app_hover(self, event):
        """鼠标悬停在应用列表上时显示应用详情"""
        # 获取鼠标位置对应的项目
        item = self.app_tree.identify_row(event.y)
        if item:
            app_name = item
            if app_name in self.apps:
                app_data = self.apps[app_name]
                self.detail_name.config(text=app_name)
                self.detail_env.config(text=app_data["env_path"] or "未设置")
                self.detail_path.config(text=app_data["app_path"])
    
    def select_app(self, app_name):
        """选择特定应用"""
        if app_name in self.apps:
            self.selected_app = app_name
            app_data = self.apps[app_name]
            
            # 更新详情显示
            self.detail_name.config(text=app_name)
            self.detail_env.config(text=app_data["env_path"] or "未设置")
            self.detail_path.config(text=app_data["app_path"])
            
            # 启用按钮
            self.launch_btn.config(state=tk.NORMAL)
            self.delete_btn.config(state=tk.NORMAL)
    
    def launch_app(self):
        """启动选中的应用"""
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
            # 根据操作系统选择启动方式
            if sys.platform == "win32":
                # Windows系统
                if env_path:
                    # 如果有环境路径，设置工作目录
                    subprocess.Popen([app_path], cwd=env_path, shell=True)
                else:
                    subprocess.Popen([app_path], shell=True)
            else:
                # macOS 或 Linux
                if env_path:
                    subprocess.Popen([app_path], cwd=env_path)
                else:
                    subprocess.Popen([app_path])
            
        except Exception as e:
            messagebox.showerror("错误", f"启动应用时出错: {str(e)}")
    
    def delete_app(self):
        """删除选中的应用"""
        if not self.selected_app:
            messagebox.showerror("错误", "没有选择应用!")
            return
        
        # 确认删除
        if messagebox.askyesno("确认", f"确定要删除应用 '{self.selected_app}' 吗?"):
            # 从字典中删除
            del self.apps[self.selected_app]
            
            # 从Treeview中删除
            self.app_tree.delete(self.selected_app)
            
            # 保存到文件
            self.save_apps()
            
            # 清空详情显示
            self.detail_name.config(text="")
            self.detail_env.config(text="")
            self.detail_path.config(text="")
            
            # 禁用启动和删除按钮
            self.launch_btn.config(state=tk.DISABLED)
            self.delete_btn.config(state=tk.DISABLED)
            
            # 清除选择
            self.selected_app = None
            
            messagebox.showinfo("成功", "应用已删除!")
    
    def save_apps(self):
        """保存应用到JSON文件"""
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(self.apps, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("错误", f"保存应用数据时出错: {str(e)}")
    
    def load_apps(self):
        """从JSON文件加载应用"""
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
    
    # 设置窗口图标
    try:
        root.iconbitmap("al.ico")
    except:
        pass
    
    app = AppLauncher(root)
    root.mainloop()

if __name__ == "__main__":
    main()