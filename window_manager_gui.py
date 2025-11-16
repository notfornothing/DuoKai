#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
可视化窗口管理工具
支持拖拽设置窗口位置，保存和加载配置
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ctypes
from ctypes import wintypes, windll
import json
import os
import subprocess
import sys
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
import uuid

def is_admin():
    """检查是否以管理员身份运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """以管理员身份重新运行程序"""
    if is_admin():
        return True
    else:
        try:
            # 以管理员身份重新运行
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
            return False
        except:
            messagebox.showerror("错误", "无法获取管理员权限")
            return False

# 现代化UI配色方案
COLORS = {
    'bg_primary': '#2b2b2b',      # 主背景色
    'bg_secondary': '#3c3c3c',    # 次要背景色
    'bg_accent': '#404040',       # 强调背景色
    'fg_primary': '#ffffff',      # 主文字色
    'fg_secondary': '#cccccc',    # 次要文字色
    'accent_blue': '#0078d4',     # 蓝色强调
    'accent_green': '#107c10',    # 绿色强调
    'accent_orange': '#ff8c00',   # 橙色强调
    'accent_red': '#d13438',      # 红色强调
    'border': '#555555',          # 边框色
    'hover': '#4a4a4a',           # 悬停色
    'selected': '#0078d4',        # 选中色
}

# Windows API 常量
SW_RESTORE = 9
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
HWND_TOP = 0

# --- Virtual Desktop (Windows 10) COM 定义 ---
class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]

def _guid_from_string(s: str) -> GUID:
    u = uuid.UUID(s)
    data = u.bytes_le
    g = GUID()
    g.Data1 = int.from_bytes(data[0:4], "little")
    g.Data2 = int.from_bytes(data[4:6], "little")
    g.Data3 = int.from_bytes(data[6:8], "little")
    g.Data4 = (ctypes.c_ubyte * 8).from_buffer_copy(data[8:16])
    return g

CLSID_VIRTUAL_DESKTOP_MANAGER = _guid_from_string("{AA509086-5CA9-4C25-8F95-589D3C07B48A}")
IID_IVIRTUAL_DESKTOP_MANAGER   = _guid_from_string("{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")

CLSCTX_INPROC_SERVER = 0x1
COINIT_APARTMENTTHREADED = 0x2

class VirtualDesktopManagerWrapper:
    """使用 ctypes 直接调用 IVirtualDesktopManager 接口"""
    def __init__(self):
        ole32 = ctypes.windll.ole32
        # 初始化 COM（APARTMENT 模式）
        ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED)

        # 创建对象实例
        self._obj = ctypes.c_void_p()
        hr = ole32.CoCreateInstance(
            ctypes.byref(CLSID_VIRTUAL_DESKTOP_MANAGER),
            None,
            CLSCTX_INPROC_SERVER,
            ctypes.byref(IID_IVIRTUAL_DESKTOP_MANAGER),
            ctypes.byref(self._obj)
        )
        if hr != 0 or not self._obj.value:
            raise OSError(f"CoCreateInstance 失败, HRESULT=0x{(hr & 0xFFFFFFFF):08X}")

        # 取 vtable 指针
        vtbl_ptr = ctypes.cast(self._obj, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents

        # IUnknown 有 3 个函数在前面，接口方法从索引 3 开始
        IsOnProto = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, wintypes.HWND, ctypes.POINTER(wintypes.BOOL))
        GetIdProto = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, wintypes.HWND, ctypes.POINTER(GUID))
        MoveProto  = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, wintypes.HWND, ctypes.POINTER(GUID))

        self._IsWindowOnCurrentVirtualDesktop = IsOnProto(vtbl_ptr[3])
        self._GetWindowDesktopId = GetIdProto(vtbl_ptr[4])
        self._MoveWindowToDesktop = MoveProto(vtbl_ptr[5])

    def IsWindowOnCurrentVirtualDesktop(self, hwnd: int) -> bool:
        flag = wintypes.BOOL()
        hr = self._IsWindowOnCurrentVirtualDesktop(self._obj.value, hwnd, ctypes.byref(flag))
        if hr != 0:
            raise OSError(f"IsWindowOnCurrentVirtualDesktop 失败, HRESULT=0x{(hr & 0xFFFFFFFF):08X}")
        return bool(flag.value)

    def GetWindowDesktopId(self, hwnd: int) -> GUID:
        gid = GUID()
        hr = self._GetWindowDesktopId(self._obj.value, hwnd, ctypes.byref(gid))
        if hr != 0:
            raise OSError(f"GetWindowDesktopId 失败, HRESULT=0x{(hr & 0xFFFFFFFF):08X}")
        return gid

    def MoveWindowToDesktop(self, hwnd: int, desktop_guid: GUID) -> None:
        hr = self._MoveWindowToDesktop(self._obj.value, hwnd, ctypes.byref(desktop_guid))
        if hr != 0:
            raise OSError(f"MoveWindowToDesktop 失败, HRESULT=0x{(hr & 0xFFFFFFFF):08X}")

class WindowInfo:
    """窗口信息类"""
    def __init__(self, hwnd: int, title: str, class_name: str):
        self.hwnd = hwnd
        self.title = title
        self.class_name = class_name
        self.assigned_position = None  # (row, col) 或 None
    
    def __str__(self):
        return f"{self.title} ({self.class_name})"

@dataclass
class SandboxConfig:
    """沙盒配置数据类"""
    sandbox_path: str = r"D:\Install\sandbox\Sandboxie-Plus\Start.exe"
    program_path: str = r"C:\Install\menghuanDesk\menghuanxiyoushikong"
    program_exe: str = "MyLauncher_x64r.exe"
    box_prefix: str = "01"
    box_count: int = 6
    enabled_boxes: List[str] = None
    
    def __post_init__(self):
        if self.enabled_boxes is None:
            self.enabled_boxes = [f"{int(self.box_prefix):02d}", f"{int(self.box_prefix)+1:02d}"]

class WindowManagerGUI:
    """可视化窗口管理器"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🪟 窗口管理器 & 沙盒多开工具")
        self.root.geometry("1200x800")
        self.root.configure(bg=COLORS['bg_primary'])
        
        # 设置现代化样式
        self.setup_styles()
        
        # 配置变量
        self.rows = tk.IntVar(value=2)
        self.columns = tk.IntVar(value=3)
        self.screen_width = tk.IntVar(value=2560)
        self.screen_height = tk.IntVar(value=1440)
        self.use_workarea = tk.BooleanVar(value=True)
        self.h_gap = tk.IntVar(value=10)  # 左右间隙
        self.v_gap = tk.IntVar(value=10)  # 上下间隙

        # 数据
        self.windows: List[WindowInfo] = []
        # 布局组（联动沙盒分组）：01-06、07-12、13-18、19-24
        self.layout_groups = ["01-06", "07-12", "13-18", "19-24"]
        self.layout_group_var = tk.StringVar(value=self.layout_groups[0])
        self.group_assignments: Dict[str, Dict[Tuple[int, int], WindowInfo]] = {
            g: {} for g in self.layout_groups
        }
        
        # 沙盒配置
        self.sandbox_config = SandboxConfig()
        self.sandbox_config_file = os.path.join("saving", "multiSandbox.json")
        self.window_config_file = os.path.join("saving", "multiWindows.json")
        
        # GUI 组件
        self.window_listbox = None
        self.grid_frame = None
        self.grid_buttons = {}  # (row, col) -> Button
        
        # 拖拽相关
        self.drag_data = {"item": None, "source": None}
        
        self.setup_ui()
        
        # 加载配置
        self.load_sandbox_config()
        
        self.refresh_windows()

    def get_current_assignments(self) -> Dict[Tuple[int, int], WindowInfo]:
        return self.group_assignments[self.layout_group_var.get()]

    def on_layout_group_change(self, *_):
        # 切换当前布局组时，更新网格显示与窗口分配标记
        self.update_grid_display()
        # 同步列表中分配状态
        assignments = self.get_current_assignments()
        for w in self.windows:
            w.assigned_position = None
        for (r, c), w in assignments.items():
            if w:
                w.assigned_position = (r, c)
        self.refresh_windows()
    
    def setup_styles(self):
        """设置现代化样式"""
        style = ttk.Style()
        
        # 配置主题样式
        style.configure('Modern.TFrame', background=COLORS['bg_secondary'])
        style.configure('Modern.TLabel', 
                       background=COLORS['bg_secondary'], 
                       foreground=COLORS['fg_primary'],
                       font=('Segoe UI', 10))
        style.configure('Modern.TEntry',
                       fieldbackground=COLORS['bg_accent'],
                       borderwidth=1,
                       insertcolor=COLORS['fg_primary'])
        style.configure('Modern.TSpinbox',
                       fieldbackground=COLORS['bg_accent'],
                       borderwidth=1,
                       arrowcolor=COLORS['fg_primary'])
        style.configure('Modern.TCheckbutton',
                       background=COLORS['bg_secondary'],
                       foreground=COLORS['fg_primary'],
                       focuscolor='none')
        
        # 简化LabelFrame样式，避免布局错误
        style.configure('Modern.TLabelFrame',
                       background=COLORS['bg_secondary'],
                       borderwidth=1,
                       relief='solid')
        style.configure('Modern.TLabelFrame.Label',
                       background=COLORS['bg_secondary'],
                       foreground=COLORS['fg_primary'],
                       font=('Segoe UI', 11, 'bold'))
    
    def setup_ui(self):
        """设置用户界面"""
        # 设置根窗口背景
        self.root.configure(bg=COLORS['bg_secondary'])
        
        # 主框架
        main_frame = tk.Frame(self.root, bg=COLORS['bg_secondary'], padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 顶部切换（替换选项卡）：两个按钮在同一页显示
        toggle_frame = tk.Frame(main_frame, bg=COLORS['bg_secondary'])
        toggle_frame.pack(fill=tk.X, pady=(0, 10))

        self.window_tab = tk.Frame(main_frame, bg=COLORS['bg_secondary'])
        self.sandbox_tab = tk.Frame(main_frame, bg=COLORS['bg_secondary'])

        def show_window_tab():
            self.sandbox_tab.pack_forget()
            self.window_tab.pack(fill=tk.BOTH, expand=True)

        def show_sandbox_tab():
            self.window_tab.pack_forget()
            self.sandbox_tab.pack(fill=tk.BOTH, expand=True)

        btn_window = tk.Button(
            toggle_frame,
            text="窗口管理",
            command=show_window_tab,
            bg=COLORS['accent_blue'],
            fg=COLORS['fg_primary'],
            font=('Segoe UI', 10, 'bold'),
            borderwidth=0,
            padx=12,
            pady=6
        )
        btn_window.pack(side=tk.LEFT, padx=(0, 8))
        self.add_hover_effect(btn_window, COLORS['accent_blue'])

        btn_sandbox = tk.Button(
            toggle_frame,
            text="沙盒多开",
            command=show_sandbox_tab,
            bg=COLORS['accent_orange'],
            fg=COLORS['fg_primary'],
            font=('Segoe UI', 10, 'bold'),
            borderwidth=0,
            padx=12,
            pady=6
        )
        btn_sandbox.pack(side=tk.LEFT)
        self.add_hover_effect(btn_sandbox, COLORS['accent_orange'])
        
        # 设置窗口管理界面
        self.setup_window_management_ui()

        # 设置沙盒多开界面
        self.setup_sandbox_ui()

        # 默认显示窗口管理
        show_window_tab()
    
    def setup_window_management_ui(self):
        """设置窗口管理界面"""
        # 配置区域
        config_frame = tk.LabelFrame(self.window_tab, text="⚙️ 配置设置", 
                                    bg=COLORS['bg_secondary'], fg=COLORS['fg_primary'],
                                    font=('Segoe UI', 11, 'bold'), padx=10, pady=10)
        config_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 网格设置
        grid_config_frame = tk.Frame(config_frame, bg=COLORS['bg_secondary'])
        grid_config_frame.pack(fill=tk.X)
        
        # 行数设置
        tk.Label(grid_config_frame, text="行数:", bg=COLORS['bg_secondary'], 
                fg=COLORS['fg_primary'], font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 5))
        rows_spinbox = tk.Spinbox(grid_config_frame, from_=1, to=5, textvariable=self.rows, 
                                 width=5, command=self.update_grid, bg=COLORS['bg_accent'],
                                 fg=COLORS['fg_primary'], font=('Segoe UI', 10))
        rows_spinbox.pack(side=tk.LEFT, padx=(0, 15))
        
        # 列数设置
        tk.Label(grid_config_frame, text="列数:", bg=COLORS['bg_secondary'], 
                fg=COLORS['fg_primary'], font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 5))
        cols_spinbox = tk.Spinbox(grid_config_frame, from_=1, to=5, textvariable=self.columns, 
                                 width=5, command=self.update_grid, bg=COLORS['bg_accent'],
                                 fg=COLORS['fg_primary'], font=('Segoe UI', 10))
        cols_spinbox.pack(side=tk.LEFT, padx=(0, 15))
        
        # 分辨率设置
        tk.Label(grid_config_frame, text="宽度:", bg=COLORS['bg_secondary'], 
                fg=COLORS['fg_primary'], font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 5))
        width_entry = tk.Entry(grid_config_frame, textvariable=self.screen_width, width=8,
                              bg=COLORS['bg_accent'], fg=COLORS['fg_primary'], font=('Segoe UI', 10))
        width_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        tk.Label(grid_config_frame, text="高度:", bg=COLORS['bg_secondary'], 
                fg=COLORS['fg_primary'], font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 5))
        height_entry = tk.Entry(grid_config_frame, textvariable=self.screen_height, width=8,
                               bg=COLORS['bg_accent'], fg=COLORS['fg_primary'], font=('Segoe UI', 10))
        height_entry.pack(side=tk.LEFT, padx=(0, 15))

        # 间隙设置
        tk.Label(grid_config_frame, text="左右间隙:", bg=COLORS['bg_secondary'],
                 fg=COLORS['fg_primary'], font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 5))
        hgap_spinbox = tk.Spinbox(grid_config_frame, from_=-50, to=200, textvariable=self.h_gap,
                                  width=5, bg=COLORS['bg_accent'], fg=COLORS['fg_primary'], font=('Segoe UI', 10))
        hgap_spinbox.pack(side=tk.LEFT, padx=(0, 15))

        tk.Label(grid_config_frame, text="上下间隙:", bg=COLORS['bg_secondary'],
                 fg=COLORS['fg_primary'], font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(0, 5))
        vgap_spinbox = tk.Spinbox(grid_config_frame, from_=-50, to=200, textvariable=self.v_gap,
                                  width=5, bg=COLORS['bg_accent'], fg=COLORS['fg_primary'], font=('Segoe UI', 10))
        vgap_spinbox.pack(side=tk.LEFT, padx=(0, 15))

        # 布局组选择（联动沙盒分组）
        tk.Label(grid_config_frame, text="布局组:", bg=COLORS['bg_secondary'],
                 fg=COLORS['fg_primary'], font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=(10, 5))
        group_menu = tk.OptionMenu(grid_config_frame, self.layout_group_var, *self.layout_groups, command=self.on_layout_group_change)
        group_menu.config(bg=COLORS['bg_accent'], fg=COLORS['fg_primary'], highlightthickness=0)
        group_menu.pack(side=tk.LEFT)
        
        # 工作区选项
        workarea_check = tk.Checkbutton(grid_config_frame, text="使用工作区(避开任务栏)", 
                                       variable=self.use_workarea, bg=COLORS['bg_secondary'],
                                       fg=COLORS['fg_primary'], font=('Segoe UI', 10),
                                       selectcolor=COLORS['bg_accent'])
        workarea_check.pack(side=tk.LEFT)
        
        # 内容区域
        content_frame = tk.Frame(self.window_tab, bg=COLORS['bg_secondary'])
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧：窗口列表
        left_frame = tk.LabelFrame(content_frame, text="📋 可用窗口 (拖拽到右侧网格)", 
                                  bg=COLORS['bg_secondary'], fg=COLORS['fg_primary'],
                                  font=('Segoe UI', 11, 'bold'), padx=10, pady=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 窗口列表容器
        list_container = tk.Frame(left_frame, bg=COLORS['bg_secondary'])
        list_container.pack(fill=tk.BOTH, expand=True)
        
        # 自定义Listbox样式
        self.window_listbox = tk.Listbox(
            list_container, 
            bg=COLORS['bg_accent'],
            fg=COLORS['fg_primary'],
            selectbackground=COLORS['selected'],
            selectforeground=COLORS['fg_primary'],
            borderwidth=0,
            highlightthickness=0,
            font=('Segoe UI', 10),
            height=20
        )
        
        scrollbar = tk.Scrollbar(list_container, orient=tk.VERTICAL, bg=COLORS['bg_accent'])
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.window_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.window_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.window_listbox.yview)
        
        # 绑定拖拽事件
        self.window_listbox.bind('<Button-1>', self.on_listbox_click)
        self.window_listbox.bind('<B1-Motion>', self.on_listbox_drag)
        self.window_listbox.bind('<ButtonRelease-1>', self.on_listbox_release)
        self.window_listbox.bind('<Double-1>', self.on_window_double_click)
        
        # 刷新按钮
        refresh_btn = tk.Button(
            left_frame, 
            text="🔄 刷新窗口列表",
            command=self.refresh_windows,
            bg=COLORS['accent_green'],
            fg=COLORS['fg_primary'],
            font=('Segoe UI', 10, 'bold'),
            borderwidth=0,
            pady=8
        )
        refresh_btn.pack(pady=(10, 0), fill=tk.X)
        
        # 右侧：网格布局
        right_frame = tk.LabelFrame(content_frame, text="🎯 网格布局 (放置窗口)", 
                                   bg=COLORS['bg_secondary'], fg=COLORS['fg_primary'],
                                   font=('Segoe UI', 11, 'bold'), padx=10, pady=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.grid_frame = tk.Frame(right_frame, bg=COLORS['bg_secondary'])
        self.grid_frame.pack(fill=tk.BOTH, expand=True)
        
        # 底部按钮区域
        button_frame = tk.Frame(self.window_tab, bg=COLORS['bg_secondary'])
        button_frame.pack(pady=(15, 0))
        
        # 创建现代化按钮
        buttons_config = [
            ("👁️ 预览布局", self.preview_layout, COLORS['accent_blue']),
            ("✅ 应用设置", self.apply_layout, COLORS['accent_green']),
            ("🗑️ 清空设置", self.clear_assignments, COLORS['accent_red']),
            ("💾 保存配置", self.save_config, COLORS['accent_orange']),
            ("📂 加载配置", self.load_config, COLORS['accent_orange'])
        ]
        
        for i, (text, command, color) in enumerate(buttons_config):
            btn = tk.Button(
                button_frame,
                text=text,
                command=command,
                bg=color,
                fg=COLORS['fg_primary'],
                font=('Segoe UI', 10, 'bold'),
                borderwidth=0,
                padx=15,
                pady=8
            )
            btn.pack(side=tk.LEFT, padx=5)
            self.add_hover_effect(btn, color)
        
        # 初始化网格
        self.update_grid()

        # 操作日志区域
        self.window_status_frame = tk.LabelFrame(
            self.window_tab,
            text="📜 操作日志",
            bg=COLORS['bg_secondary'],
            fg=COLORS['fg_primary'],
            font=('Segoe UI', 11, 'bold'),
            padx=10,
            pady=10
        )
        self.window_status_frame.pack(fill=tk.BOTH, expand=True, pady=(15, 0))

        self.window_status_text = tk.Text(
            self.window_status_frame,
            height=8,
            bg=COLORS['bg_accent'],
            fg=COLORS['fg_primary'],
            font=('Consolas', 9),
            borderwidth=0,
            wrap=tk.WORD
        )
        window_status_scrollbar = tk.Scrollbar(self.window_status_frame, orient=tk.VERTICAL)
        window_status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.window_status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.window_status_text.config(yscrollcommand=window_status_scrollbar.set)
        window_status_scrollbar.config(command=self.window_status_text.yview)
    
    def setup_sandbox_ui(self):
        """设置沙盒多开界面"""
        # 配置区域
        config_frame = tk.LabelFrame(self.sandbox_tab, text="⚙️ 沙盒配置", 
                                    bg=COLORS['bg_secondary'], fg=COLORS['fg_primary'],
                                    font=('Segoe UI', 11, 'bold'), padx=15, pady=15)
        config_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 沙盒路径设置
        path_frame = tk.Frame(config_frame, bg=COLORS['bg_secondary'])
        path_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(path_frame, text="沙盒程序路径:", bg=COLORS['bg_secondary'], 
                fg=COLORS['fg_primary'], font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W)
        
        sandbox_path_frame = tk.Frame(path_frame, bg=COLORS['bg_secondary'])
        sandbox_path_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.sandbox_path_var = tk.StringVar(value=self.sandbox_config.sandbox_path)
        sandbox_path_entry = tk.Entry(sandbox_path_frame, textvariable=self.sandbox_path_var,
                                     bg=COLORS['bg_accent'], fg=COLORS['fg_primary'], 
                                     font=('Segoe UI', 10))
        sandbox_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        browse_sandbox_btn = tk.Button(sandbox_path_frame, text="浏览...",
                                      command=self.browse_sandbox_path,
                                      bg=COLORS['accent_blue'], fg=COLORS['fg_primary'],
                                      font=('Segoe UI', 9), borderwidth=0, padx=10)
        browse_sandbox_btn.pack(side=tk.RIGHT)
        
        # 程序路径设置
        program_frame = tk.Frame(config_frame, bg=COLORS['bg_secondary'])
        program_frame.pack(fill=tk.X, pady=(10, 10))
        
        tk.Label(program_frame, text="目标程序目录:", bg=COLORS['bg_secondary'], 
                fg=COLORS['fg_primary'], font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W)
        
        program_path_frame = tk.Frame(program_frame, bg=COLORS['bg_secondary'])
        program_path_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.program_path_var = tk.StringVar(value=self.sandbox_config.program_path)
        program_path_entry = tk.Entry(program_path_frame, textvariable=self.program_path_var,
                                     bg=COLORS['bg_accent'], fg=COLORS['fg_primary'], 
                                     font=('Segoe UI', 10))
        program_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        browse_program_btn = tk.Button(program_path_frame, text="浏览...",
                                      command=self.browse_program_path,
                                      bg=COLORS['accent_blue'], fg=COLORS['fg_primary'],
                                      font=('Segoe UI', 9), borderwidth=0, padx=10)
        browse_program_btn.pack(side=tk.RIGHT)
        
        # 程序可执行文件
        exe_frame = tk.Frame(config_frame, bg=COLORS['bg_secondary'])
        exe_frame.pack(fill=tk.X, pady=(10, 10))
        
        tk.Label(exe_frame, text="可执行文件名:", bg=COLORS['bg_secondary'], 
                fg=COLORS['fg_primary'], font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=(0, 10))
        
        self.program_exe_var = tk.StringVar(value=self.sandbox_config.program_exe)
        exe_entry = tk.Entry(exe_frame, textvariable=self.program_exe_var, width=30,
                            bg=COLORS['bg_accent'], fg=COLORS['fg_primary'], 
                            font=('Segoe UI', 10))
        exe_entry.pack(side=tk.LEFT)
        
        # Box配置区域
        box_frame = tk.LabelFrame(self.sandbox_tab, text="📦 Box配置", 
                                 bg=COLORS['bg_secondary'], fg=COLORS['fg_primary'],
                                 font=('Segoe UI', 11, 'bold'), padx=15, pady=15)
        box_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Box选择区域
        box_select_frame = tk.Frame(box_frame, bg=COLORS['bg_secondary'])
        box_select_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(box_select_frame, text="选择要启动的 Box (01-24):", 
                bg=COLORS['bg_secondary'], fg=COLORS['fg_primary'], 
                font=('Segoe UI', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))

        # 创建分组 Box 复选框（01-06、07-12、13-18、19-24），每组一行并带批量按钮
        self.box_vars = {}

        def create_group_row(parent, start_idx: int, end_idx: int):
            group_frame = tk.Frame(parent, bg=COLORS['bg_secondary'])
            group_frame.pack(fill=tk.X, pady=(0, 8))

            # 左侧：组标签 + 复选框们
            left_frame = tk.Frame(group_frame, bg=COLORS['bg_secondary'])
            left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

            tk.Label(
                left_frame,
                text=f"组 {start_idx:02d}-{end_idx:02d}",
                bg=COLORS['bg_secondary'],
                fg=COLORS['fg_primary'],
                font=('Segoe UI', 10, 'bold')
            ).pack(side=tk.LEFT, padx=(0, 10))

            group_ids = [f"{i:02d}" for i in range(start_idx, end_idx + 1)]
            for box_id in group_ids:
                var = tk.BooleanVar(value=box_id in getattr(self.sandbox_config, 'enabled_boxes', []))
                self.box_vars[box_id] = var
                cb = tk.Checkbutton(
                    left_frame,
                    text=f"Box {box_id}",
                    variable=var,
                    bg=COLORS['bg_secondary'],
                    fg=COLORS['fg_primary'],
                    font=('Segoe UI', 10),
                    selectcolor=COLORS['bg_accent']
                )
                cb.pack(side=tk.LEFT, padx=(0, 14))

            # 右侧：批量选择按钮（全选 / 全不选 / 只选前五个）
            right_frame = tk.Frame(group_frame, bg=COLORS['bg_secondary'])
            right_frame.pack(side=tk.RIGHT)

            def apply_group_selection(ids: List[str], mode: str):
                if mode == 'all':
                    for bid in ids:
                        self.box_vars[bid].set(True)
                elif mode == 'none':
                    for bid in ids:
                        self.box_vars[bid].set(False)
                elif mode == 'first5':
                    for idx, bid in enumerate(ids):
                        self.box_vars[bid].set(idx < 5)
                # 联动：切换窗口布局到该组
                group_key = f"{start_idx:02d}-{end_idx:02d}"
                self.layout_group_var.set(group_key)
                self.on_layout_group_change()

            btn_first5 = tk.Button(
                right_frame,
                text="只选前五个",
                command=lambda ids=group_ids: apply_group_selection(ids, 'first5'),
                bg=COLORS['accent_blue'],
                fg=COLORS['fg_primary'],
                font=('Segoe UI', 9),
                borderwidth=0,
                padx=10
            )
            btn_first5.pack(side=tk.RIGHT, padx=(5, 0))
            self.add_hover_effect(btn_first5, COLORS['accent_blue'])

            btn_none = tk.Button(
                right_frame,
                text="全不选",
                command=lambda ids=group_ids: apply_group_selection(ids, 'none'),
                bg=COLORS['accent_red'],
                fg=COLORS['fg_primary'],
                font=('Segoe UI', 9),
                borderwidth=0,
                padx=10
            )
            btn_none.pack(side=tk.RIGHT, padx=(5, 0))
            self.add_hover_effect(btn_none, COLORS['accent_red'])

            btn_all = tk.Button(
                right_frame,
                text="全选",
                command=lambda ids=group_ids: apply_group_selection(ids, 'all'),
                bg=COLORS['accent_green'],
                fg=COLORS['fg_primary'],
                font=('Segoe UI', 9),
                borderwidth=0,
                padx=10
            )
            btn_all.pack(side=tk.RIGHT, padx=(5, 0))
            self.add_hover_effect(btn_all, COLORS['accent_green'])

        # 组行：01-06、07-12、13-18、19-24
        create_group_row(box_select_frame, 1, 6)
        create_group_row(box_select_frame, 7, 12)
        create_group_row(box_select_frame, 13, 18)
        create_group_row(box_select_frame, 19, 24)
        
        # 快速选择按钮
        quick_select_frame = tk.Frame(box_frame, bg=COLORS['bg_secondary'])
        quick_select_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(quick_select_frame, text="快速选择:", bg=COLORS['bg_secondary'], 
                fg=COLORS['fg_primary'], font=('Segoe UI', 10, 'bold')).pack(side=tk.LEFT, padx=(0, 10))
        
        select_all_btn = tk.Button(quick_select_frame, text="全选",
                                  command=self.select_all_boxes,
                                  bg=COLORS['accent_green'], fg=COLORS['fg_primary'],
                                  font=('Segoe UI', 9), borderwidth=0, padx=10)
        select_all_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        select_none_btn = tk.Button(quick_select_frame, text="全不选",
                                   command=self.select_no_boxes,
                                   bg=COLORS['accent_red'], fg=COLORS['fg_primary'],
                                   font=('Segoe UI', 9), borderwidth=0, padx=10)
        select_none_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        select_first_two_btn = tk.Button(quick_select_frame, text="选择前两个",
                                        command=self.select_first_two_boxes,
                                        bg=COLORS['accent_blue'], fg=COLORS['fg_primary'],
                                        font=('Segoe UI', 9), borderwidth=0, padx=10)
        select_first_two_btn.pack(side=tk.LEFT)
        
        # 启动按钮区域
        launch_frame = tk.Frame(self.sandbox_tab, bg=COLORS['bg_secondary'])
        launch_frame.pack(pady=(0, 15))
        
        # 按钮顺序：启动选中的沙盒 -> 关闭所有沙盒 -> 保存配置 -> 加载配置

        launch_btn = tk.Button(
            launch_frame,
            text="🚀 启动选中的沙盒",
            command=self.launch_sandboxes,
            bg=COLORS['accent_green'],
            fg=COLORS['fg_primary'],
            font=('Segoe UI', 12, 'bold'),
            borderwidth=0,
            padx=30,
            pady=10
        )
        launch_btn.pack(side=tk.LEFT, padx=(0, 15))
        self.add_hover_effect(launch_btn, COLORS['accent_green'])

        terminate_all_btn = tk.Button(
            launch_frame,
            text="🛑 关闭所有沙盒",
            command=self.terminate_all_sandboxes,
            bg=COLORS['accent_red'],
            fg=COLORS['fg_primary'],
            font=('Segoe UI', 10, 'bold'),
            borderwidth=0,
            padx=20,
            pady=10
        )
        terminate_all_btn.pack(side=tk.LEFT, padx=(0, 15))
        self.add_hover_effect(terminate_all_btn, COLORS['accent_red'])

        save_sandbox_config_btn = tk.Button(
            launch_frame,
            text="💾 保存配置",
            command=self.save_sandbox_config,
            bg=COLORS['accent_orange'],
            fg=COLORS['fg_primary'],
            font=('Segoe UI', 10, 'bold'),
            borderwidth=0,
            padx=20,
            pady=10
        )
        save_sandbox_config_btn.pack(side=tk.LEFT, padx=(0, 15))
        self.add_hover_effect(save_sandbox_config_btn, COLORS['accent_orange'])

        load_sandbox_config_btn = tk.Button(
            launch_frame,
            text="📂 加载配置",
            command=self.browse_and_load_sandbox_config,
            bg=COLORS['accent_orange'],
            fg=COLORS['fg_primary'],
            font=('Segoe UI', 10, 'bold'),
            borderwidth=0,
            padx=20,
            pady=10
        )
        load_sandbox_config_btn.pack(side=tk.LEFT, padx=(0, 0))
        self.add_hover_effect(load_sandbox_config_btn, COLORS['accent_orange'])

        # 状态显示区域
        self.sandbox_status_frame = tk.LabelFrame(self.sandbox_tab, text="📊 启动状态", 
                                                 bg=COLORS['bg_secondary'], fg=COLORS['fg_primary'],
                                                 font=('Segoe UI', 11, 'bold'), padx=15, pady=15)
        self.sandbox_status_frame.pack(fill=tk.BOTH, expand=True)
        
        self.sandbox_status_text = tk.Text(self.sandbox_status_frame, height=8,
                                          bg=COLORS['bg_accent'], fg=COLORS['fg_primary'],
                                          font=('Consolas', 9), borderwidth=0,
                                          wrap=tk.WORD)
        
        status_scrollbar = tk.Scrollbar(self.sandbox_status_frame, orient=tk.VERTICAL)
        status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.sandbox_status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sandbox_status_text.config(yscrollcommand=status_scrollbar.set)
        status_scrollbar.config(command=self.sandbox_status_text.yview)
        
        # 初始状态信息
        self.sandbox_status_text.insert(tk.END, "准备启动沙盒...\n")
        self.sandbox_status_text.insert(tk.END, f"当前配置:\n")
        self.sandbox_status_text.insert(tk.END, f"  沙盒路径: {self.sandbox_config.sandbox_path}\n")
        self.sandbox_status_text.insert(tk.END, f"  程序路径: {self.sandbox_config.program_path}\n")
        self.sandbox_status_text.insert(tk.END, f"  可执行文件: {self.sandbox_config.program_exe}\n\n")
    
    # 沙盒相关方法
    def browse_sandbox_path(self):
        """浏览沙盒程序路径"""
        filename = filedialog.askopenfilename(
            title="选择沙盒程序",
            filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
        )
        if filename:
            self.sandbox_path_var.set(filename)
    
    def browse_program_path(self):
        """浏览目标程序目录"""
        dirname = filedialog.askdirectory(title="选择目标程序目录")
        if dirname:
            self.program_path_var.set(dirname)

    def browse_and_load_sandbox_config(self):
        """浏览并加载沙盒配置文件(JSON)，并记录加载路径"""
        filename = filedialog.askopenfilename(
            title="选择沙盒配置文件",
            filetypes=[("JSON 配置", "*.json"), ("所有文件", "*.*")]
        )
        if not filename:
            return
        try:
            self.sandbox_config_file = filename
            self.load_sandbox_config()
            if hasattr(self, 'sandbox_status_text') and self.sandbox_status_text:
                self.sandbox_status_text.insert(tk.END, f"📂 已加载沙盒配置文件: {filename}\n")
                self.sandbox_status_text.see(tk.END)
            self.show_status_message("沙盒配置已加载")
        except Exception as e:
            messagebox.showerror("加载失败", f"加载沙盒配置时出错: {e}")
    
    def select_all_boxes(self):
        """选择所有Box"""
        for var in self.box_vars.values():
            var.set(True)
    
    def select_no_boxes(self):
        """取消选择所有Box"""
        for var in self.box_vars.values():
            var.set(False)
    
    def select_first_two_boxes(self):
        """选择前两个Box"""
        self.select_no_boxes()
        for i, box_id in enumerate(sorted(self.box_vars.keys())):
            if i < 2:
                self.box_vars[box_id].set(True)
    
    def get_selected_boxes(self) -> List[str]:
        """获取选中的Box列表"""
        return [box_id for box_id, var in self.box_vars.items() if var.get()]
    
    def launch_sandboxes(self):
        """启动选中的沙盒"""
        selected_boxes = self.get_selected_boxes()
        if not selected_boxes:
            messagebox.showwarning("警告", "请至少选择一个Box!")
            return
        
        # 更新配置
        self.sandbox_config.sandbox_path = self.sandbox_path_var.get()
        self.sandbox_config.program_path = self.program_path_var.get()
        self.sandbox_config.program_exe = self.program_exe_var.get()
        self.sandbox_config.enabled_boxes = selected_boxes
        
        # 清空状态显示
        self.sandbox_status_text.delete(1.0, tk.END)
        self.sandbox_status_text.insert(tk.END, f"开始启动 {len(selected_boxes)} 个沙盒...\n\n")
        self.sandbox_status_text.update()
        
        success_count = 0
        for box_id in selected_boxes:
            try:
                # 构建完整的程序路径
                full_program_path = os.path.join(self.sandbox_config.program_path, 
                                               self.sandbox_config.program_exe)
                # 基本存在性检查，避免误报“未知错误”
                if not os.path.isfile(full_program_path):
                    self.sandbox_status_text.insert(tk.END, f"❌ 可执行文件不存在: {full_program_path}\n\n")
                    continue
                
                # 构建启动命令
                command = [
                    self.sandbox_config.sandbox_path,
                    f"/box:{box_id}",
                    full_program_path
                ]
                
                self.sandbox_status_text.insert(tk.END, f"启动 Box {box_id}...\n")
                self.sandbox_status_text.insert(tk.END, f"命令: {' '.join(command)}\n")
                self.sandbox_status_text.update()
                
                # 启动并依据返回码判断（Start.exe 通常快速返回）
                result = subprocess.run(
                    command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                rc = result.returncode
                if rc == 0:
                    self.sandbox_status_text.insert(tk.END, f"✅ Box {box_id} 启动成功!\n\n")
                    success_count += 1
                else:
                    # 更准确的错误信息：优先 stderr，其次 stdout，最后返回码
                    stderr_msg = result.stderr.decode('utf-8', errors='ignore') if result.stderr else ""
                    stdout_msg = result.stdout.decode('utf-8', errors='ignore') if result.stdout else ""
                    error_msg = stderr_msg or stdout_msg or f"返回码 {rc}"
                    self.sandbox_status_text.insert(tk.END, f"❌ Box {box_id} 启动失败: {error_msg}\n\n")
                
            except Exception as e:
                self.sandbox_status_text.insert(tk.END, f"❌ Box {box_id} 启动异常: {str(e)}\n\n")
            
            self.sandbox_status_text.update()
        
        # 显示总结
        self.sandbox_status_text.insert(tk.END, f"启动完成! 成功: {success_count}/{len(selected_boxes)}\n")
        self.sandbox_status_text.see(tk.END)
        
        if success_count > 0:
            self.show_status_message("成功启动 {success_count} 个沙盒!")
        else:
            self.show_status_message("没有成功启动任何沙盒，请检查配置!")

    def terminate_all_sandboxes(self):
        """一键关闭所有沙盒窗口(通过 Sandboxie Start.exe /terminate_all)"""
        try:
            self.sandbox_config.sandbox_path = self.sandbox_path_var.get()
            cmd = [self.sandbox_config.sandbox_path, "/terminate_all"]
            if hasattr(self, 'sandbox_status_text') and self.sandbox_status_text:
                self.sandbox_status_text.insert(tk.END, f"🛑 发送终止命令: {' '.join(cmd)}\n")
                self.sandbox_status_text.see(tk.END)
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            _, stderr = process.communicate(timeout=5)
            if process.returncode == 0:
                if hasattr(self, 'sandbox_status_text') and self.sandbox_status_text:
                    self.sandbox_status_text.insert(tk.END, "✅ 已请求终止所有沙盒进程\n")
                    self.sandbox_status_text.see(tk.END)
                self.show_status_message("已请求关闭所有沙盒")
            else:
                err = stderr.decode('utf-8', errors='ignore') if stderr else f"错误码 {process.returncode}"
                if hasattr(self, 'sandbox_status_text') and self.sandbox_status_text:
                    self.sandbox_status_text.insert(tk.END, f"❌ 终止失败: {err}\n")
                    self.sandbox_status_text.see(tk.END)
                messagebox.showerror("终止失败", err)
        except Exception as e:
            if hasattr(self, 'sandbox_status_text') and self.sandbox_status_text:
                self.sandbox_status_text.insert(tk.END, f"❌ 终止异常: {e}\n")
                self.sandbox_status_text.see(tk.END)
            messagebox.showerror("终止异常", str(e))
    
    def save_sandbox_config(self):
        """保存沙盒配置"""
        try:
            # 更新配置
            self.sandbox_config.sandbox_path = self.sandbox_path_var.get()
            self.sandbox_config.program_path = self.program_path_var.get()
            self.sandbox_config.program_exe = self.program_exe_var.get()
            self.sandbox_config.enabled_boxes = self.get_selected_boxes()
            
            # 保存到文件
            os.makedirs("saving", exist_ok=True)
            config_data = {
                "sandbox": {
                    "sandbox_path": self.sandbox_config.sandbox_path,
                    "program_path": self.sandbox_config.program_path,
                    "program_exe": self.sandbox_config.program_exe,
                    "enabled_boxes": self.sandbox_config.enabled_boxes
                }
            }
            
            with open(self.sandbox_config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2, ensure_ascii=False)
            
            # 不显示弹窗，只在状态栏显示
            self.show_status_message("沙盒配置已保存")
            
        except Exception as e:
            messagebox.showerror("保存失败", f"保存配置时出错: {str(e)}")
    
    def load_sandbox_config(self):
        """加载沙盒配置"""
        try:
            if os.path.exists(self.sandbox_config_file):
                with open(self.sandbox_config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                
                if "sandbox" in config_data:
                    sandbox_config = config_data["sandbox"]
                    
                    # 更新界面
                    self.sandbox_path_var.set(sandbox_config.get("sandbox_path", ""))
                    self.program_path_var.set(sandbox_config.get("program_path", ""))
                    self.program_exe_var.set(sandbox_config.get("program_exe", ""))
                    
                    # 更新Box选择
                    enabled_boxes = sandbox_config.get("enabled_boxes", [])
                    for box_id, var in self.box_vars.items():
                        var.set(box_id in enabled_boxes)
                    
                    # 更新配置对象
                    self.sandbox_config.sandbox_path = sandbox_config.get("sandbox_path", "")
                    self.sandbox_config.program_path = sandbox_config.get("program_path", "")
                    self.sandbox_config.program_exe = sandbox_config.get("program_exe", "")
                    self.sandbox_config.enabled_boxes = enabled_boxes

                    if hasattr(self, 'sandbox_status_text') and self.sandbox_status_text:
                        self.sandbox_status_text.insert(tk.END, f"📂 从文件加载沙盒配置: {self.sandbox_config_file}\n")
                        self.sandbox_status_text.see(tk.END)
            
        except Exception as e:
            messagebox.showerror("加载失败", f"加载配置时出错: {str(e)}")
    
    # 窗口管理相关方法
    def add_hover_effect(self, button, original_color):
        """为按钮添加悬停效果"""
        def on_enter(e):
            button.config(bg=COLORS['hover'])
        
        def on_leave(e):
            button.config(bg=original_color)
        
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    # （已移除）虚拟桌面移动相关功能
    
    def get_windows(self) -> List[WindowInfo]:
        """获取所有可见窗口"""
        windows = []
        current_pid = os.getpid()
        
        # 定义回调函数类型
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        
        def enum_windows_proc(hwnd, lParam):
            if not windll.user32.IsWindowVisible(hwnd):
                return True
            if windll.user32.IsIconic(hwnd):  # 最小化
                return True
            
            # 获取窗口信息
            length = windll.user32.GetWindowTextLengthW(hwnd)
            if length == 0:
                return True
            
            buffer = ctypes.create_unicode_buffer(length + 1)
            windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
            title = buffer.value
            
            # 获取类名
            class_buffer = ctypes.create_unicode_buffer(256)
            windll.user32.GetClassNameW(hwnd, class_buffer, 256)
            class_name = class_buffer.value
            
            # 过滤系统窗口和工具窗口
            ex_style = windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            if ex_style & WS_EX_TOOLWINDOW:
                return True
            
            # 过滤当前程序窗口
            pid = wintypes.DWORD()
            windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value == current_pid:
                return True
            
            windows.append(WindowInfo(hwnd, title, class_name))
            return True
        
        enum_proc = WNDENUMPROC(enum_windows_proc)
        windll.user32.EnumWindows(enum_proc, 0)
        return windows
    
    def refresh_windows(self):
        """刷新窗口列表"""
        self.windows = self.get_windows()
        
        # 更新列表框
        self.window_listbox.delete(0, tk.END)
        for window in self.windows:
            status = ""
            if window.assigned_position:
                row, col = window.assigned_position
                status = f" → [{row+1},{col+1}]"
            self.window_listbox.insert(tk.END, f"{window.title}{status}")
    
    def update_grid(self):
        """更新网格显示"""
        # 清空现有网格
        for widget in self.grid_frame.winfo_children():
            widget.destroy()
        self.grid_buttons.clear()
        
        rows = self.rows.get()
        cols = self.columns.get()
        
        # 创建网格按钮
        for r in range(rows):
            for c in range(cols):
                btn = tk.Button(
                    self.grid_frame,
                    text=f"位置 {r+1},{c+1}\n(空)",
                    width=15,
                    height=4,
                    relief=tk.RAISED,
                    bg='lightgray',
                    command=lambda row=r, col=c: self.on_grid_click(row, col)
                )
                btn.grid(row=r, column=c, padx=2, pady=2, sticky=(tk.W, tk.E, tk.N, tk.S))
                btn.grid_position = (r, c)  # 添加位置属性用于拖拽
                self.grid_buttons[(r, c)] = btn
                
                # 添加悬停效果
                self.add_grid_hover_effect(btn)
                
                # 配置网格权重
                self.grid_frame.columnconfigure(c, weight=1)
                self.grid_frame.rowconfigure(r, weight=1)
        
        # 更新按钮显示
        self.update_grid_display()
    
    def update_grid_display(self):
        """更新网格按钮显示"""
        assignments = self.get_current_assignments()
        for (r, c), btn in self.grid_buttons.items():
            if (r, c) in assignments:
                window = assignments[(r, c)]
                btn.config(
                    text=f"🪟 位置 {r+1},{c+1}\n{window.title[:15]}...",
                    bg=COLORS['selected'],
                    fg=COLORS['fg_primary'],
                    font=('Segoe UI', 9, 'bold')
                )
            else:
                btn.config(
                    text=f"📍 位置 {r+1},{c+1}\n(空)",
                    bg=COLORS['bg_accent'],
                    fg=COLORS['fg_secondary'],
                    font=('Segoe UI', 9)
                )
    
    def add_grid_hover_effect(self, button):
        """为网格按钮添加悬停效果"""
        def on_enter(e):
            current_bg = button.cget('bg')
            if current_bg == COLORS['selected']:
                button.config(bg=COLORS['hover'])
            else:
                button.config(bg=COLORS['accent_blue'], fg=COLORS['fg_primary'])
        
        def on_leave(e):
            # 恢复原始颜色
            pos = button.grid_position
            if pos in self.grid_assignments:
                button.config(bg=COLORS['selected'], fg=COLORS['fg_primary'])
            else:
                button.config(bg=COLORS['bg_accent'], fg=COLORS['fg_secondary'])
        
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)
    
    def on_listbox_click(self, event):
        """处理列表框点击事件"""
        index = self.window_listbox.nearest(event.y)
        if index >= 0 and index < self.window_listbox.size():
            self.window_listbox.selection_clear(0, tk.END)
            self.window_listbox.selection_set(index)
            self.drag_data['start_index'] = index
            self.drag_data['dragging'] = False
    
    def on_listbox_drag(self, event):
        """处理拖拽事件"""
        if 'start_index' in self.drag_data:
            self.drag_data['dragging'] = True
            # 创建拖拽视觉反馈
            if not hasattr(self, 'drag_label'):
                self.drag_label = tk.Toplevel(self.root)
                self.drag_label.wm_overrideredirect(True)
                self.drag_label.configure(bg=COLORS['accent_blue'])
                
                # 获取被拖拽的窗口名称
                index = self.drag_data['start_index']
                if index < len(self.windows):
                    window_title = self.windows[index].title[:30] + "..." if len(self.windows[index].title) > 30 else self.windows[index].title
                    label = tk.Label(self.drag_label, text=f"📋 {window_title}", 
                                   bg=COLORS['accent_blue'], fg=COLORS['fg_primary'],
                                   font=('Segoe UI', 9, 'bold'), padx=10, pady=5)
                    label.pack()
            
            # 更新拖拽标签位置
            x = self.root.winfo_pointerx() + 10
            y = self.root.winfo_pointery() + 10
            self.drag_label.geometry(f"+{x}+{y}")
    
    def on_listbox_release(self, event):
        """处理拖拽释放事件"""
        if hasattr(self, 'drag_label'):
            self.drag_label.destroy()
            delattr(self, 'drag_label')
        
        if self.drag_data.get('dragging', False):
            # 检查是否释放在网格上
            widget = event.widget.winfo_containing(self.root.winfo_pointerx(), 
                                                  self.root.winfo_pointery())
            
            # 查找网格按钮
            target_button = None
            while widget and widget != self.root:
                if hasattr(widget, 'grid_position'):
                    target_button = widget
                    break
                widget = widget.master
            
            if target_button and 'start_index' in self.drag_data:
                # 执行拖拽分配
                index = self.drag_data['start_index']
                if index < len(self.windows):
                    row, col = target_button.grid_position
                    self.assign_window_to_position(self.windows[index], row, col)
        
        # 重置拖拽数据
        self.drag_data = {}
    
    def show_status_message(self, message, duration=2000):
        """显示状态消息"""
        if hasattr(self, 'status_label'):
            self.status_label.destroy()
        
        self.status_label = tk.Label(
            self.root,
            text=message,
            bg=COLORS['accent_green'],
            fg=COLORS['fg_primary'],
            font=('Segoe UI', 10, 'bold'),
            padx=15,
            pady=8
        )
        
        # 计算位置（屏幕中央下方）
        self.root.update_idletasks()
        x = self.root.winfo_x() + self.root.winfo_width() // 2 - 150
        y = self.root.winfo_y() + self.root.winfo_height() - 100
        
        self.status_label.place(x=x-self.root.winfo_x(), y=y-self.root.winfo_y())
        
        # 自动隐藏
        self.root.after(duration, lambda: self.status_label.destroy() if hasattr(self, 'status_label') else None)
    
    def on_window_double_click(self, event):
        """窗口列表双击事件"""
        selection = self.window_listbox.curselection()
        if not selection:
            return
        
        window_idx = selection[0]
        if window_idx >= len(self.windows):
            return
        
        window = self.windows[window_idx]
        
        # 弹出位置选择对话框
        self.select_position_for_window(window)
    
    def select_position_for_window(self, window: WindowInfo):
        """为窗口选择位置"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"选择位置 - {window.title}")
        dialog.geometry("300x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text=f"为窗口选择位置:\n{window.title}").pack(pady=10)
        
        # 位置选择
        pos_frame = ttk.Frame(dialog)
        pos_frame.pack(pady=10)
        
        ttk.Label(pos_frame, text="行:").grid(row=0, column=0, padx=5)
        row_var = tk.IntVar(value=1)
        row_spin = ttk.Spinbox(pos_frame, from_=1, to=self.rows.get(), textvariable=row_var, width=5)
        row_spin.grid(row=0, column=1, padx=5)
        
        ttk.Label(pos_frame, text="列:").grid(row=0, column=2, padx=5)
        col_var = tk.IntVar(value=1)
        col_spin = ttk.Spinbox(pos_frame, from_=1, to=self.columns.get(), textvariable=col_var, width=5)
        col_spin.grid(row=0, column=3, padx=5)
        
        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        
        def assign():
            row = row_var.get() - 1
            col = col_var.get() - 1
            self.assign_window_to_position(window, row, col)
            dialog.destroy()
        
        def remove():
            self.remove_window_assignment(window)
            dialog.destroy()
        
        ttk.Button(btn_frame, text="分配", command=assign).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="移除", command=remove).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def on_grid_click(self, row: int, col: int):
        """网格按钮点击事件"""
        assignments = self.get_current_assignments()
        if (row, col) in assignments:
            # 已有窗口，询问是否移除
            window = assignments[(row, col)]
            if messagebox.askyesno("移除窗口", f"是否移除窗口 '{window.title}' 从位置 [{row+1},{col+1}]?"):
                self.remove_window_assignment(window)
        else:
            # 空位置，选择窗口分配
            self.select_window_for_position(row, col)
    
    def select_window_for_position(self, row: int, col: int):
        """为位置选择窗口"""
        available_windows = [w for w in self.windows if w.assigned_position is None]
        if not available_windows:
            messagebox.showinfo("提示", "没有可用的窗口")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title(f"选择窗口 - 位置 [{row+1},{col+1}]")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text=f"为位置 [{row+1},{col+1}] 选择窗口:").pack(pady=10)
        
        # 窗口列表
        listbox = tk.Listbox(dialog, height=10)
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        for window in available_windows:
            listbox.insert(tk.END, window.title)
        
        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        def assign():
            selection = listbox.curselection()
            if selection:
                window = available_windows[selection[0]]
                self.assign_window_to_position(window, row, col)
                dialog.destroy()
        
        ttk.Button(btn_frame, text="分配", command=assign).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def assign_window_to_position(self, window: WindowInfo, row: int, col: int):
        """分配窗口到位置"""
        assignments = self.get_current_assignments()
        # 移除窗口的旧分配
        if window.assigned_position:
            old_pos = window.assigned_position
            if old_pos in assignments:
                del assignments[old_pos]
        
        # 移除位置的旧窗口
        if (row, col) in assignments:
            old_window = assignments[(row, col)]
            old_window.assigned_position = None
        
        # 新分配
        window.assigned_position = (row, col)
        assignments[(row, col)] = window
        
        # 更新显示
        self.update_grid_display()
        self.refresh_windows()
        
        # 显示成功提示
        self.show_status_message(f"✅ 已将 '{window.title[:20]}...' 分配到位置 ({row+1}, {col+1})")
    
    def remove_window_assignment(self, window: WindowInfo):
        """移除窗口分配"""
        assignments = self.get_current_assignments()
        if window.assigned_position:
            pos = window.assigned_position
            if pos in assignments:
                del assignments[pos]
            window.assigned_position = None
            
            self.update_grid_display()
            self.refresh_windows()
    
    def clear_assignments(self):
        """清空所有分配"""
        if messagebox.askyesno("确认", "是否清空当前布局组的窗口分配?"):
            assignments = self.get_current_assignments()
            assignments.clear()
            for window in self.windows:
                window.assigned_position = None
            self.update_grid_display()
            self.refresh_windows()
    
    def preview_layout(self):
        """预览布局"""
        if not self.get_current_assignments():
            messagebox.showinfo("提示", "没有分配任何窗口")
            return
        
        # 计算位置
        positions = self.calculate_positions()
        
        # 显示预览信息
        preview_text = "布局预览:\n\n"
        for (row, col), window in self.get_current_assignments().items():
            x, y, w, h = positions[(row, col)]
            preview_text += f"位置 [{row+1},{col+1}]: {window.title}\n"
            preview_text += f"  坐标: ({x}, {y}), 大小: {w}×{h}\n\n"
        
        messagebox.showinfo("布局预览", preview_text)
    
    def calculate_positions(self) -> Dict[Tuple[int, int], Tuple[int, int, int, int]]:
        """计算网格位置"""
        if self.use_workarea.get():
            # 使用工作区
            rect = wintypes.RECT()
            windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0)  # SPI_GETWORKAREA
            screen_w = rect.right - rect.left
            screen_h = rect.bottom - rect.top
            offset_x = rect.left
            offset_y = rect.top
        else:
            # 使用全屏
            screen_w = self.screen_width.get()
            screen_h = self.screen_height.get()
            offset_x = 0
            offset_y = 0
        
        rows = self.rows.get()
        cols = self.columns.get()
        
        # 计算每个格子的大小
        cell_w = screen_w // cols
        cell_h = screen_h // rows
        
        # 处理余数
        extra_w = screen_w % cols
        extra_h = screen_h % rows
        
        positions = {}
        # 允许负值间隙用于抵消窗口边框/阴影造成的视觉缝隙
        h_gap = int(self.h_gap.get())
        v_gap = int(self.v_gap.get())
        half_h = h_gap // 2
        half_v = v_gap // 2
        for row in range(rows):
            for col in range(cols):
                x = offset_x + col * cell_w + min(col, extra_w)
                y = offset_y + row * cell_h + min(row, extra_h)
                w = cell_w + (1 if col < extra_w else 0)
                h = cell_h + (1 if row < extra_h else 0)
                # 仅在相邻边缘应用间隙，避免影响外边界
                left_off = half_h if col > 0 else 0
                right_off = half_h if col < cols - 1 else 0
                top_off = half_v if row > 0 else 0
                bottom_off = half_v if row < rows - 1 else 0

                adj_x = x + left_off
                adj_y = y + top_off
                adj_w = max(0, w - (left_off + right_off))
                adj_h = max(0, h - (top_off + bottom_off))
                positions[(row, col)] = (adj_x, adj_y, adj_w, adj_h)
        
        return positions
    
    def apply_layout(self):
        """应用布局"""
        if not self.get_current_assignments():
            self.show_status_message("没有分配任何窗口")
            return
        
        positions = self.calculate_positions()
        success_count = 0
        
        for (row, col), window in self.get_current_assignments().items():
            x, y, w, h = positions[(row, col)]
            
            try:
                # 恢复窗口（如果最小化）
                windll.user32.ShowWindow(window.hwnd, SW_RESTORE)
                
                # 移动和调整大小
                windll.user32.SetWindowPos(
                    window.hwnd, HWND_TOP, x, y, w, h,
                    0x0040  # SWP_SHOWWINDOW
                )
                success_count += 1
            except Exception as e:
                print(f"移动窗口失败 {window.title}: {e}")
        
        self.show_status_message(f"成功应用 {success_count}/{len(self.get_current_assignments())} 个窗口的布局")
    
    def save_config(self):
        """保存配置"""
        # 允许保存即使当前组为空，以便记录多组配置
        if not any(self.group_assignments.values()):
            self.show_status_message("没有配置可保存")
            return
        
        config = {
            'rows': self.rows.get(),
            'columns': self.columns.get(),
            'screen_width': self.screen_width.get(),
            'screen_height': self.screen_height.get(),
            'use_workarea': self.use_workarea.get(),
            'h_gap': self.h_gap.get(),
            'v_gap': self.v_gap.get(),
            # 新增：按组保存布局
            'group_assignments': {}
        }

        for group_key, assignments in self.group_assignments.items():
            group_data = {}
            for (row, col), window in assignments.items():
                group_data[f"{row},{col}"] = {
                    'title': window.title,
                    'class_name': window.class_name
                }
            config['group_assignments'][group_key] = group_data

        # 兼容旧版：同时写当前组到旧字段 'assignments'
        current_assignments = self.get_current_assignments()
        legacy = {}
        for (row, col), window in current_assignments.items():
            legacy[f"{row},{col}"] = {
                'title': window.title,
                'class_name': window.class_name
            }
        config['assignments'] = legacy
        
        try:
            os.makedirs("saving", exist_ok=True)
            with open(self.window_config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.show_status_message("窗口配置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

    def load_config(self):
        """加载配置"""
        try:
            if not os.path.exists(self.window_config_file):
                self.show_status_message("配置文件不存在")
                return
                
            # 记录加载路径日志
            if hasattr(self, 'window_status_text') and self.window_status_text:
                self.window_status_text.insert(tk.END, f"📂 加载窗口配置文件: {self.window_config_file}\n")
                self.window_status_text.see(tk.END)

            with open(self.window_config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 加载设置
            self.rows.set(config.get('rows', 2))
            self.columns.set(config.get('columns', 3))
            self.screen_width.set(config.get('screen_width', 2560))
            self.screen_height.set(config.get('screen_height', 1440))
            self.use_workarea.set(config.get('use_workarea', True))
            self.h_gap.set(config.get('h_gap', 10))
            self.v_gap.set(config.get('v_gap', 10))
            
            # 更新网格
            self.update_grid()
            
            # 加载分组布局（新格式）或旧格式
            matched_count = 0
            group_assignments_cfg = config.get('group_assignments')
            if group_assignments_cfg:
                # 先清空所有组
                for g in self.layout_groups:
                    self.group_assignments[g].clear()
                # 逐组匹配
                for g, assignments in group_assignments_cfg.items():
                    if g not in self.layout_groups:
                        continue
                    for pos_str, window_info in assignments.items():
                        row, col = map(int, pos_str.split(','))
                        title = window_info.get('title')
                        class_name = window_info.get('class_name')
                        for window in self.windows:
                            if (window.title == title or window.class_name == class_name):
                                # 临时设置到该组（不改变当前组）
                                self.group_assignments[g][(row, col)] = window
                                matched_count += 1
                                break
                # 切回当前组显示
                self.on_layout_group_change()
            else:
                # 旧格式：仅当前组
                assignments = config.get('assignments', {})
                self.get_current_assignments().clear()
                for pos_str, window_info in assignments.items():
                    row, col = map(int, pos_str.split(','))
                    title = window_info.get('title')
                    class_name = window_info.get('class_name')
                    for window in self.windows:
                        if (window.title == title or window.class_name == class_name):
                            self.assign_window_to_position(window, row, col)
                            matched_count += 1
            
            self.show_status_message(f"配置已加载，匹配到 {matched_count} 个窗口")
            if hasattr(self, 'window_status_text') and self.window_status_text:
                self.window_status_text.insert(tk.END, f"✅ 配置已加载，匹配到 {matched_count} 个窗口\n")
                self.window_status_text.see(tk.END)
        except Exception as e:
            messagebox.showerror("错误", f"加载失败: {e}")
    
    def run(self):
        """运行程序"""
        self.root.mainloop()

def main():
    """主函数"""
    # 检查管理员权限
    if not is_admin():
        result = messagebox.askyesno(
            "权限提示", 
            "程序需要管理员权限才能正常管理窗口。\n\n是否以管理员身份重新运行？"
        )
        if result:
            if not run_as_admin():
                sys.exit(1)
            else:
                sys.exit(0)  # 当前进程退出，管理员进程接管
        else:
            messagebox.showwarning("警告", "没有管理员权限可能导致部分功能无法正常使用")
    
    app = WindowManagerGUI()
    app.run()

if __name__ == "__main__":
    main()