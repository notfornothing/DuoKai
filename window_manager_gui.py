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
from typing import List, Dict, Tuple, Optional

# Windows API 常量
SW_RESTORE = 9
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
HWND_TOP = 0

class WindowInfo:
    """窗口信息类"""
    def __init__(self, hwnd: int, title: str, class_name: str):
        self.hwnd = hwnd
        self.title = title
        self.class_name = class_name
        self.assigned_position = None  # (row, col) 或 None
    
    def __str__(self):
        return f"{self.title} ({self.class_name})"

class WindowManagerGUI:
    """可视化窗口管理器"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("窗口管理器 - 可视化设置")
        self.root.geometry("900x700")
        
        # 配置变量
        self.rows = tk.IntVar(value=2)
        self.columns = tk.IntVar(value=3)
        self.screen_width = tk.IntVar(value=2560)
        self.screen_height = tk.IntVar(value=1440)
        self.use_workarea = tk.BooleanVar(value=True)
        
        # 数据
        self.windows: List[WindowInfo] = []
        self.grid_assignments: Dict[Tuple[int, int], WindowInfo] = {}  # (row, col) -> WindowInfo
        
        # GUI 组件
        self.window_listbox = None
        self.grid_frame = None
        self.grid_buttons = {}  # (row, col) -> Button
        
        self.setup_ui()
        self.refresh_windows()
    
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置区域
        config_frame = ttk.LabelFrame(main_frame, text="配置设置", padding="5")
        config_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 网格设置
        ttk.Label(config_frame, text="行数:").grid(row=0, column=0, padx=(0, 5))
        ttk.Spinbox(config_frame, from_=1, to=5, textvariable=self.rows, width=5,
                   command=self.update_grid).grid(row=0, column=1, padx=(0, 10))
        
        ttk.Label(config_frame, text="列数:").grid(row=0, column=2, padx=(0, 5))
        ttk.Spinbox(config_frame, from_=1, to=5, textvariable=self.columns, width=5,
                   command=self.update_grid).grid(row=0, column=3, padx=(0, 10))
        
        # 分辨率设置
        ttk.Label(config_frame, text="宽度:").grid(row=0, column=4, padx=(0, 5))
        ttk.Entry(config_frame, textvariable=self.screen_width, width=6).grid(row=0, column=5, padx=(0, 5))
        
        ttk.Label(config_frame, text="高度:").grid(row=0, column=6, padx=(0, 5))
        ttk.Entry(config_frame, textvariable=self.screen_height, width=6).grid(row=0, column=7, padx=(0, 10))
        
        ttk.Checkbutton(config_frame, text="使用工作区(避开任务栏)", 
                       variable=self.use_workarea).grid(row=0, column=8)
        
        # 左侧：窗口列表
        left_frame = ttk.LabelFrame(main_frame, text="可用窗口", padding="5")
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        # 窗口列表
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.window_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=15)
        self.window_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.window_listbox.yview)
        
        # 绑定双击事件
        self.window_listbox.bind('<Double-1>', self.on_window_double_click)
        
        # 刷新按钮
        ttk.Button(left_frame, text="刷新窗口列表", command=self.refresh_windows).pack(pady=(5, 0))
        
        # 右侧：网格布局
        right_frame = ttk.LabelFrame(main_frame, text="网格布局 (拖拽窗口到这里)", padding="5")
        right_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.grid_frame = ttk.Frame(right_frame)
        self.grid_frame.pack(fill=tk.BOTH, expand=True)
        
        # 底部按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(button_frame, text="预览布局", command=self.preview_layout).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="应用设置", command=self.apply_layout).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="清空设置", command=self.clear_assignments).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="加载配置", command=self.load_config).pack(side=tk.LEFT)
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=2)
        main_frame.rowconfigure(1, weight=1)
        
        self.update_grid()
    
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
                self.grid_buttons[(r, c)] = btn
                
                # 配置网格权重
                self.grid_frame.columnconfigure(c, weight=1)
                self.grid_frame.rowconfigure(r, weight=1)
        
        # 更新按钮显示
        self.update_grid_display()
    
    def update_grid_display(self):
        """更新网格按钮显示"""
        for (r, c), btn in self.grid_buttons.items():
            if (r, c) in self.grid_assignments:
                window = self.grid_assignments[(r, c)]
                btn.config(
                    text=f"位置 {r+1},{c+1}\n{window.title[:20]}",
                    bg='lightblue'
                )
            else:
                btn.config(
                    text=f"位置 {r+1},{c+1}\n(空)",
                    bg='lightgray'
                )
    
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
        if (row, col) in self.grid_assignments:
            # 已有窗口，询问是否移除
            window = self.grid_assignments[(row, col)]
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
        # 移除窗口的旧分配
        if window.assigned_position:
            old_pos = window.assigned_position
            if old_pos in self.grid_assignments:
                del self.grid_assignments[old_pos]
        
        # 移除位置的旧窗口
        if (row, col) in self.grid_assignments:
            old_window = self.grid_assignments[(row, col)]
            old_window.assigned_position = None
        
        # 新分配
        window.assigned_position = (row, col)
        self.grid_assignments[(row, col)] = window
        
        # 更新显示
        self.update_grid_display()
        self.refresh_windows()
    
    def remove_window_assignment(self, window: WindowInfo):
        """移除窗口分配"""
        if window.assigned_position:
            pos = window.assigned_position
            if pos in self.grid_assignments:
                del self.grid_assignments[pos]
            window.assigned_position = None
            
            self.update_grid_display()
            self.refresh_windows()
    
    def clear_assignments(self):
        """清空所有分配"""
        if messagebox.askyesno("确认", "是否清空所有窗口分配?"):
            self.grid_assignments.clear()
            for window in self.windows:
                window.assigned_position = None
            self.update_grid_display()
            self.refresh_windows()
    
    def preview_layout(self):
        """预览布局"""
        if not self.grid_assignments:
            messagebox.showinfo("提示", "没有分配任何窗口")
            return
        
        # 计算位置
        positions = self.calculate_positions()
        
        # 显示预览信息
        preview_text = "布局预览:\n\n"
        for (row, col), window in self.grid_assignments.items():
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
        for row in range(rows):
            for col in range(cols):
                x = offset_x + col * cell_w + min(col, extra_w)
                y = offset_y + row * cell_h + min(row, extra_h)
                w = cell_w + (1 if col < extra_w else 0)
                h = cell_h + (1 if row < extra_h else 0)
                positions[(row, col)] = (x, y, w, h)
        
        return positions
    
    def apply_layout(self):
        """应用布局"""
        if not self.grid_assignments:
            messagebox.showinfo("提示", "没有分配任何窗口")
            return
        
        if not messagebox.askyesno("确认", "是否应用当前布局设置?"):
            return
        
        positions = self.calculate_positions()
        success_count = 0
        
        for (row, col), window in self.grid_assignments.items():
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
        
        messagebox.showinfo("完成", f"成功应用 {success_count}/{len(self.grid_assignments)} 个窗口的布局")
    
    def save_config(self):
        """保存配置"""
        if not self.grid_assignments:
            messagebox.showinfo("提示", "没有配置可保存")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            config = {
                'rows': self.rows.get(),
                'columns': self.columns.get(),
                'screen_width': self.screen_width.get(),
                'screen_height': self.screen_height.get(),
                'use_workarea': self.use_workarea.get(),
                'assignments': {}
            }
            
            for (row, col), window in self.grid_assignments.items():
                config['assignments'][f"{row},{col}"] = {
                    'title': window.title,
                    'class_name': window.class_name
                }
            
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("成功", f"配置已保存到 {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {e}")
    
    def load_config(self):
        """加载配置"""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # 加载设置
                self.rows.set(config.get('rows', 2))
                self.columns.set(config.get('columns', 3))
                self.screen_width.set(config.get('screen_width', 2560))
                self.screen_height.set(config.get('screen_height', 1440))
                self.use_workarea.set(config.get('use_workarea', True))
                
                # 更新网格
                self.update_grid()
                
                # 尝试匹配窗口
                assignments = config.get('assignments', {})
                matched_count = 0
                
                for pos_str, window_info in assignments.items():
                    row, col = map(int, pos_str.split(','))
                    title = window_info['title']
                    class_name = window_info['class_name']
                    
                    # 查找匹配的窗口
                    for window in self.windows:
                        if (window.title == title or window.class_name == class_name) and window.assigned_position is None:
                            self.assign_window_to_position(window, row, col)
                            matched_count += 1
                            break
                
                messagebox.showinfo("成功", f"配置已加载，匹配到 {matched_count}/{len(assignments)} 个窗口")
                
            except Exception as e:
                messagebox.showerror("错误", f"加载失败: {e}")
    
    def run(self):
        """运行程序"""
        self.root.mainloop()

def main():
    """主函数"""
    app = WindowManagerGUI()
    app.run()

if __name__ == "__main__":
    main()