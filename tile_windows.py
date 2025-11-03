import argparse
import ctypes
import os
import sys

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

GW_OWNER = 4
GWL_STYLE = -16
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
SW_RESTORE = 9
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010
SWP_NOOWNERZORDER = 0x0200

SM_CXSCREEN = 0
SM_CYSCREEN = 1
SPI_GETWORKAREA = 0x0030

WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)


def get_primary_resolution(use_work_area: bool):
    if use_work_area:
        rect = (ctypes.c_long * 4)()
        user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, rect, 0)
        left, top, right, bottom = rect
        return int(right - left), int(bottom - top), int(left), int(top)
    else:
        width = user32.GetSystemMetrics(SM_CXSCREEN)
        height = user32.GetSystemMetrics(SM_CYSCREEN)
        return int(width), int(height), 0, 0


def is_window_valid(hwnd, current_pid):
    # 可见、非最小化、有标题、非工具窗口、无所有者窗口
    if not user32.IsWindowVisible(hwnd):
        return False
    if user32.IsIconic(hwnd):
        return False
    # 过滤工具窗口
    ex_style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    if ex_style & WS_EX_TOOLWINDOW:
        return False
    # 过滤有所有者的窗口（通常是对话框等）
    owner = user32.GetWindow(hwnd, GW_OWNER)
    if owner:
        return False
    # 过滤空标题
    length = user32.GetWindowTextLengthW(hwnd)
    if length == 0:
        return False
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    title = buf.value.strip()
    if not title:
        return False
    # 过滤程序管理器等特殊窗口
    if title.lower() in ("program manager",):
        return False
    # 过滤当前脚本的控制台窗口
    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if pid.value == current_pid:
        return False
    return True


def enumerate_windows(max_count=None):
    result = []
    current_pid = os.getpid()

    def callback(hwnd, lparam):
        nonlocal result
        if is_window_valid(hwnd, current_pid):
            # 标题
            length = user32.GetWindowTextLengthW(hwnd)
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            title = buf.value
            result.append((hwnd, title))
            if max_count and len(result) >= max_count:
                return False  # 停止枚举
        return True

    user32.EnumWindows(WNDENUMPROC(callback), 0)
    return result


def compute_grid_positions(screen_w, screen_h, offset_x, offset_y, columns, rows, count):
    # 按列和行均匀分配像素，余数从左/上开始逐列逐行分配，避免缝隙
    base_w = screen_w // columns
    rem_w = screen_w % columns
    col_widths = [base_w + (1 if i < rem_w else 0) for i in range(columns)]
    col_x = [offset_x]
    for i in range(1, columns):
        col_x.append(col_x[i - 1] + col_widths[i - 1])

    base_h = screen_h // rows
    rem_h = screen_h % rows
    row_heights = [base_h + (1 if i < rem_h else 0) for i in range(rows)]
    row_y = [offset_y]
    for i in range(1, rows):
        row_y.append(row_y[i - 1] + row_heights[i - 1])

    rects = []
    for idx in range(count):
        r = idx // columns
        c = idx % columns
        if r >= rows:
            break
        x = col_x[c]
        y = row_y[r]
        w = col_widths[c]
        h = row_heights[r]
        rects.append((x, y, w, h))
    return rects


def move_resize_window(hwnd, x, y, w, h):
    user32.ShowWindow(hwnd, SW_RESTORE)
    user32.SetWindowPos(hwnd, 0, int(x), int(y), int(w), int(h),
                        SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOOWNERZORDER)


def tile_windows(columns=3, rows=2, count=6, use_work_area=True, force_resolution=None, dry_run=False):
    if force_resolution:
        screen_w, screen_h = force_resolution
        offset_x, offset_y = 0, 0
    else:
        screen_w, screen_h, offset_x, offset_y = get_primary_resolution(use_work_area)

    windows = enumerate_windows(max_count=None)
    if not windows:
        print("未找到可平铺的窗口。")
        return

    n = min(count, len(windows), columns * rows)
    rects = compute_grid_positions(screen_w, screen_h, offset_x, offset_y, columns, rows, n)

    print(f"检测到 {len(windows)} 个候选窗口，将平铺其中 {n} 个：")
    for i in range(n):
        hwnd, title = windows[i]
        x, y, w, h = rects[i]
        print(f"[{i+1}] '{title}' -> x={x}, y={y}, w={w}, h={h}")
        if not dry_run:
            move_resize_window(hwnd, x, y, w, h)


def parse_args():
    parser = argparse.ArgumentParser(description="将桌面窗口平铺为两行三列（或自定义）。")
    parser.add_argument("--columns", type=int, default=3, help="列数，默认3")
    parser.add_argument("--rows", type=int, default=2, help="行数，默认2")
    parser.add_argument("--count", type=int, default=6, help="平铺窗口数量，默认6")
    parser.add_argument("--full-screen", action="store_true", help="使用整屏分辨率而非工作区")
    parser.add_argument("--resolution", nargs=2, type=int, metavar=("W", "H"), help="强制使用给定分辨率，例如 2560 1440")
    parser.add_argument("--dry-run", action="store_true", help="仅打印布局结果，不实际移动窗口")
    return parser.parse_args()


def main():
    args = parse_args()
    force_res = None
    if args.resolution:
        force_res = (args.resolution[0], args.resolution[1])

    tile_windows(columns=args.columns,
                 rows=args.rows,
                 count=args.count,
                 use_work_area=not args.full_screen,
                 force_resolution=force_res,
                 dry_run=args.dry_run)


if __name__ == "__main__":
    if os.name != "nt":
        print("该脚本仅支持在 Windows 上运行。")
        sys.exit(1)
    main()