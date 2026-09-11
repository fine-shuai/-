#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import datetime
import ctypes
from ctypes import wintypes
import openpyxl
import glob

# ---------- 设置控制台编码为 UTF-8 ----------
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# ---------- 常量 ----------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TABLE_PATTERNS = ["文件时间修改.xlsx", "文件时间修改.xls"]
TIME_FORMATS_FULL = [
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%Y年%m月%d日 %H:%M:%S",
    "%m/%d/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M:%S",
]
TIME_FORMATS_DATE = [
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y.%m.%d",
    "%Y年%m月%d日",
    "%m/%d/%Y",
    "%d/%m/%Y",
]

def time_to_str(val):
    """将各种时间类型转为字符串，如果时间为 00:00:00 则只返回日期部分"""
    if val is None:
        return ""
    if isinstance(val, datetime.datetime):
        if val.hour == 0 and val.minute == 0 and val.second == 0 and val.microsecond == 0:
            return val.strftime("%Y-%m-%d")
        else:
            return val.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(val, (int, float)):
        dt = datetime.datetime.fromtimestamp(val)
        if dt.hour == 0 and dt.minute == 0 and dt.second == 0 and dt.microsecond == 0:
            return dt.strftime("%Y-%m-%d")
        else:
            return dt.strftime("%Y-%m-%d %H:%M:%S")
    return str(val).strip()

def find_table_file():
    for pattern in TABLE_PATTERNS:
        full_pattern = os.path.join(SCRIPT_DIR, pattern)
        files = glob.glob(full_pattern)
        if files:
            return files[0]
    return None

def read_excel_table(filepath):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception as e:
        print(f"无法打开表格文件: {e}")
        return None

    sheet = wb.active
    headers = [cell.value for cell in sheet[1]]
    try:
        idx_name = headers.index("文件名")
        idx_mtime = headers.index("修改时间")
        idx_ctime = headers.index("创建时间")
        idx_atime = headers.index("访问时间")
    except ValueError as e:
        print(f"表格列名不完整，需要包含：文件名、修改时间、创建时间、访问时间。错误：{e}")
        return None

    mapping = {}          # 完全匹配映射
    prefix_list = []      # 前缀列表，用于部分匹配（存储 (前缀, (mtime, ctime, atime))）
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row[idx_name]:
            continue
        name = str(row[idx_name])
        mtime = time_to_str(row[idx_mtime])
        ctime = time_to_str(row[idx_ctime])
        atime = time_to_str(row[idx_atime])
        mapping[name] = (mtime, ctime, atime)
        # 将表格中的名称作为前缀，加入前缀列表（支持部分匹配）
        prefix_list.append((name, (mtime, ctime, atime)))

    wb.close()
    return mapping, prefix_list

def parse_time_str(time_str, original_timestamp=None):
    """
    解析时间字符串，返回新时间戳（float）。
    - 完整格式（含非零时间）：完全替换
    - 完整格式但时间为 00:00:00：视为日期格式，仅替换日期，保留原有时分秒
    - 纯日期格式：仅替换日期，保留原有时分秒
    - 空字符串：返回 None（不修改）
    """
    if not time_str or time_str.strip() == "":
        return None
    time_str = time_str.strip()

    # 先尝试完整格式
    for fmt in TIME_FORMATS_FULL:
        try:
            dt = datetime.datetime.strptime(time_str, fmt)
            # 检查时间部分是否全为零
            if dt.hour == 0 and dt.minute == 0 and dt.second == 0 and dt.microsecond == 0:
                # 视为日期格式，使用日期保留逻辑
                if original_timestamp is None:
                    print(f"警告：日期格式 '{time_str}' 需要原始时间，但未提供，将不修改。")
                    return None
                orig_dt = datetime.datetime.fromtimestamp(original_timestamp)
                new_dt = datetime.datetime.combine(dt.date(), orig_dt.time())
                new_dt = new_dt.replace(microsecond=orig_dt.microsecond)
                return new_dt.timestamp()
            else:
                # 非零时间，完全替换
                return dt.timestamp()
        except ValueError:
            continue

    # 尝试纯日期格式
    for fmt in TIME_FORMATS_DATE:
        try:
            dt_date = datetime.datetime.strptime(time_str, fmt)
            if original_timestamp is None:
                print(f"警告：日期格式 '{time_str}' 需要原始时间，但未提供，将不修改。")
                return None
            orig_dt = datetime.datetime.fromtimestamp(original_timestamp)
            new_dt = datetime.datetime.combine(dt_date.date(), orig_dt.time())
            new_dt = new_dt.replace(microsecond=orig_dt.microsecond)
            return new_dt.timestamp()
        except ValueError:
            continue

    print(f"错误：无法解析时间字符串 '{time_str}'，请使用常见格式（如 2022-12-15 或 2022/12/15 18:04:00）")
    return None

# ---------- Windows API 时间修改 ----------
class FILETIME(ctypes.Structure):
    _fields_ = [("dwLowDateTime", wintypes.DWORD),
                ("dwHighDateTime", wintypes.DWORD)]

GENERIC_WRITE = 0x40000000
OPEN_EXISTING = 3
FILE_ATTRIBUTE_NORMAL = 0x80
INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value
kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)

def timestamp_to_filetime(timestamp):
    if timestamp is None:
        return None
    EPOCH_DIFF = 11644473600
    ft_seconds = int(timestamp) + EPOCH_DIFF
    ft_100ns = ft_seconds * 10**7
    frac = timestamp - int(timestamp)
    ft_100ns += int(frac * 10**7)
    ft = FILETIME()
    ft.dwLowDateTime = ft_100ns & 0xFFFFFFFF
    ft.dwHighDateTime = (ft_100ns >> 32) & 0xFFFFFFFF
    return ft

def set_file_times(filepath, atime, mtime, ctime):
    if atime is None and mtime is None and ctime is None:
        return True
    handle = kernel32.CreateFileW(
        filepath,
        GENERIC_WRITE,
        0,
        None,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        None
    )
    if handle == INVALID_HANDLE_VALUE:
        error = ctypes.get_last_error()
        print(f"无法打开文件 '{filepath}'，错误码: {error}")
        return False
    ft_ctime = timestamp_to_filetime(ctime) if ctime is not None else None
    ft_atime = timestamp_to_filetime(atime) if atime is not None else None
    ft_mtime = timestamp_to_filetime(mtime) if mtime is not None else None
    result = kernel32.SetFileTime(
        handle,
        ctypes.byref(ft_ctime) if ft_ctime else None,
        ctypes.byref(ft_atime) if ft_atime else None,
        ctypes.byref(ft_mtime) if ft_mtime else None
    )
    kernel32.CloseHandle(handle)
    if not result:
        error = ctypes.get_last_error()
        print(f"设置文件时间失败 '{filepath}'，错误码: {error}")
        return False
    return True

def find_files_from_paths(paths):
    all_files = []
    for p in paths:
        p = os.path.abspath(p)
        if not os.path.exists(p):
            print(f"警告：路径不存在，跳过: {p}")
            continue
        if os.path.isfile(p):
            all_files.append(p)
        elif os.path.isdir(p):
            for root, dirs, files in os.walk(p):
                for f in files:
                    all_files.append(os.path.join(root, f))
    return all_files

def match_filename(filename, mapping, prefix_list):
    """匹配文件名，返回 (mtime_str, ctime_str, atime_str) 和匹配类型（'exact' 或 'prefix'）"""
    # 完全匹配优先
    if filename in mapping:
        return mapping[filename], 'exact'
    # 前缀匹配：查找最长前缀
    best_prefix = None
    best_data = None
    best_len = -1
    for prefix, data in prefix_list:
        if filename.startswith(prefix) and len(prefix) > best_len:
            best_len = len(prefix)
            best_prefix = prefix
            best_data = data
    if best_data:
        return best_data, f'prefix ("{best_prefix}")'
    return None, None

def main():
    table_path = find_table_file()
    if not table_path:
        print("错误：未找到表格文件。请将表格命名为“文件时间修改.xlsx”或“文件时间修改.xls”并放在脚本同级目录。")
        input("按回车键退出...")
        sys.exit(1)

    print(f"找到表格文件: {table_path}")
    mapping, prefix_list = read_excel_table(table_path)
    if mapping is None:
        input("按回车键退出...")
        sys.exit(1)
    print(f"已加载 {len(mapping)} 条记录。")
    print(f"表格中的文件名/前缀列表: {list(mapping.keys())}")

    if len(sys.argv) < 2:
        print("未提供文件/文件夹。请将文件或文件夹拖拽到批处理文件上运行。")
        input("按回车键退出...")
        sys.exit(1)

    raw_paths = sys.argv[1:]
    print(f"拖拽的路径: {raw_paths}")
    all_files = find_files_from_paths(raw_paths)
    if not all_files:
        print("没有找到有效的文件。")
        input("按回车键退出...")
        sys.exit(1)

    print(f"扫描到 {len(all_files)} 个文件。")
    if len(all_files) > 0:
        print("前5个文件示例:")
        for f in all_files[:5]:
            print(f"  {f}")

    print("\n请选择运行模式：")
    print("1) 试运行（仅显示即将执行的操作）")
    print("2) 正式运行（修改文件时间）")
    mode = input("请输入 1 或 2: ").strip()
    if mode not in ("1", "2"):
        print("无效输入，退出。")
        input("按回车键退出...")
        sys.exit(1)
    dry_run = (mode == "1")

    processed = 0
    success_count = 0
    fail_count = 0

    for filepath in all_files:
        filename = os.path.basename(filepath)
        times, match_type = match_filename(filename, mapping, prefix_list)
        if times is None:
            continue

        processed += 1
        mtime_str, ctime_str, atime_str = times
        print(f"\n处理文件: {filename}")
        print(f"  匹配方式: {match_type}")

        try:
            orig_stat = os.stat(filepath)
            orig_mtime = orig_stat.st_mtime
            orig_ctime = orig_stat.st_ctime
            orig_atime = orig_stat.st_atime
        except Exception as e:
            print(f"  无法获取文件时间: {e}")
            continue

        new_mtime = parse_time_str(mtime_str, orig_mtime)
        new_ctime = parse_time_str(ctime_str, orig_ctime)
        new_atime = parse_time_str(atime_str, orig_atime)

        if new_mtime is not None:
            old = datetime.datetime.fromtimestamp(orig_mtime).strftime("%Y-%m-%d %H:%M:%S")
            new = datetime.datetime.fromtimestamp(new_mtime).strftime("%Y-%m-%d %H:%M:%S")
            print(f"  修改时间: {old} -> {new}")
        else:
            print(f"  修改时间: 不修改")
        if new_ctime is not None:
            old = datetime.datetime.fromtimestamp(orig_ctime).strftime("%Y-%m-%d %H:%M:%S")
            new = datetime.datetime.fromtimestamp(new_ctime).strftime("%Y-%m-%d %H:%M:%S")
            print(f"  创建时间: {old} -> {new}")
        else:
            print(f"  创建时间: 不修改")
        if new_atime is not None:
            old = datetime.datetime.fromtimestamp(orig_atime).strftime("%Y-%m-%d %H:%M:%S")
            new = datetime.datetime.fromtimestamp(new_atime).strftime("%Y-%m-%d %H:%M:%S")
            print(f"  访问时间: {old} -> {new}")
        else:
            print(f"  访问时间: 不修改")

        if dry_run:
            print("  [试运行] 未实际修改")
        else:
            if set_file_times(filepath, new_atime, new_mtime, new_ctime):
                print("  ✓ 修改成功")
                success_count += 1
            else:
                print("  ✗ 修改失败")
                fail_count += 1

    # 输出统计信息
    if processed == 0:
        print("\n警告：没有找到与表格匹配的文件。")
        print("请确保表格中的文件名/前缀与实际文件名（不含扩展名）一致。")
        print("支持完全匹配和前缀匹配（最长前缀优先）。")
        print("如果文件在子文件夹中，脚本会递归搜索所有文件。")
    elif dry_run:
        print(f"\n试运行结束。匹配到 {processed} 个文件，未实际修改。")
    else:
        print(f"\n正式运行完成。成功修改 {success_count} 个文件，失败 {fail_count} 个文件（共匹配 {processed} 个文件）。")

    input("按回车键退出...")

if __name__ == "__main__":
    main()