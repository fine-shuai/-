#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
文件重命名工具
根据Excel表格中的映射关系，批量重命名文件/文件夹。
支持拖拽多个文件/文件夹，递归处理子项。
提供试运行和正式运行两种模式。
"""

import sys
import os
from pathlib import Path
from collections import Counter
import openpyxl  # 需要安装：pip install openpyxl

def find_table(script_dir):
    """在脚本所在目录查找表格文件（文件名修改.xlsx）"""
    table_path = script_dir / "文件名修改.xlsx"
    if table_path.exists():
        return table_path
    # 尝试其他常见扩展名
    for ext in ['.xls', '.csv']:
        alt_path = script_dir / f"文件名修改{ext}"
        if alt_path.exists():
            return alt_path
    raise FileNotFoundError("未找到表格文件，请确保脚本目录下存在“文件名修改.xlsx”")

def load_mapping(table_path):
    """从Excel读取映射表，返回字典 {旧文件名: 新文件名}"""
    wb = openpyxl.load_workbook(table_path, data_only=True)
    ws = wb.active
    mapping = {}
    # 假设第一行是表头，从第二行开始读取，第一列为旧文件名，第二列为新文件名
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        old_name = str(row[0]).strip()
        new_name = str(row[1]).strip() if row[1] is not None else ""
        if old_name and new_name:
            mapping[old_name] = new_name
    return mapping

def collect_items(paths):
    """
    收集需要处理的所有文件/文件夹路径。
    对于传入的每个路径：
    - 如果是文件，直接加入列表
    - 如果是文件夹，递归遍历其下所有文件和文件夹（但不包括文件夹本身）
    """
    items = []
    for p in paths:
        p = Path(p)
        if not p.exists():
            print(f"警告：路径不存在，跳过：{p}")
            continue
        if p.is_file():
            items.append(p)
        elif p.is_dir():
            # 递归遍历所有子文件和子文件夹
            for child in p.rglob('*'):
                items.append(child)
    return items

def rename_item(item, mapping, dry_run):
    """
    根据映射表重命名单个文件/文件夹。
    返回 (success, new_name, ext, reason)
    """
    # 获取文件名（不含扩展名）和扩展名
    full_name = item.name
    if item.is_file():
        stem = item.stem          # 不含扩展名
        suffix = item.suffix      # 扩展名（含点）
    else:
        stem = full_name
        suffix = ""

    # 查找映射
    if stem not in mapping:
        return False, None, suffix, "未在映射表中找到旧文件名"

    new_stem = mapping[stem]
    if suffix:
        new_name = new_stem + suffix
    else:
        new_name = new_stem

    target = item.parent / new_name

    # 检查目标是否已存在
    if target.exists():
        return False, None, suffix, f"目标已存在：{target}"

    if dry_run:
        print(f"[试运行] 将重命名：{item} -> {target}")
        return True, new_name, suffix, "试运行"
    else:
        try:
            item.rename(target)
            print(f"[已重命名] {item} -> {target}")
            return True, new_name, suffix, "成功"
        except Exception as e:
            return False, None, suffix, f"重命名失败：{e}"

def main():
    # 获取传入的路径（拖拽的文件/文件夹）
    if len(sys.argv) < 2:
        print("用法：将文件/文件夹拖拽到此脚本上，或通过命令行传递路径")
        print("示例：rename_files.py C:\\folder1 D:\\file.txt")
        input("按回车键退出...")
        sys.exit(1)

    paths = sys.argv[1:]

    # 定位表格
    script_dir = Path(__file__).parent
    try:
        table_path = find_table(script_dir)
        print(f"已加载表格：{table_path}")
    except FileNotFoundError as e:
        print(e)
        input("按回车键退出...")
        sys.exit(1)

    # 加载映射关系
    try:
        mapping = load_mapping(table_path)
        print(f"共加载 {len(mapping)} 条映射记录")
    except Exception as e:
        print(f"读取表格失败：{e}")
        input("按回车键退出...")
        sys.exit(1)

    # 收集所有待处理项
    print("正在扫描文件/文件夹...")
    items = collect_items(paths)
    print(f"共发现 {len(items)} 个文件/文件夹")

    # 询问运行模式
    print("\n请选择运行模式：")
    print("1) 试运行（仅显示将执行的操作）")
    print("2) 正式运行（实际修改文件名）")
    choice = input("请输入数字（1或2）：").strip()
    if choice not in ('1', '2'):
        print("无效选择，退出。")
        input("按回车键退出...")
        sys.exit(1)

    dry_run = (choice == '1')
    print(f"\n{'【试运行模式】' if dry_run else '【正式运行模式】'}开始处理...\n")

    # 统计信息
    success_count = 0
    ext_counter = Counter()  # 统计文件类型（扩展名）
    failed_details = []

    # 为避免重命名父文件夹后子路径失效，按深度从深到浅排序
    items_sorted = sorted(items, key=lambda x: len(x.parents), reverse=True)

    for item in items_sorted:
        success, new_name, ext, reason = rename_item(item, mapping, dry_run)
        if success:
            success_count += 1
            if ext:
                ext_counter[ext] += 1
            else:
                ext_counter["(文件夹)"] += 1
        else:
            if reason not in ("未在映射表中找到旧文件名", "试运行"):
                failed_details.append((item, reason))

    # 输出统计结果
    print("\n" + "="*50)
    print(f"处理完成！")
    print(f"共处理 {len(items)} 个项，成功重命名 {success_count} 个")
    if not dry_run and failed_details:
        print(f"失败 {len(failed_details)} 个：")
        for item, reason in failed_details:
            print(f"  - {item} : {reason}")

    print("\n文件类型统计：")
    for ext, count in ext_counter.most_common():
        print(f"  {ext}: {count} 个")

    input("\n按回车键退出...")

if __name__ == "__main__":
    main()