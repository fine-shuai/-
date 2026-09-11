#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import shutil
import pandas as pd

def get_name_list_from_excel(file_path):
    """读取 Excel 中的“复制粘贴重命名”列，返回非空名称列表"""
    try:
        df = pd.read_excel(file_path, dtype=str)
    except Exception as e:
        print(f"读取 Excel 文件失败：{e}")
        sys.exit(1)

    if "复制粘贴重命名" not in df.columns:
        print("错误：Excel 中未找到“复制粘贴重命名”列。")
        sys.exit(1)

    names = df["复制粘贴重命名"].dropna().tolist()
    if not names:
        print("错误：表格中“复制粘贴重命名”列没有有效数据。")
        sys.exit(1)

    return names

def copy_file(src, dst):
    """复制文件，若目标目录不存在则创建"""
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)

def copy_folder_all(src, dst):
    """复制整个文件夹（包括子文件夹和文件）"""
    if os.path.exists(dst):
        # 如果目标已存在，则先删除（避免 copytree 报错）
        shutil.rmtree(dst)
    shutil.copytree(src, dst)

def copy_folder_with_ext_filter(src, dst, extensions):
    """递归复制文件夹中扩展名匹配的文件，保持目录结构"""
    for root, dirs, files in os.walk(src):
        rel_path = os.path.relpath(root, src)
        target_dir = os.path.join(dst, rel_path) if rel_path != '.' else dst
        os.makedirs(target_dir, exist_ok=True)
        for file in files:
            if any(file.lower().endswith(ext.lower()) for ext in extensions):
                src_file = os.path.join(root, file)
                dst_file = os.path.join(target_dir, file)
                shutil.copy2(src_file, dst_file)

def main():
    # 检查命令行参数（拖拽的文件/文件夹）
    if len(sys.argv) < 2:
        print("请将文件或文件夹拖拽到此脚本上运行。")
        sys.exit(0)

    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    excel_xlsx = os.path.join(script_dir, "批量复制表.xlsx")
    excel_xls = os.path.join(script_dir, "批量复制表.xls")

    # 检查表格文件
    excel_file = None
    if os.path.exists(excel_xlsx):
        excel_file = excel_xlsx
    elif os.path.exists(excel_xls):
        excel_file = excel_xls
    else:
        print("错误：未找到“批量复制表.xlsx”或“批量复制表.xls”。")
        sys.exit(1)

    # 读取重命名列表
    print("正在读取表格...")
    names = get_name_list_from_excel(excel_file)
    print(f"找到 {len(names)} 个重命名名称：{', '.join(names)}")

    # 源文件/文件夹列表
    sources = sys.argv[1:]
    print(f"拖拽了 {len(sources)} 个源项目。")
    print(f"总共将执行 {len(sources) * len(names)} 次复制操作。")

    # 交互选项
    print("\n请选择复制模式：")
    print("  1. 复制所有内容（文件/文件夹及其全部子内容）")
    print("  2. 仅复制文件（若源为文件夹则复制其中所有文件并保持原结构）")
    print("  3. 仅复制特定扩展名的文件")
    mode = input("请输入选项 [1/2/3]：").strip()
    extensions = []
    if mode == "3":
        ext_str = input("请输入扩展名（多个用空格隔开，例如 .txt .jpg）：").strip()
        extensions = [e.strip().lower() for e in ext_str.split() if e.strip()]
        if not extensions:
            print("未输入有效扩展名，将按模式1执行。")
            mode = "1"

    # 目标保存位置
    print("\n请输入目标文件夹路径（可直接拖拽文件夹到此窗口）：")
    dest_input = input("目标路径（直接回车使用默认目录）：").strip()
    if dest_input:
        dest_path = dest_input.strip('"')
    else:
        dest_path = os.path.join(script_dir, "CopyOutput")
        print(f"将使用默认路径：{dest_path}")

    dest_path = os.path.abspath(dest_path)
    os.makedirs(dest_path, exist_ok=True)

    # 开始复制
    success = 0
    fail = 0

    for src in sources:
        src = src.strip('"')
        if not os.path.exists(src):
            print(f"警告：源项目不存在，跳过 - {src}")
            fail += 1
            continue

        is_folder = os.path.isdir(src)
        for name in names:
            target = os.path.join(dest_path, name)
            try:
                if is_folder:
                    if mode == "1":
                        copy_folder_all(src, target)
                        print(f"成功复制文件夹：{src} -> {target}")
                    elif mode == "2":
                        copy_folder_all(src, target)
                        print(f"成功复制文件夹（所有文件）：{src} -> {target}")
                    elif mode == "3":
                        copy_folder_with_ext_filter(src, target, extensions)
                        print(f"成功复制文件夹（过滤扩展名）：{src} -> {target}")
                else:  # 文件
                    if mode == "1" or mode == "2":
                        copy_file(src, target)
                        print(f"成功复制文件：{src} -> {target}")
                    elif mode == "3":
                        ext = os.path.splitext(src)[1].lower()
                        if ext in extensions:
                            copy_file(src, target)
                            print(f"成功复制文件：{src} -> {target}")
                        else:
                            print(f"跳过文件（扩展名不匹配）：{src}")
                            success += 1  # 跳过也算成功（未失败）
                            continue
                success += 1
            except Exception as e:
                print(f"失败：{src} -> {target}，错误：{e}")
                fail += 1

    print("\n========== 处理完成 ==========")
    print(f"成功操作数：{success}")
    print(f"失败/跳过数：{fail}")
    print(f"目标位置：{dest_path}")

if __name__ == "__main__":
    main()