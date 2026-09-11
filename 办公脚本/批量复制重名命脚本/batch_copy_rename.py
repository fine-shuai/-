#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import shutil
import subprocess
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
    """复制单个文件（使用 shutil）"""
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)

def copy_folder_with_robocopy(src, dst, mode, extensions=None):
    """
    使用 robocopy 复制文件夹
    mode: 1-全部内容(E), 2-全部内容(E)（与模式1相同，因为文件复制模式2也是复制所有文件），
          3-仅复制指定扩展名文件（使用通配符）
    extensions: 模式3时需提供扩展名列表，如 ['.txt', '.jpg']
    """
    # 构建 robocopy 命令
    cmd = ["robocopy", src, dst]

    # 根据模式添加参数
    if mode == "1" or mode == "2":
        # 复制所有内容，包含空目录，多线程，失败重试3次，等待5秒
        cmd.extend(["/E", "/MT:8", "/R:3", "/W:5"])
    elif mode == "3" and extensions:
        # 仅复制指定扩展名的文件，使用 /S 复制子目录（非空），多线程等
        # 添加文件过滤：例如 *.txt *.jpg
        for ext in extensions:
            cmd.append(f"*{ext}")
        cmd.extend(["/S", "/MT:8", "/R:3", "/W:5"])
    else:
        # 默认模式（应不会走到这里）
        cmd.append("/E")

    # 添加静默输出选项（不显示详细进度和头部信息，仅保留错误）
    cmd.extend(["/NFL", "/NDL", "/NJH", "/NJS", "/NP"])

    # 执行 robocopy
    result = subprocess.run(cmd, capture_output=True, text=True, shell=True)

    # robocopy 返回码：0-3 成功，>=4 失败
    if result.returncode >= 4:
        # 失败，打印错误信息
        print(f"robocopy 错误：{result.stderr}")
        return False
    else:
        # 成功，可选择输出简要信息
        if result.stdout:
            # 输出可能包含一些信息（如跳过信息），但已屏蔽，可忽略
            pass
        return True

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
                    # 文件夹使用 robocopy
                    print(f"正在复制文件夹：{src} -> {target}")
                    if copy_folder_with_robocopy(src, target, mode, extensions):
                        print(f"成功复制文件夹：{src} -> {target}")
                        success += 1
                    else:
                        print(f"失败：复制文件夹 {src} -> {target}")
                        fail += 1
                else:
                    # 文件使用 shutil
                    if mode == "1" or mode == "2":
                        copy_file(src, target)
                        print(f"成功复制文件：{src} -> {target}")
                        success += 1
                    elif mode == "3":
                        ext = os.path.splitext(src)[1].lower()
                        if ext in extensions:
                            copy_file(src, target)
                            print(f"成功复制文件：{src} -> {target}")
                            success += 1
                        else:
                            print(f"跳过文件（扩展名不匹配）：{src}")
                            success += 1  # 跳过也算成功（未失败）
            except Exception as e:
                print(f"失败：{src} -> {target}，错误：{e}")
                fail += 1

    print("\n========== 处理完成 ==========")
    print(f"成功操作数：{success}")
    print(f"失败/跳过数：{fail}")
    print(f"目标位置：{dest_path}")

if __name__ == "__main__":
    main()