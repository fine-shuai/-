#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
批量汇总多个 Excel 表格中的指定列数据（支持多级子文件夹）
输出列顺序与需求文件严格一致，来源表列放在最后
作者：Assistant
日期：2026-03-21
"""

import os
import sys
import glob
import pandas as pd

def get_requirement_columns(requirement_file):
    ext = os.path.splitext(requirement_file)[1].lower()
    try:
        if ext == '.csv':
            df = pd.read_csv(requirement_file, header=None)
        elif ext in ['.xls', '.xlsx']:
            df = pd.read_excel(requirement_file, header=None)
        else:
            print(f"不支持的需求文件格式: {ext}")
            sys.exit(1)
    except Exception as e:
        print(f"读取需求文件失败: {e}")
        sys.exit(1)

    columns = df.iloc[:, 0].dropna().astype(str).tolist()
    if not columns:
        print("需求文件中未找到有效的列名（第一列无数据）")
        sys.exit(1)
    return columns

def choose_file_format():
    print("\n请选择要汇总的表格格式：")
    print("1 - 只汇总 .xls 文件")
    print("2 - 只汇总 .xlsx 文件")
    print("3 - 同时汇总 .xls 和 .xlsx 文件")
    choice = input("请输入数字 (1/2/3): ").strip()
    if choice == '1':
        return ['.xls']
    elif choice == '2':
        return ['.xlsx']
    elif choice == '3':
        return ['.xls', '.xlsx']
    else:
        print("输入无效，默认选择 3 (两者都汇总)")
        return ['.xls', '.xlsx']

def find_excel_files_recursively(root_folder, extensions):
    files = []
    for ext in extensions:
        pattern = os.path.join(root_folder, "**", f"*{ext}")
        files.extend(glob.glob(pattern, recursive=True))
    files = [f for f in files if not os.path.basename(f).startswith('列汇总需求')]
    files = [f for f in files if not os.path.basename(f).startswith('汇总结果')]
    return files

def main():
    print("=== 批量汇总表格指定列数据工具（支持多级子文件夹）===\n")

    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 1. 获取目标文件夹路径
    if len(sys.argv) > 1:
        folder = sys.argv[1]
        print(f"使用拖拽的文件夹: {folder}")
    else:
        folder = input("请输入存放待汇总表格的文件夹路径（支持拖拽）: ").strip()
        if not folder:
            folder = os.getcwd()
            print(f"使用当前目录: {folder}")

    if not os.path.isdir(folder):
        print(f"文件夹不存在: {folder}")
        sys.exit(1)

    # 2. 选择文件格式
    extensions = choose_file_format()

    # 3. 获取列汇总需求文件路径（优先从脚本所在目录查找）
    req_file = input("请输入“列汇总需求”表格的路径 (直接回车则使用脚本所在目录下的“列汇总需求”文件): ").strip()
    if not req_file:
        possible_names = ['列汇总需求.xlsx', '列汇总需求.xls', '列汇总需求.csv']
        found = None
        for name in possible_names:
            file_path = os.path.join(script_dir, name)
            if os.path.exists(file_path):
                found = file_path
                break
        if found:
            req_file = found
            print(f"自动使用需求文件: {req_file}")
        else:
            print(f"未在脚本所在目录 {script_dir} 中找到“列汇总需求”文件，请重新运行脚本并输入正确路径。")
            sys.exit(1)
    else:
        if not os.path.exists(req_file):
            print(f"文件不存在: {req_file}")
            sys.exit(1)

    required_columns = get_requirement_columns(req_file)
    print(f"\n需要汇总的列: {required_columns}")

    # 4. 递归查找所有符合条件的 Excel 文件
    excel_files = find_excel_files_recursively(folder, extensions)
    if not excel_files:
        print(f"在 {folder} 及其子文件夹中没有找到 {extensions} 格式的 Excel 文件")
        sys.exit(1)

    print(f"\n找到 {len(excel_files)} 个待汇总的文件:")
    for f in excel_files[:10]:
        print(f"  {os.path.relpath(f, folder)}")
    if len(excel_files) > 10:
        print(f"  ... 共 {len(excel_files)} 个文件")

    # 5. 逐个处理文件，汇总数据
    all_data = []
    for file_path in excel_files:
        try:
            df = pd.read_excel(file_path, sheet_name=0, dtype=str)
        except Exception as e:
            print(f"读取文件失败，跳过: {file_path}, 错误: {e}")
            continue

        existing_cols = [col for col in required_columns if col in df.columns]
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            print(f"  警告: 文件 {os.path.basename(file_path)} 中缺少列 {missing_cols}，对应位置将为空")

        if not existing_cols:
            print(f"  警告: 文件 {os.path.basename(file_path)} 中没有任何需要的列，跳过")
            continue

        sub_df = df[existing_cols].copy()
        sub_df['来源表'] = os.path.basename(file_path)
        all_data.append(sub_df)

    if not all_data:
        print("\n没有成功读取到任何有效数据，程序结束。")
        sys.exit(1)

    # 6. 合并所有数据
    result = pd.concat(all_data, ignore_index=True, sort=False)

    # 7. 确保输出列顺序与需求文件一致，并将“来源表”列放在最后
    final_columns = []
    for col in required_columns:
        if col in result.columns:
            final_columns.append(col)
    for col in result.columns:
        if col not in required_columns:
            final_columns.append(col)
    result = result[final_columns]

    # 8. 输出结果
    output_file = "汇总结果.xlsx"
    try:
        result.to_excel(output_file, index=False, engine='openpyxl')
        print(f"\n汇总完成！结果已保存至: {os.path.abspath(output_file)}")
        print(f"共汇总 {len(result)} 行数据，包含列: {list(result.columns)}")
    except Exception as e:
        print(f"保存结果文件失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()