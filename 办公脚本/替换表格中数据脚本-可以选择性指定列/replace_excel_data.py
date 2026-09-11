#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
from pathlib import Path
import argparse
from openpyxl import load_workbook
import xlrd
from xlutils.copy import copy as xl_copy
from xlwt import easyxf

# ---------- 配置 ----------
REPLACE_TABLE_NAME = "替换表.xlsx"   # 放在脚本同目录下
SUPPORTED_EXT = ('.xlsx', '.xls')

# ---------- 工具函数 ----------
def load_replace_rules(replace_file_path):
    """从替换表.xlsx读取规则，返回规则列表，每条规则为(列名, 原数据, 新数据)"""
    wb = load_workbook(replace_file_path, data_only=True)
    sheet = wb.active
    rules = []
    for row in sheet.iter_rows(min_row=2, values_only=True):  # 假设第一行是标题
        col_name = row[0] if row[0] is not None else ""
        old_val = row[1]
        new_val = row[2]
        if old_val is None:
            continue
        # 统一转为字符串（excel中数字会被读成int/float，需转成字符串比较）
        old_val = str(old_val).strip()
        new_val = str(new_val).strip() if new_val is not None else ""
        rules.append((col_name.strip() if col_name else "", old_val, new_val))
    wb.close()
    return rules

def find_column_index(sheet, col_name):
    """在openpyxl工作表中搜索列名，返回列号(1-based)，若找不到返回None"""
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None and str(cell.value).strip() == col_name:
                return cell.column
    return None

def find_column_index_xls(sheet, col_name):
    """在xlrd工作表中搜索列名，返回列索引(0-based)，若找不到返回None"""
    for row_idx in range(sheet.nrows):
        for col_idx in range(sheet.ncols):
            cell_val = sheet.cell_value(row_idx, col_idx)
            if cell_val is not None and str(cell_val).strip() == col_name:
                return col_idx
    return None

def process_xlsx(file_path, rules, dry_run=False):
    """处理.xlsx文件，返回该文件修改的单元格数量"""
    wb = load_workbook(file_path)
    modified_count = 0

    for sheet in wb.worksheets:
        # 1. 构建规则索引：列名 -> [(old, new), ...]
        col_rules = {}
        global_rules = []  # 列名为空的规则
        for col_name, old_val, new_val in rules:
            if col_name == "":
                global_rules.append((old_val, new_val))
            else:
                col_rules.setdefault(col_name, []).append((old_val, new_val))

        # 2. 处理有列名的规则
        for col_name, replacements in col_rules.items():
            col_idx = find_column_index(sheet, col_name)
            if col_idx is None:
                continue  # 列名不存在，跳过

            # 找到列名所在行号
            header_row = None
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None and str(cell.value).strip() == col_name:
                        header_row = cell.row
                        break
                if header_row:
                    break

            if header_row is None:
                continue

            # 遍历该列从下一行开始的数据
            for row in sheet.iter_rows(min_row=header_row + 1, max_col=col_idx, max_row=sheet.max_row):
                cell = row[col_idx - 1]  # iter_rows返回的row是元组，索引从0开始
                if cell.value is None:
                    continue
                cell_str = str(cell.value).strip()
                for old_val, new_val in replacements:
                    if cell_str == old_val:
                        if not dry_run:
                            cell.value = new_val
                        modified_count += 1
                        break  # 一个单元格只替换一次

        # 3. 处理全局替换（列名为空）
        if global_rules:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    cell_str = str(cell.value).strip()
                    for old_val, new_val in global_rules:
                        if cell_str == old_val:
                            if not dry_run:
                                cell.value = new_val
                            modified_count += 1
                            break

    if not dry_run and modified_count > 0:
        wb.save(file_path)
    wb.close()
    return modified_count

def process_xls(file_path, rules, dry_run=False):
    """处理.xls文件，返回修改的单元格数量"""
    # 读取工作簿
    rb = xlrd.open_workbook(file_path, formatting_info=True)
    wb = xl_copy(rb)  # 可写的副本，保留格式
    modified_count = 0

    for sheet_idx in range(rb.nsheets):
        r_sheet = rb.sheet_by_index(sheet_idx)
        w_sheet = wb.get_sheet(sheet_idx)

        col_rules = {}
        global_rules = []
        for col_name, old_val, new_val in rules:
            if col_name == "":
                global_rules.append((old_val, new_val))
            else:
                col_rules.setdefault(col_name, []).append((old_val, new_val))

        # 处理列名规则
        for col_name, replacements in col_rules.items():
            col_idx = find_column_index_xls(r_sheet, col_name)
            if col_idx is None:
                continue

            # 找列名所在行号
            header_row = None
            for rowx in range(r_sheet.nrows):
                for colx in range(r_sheet.ncols):
                    val = r_sheet.cell_value(rowx, colx)
                    if val is not None and str(val).strip() == col_name:
                        header_row = rowx
                        break
                if header_row is not None:
                    break

            if header_row is None:
                continue

            # 遍历该列数据行
            for rowx in range(header_row + 1, r_sheet.nrows):
                cell_val = r_sheet.cell_value(rowx, col_idx)
                if cell_val is None or cell_val == "":
                    continue
                cell_str = str(cell_val).strip()
                for old_val, new_val in replacements:
                    if cell_str == old_val:
                        if not dry_run:
                            w_sheet.write(rowx, col_idx, new_val)
                        modified_count += 1
                        break

        # 处理全局替换
        if global_rules:
            for rowx in range(r_sheet.nrows):
                for colx in range(r_sheet.ncols):
                    cell_val = r_sheet.cell_value(rowx, colx)
                    if cell_val is None or cell_val == "":
                        continue
                    cell_str = str(cell_val).strip()
                    for old_val, new_val in global_rules:
                        if cell_str == old_val:
                            if not dry_run:
                                w_sheet.write(rowx, colx, new_val)
                            modified_count += 1
                            break

    if not dry_run and modified_count > 0:
        wb.save(file_path)
    return modified_count

def process_file(file_path, rules, dry_run):
    """根据扩展名分发处理"""
    ext = file_path.suffix.lower()
    if ext == '.xlsx':
        return process_xlsx(file_path, rules, dry_run)
    elif ext == '.xls':
        return process_xls(file_path, rules, dry_run)
    else:
        return 0

def main():
    parser = argparse.ArgumentParser(description='批量替换Excel表格内容（保留格式）')
    parser.add_argument('target', help='要处理的Excel文件或文件夹路径（支持拖拽）')
    parser.add_argument('--dry-run', action='store_true', help='试运行，只统计不实际修改')
    args = parser.parse_args()

    target_path = Path(args.target)
    if not target_path.exists():
        print(f"错误：路径不存在 - {target_path}")
        sys.exit(1)

    # 读取替换表
    script_dir = Path(__file__).parent
    replace_file = script_dir / REPLACE_TABLE_NAME
    if not replace_file.exists():
        print(f"错误：未找到替换表文件 {REPLACE_TABLE_NAME}，请将其放在脚本同目录下。")
        sys.exit(1)

    print("正在加载替换规则...")
    rules = load_replace_rules(replace_file)
    print(f"已加载 {len(rules)} 条规则。")

    # 收集所有要处理的Excel文件
    files_to_process = []
    if target_path.is_file():
        if target_path.suffix.lower() in SUPPORTED_EXT:
            files_to_process.append(target_path)
        else:
            print(f"错误：不支持的文件类型 - {target_path}")
            sys.exit(1)
    elif target_path.is_dir():
        for ext in SUPPORTED_EXT:
            files_to_process.extend(target_path.rglob(f'*{ext}'))
        if not files_to_process:
            print("未找到任何 Excel 文件。")
            sys.exit(0)
    else:
        print("错误：目标必须是文件或文件夹。")
        sys.exit(1)

    print(f"找到 {len(files_to_process)} 个 Excel 文件待处理。")
    if args.dry_run:
        print("【试运行模式】将不会实际修改文件。")

    total_files_modified = 0
    total_cells_modified = 0

    for file_path in files_to_process:
        print(f"正在处理: {file_path}")
        try:
            changes = process_file(file_path, rules, args.dry_run)
            if changes > 0:
                total_files_modified += 1
                total_cells_modified += changes
                print(f"  -> 修改了 {changes} 个单元格")
            else:
                print(f"  -> 无修改")
        except Exception as e:
            print(f"  -> 处理失败: {e}")

    print("\n===== 统计结果 =====")
    print(f"处理文件总数: {len(files_to_process)}")
    print(f"实际修改文件数: {total_files_modified}")
    print(f"总共替换单元格数: {total_cells_modified}")
    if args.dry_run:
        print("（试运行，未写入任何更改）")
    else:
        print("所有修改已保存。")

if __name__ == '__main__':
    main()