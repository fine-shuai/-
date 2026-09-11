#!/usr/bin/env python3
"""
二代测序数据文件批量重命名脚本
支持华大MGI T7和Illumina两种数据格式
"""

import os
import sys
import pandas as pd
import argparse
import re
import shutil
from pathlib import Path

# 尝试导入 openpyxl（仅华大模式需要）
try:
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("警告: 未安装 openpyxl，Excel 格式保留功能可能受限。建议安装: pip install openpyxl")


# ==================== 公共函数 ====================
def parse_arguments():
    parser = argparse.ArgumentParser(description='二代测序数据文件批量重命名脚本')
    parser.add_argument('-m', '--mapping-file', required=True,
                        help='映射表格文件路径 (Excel格式)')
    parser.add_argument('-d', '--data-dir', required=True,
                        help='包含测序数据的目录路径')
    parser.add_argument('--mode', choices=['mgi', 'illumina'],
                        help='处理模式: mgi(华大) 或 illumina')
    parser.add_argument('--dry-run', action='store_true',
                        help='试运行，不实际修改文件')
    parser.add_argument('--overwrite', action='store_true',
                        help='覆盖已存在的目标文件')
    parser.add_argument('--batch', action='store_true',
                        help='批量模式：将-d指定的目录视为根文件夹，处理其下所有子文件夹（华大模式有效）')
    parser.add_argument('-c', '--original-col', default='原文件名',
                        help='映射表中原文件名列名 (默认: 原文件名)')
    parser.add_argument('-n', '--new-col', default='新文件名',
                        help='映射表中新文件名列名 (默认: 新文件名)')
    parser.add_argument('--sample-barcode-col', default='样本条码',
                        help='华大模式中Excel表格样本条码列名 (默认: 样本条码)')
    return parser.parse_args()


def read_mapping_file(file_path, original_col='原文件名', new_col='新文件名'):
    """读取映射表格文件，返回字典 {原名称: 新名称}"""
    try:
        df = pd.read_excel(file_path)
        if original_col not in df.columns:
            raise ValueError(f"映射表中找不到列: {original_col}")
        if new_col not in df.columns:
            raise ValueError(f"映射表中找不到列: {new_col}")

        mapping_dict = {}
        for _, row in df.iterrows():
            original_name = str(row[original_col]).strip()
            new_name = str(row[new_col]).strip()
            if original_name and new_name and original_name != 'nan' and new_name != 'nan':
                mapping_dict[original_name] = new_name
        return mapping_dict
    except Exception as e:
        print(f"读取映射文件时出错: {e}")
        return {}


def get_user_mode():
    """交互式获取模式选择"""
    print("\n请选择处理模式：")
    print("1. 华大 MGI T7 数据 (支持子文件夹、Excel更新)")
    print("2. Illumina 数据 (双端fastq文件，支持子文件夹)")
    while True:
        choice = input("请输入数字 (1 或 2): ").strip()
        if choice == '1':
            return 'mgi'
        elif choice == '2':
            return 'illumina'
        else:
            print("输入无效，请输入 1 或 2")


# ==================== 华大 MGI 模式 ====================
def rename_fq_files_mgi(data_dir, mapping_dict, overwrite=False, dry_run=False):
    """重命名华大 MGI fq 文件（去除_raw）"""
    fq_files = []
    for ext in ['*.fq.gz', '*.fastq.gz', '*.fq', '*.fastq']:
        fq_files.extend(Path(data_dir).glob(ext))

    if not fq_files:
        print(f"在目录 {data_dir} 中未找到fq文件")
        return 0

    print(f"找到 {len(fq_files)} 个fq文件")
    renamed_count = 0

    for fq_file in fq_files:
        file_name = fq_file.name
        new_file_name = None

        for old_name, new_name in mapping_dict.items():
            if old_name in file_name:
                # 生成新文件名：替换原名称，然后去除_raw
                new_file_name = file_name.replace(old_name, new_name)
                new_file_name = re.sub(r'_raw(?=_[12]\.fq\.gz)', '', new_file_name)
                break

        if not new_file_name:
            print(f"文件 {file_name} 不匹配任何映射关系，跳过")
            continue

        new_file_path = os.path.join(data_dir, new_file_name)
        print(f"原文件: {file_name}")
        print(f"新文件: {new_file_name}")

        if not dry_run:
            try:
                if os.path.exists(new_file_path):
                    if overwrite:
                        os.remove(new_file_path)
                        print(f"已删除已存在的文件: {new_file_name}")
                    else:
                        print(f"警告: 目标文件已存在，跳过重命名 (使用 --overwrite 覆盖)")
                        continue

                os.rename(fq_file, new_file_path)
                print(f"重命名成功")
                renamed_count += 1
            except Exception as e:
                print(f"重命名失败: {e}")
        else:
            renamed_count += 1

        print("-" * 50)

    return renamed_count


def update_excel_file_mgi(data_dir, mapping_dict, sample_barcode_col='样本条码', dry_run=False):
    """更新华大 MGI 目录中的 Excel 文件（.xls 或 .xlsx）"""
    excel_files = list(Path(data_dir).glob("*.xls")) + list(Path(data_dir).glob("*.xlsx"))
    excel_files = [f for f in excel_files if not f.name.startswith('~$')]

    if not excel_files:
        print(f"在目录 {data_dir} 中未找到Excel文件")
        return 0

    if len(excel_files) > 1:
        print(f"警告: 找到多个Excel文件，将处理第一个: {excel_files[0].name}")

    excel_file = excel_files[0]
    file_name = excel_file.name
    print(f"处理Excel文件: {file_name}")

    try:
        if HAS_OPENPYXL and excel_file.suffix == '.xlsx':
            wb = load_workbook(filename=excel_file)
            if 'Sheet1' not in wb.sheetnames:
                print(f"文件 {file_name} 中没有 'Sheet1' 工作表，跳过")
                return 0

            ws = wb['Sheet1']
            barcode_col = None
            for col in range(1, ws.max_column + 1):
                if ws.cell(row=1, column=col).value == sample_barcode_col:
                    barcode_col = col
                    break

            if barcode_col is None:
                print(f"文件 {file_name} 的Sheet1工作表中没有 '{sample_barcode_col}' 列，跳过")
                return 0

            changes_made = False
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=barcode_col)
                if cell.value is None:
                    continue
                value_str = str(cell.value).strip()
                for old_name, new_name in mapping_dict.items():
                    # 去除原始名称中的_raw（如果有）
                    clean_old = re.sub(r'_raw$', '', old_name)
                    if value_str == old_name or old_name in value_str or value_str == clean_old:
                        new_value = value_str.replace(old_name, new_name)
                        new_value = re.sub(r'_raw$', '', new_value)
                        cell.value = new_value
                        print(f"更新样本条码: {value_str} -> {new_value}")
                        changes_made = True
                        break

            if changes_made:
                if not dry_run:
                    backup_file = excel_file.with_suffix(excel_file.suffix + '.bak')
                    shutil.copy2(excel_file, backup_file)
                    print(f"已创建备份文件: {backup_file.name}")
                    wb.save(excel_file)
                    print(f"Excel文件更新成功: {file_name}")
                    return 1
                else:
                    print(f"计划更新Excel文件: {file_name}")
                    return 1
            else:
                print(f"Excel文件无需更新: {file_name}")
                return 0

        else:
            # 回退方案：使用 pandas（对 .xls 或未安装 openpyxl 时使用）
            excel_data = pd.read_excel(excel_file, sheet_name=None)
            if 'Sheet1' not in excel_data:
                print(f"文件 {file_name} 中没有 'Sheet1' 工作表，跳过")
                return 0

            df = excel_data['Sheet1']
            if sample_barcode_col not in df.columns:
                print(f"文件 {file_name} 的Sheet1工作表中没有 '{sample_barcode_col}' 列，跳过")
                return 0

            changes_made = False
            for i, value in enumerate(df[sample_barcode_col]):
                if pd.isna(value):
                    continue
                value_str = str(value).strip()
                for old_name, new_name in mapping_dict.items():
                    clean_old = re.sub(r'_raw$', '', old_name)
                    if value_str == old_name or old_name in value_str or value_str == clean_old:
                        new_value = value_str.replace(old_name, new_name)
                        new_value = re.sub(r'_raw$', '', new_value)
                        df.at[i, sample_barcode_col] = new_value
                        print(f"更新样本条码: {value_str} -> {new_value}")
                        changes_made = True
                        break

            if changes_made:
                if not dry_run:
                    backup_file = excel_file.with_suffix(excel_file.suffix + '.bak')
                    shutil.copy2(excel_file, backup_file)
                    print(f"已创建备份文件: {backup_file.name}")

                    excel_data['Sheet1'] = df
                    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                        for sheet_name, sheet_data in excel_data.items():
                            sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

                    print(f"Excel文件更新成功: {file_name}")
                    return 1
                else:
                    print(f"计划更新Excel文件: {file_name}")
                    return 1
            else:
                print(f"Excel文件无需更新: {file_name}")
                return 0

    except Exception as e:
        print(f"处理Excel文件 {file_name} 时出错: {e}")
        return 0


def process_single_directory_mgi(data_dir, mapping_dict, sample_barcode_col, overwrite, dry_run):
    """处理单个目录：重命名fq文件并更新Excel文件（华大模式）"""
    print(f"\n{'='*60}")
    print(f"处理目录: {data_dir}")
    print('='*60)

    fq_renamed = rename_fq_files_mgi(data_dir, mapping_dict, overwrite, dry_run)
    excel_updated = update_excel_file_mgi(data_dir, mapping_dict, sample_barcode_col, dry_run)

    if dry_run:
        print(f"试运行完成，计划重命名 {fq_renamed} 个fq文件，更新 {excel_updated} 个Excel文件")
    else:
        print(f"处理完成，重命名 {fq_renamed} 个fq文件，更新 {excel_updated} 个Excel文件")
    return fq_renamed, excel_updated


# ==================== Illumina 模式 ====================
def rename_fq_files_illumina(data_dir, mapping_dict, overwrite=False, dry_run=False):
    """
    重命名 Illumina 双端 fq 文件
    文件名格式: {prefix}_R1_001.fastq, {prefix}_R2_001.fastq
    映射表中原文件名为公共前缀，替换为新前缀，保留 _R1_001.fastq 等后缀
    """
    fq_files = []
    for ext in ['*.fastq', '*.fq', '*.fastq.gz', '*.fq.gz']:
        fq_files.extend(Path(data_dir).glob(ext))

    if not fq_files:
        print(f"在目录 {data_dir} 中未找到fq文件")
        return 0

    print(f"找到 {len(fq_files)} 个fq文件")
    renamed_count = 0

    for fq_file in fq_files:
        file_name = fq_file.name
        new_file_name = None

        # 尝试匹配映射表中的每个原名称（公共前缀）
        for old_name, new_name in mapping_dict.items():
            # 检查文件名是否以原名称开头
            if file_name.startswith(old_name):
                # 新文件名 = 新前缀 + 原文件名去除原前缀后的部分
                suffix = file_name[len(old_name):]
                new_file_name = new_name + suffix
                break

        if not new_file_name:
            print(f"文件 {file_name} 不匹配任何映射关系，跳过")
            continue

        new_file_path = os.path.join(data_dir, new_file_name)
        print(f"原文件: {file_name}")
        print(f"新文件: {new_file_name}")

        if not dry_run:
            try:
                if os.path.exists(new_file_path):
                    if overwrite:
                        os.remove(new_file_path)
                        print(f"已删除已存在的文件: {new_file_name}")
                    else:
                        print(f"警告: 目标文件已存在，跳过重命名 (使用 --overwrite 覆盖)")
                        continue

                os.rename(fq_file, new_file_path)
                print(f"重命名成功")
                renamed_count += 1
            except Exception as e:
                print(f"重命名失败: {e}")
        else:
            renamed_count += 1

        print("-" * 50)

    return renamed_count


def process_single_directory_illumina(data_dir, mapping_dict, overwrite, dry_run):
    """处理单个目录：重命名配对fq文件（Illumina模式）"""
    print(f"\n{'='*60}")
    print(f"处理目录: {data_dir}")
    print('='*60)

    fq_renamed = rename_fq_files_illumina(data_dir, mapping_dict, overwrite, dry_run)

    if dry_run:
        print(f"试运行完成，计划重命名 {fq_renamed} 个fq文件")
    else:
        print(f"处理完成，重命名 {fq_renamed} 个fq文件")
    return fq_renamed


# ==================== 主函数 ====================
def main():
    # 处理拖放参数：如果第一个参数不是以 - 开头，视为数据目录
    if len(sys.argv) > 1 and not sys.argv[1].startswith('-'):
        data_dir = sys.argv[1]
        sys.argv = [sys.argv[0]] + sys.argv[2:]
        sys.argv.extend(['-d', data_dir])

    args = parse_arguments()

    # 确定模式：优先使用命令行参数，否则交互式选择
    if args.mode is None:
        mode = get_user_mode()
    else:
        mode = args.mode

    print(f"\n当前模式: {'华大 MGI' if mode == 'mgi' else 'Illumina'}")
    print("=" * 70)

    # 读取映射文件
    mapping_dict = read_mapping_file(args.mapping_file, args.original_col, args.new_col)
    if not mapping_dict:
        print("未找到有效的映射关系，请检查映射文件")
        return
    print(f"成功读取 {len(mapping_dict)} 条映射关系")

    data_dir = Path(args.data_dir)
    if not data_dir.exists():
        print(f"错误: 目录不存在 - {data_dir}")
        return

    # 根据模式决定处理逻辑
    if mode == 'mgi':
        # 华大模式：自动判断是否批量处理
        use_batch = args.batch
        if not use_batch:
            # 检查当前目录是否有 fq 文件
            has_fq = any(data_dir.glob('*.fq*')) or any(data_dir.glob('*.fastq*'))
            if not has_fq:
                subdirs = [d for d in data_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]
                if subdirs:
                    print("检测到当前目录无fq文件，但存在子文件夹，自动进入批量处理模式。")
                    use_batch = True
                else:
                    print("当前目录无fq文件且无子文件夹，无法处理。")
                    return
            else:
                subdirs = [d for d in data_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]
                if subdirs:
                    print("警告: 当前目录既有fq文件又有子文件夹，将只处理当前目录。如需批量处理子文件夹，请使用 --batch 参数。")

        if use_batch:
            print("模式: 批量处理根目录下的所有子文件夹")
            subdirs = [d for d in data_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]
            if not subdirs:
                print(f"在根目录 {data_dir} 中未找到任何子文件夹")
                return
            print(f"找到 {len(subdirs)} 个子文件夹")
            dirs_to_process = subdirs
        else:
            print("模式: 单目录处理")
            dirs_to_process = [data_dir]

        total_fq = 0
        total_excel = 0
        for sub_dir in dirs_to_process:
            fq_count, excel_count = process_single_directory_mgi(
                sub_dir, mapping_dict, args.sample_barcode_col, args.overwrite, args.dry_run)
            total_fq += fq_count
            total_excel += excel_count

        print("\n" + "="*70)
        if args.dry_run:
            print(f"试运行总结: 共计划重命名 {total_fq} 个fq文件，更新 {total_excel} 个Excel文件")
        else:
            print(f"处理总结: 共重命名 {total_fq} 个fq文件，更新 {total_excel} 个Excel文件")
        print("="*70)

    else:  # illumina 模式
        # Illumina 模式：支持批量处理子文件夹（自动检测或使用 --batch 参数）
        use_batch = args.batch
        if not use_batch:
            # 检查当前目录是否有 fq 文件
            has_fq = any(data_dir.glob('*.fastq*')) or any(data_dir.glob('*.fq*'))
            if not has_fq:
                subdirs = [d for d in data_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]
                if subdirs:
                    print("检测到当前目录无fq文件，但存在子文件夹，自动进入批量处理模式。")
                    use_batch = True
                else:
                    print("当前目录无fq文件且无子文件夹，无法处理。")
                    return
            else:
                subdirs = [d for d in data_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]
                if subdirs:
                    print("警告: 当前目录既有fq文件又有子文件夹，将只处理当前目录。如需批量处理子文件夹，请使用 --batch 参数。")

        if use_batch:
            print("模式: 批量处理根目录下的所有子文件夹")
            subdirs = [d for d in data_dir.iterdir() if d.is_dir() and not d.name.startswith('.')]
            if not subdirs:
                print(f"在根目录 {data_dir} 中未找到任何子文件夹")
                return
            print(f"找到 {len(subdirs)} 个子文件夹")
            dirs_to_process = subdirs
        else:
            print("模式: 单目录处理")
            dirs_to_process = [data_dir]

        total_fq = 0
        for sub_dir in dirs_to_process:
            fq_count = process_single_directory_illumina(
                sub_dir, mapping_dict, args.overwrite, args.dry_run)
            total_fq += fq_count

        print("\n" + "="*70)
        if args.dry_run:
            print(f"试运行总结: 共计划重命名 {total_fq} 个fq文件")
        else:
            print(f"处理总结: 共重命名 {total_fq} 个fq文件")
        print("="*70)

    if os.name == 'nt':
        input("\n按Enter键退出...")


if __name__ == '__main__':
    main()