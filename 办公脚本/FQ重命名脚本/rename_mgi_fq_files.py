#!/usr/bin/env python3
"""
华大MGI T7测序下机数据文件批量重命名脚本
根据Excel表格中的映射关系重命名fq文件，并去除_raw部分
支持拖放方式指定目录
"""

import os
import sys
import pandas as pd
import argparse
import re
from pathlib import Path


def parse_arguments():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(description='华大MGI T7测序下机数据文件批量重命名脚本')
    parser.add_argument('-m', '--mapping-file', required=True,
                        help='映射表格文件路径 (Excel格式)')
    parser.add_argument('-d', '--data-dir', required=True,
                        help='包含fq文件的目录路径')
    parser.add_argument('-c', '--original-col', default='原文件名',
                        help='映射表中原文件名列名 (默认: 原文件名)')
    parser.add_argument('-n', '--new-col', default='新文件名',
                        help='映射表中新文件名列名 (默认: 新文件名)')
    parser.add_argument('--dry-run', action='store_true',
                        help='试运行，不实际重命名文件')
    return parser.parse_args()


def read_mapping_file(file_path, original_col, new_col):
    """读取映射表格文件"""
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)

        # 检查必需的列是否存在
        if original_col not in df.columns:
            raise ValueError(f"映射表中找不到列: {original_col}")
        if new_col not in df.columns:
            raise ValueError(f"映射表中找不到列: {new_col}")

        # 创建映射字典
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


def remove_raw_from_filename(filename):
    """从文件名中去除_raw部分"""
    # 使用正则表达式去除_raw
    return re.sub(r'_raw(?=_[12]\.fq\.gz)', '', filename)


def rename_fq_files(data_dir, mapping_dict, dry_run=False):
    """重命名fq文件"""
    # 获取所有fq文件
    fq_files = []
    for ext in ['*.fq.gz', '*.fastq.gz', '*.fq', '*.fastq']:
        fq_files.extend(Path(data_dir).glob(ext))

    if not fq_files:
        print(f"在目录 {data_dir} 中未找到fq文件")
        return

    print(f"找到 {len(fq_files)} 个fq文件")

    # 处理每个fq文件
    renamed_count = 0
    for fq_file in fq_files:
        file_name = fq_file.name
        new_file_name = None

        # 尝试匹配映射表中的每个原文件名
        for old_name, new_name in mapping_dict.items():
            if old_name in file_name:
                # 生成新文件名：先替换原名称，然后去除_raw
                new_file_name = file_name.replace(old_name, new_name)
                new_file_name = remove_raw_from_filename(new_file_name)
                break

        # 如果没有匹配到映射表中的任何名称，跳过此文件
        if not new_file_name:
            print(f"文件 {file_name} 不匹配任何映射关系，跳过")
            continue

        # 构建新文件路径
        new_file_path = os.path.join(data_dir, new_file_name)

        print(f"原文件: {file_name}")
        print(f"新文件: {new_file_name}")

        if not dry_run:
            try:
                # 重命名文件
                os.rename(fq_file, new_file_path)
                print(f"重命名成功: {file_name} -> {new_file_name}")
                renamed_count += 1
            except Exception as e:
                print(f"重命名失败: {e}")
        else:
            renamed_count += 1

        print("-" * 50)

    return renamed_count


def main():
    """主函数"""
    # 检查是否通过拖放方式传递参数
    if len(sys.argv) > 1 and not sys.argv[1].startswith('-'):
        # 假设第一个参数是数据目录
        data_dir = sys.argv[1]
        # 移除这个参数，让argparse正常处理其他参数
        sys.argv = [sys.argv[0]] + sys.argv[2:]

        # 手动设置数据目录参数
        sys.argv.extend(['-d', data_dir])

    args = parse_arguments()

    print("华大MGI T7测序下机数据文件批量重命名脚本")
    print("=" * 60)
    print("功能: 根据映射表重命名fq文件，并去除文件名中的'_raw'部分")
    print("=" * 60)

    # 读取映射文件
    mapping_dict = read_mapping_file(args.mapping_file, args.original_col, args.new_col)

    if not mapping_dict:
        print("未找到有效的映射关系，请检查映射文件")
        return

    print(f"成功读取 {len(mapping_dict)} 条映射关系")

    # 执行重命名操作
    renamed_count = rename_fq_files(args.data_dir, mapping_dict, args.dry_run)

    if args.dry_run:
        print(f"\n试运行完成，共计划重命名 {renamed_count} 个文件")
    else:
        print(f"\n文件重命名完成，共重命名 {renamed_count} 个文件")

    # 在Windows中保持窗口打开，直到用户按键
    if os.name == 'nt':  # Windows系统
        input("\n按Enter键退出...")


if __name__ == '__main__':
    main()