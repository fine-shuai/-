import pandas as pd
import os
import sys

def main():
    config_path = '汇总需求.xlsx'
    if not os.path.exists(config_path):
        print(f"找不到配置文件：{config_path}")
        input("按回车键退出...")
        return

    # 读取透视列配置
    try:
        config_pivot = pd.read_excel(config_path, sheet_name='透视列')
        pivot_col_name = config_pivot.iloc[0]['新列']          # 样本编号(Sample_ID)
        value_col_name = config_pivot.iloc[0]['值列']          # 替代等位基因频率(Alternative Allele Frequency, AAF)
        index_col_name = '位点编号(LOCUS_ID)'                   # 保持位点编号为行
    except Exception as e:
        print(f"读取透视列配置出错: {e}")
        input("按回车键退出...")
        return

    print("读取配置成功。")
    print(f"透视配置: 列为[{pivot_col_name}], 值为[{value_col_name}], 行为[{index_col_name}]")

    # 第一步提取规则
    target_sheet = 'Total_SNP'
    required_cols = ['样本编号(Sample_ID)', '位点编号(LOCUS_ID)', '替代等位基因频率(Alternative Allele Frequency, AAF)']

    # 获取待处理路径
    if len(sys.argv) > 1:
        file_paths = sys.argv[1:]
    else:
        file_paths = [input("请拖入Excel文件或文件夹路径: ").strip('"')]

    total_files = 0
    output_dir = '输出结果'
    os.makedirs(output_dir, exist_ok=True)

    # 遍历用户输入的路径，处理不同的根目录
    for path in file_paths:
        path = path.strip('"')
        if os.path.isdir(path):
            root_path = path
            print(f"正在处理文件夹: {root_path}")
            for root, dirs, files in os.walk(path):
                for file in files:
                    if file.endswith('.xlsx') and not file.startswith('~$'):
                        file_path = os.path.join(root, file)
                        process_and_save(file_path, root_path, target_sheet, required_cols, pivot_col_name, value_col_name, index_col_name, output_dir)
                        total_files += 1
        elif os.path.isfile(path):
            root_path = os.path.dirname(path)
            process_and_save(path, root_path, target_sheet, required_cols, pivot_col_name, value_col_name, index_col_name, output_dir)
            total_files += 1
        else:
            print(f"路径不存在: {path}")

    print(f"共处理了 {total_files} 个文件。")
    print(f"所有处理完成，结果已按原目录结构保存在：{os.path.abspath(output_dir)}")
    input("按回车键退出...")

def process_and_save(file_path, root_path, target_sheet, required_cols, pivot_col_name, value_col_name, index_col_name, output_dir):
    print(f"正在处理源文件：{file_path}")
    try:
        xls = pd.ExcelFile(file_path)
        if target_sheet not in xls.sheet_names:
            print(f"  找不到子表 [{target_sheet}]，跳过。")
            return

        df = pd.read_excel(xls, sheet_name=target_sheet)
        
        # 检查列是否存在
        if not all(col in df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in df.columns]
            print(f"  子表 [{target_sheet}] 缺少列：{missing}，跳过。")
            return

        extracted_df = df[required_cols].copy()
        
        # 单个文件内部的重复检查
        duplicates = extracted_df.duplicated(subset=[index_col_name, pivot_col_name], keep=False)
        if duplicates.any():
            print(f"  警告：文件内部存在 {duplicates.sum()} 行重复 [位点编号 + 样本编号] 的数据！")
            print(f"  由于配置了'不要聚合'，无法生成该文件，请检查该源表内部是否有重复行。")
            return

        # 透视
        df_wide = extracted_df.pivot(index=index_col_name, columns=pivot_col_name, values=value_col_name)
        df_wide = df_wide.fillna('NA')
        df_wide.reset_index(inplace=True)
        df_wide.columns.name = None

        # 计算相对路径，保留源目录结构
        rel_path = os.path.relpath(file_path, root_path)
        rel_dir = os.path.dirname(rel_path)
        
        # 构建目标目录并创建
        target_dir = os.path.join(output_dir, rel_dir)
        os.makedirs(target_dir, exist_ok=True)
        
        # 生成文件
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_filename = f"{base_name}_透视结果.xlsx"
        output_path = os.path.join(target_dir, output_filename)
        
        df_wide.to_excel(output_path, index=False)
        print(f"  提取 {len(extracted_df)} 行，透视成功，保存至：{output_path}")

    except Exception as e:
        print(f"  处理文件时出错: {e}")

if __name__ == "__main__":
    main()