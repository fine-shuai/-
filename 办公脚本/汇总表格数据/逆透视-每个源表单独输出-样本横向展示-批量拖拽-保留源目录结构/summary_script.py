import pandas as pd
import os
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed

def main():
    config_path = '汇总需求.xlsx'
    if not os.path.exists(config_path):
        print(f"找不到配置文件：{config_path}")
        input("按回车键退出...")
        return

    # 读取透视列配置
    try:
        config_pivot = pd.read_excel(config_path, sheet_name='透视列')
        pivot_col_name = config_pivot.iloc[0]['新列']
        pivot_value_col = config_pivot.iloc[0]['值列']
        index_col_name = '位点编号(LOCUS_ID)'
    except Exception as e:
        print(f"读取透视列配置出错: {e}")
        input("按回车键退出...")
        return

    # 读取 Sheet1 全部有效条件
    try:
        config_sheet1 = pd.read_excel(config_path, sheet_name='Sheet1')
        config_sheet1 = config_sheet1[config_sheet1['子表名'].notna()]
        configs = config_sheet1.to_dict('records')
        print(f"读取配置成功，共 {len(configs)} 个查询条件。")
    except Exception as e:
        print(f"读取 Sheet1 配置出错: {e}")
        input("按回车键退出...")
        return

    # 获取待处理路径
    if len(sys.argv) > 1:
        file_paths = sys.argv[1:]
    else:
        file_paths = [input("请拖入Excel文件或文件夹路径: ").strip('"')]

    # 收集所有待处理的文件列表
    all_files = []
    for path in file_paths:
        path = path.strip('"')
        if os.path.isdir(path):
            for root, dirs, files in os.walk(path):
                for file in files:
                    if file.endswith('.xlsx') and not file.startswith('~$'):
                        all_files.append((os.path.join(root, file), path))
        elif os.path.isfile(path):
            all_files.append((path, os.path.dirname(path)))
        else:
            print(f"路径不存在: {path}")

    print(f"共找到 {len(all_files)} 个待处理文件。")
    
    if not all_files:
        input("按回车键退出...")
        return

    # 开启多线程处理（建议设为 CPU 核心数的 2 倍，最高不超过 8，防止内存溢出）
    max_workers = min(8, os.cpu_count() or 1)
    print(f"正在启用 {max_workers} 线程并行处理，速度大幅提升...")
    
    output_dir = '输出结果'
    os.makedirs(output_dir, exist_ok=True)
    
    success_count = 0
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # 提交所有任务
        future_to_file = {
            executor.submit(process_file, file_path, root_path, configs, pivot_col_name, pivot_value_col, index_col_name, output_dir): file_path
            for file_path, root_path in all_files
        }
        
        # 等待完成
        for future in as_completed(future_to_file):
            file_path = future_to_file[future]
            try:
                success = future.result()
                if success:
                    success_count += 1
            except Exception as e:
                print(f"处理文件 {file_path} 时发生未知异常: {e}")

    print(f"\n全部处理完成！成功生成 {success_count} 个结果文件。")
    print(f"结果已按原目录结构保存至：{os.path.abspath(output_dir)}")
    input("按回车键退出...")

def process_file(file_path, root_path, configs, pivot_col_name, pivot_value_col, index_col_name, output_dir):
    # 核心性能优化：只将文件打开一次！
    try:
        xls = pd.ExcelFile(file_path)
    except:
        return False

    # 1. 收集所有需要用到的子表名和列名
    needed_sheets = set([cfg['子表名'] for cfg in configs if pd.notna(cfg['子表名'])])
    # 只需要读取这3列，极大提高读取速度！
    needed_cols = [pivot_col_name, index_col_name, pivot_value_col]
    
    # 2. 将读取到的数据缓存到内存中
    sheet_cache = {}
    for sheet in needed_sheets:
        if sheet in xls.sheet_names:
            # 使用 usecols 参数：只提取需要的列
            sheet_cache[sheet] = pd.read_excel(xls, sheet_name=sheet, usecols=lambda c: c in needed_cols)

    combined_dfs = []

    # 3. 遍历配置，直接从缓存中拿数据，不再重复读文件
    for cfg in configs:
        target_sheet = cfg['子表名']
        filter_sample = cfg['样本编号(Sample_ID)']
        filter_locus = cfg['位点编号(LOCUS_ID)']

        if target_sheet not in sheet_cache:
            continue

        df = sheet_cache[target_sheet].copy()

        # 精确过滤
        if pd.notna(filter_sample) and str(filter_sample).strip() != '':
            df = df[df[pivot_col_name] == str(filter_sample).strip()]
        if pd.notna(filter_locus) and str(filter_locus).strip() != '':
            df = df[df[index_col_name] == str(filter_locus).strip()]

        if not df.empty:
            combined_dfs.append(df)

    if not combined_dfs:
        return False

    # 合并与透视
    combined_df = pd.concat(combined_dfs, ignore_index=True)
    
    # 检查重复
    duplicates = combined_df.duplicated(subset=[index_col_name, pivot_col_name], keep=False)
    if duplicates.any():
        print(f"警告：{os.path.basename(file_path)} 存在 {duplicates.sum()} 行重复数据。请检查配置或源文件。")
        return False

    df_wide = combined_df.pivot(index=index_col_name, columns=pivot_col_name, values=pivot_value_col)
    df_wide = df_wide.fillna('NA')
    df_wide.reset_index(inplace=True)
    df_wide.columns.name = None

    # 4. 输出文件（保留目录结构）
    rel_path = os.path.relpath(file_path, root_path)
    rel_dir = os.path.dirname(rel_path)
    target_dir = os.path.join(output_dir, rel_dir)
    os.makedirs(target_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_filename = f"{base_name}_透视结果.xlsx"
    output_path = os.path.join(target_dir, output_filename)

    df_wide.to_excel(output_path, index=False)
    print(f"处理完成：{base_name} (提取 {len(combined_df)} 行)")
    return True

if __name__ == "__main__":
    main()