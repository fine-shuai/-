import pandas as pd
import os
import sys
from pathlib import Path

def find_config_file(script_dir):
    config_files = list(Path(script_dir).glob("*.xlsx")) + list(Path(script_dir).glob("*.xls"))
    if not config_files:
        raise FileNotFoundError("未找到配置文件（.xlsx或.xls），请将配置文件放在脚本同目录下。")
    if len(config_files) > 1:
        print("找到多个Excel文件，将使用第一个：", config_files[0].name)
    return config_files[0]

def parse_config(config_path):
    df_config = pd.read_excel(config_path, dtype=str, keep_default_na=False, sheet_name=0)
    sheet_col = None
    for col in df_config.columns:
        if '子表' in col or 'sheet' in col.lower():
            sheet_col = col
            break
    if sheet_col is None:
        raise ValueError("配置文件中缺少'子表名'列，请确保第一行包含该列。")
    
    output_cols = [c for c in df_config.columns if c != sheet_col]
    
    queries = []
    for idx, row in df_config.iterrows():
        sheet_name = str(row[sheet_col]).strip()
        if sheet_name == '' or sheet_name == 'nan':
            continue
        conditions = {}
        for col in output_cols:
            val = row[col]
            if val != '' and val != 'nan':
                conditions[col] = str(val)
        queries.append({
            'sheet_name': sheet_name,
            'conditions': conditions,
            'output_cols': output_cols,
            'row_id': idx + 2
        })
    return queries, output_cols

def read_sort_config(config_path):
    try:
        df_sort = pd.read_excel(config_path, sheet_name="排序设置", dtype=str, keep_default_na=False)
    except ValueError:
        print("未找到“排序设置”工作表，将不进行排序。")
        return []
    except Exception as e:
        print(f"读取排序设置失败：{e}，将不进行排序。")
        return []
    
    if df_sort.empty:
        return []
    
    sort_col = None
    order_col = None
    for col in df_sort.columns:
        if '排序列' in col or 'sort' in col.lower():
            sort_col = col
        if '升序' in col or '降序' in col or 'order' in col.lower():
            order_col = col
    if sort_col is None or order_col is None:
        print("排序设置工作表需要包含“排序列”和“升序/降序”两列，将不进行排序。")
        return []
    
    sort_rules = []
    for _, row in df_sort.iterrows():
        col_name = str(row[sort_col]).strip()
        if col_name == '' or col_name == 'nan':
            continue
        order_text = str(row[order_col]).strip().lower()
        ascending = True if order_text in ('升序', 'asc') else False
        sort_rules.append((col_name, ascending))
    return sort_rules

def try_convert_numeric(series):
    """尝试将series转换为数值类型（float/int），转换失败则返回原series"""
    try:
        # 先去除空格，空字符串转为NaN
        cleaned = series.replace('', pd.NA).astype(str).str.strip()
        # 尝试转换为数字
        numeric = pd.to_numeric(cleaned, errors='raise')
        return numeric
    except (ValueError, TypeError):
        return series

def apply_sort(df, sort_rules):
    """
    对DataFrame应用排序规则，自动识别数值列并转换。
    返回排序后的DataFrame。
    """
    if not sort_rules or df.empty:
        return df
    
    sort_cols = []
    ascending_list = []
    for col, asc in sort_rules:
        if col not in df.columns:
            print(f"警告：排序列 '{col}' 不在结果表中，将忽略。可用列：{list(df.columns)}")
            continue
        sort_cols.append(col)
        ascending_list.append(asc)
    
    if not sort_cols:
        print("没有有效的排序列，跳过排序。")
        return df
    
    # 为每个排序列生成排序键（尝试数值转换）
    sort_keys = []
    for col in sort_cols:
        key_series = try_convert_numeric(df[col])
        sort_keys.append(key_series)
    
    # 使用numpy的lexsort进行稳定多列排序（性能好且支持混合类型）
    # 注意：需要处理NaN，我们使用na_position='last'，但需要手动实现
    # 更简单：使用sort_values，但传入自定义key函数（pandas 1.3+支持key参数）
    try:
        # 新版本pandas支持key参数
        def make_key(col):
            return lambda s: try_convert_numeric(s)
        df_sorted = df.sort_values(by=sort_cols, ascending=ascending_list, 
                                   na_position='last', key=make_key)
    except TypeError:
        # 旧版本pandas不支持key，采用临时转换列的方式
        temp_cols = []
        for col in sort_cols:
            temp_col = f"__sort_{col}"
            df[temp_col] = try_convert_numeric(df[col])
            temp_cols.append(temp_col)
        df_sorted = df.sort_values(by=temp_cols, ascending=ascending_list, na_position='last')
        df_sorted = df_sorted.drop(columns=temp_cols)
    
    print(f"已按规则排序：{', '.join([f'{col}({"升序" if asc else "降序"})' for col, asc in zip(sort_cols, ascending_list)])}")
    return df_sorted

def get_matched_data(file_path, query):
    sheet_name = query['sheet_name']
    output_cols = query['output_cols']
    conditions = query['conditions']
    
    try:
        xl = pd.ExcelFile(file_path)
        if sheet_name not in xl.sheet_names:
            return pd.DataFrame(columns=output_cols)
    except Exception:
        return pd.DataFrame(columns=output_cols)
    
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
    except Exception:
        return pd.DataFrame(columns=output_cols)
    
    mask = pd.Series([True] * len(df))
    for col, val in conditions.items():
        if col not in df.columns:
            return pd.DataFrame(columns=output_cols)
        mask &= (df[col].astype(str).str.strip() == val)
    
    matched = df[mask]
    if matched.empty:
        return pd.DataFrame(columns=output_cols)
    
    for col in output_cols:
        if col not in matched.columns:
            matched[col] = ''
    
    return matched[output_cols].reset_index(drop=True)

def main(input_path):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    try:
        config_path = find_config_file(script_dir)
        print(f"使用配置文件：{config_path.name}")
    except FileNotFoundError as e:
        print(e)
        input("按回车键退出...")
        return

    try:
        queries, output_cols = parse_config(config_path)
    except Exception as e:
        print(f"解析配置文件失败：{e}")
        input("按回车键退出...")
        return

    if not queries:
        print("配置文件中没有有效的查询行（缺少子表名）。")
        input("按回车键退出...")
        return
    print(f"共解析出 {len(queries)} 个查询条件")

    input_path_obj = Path(input_path)
    if not input_path_obj.exists():
        print(f"错误：路径 {input_path} 不存在")
        input("按回车键退出...")
        return

    if input_path_obj.is_file():
        if input_path_obj.suffix.lower() in ('.xlsx', '.xls'):
            excel_files = [input_path_obj]
        else:
            print("错误：仅支持 .xlsx 或 .xls 文件")
            input("按回车键退出...")
            return
    else:
        excel_files = list(input_path_obj.glob("*.xlsx")) + list(input_path_obj.glob("*.xls"))
        excel_files = [f for f in excel_files if f.resolve() != config_path.resolve()]
        excel_files.sort(key=lambda x: x.name)
        if not excel_files:
            print("目标文件夹中没有Excel文件。")
            input("按回车键退出...")
            return

    print(f"找到 {len(excel_files)} 个待处理文件")

    all_query_results = []
    for q in queries:
        print(f"处理查询条件：子表={q['sheet_name']}, 条件={q['conditions']}")
        file_data_frames = []
        max_rows = 0
        
        for file_path in excel_files:
            matched_df = get_matched_data(file_path, q)
            prefix = file_path.stem
            renamed_df = matched_df.rename(columns={col: f"{prefix}_{col}" for col in matched_df.columns})
            file_data_frames.append(renamed_df)
            if len(renamed_df) > max_rows:
                max_rows = len(renamed_df)
        
        if max_rows == 0:
            print(f"  该查询条件在所有文件中均无匹配，跳过")
            continue
        
        aligned_dfs = []
        for df in file_data_frames:
            if len(df) < max_rows:
                padding = pd.DataFrame('', index=range(max_rows - len(df)), columns=df.columns)
                df_aligned = pd.concat([df, padding], ignore_index=True)
            else:
                df_aligned = df
            aligned_dfs.append(df_aligned)
        
        combined = pd.concat(aligned_dfs, axis=1)
        combined.insert(0, '子表名', q['sheet_name'])
        combined.insert(1, '配置行号', q['row_id'])
        all_query_results.append(combined)
    
    if not all_query_results:
        print("未找到任何匹配数据，结果文件将仅包含表头。")
        final_df = pd.DataFrame()
    else:
        final_df = pd.concat(all_query_results, ignore_index=True)
    
    # 排序
    sort_rules = read_sort_config(config_path)
    if sort_rules and not final_df.empty:
        final_df = apply_sort(final_df, sort_rules)
    
    output_path = Path(script_dir) / "汇总结果.xlsx"
    final_df.to_excel(output_path, index=False, na_rep='')
    print(f"汇总完成！结果已保存至：{output_path}")
    input("按回车键退出...")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("请将Excel文件或文件夹拖拽到此批处理文件上。")
        input("按回车键退出...")
        sys.exit(1)
    main(sys.argv[1])