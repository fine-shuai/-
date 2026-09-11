import pandas as pd
import os
import sys
from pathlib import Path

def find_config_file(script_dir):
    """在脚本所在目录查找Excel配置文件（.xlsx或.xls）"""
    config_files = list(Path(script_dir).glob("*.xlsx")) + list(Path(script_dir).glob("*.xls"))
    if not config_files:
        raise FileNotFoundError("未找到配置文件（.xlsx或.xls），请将配置文件放在脚本同目录下。")
    if len(config_files) > 1:
        print("找到多个Excel文件，将使用第一个：", config_files[0].name)
    return config_files[0]

def parse_config(config_path):
    """
    解析配置文件，返回查询列表。
    每个查询项格式：{
        'sheet_name': str,
        'conditions': {col: value},   # 筛选条件（非空单元格）
        'output_cols': list           # 需要输出的列名（配置表的所有非“子表名”列）
    }
    """
    df_config = pd.read_excel(config_path, dtype=str, keep_default_na=False)
    # 定位“子表名”列（支持列名变体）
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
            'row_id': idx + 2              # 配置行号（Excel行号）
        })
    return queries, output_cols

def process_source_file(file_path, queries, source_file_name):
    """
    处理单个源文件：根据查询条件提取数据。
    返回结果列表，每个元素为字典（包含源文件名、子表名、查询行号及各输出列的值）
    """
    results = []
    try:
        xl = pd.ExcelFile(file_path)
        available_sheets = xl.sheet_names
    except Exception as e:
        print(f"  无法读取文件 {file_path}: {e}")
        return results

    for q in queries:
        sheet_name = q['sheet_name']
        if sheet_name not in available_sheets:
            print(f"  警告：文件 {source_file_name} 中不存在子表 '{sheet_name}'，跳过该条件")
            continue
        
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
        except Exception as e:
            print(f"  读取子表 {sheet_name} 失败: {e}")
            continue
        
        # 构建筛选条件
        mask = pd.Series([True] * len(df))
        for col, val in q['conditions'].items():
            if col not in df.columns:
                print(f"  警告：子表 {sheet_name} 中缺少列 '{col}'，该条件被忽略")
                continue
            mask &= (df[col].astype(str).str.strip() == val)
        
        matched = df[mask]
        if matched.empty:
            continue
        
        for out_col in q['output_cols']:
            if out_col not in df.columns:
                matched[out_col] = ''
        
        # 按原表顺序逐行记录结果
        for _, row in matched.iterrows():
            record = {
                '来源文件': source_file_name,
                '子表名': sheet_name,
                '配置行号': q['row_id']
            }
            for col in q['output_cols']:
                record[col] = row[col] if pd.notna(row[col]) else ''
            results.append(record)
    return results

def main(input_path):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # 1. 查找配置文件
    try:
        config_path = find_config_file(script_dir)
        print(f"使用配置文件：{config_path.name}")
    except FileNotFoundError as e:
        print(e)
        input("按回车键退出...")
        return

    # 2. 解析查询条件
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

    # 3. 获取待处理的Excel文件列表
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
        # 按文件名排序，确保文件顺序可预测（不破坏内部行顺序）
        excel_files.sort(key=lambda x: x.name)
        if not excel_files:
            print("目标文件夹中没有Excel文件。")
            input("按回车键退出...")
            return

    print(f"找到 {len(excel_files)} 个待处理文件")

    # 4. 遍历每个源文件并收集结果（保持处理顺序，不进行额外排序）
    all_results = []
    for file_path in excel_files:
        print(f"正在处理：{file_path.name}")
        res = process_source_file(file_path, queries, file_path.name)
        all_results.extend(res)

    if not all_results:
        print("未找到任何匹配数据，结果文件将仅包含表头。")

    # 5. 输出结果到新Excel文件（不排序，直接按 all_results 顺序）
    result_df = pd.DataFrame(all_results)
    output_path = Path(script_dir) / "汇总结果.xlsx"
    result_df.to_excel(output_path, index=False, na_rep='')
    print(f"汇总完成！结果已保存至：{output_path}")
    input("按回车键退出...")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("请将Excel文件或文件夹拖拽到此批处理文件上。")
        input("按回车键退出...")
        sys.exit(1)
    main(sys.argv[1])