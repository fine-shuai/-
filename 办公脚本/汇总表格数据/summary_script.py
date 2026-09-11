import pandas as pd
import os
import sys
from pathlib import Path
from typing import List, Dict, Tuple, Optional

# ---------- 智能读取文件（支持 Excel / TSV / CSV）----------
def read_file_to_dataframes(file_path: Path) -> Dict[str, pd.DataFrame]:
    """
    尝试以多种方式读取文件，返回 {sheet_name: DataFrame} 字典。
    优先尝试 Excel，若失败则尝试 TSV，再尝试 CSV。
    对于非 Excel 文件，默认工作表名为 'Sheet1'。
    """
    # 1. 尝试作为 Excel 读取
    try:
        xl = pd.ExcelFile(file_path)
        sheets = {}
        for sheet in xl.sheet_names:
            sheets[sheet] = pd.read_excel(file_path, sheet_name=sheet, dtype=str, keep_default_na=False)
        return sheets
    except Exception as e:
        print(f"  非 Excel 格式，尝试作为文本文件读取...")

    # 2. 尝试作为 TSV 读取（制表符分隔）
    encodings = ['utf-8', 'gbk', 'latin1']
    for enc in encodings:
        try:
            df = pd.read_csv(file_path, sep='\t', dtype=str, keep_default_na=False, encoding=enc)
            return {'Sheet1': df}
        except UnicodeDecodeError:
            continue
        except Exception:
            continue

    # 3. 尝试作为 CSV 读取（逗号分隔）
    for enc in encodings:
        try:
            df = pd.read_csv(file_path, dtype=str, keep_default_na=False, encoding=enc)
            return {'Sheet1': df}
        except Exception:
            continue

    # 所有方式都失败
    raise ValueError(f"无法读取文件 {file_path.name}，请检查文件格式是否为 Excel/TSV/CSV。")

# ---------- 配置查找 ----------
def find_config_file(script_dir: Path) -> Path:
    config_files = list(script_dir.glob("*.xlsx")) + list(script_dir.glob("*.xls"))
    if not config_files:
        raise FileNotFoundError("未找到配置文件（.xlsx或.xls），请将配置文件放在脚本同目录下。")
    if len(config_files) > 1:
        print(f"找到多个Excel文件，将使用第一个：{config_files[0].name}")
    return config_files[0]

# ---------- 解析配置 ----------
def parse_config(config_path: Path) -> Tuple[List[Dict], List[str]]:
    df_config = pd.read_excel(config_path, dtype=str, keep_default_na=False)
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
        if sheet_name == '' or sheet_name.lower() == 'nan':
            sheet_name = None  # 通配所有 sheet

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

# ---------- 处理单个源文件 ----------
def process_source_file(file_path: Path, queries: List[Dict], source_file_name: str) -> List[Dict]:
    results = []
    try:
        sheets_dict = read_file_to_dataframes(file_path)
    except Exception as e:
        print(f"  读取文件失败: {e}")
        return results

    available_sheets = list(sheets_dict.keys())

    for q in queries:
        # 确定目标 sheet 列表
        if q['sheet_name'] is None:
            target_sheets = available_sheets
        else:
            target_sheets = [q['sheet_name']] if q['sheet_name'] in available_sheets else []

        for sheet_name in target_sheets:
            df = sheets_dict[sheet_name]

            # 构建筛选条件掩码
            mask = pd.Series([True] * len(df))
            for col, val in q['conditions'].items():
                if col not in df.columns:
                    print(f"  警告：子表 {sheet_name} 中缺少列 '{col}'，该条件被忽略")
                    continue
                mask &= (df[col].astype(str).str.strip() == val)

            matched = df[mask]
            if matched.empty:
                continue

            # 确保输出列存在
            for out_col in q['output_cols']:
                if out_col not in df.columns:
                    matched[out_col] = ''

            # 逐行记录
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

# ---------- 主函数 ----------
def main(input_path: str):
    script_dir = Path(__file__).parent.resolve()

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
        print("配置文件中没有有效的查询行。")
        input("按回车键退出...")
        return
    print(f"共解析出 {len(queries)} 个查询条件")

    # 3. 获取待处理的文件列表（递归遍历所有子文件夹）
    input_path_obj = Path(input_path)
    if not input_path_obj.exists():
        print(f"错误：路径 {input_path} 不存在")
        input("按回车键退出...")
        return

    if input_path_obj.is_file():
        # 如果拖入的是单个文件，只处理该文件（不递归）
        if input_path_obj.suffix.lower() in ('.xlsx', '.xls'):
            excel_files = [input_path_obj]
        else:
            print("错误：仅支持 .xlsx 或 .xls 文件")
            input("按回车键退出...")
            return
    else:
        # 如果拖入的是文件夹，递归搜索所有子文件夹中的 Excel 文件
        excel_files = []
        for ext in ['*.xlsx', '*.xls']:
            excel_files.extend(input_path_obj.rglob(ext))
        # 排除配置文件本身（避免将配置文件作为数据源）
        excel_files = [f for f in excel_files if f.resolve() != config_path.resolve()]
        excel_files.sort(key=lambda x: x.name)  # 按文件名排序，保证顺序可复现
        if not excel_files:
            print("目标文件夹及其子文件夹中没有Excel文件。")
            input("按回车键退出...")
            return

    print(f"找到 {len(excel_files)} 个待处理文件")

    # 4. 遍历每个源文件并收集结果
    all_results = []
    for file_path in excel_files:
        print(f"正在处理：{file_path}")
        res = process_source_file(file_path, queries, file_path.name)
        all_results.extend(res)

    if not all_results:
        print("未找到任何匹配数据，结果文件将仅包含表头。")

    # 5. 输出结果
    result_df = pd.DataFrame(all_results)
    output_path = script_dir / "汇总结果.xlsx"
    result_df.to_excel(output_path, index=False, na_rep='')
    print(f"汇总完成！结果已保存至：{output_path}")
    input("按回车键退出...")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("请将Excel文件或文件夹拖拽到此批处理文件上。")
        input("按回车键退出...")
        sys.exit(1)
    main(sys.argv[1])