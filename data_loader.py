"""
数据读取模块

负责读取Excel文件，解析为统一的DataFrame格式
新增：读取历史总表，识别姓名和上次排名列
兼容性：自动检测文件扩展名并选择合适引擎，处理导入错误
"""

import pandas as pd
import os
import re
from typing import List, Dict, Optional, Any

# 导入配置
from config import HISTORY_NAME_VARIANTS, HISTORY_RANK_VARIANTS

def _read_excel_file(file_path: str, sheet_name=0, dtype=None, keep_default_na=False):
    """
    读取Excel文件，根据文件扩展名自动选择引擎，并处理缺失库的情况
    """
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.xlsx':
            # .xlsx 文件使用 openpyxl 引擎
            return pd.read_excel(file_path, sheet_name=sheet_name, dtype=dtype,
                                 keep_default_na=keep_default_na, engine='openpyxl')
        elif ext == '.xls':
            # .xls 文件使用 xlrd 引擎
            return pd.read_excel(file_path, sheet_name=sheet_name, dtype=dtype,
                                 keep_default_na=keep_default_na, engine='xlrd')
        else:
            # 未知扩展名，让 pandas 自动选择（可能报错）
            return pd.read_excel(file_path, sheet_name=sheet_name, dtype=dtype,
                                 keep_default_na=keep_default_na)
    except ImportError as e:
        # 捕获缺失引擎的错误，给出明确提示
        if 'xlrd' in str(e):
            raise ImportError("读取.xls文件需要安装xlrd库。请运行: pip install xlrd>=2.0.1")
        elif 'openpyxl' in str(e):
            raise ImportError("读取.xlsx文件需要安装openpyxl库。请运行: pip install openpyxl")
        else:
            raise e

def load_excel_files(file_paths: List[str], config: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
    """
    读取多个Excel文件，合并为一个DataFrame
    （增强兼容性：自动处理不同Excel格式）
    """
    if config is None:
        config = {}

    class_col_given = config.get('class_column', None)
    name_col_given = config.get('name_column', None)
    zuohao_col_given = config.get('zuohao_column', None)
    subject_cols = config.get('subject_columns', None)
    infer_class = config.get('infer_class_from_filename', True)
    sheet_name = config.get('sheet_name', 0)
    na_vals = config.get('na_values', ['', 'NULL', '缺考', '缺席', 'nan', 'NaN', 'None'])

    all_dfs = []
    for file_path in file_paths:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")

        try:
            # 使用增强的读取函数
            df = _read_excel_file(file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
        except Exception as e:
            raise ValueError(f"无法读取文件 {file_path}: {str(e)}")

        if df.empty:
            print(f"警告：文件 {file_path} 为空，已跳过。")
            continue

        # 清理列名
        df.columns = df.columns.str.strip()
        df.columns = [re.sub(r'[（）]', '()', col) for col in df.columns]

        class_col = class_col_given
        name_col = name_col_given
        zuohao_col = zuohao_col_given

        if class_col is None:
            class_col = _find_column_flexible(df, _get_class_name_variants())
            if class_col is None and infer_class:
                base = os.path.basename(file_path)
                class_name = os.path.splitext(base)[0]
                df['班级'] = class_name
                class_col = '班级'

        if name_col is None:
            name_col = _find_column_flexible(df, _get_name_variants())

        if zuohao_col is None:
            zuohao_col = _find_column_flexible(df, _get_zuohao_variants())

        if class_col is None:
            cols_list = ', '.join(df.columns.tolist())
            raise ValueError(f"文件 {file_path} 中无法识别班级列。找到的列有：{cols_list}\n"
                             f"请确保表头包含'班级'、'班'等字样，或在配置中指定class_column。")

        if name_col is None:
            cols_list = ', '.join(df.columns.tolist())
            raise ValueError(f"文件 {file_path} 中无法识别姓名列。找到的列有：{cols_list}\n"
                             f"请确保表头包含'姓名'、'名字'等字样，或在配置中指定name_column。")

        if subject_cols is None:
            exclude_cols = [class_col, name_col]
            if zuohao_col:
                exclude_cols.append(zuohao_col)
            potential_subjects = [col for col in df.columns if col not in exclude_cols]

            subject_cols = []
            for col in potential_subjects:
                series = df[col].replace(na_vals, pd.NA)
                numeric_series = pd.to_numeric(series, errors='coerce')
                if numeric_series.notna().any():
                    subject_cols.append(col)

            if not subject_cols:
                print(f"警告：文件 {file_path} 中未识别出任何成绩列，将只包含班级、姓名和座号（如果有）。")
        else:
            missing = [col for col in subject_cols if col not in df.columns]
            if missing:
                raise ValueError(f"文件 {file_path} 缺少指定的科目列: {missing}")

        needed_cols = [class_col, name_col]
        if zuohao_col:
            needed_cols.append(zuohao_col)
        needed_cols += subject_cols

        df_sub = df[needed_cols].copy()

        rename_dict = {class_col: '班级', name_col: '姓名'}
        if zuohao_col:
            rename_dict[zuohao_col] = '座号'
        df_sub.rename(columns=rename_dict, inplace=True)

        # 批量转换科目列为数值类型
        for col in subject_cols:
            df_sub[col] = df_sub[col].replace(na_vals, pd.NA)
            df_sub[col] = pd.to_numeric(df_sub[col], errors='coerce')

        all_dfs.append(df_sub)

    if not all_dfs:
        base_columns = ['班级', '姓名']
        return pd.DataFrame(columns=base_columns)

    # 批量合并所有DataFrame
    combined = pd.concat(all_dfs, ignore_index=True)
    
    # 清理数据
    cols_to_check = ['班级', '姓名']
    if '座号' in combined.columns:
        cols_to_check.append('座号')
    cols_to_check += [c for c in combined.columns if c not in ['班级', '姓名', '座号']]
    combined.dropna(how='all', subset=cols_to_check, inplace=True)

    return combined

def load_history_file(file_path: str, config: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
    """
    读取历史总表Excel，返回包含'姓名'、'座号'和'上次排名'的DataFrame
    （增强兼容性：自动处理不同Excel格式）
    """
    if config is None:
        config = {}

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    name_col_given = config.get('name_column', None)
    rank_col_given = config.get('rank_column', None)
    zuohao_col_given = config.get('zuohao_column', None)
    sheet_name = config.get('sheet_name', 0)
    na_vals = config.get('na_values', ['', 'NULL', '缺考', '缺席', 'nan', 'NaN', 'None'])

    try:
        df = _read_excel_file(file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
    except Exception as e:
        raise ValueError(f"无法读取文件 {file_path}: {str(e)}")

    if df.empty:
        return pd.DataFrame(columns=['姓名', '座号', '上次排名'])

    df.columns = df.columns.str.strip()
    df.columns = [re.sub(r'[（）]', '()', col) for col in df.columns]

    name_col = name_col_given
    if name_col is None:
        name_col = _find_column_flexible(df, HISTORY_NAME_VARIANTS)
    if name_col is None:
        cols_list = ', '.join(df.columns.tolist())
        raise ValueError(f"历史文件中无法识别姓名列。找到的列有：{cols_list}\n"
                         f"请确保表头包含'姓名'、'名字'等字样，或在配置中指定name_column。")

    rank_col = rank_col_given
    if rank_col is None:
        rank_col = _find_column_flexible(df, HISTORY_RANK_VARIANTS)
    if rank_col is None:
        cols_list = ', '.join(df.columns.tolist())
        raise ValueError(f"历史文件中无法识别排名列。找到的列有：{cols_list}\n"
                         f"请确保表头包含'排名'、'名次'、'段名'等字样，或在配置中指定rank_column。")

    # 识别座号列
    zuohao_col = config.get('zuohao_column', None)
    if zuohao_col is None:
        zuohao_col = _find_column_flexible(df, _get_zuohao_variants())

    # 构建需要的列列表
    needed_cols = [name_col, rank_col]
    if zuohao_col:
        needed_cols.append(zuohao_col)

    df_sub = df[needed_cols].copy()
    
    # 重命名列
    rename_dict = {name_col: '姓名', rank_col: '上次排名'}
    if zuohao_col:
        rename_dict[zuohao_col] = '座号'
    df_sub.rename(columns=rename_dict, inplace=True)

    df_sub['上次排名'] = pd.to_numeric(df_sub['上次排名'], errors='coerce')

    # 处理座号列
    if '座号' in df_sub.columns:
        df_sub['座号'] = pd.to_numeric(df_sub['座号'], errors='coerce')

    # 删除姓名或上次排名为NaN的行，确保只有有效的数据被保留
    df_sub.dropna(subset=['姓名', '上次排名'], how='any', inplace=True)

    # 确保上次排名是正整数
    df_sub = df_sub[df_sub['上次排名'] > 0]

    return df_sub

def load_total_score_file(file_path: str, config: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
    """
    从年段总表导入数据，自动根据班级列分好班级
    
    Args:
        file_path: Excel文件路径
        config: 配置字典，可选
        
    Returns:
        包含班级、姓名、座号和所有成绩列的DataFrame
    """
    if config is None:
        config = {}

    class_col_given = config.get('class_column', None)
    name_col_given = config.get('name_column', None)
    zuohao_col_given = config.get('zuohao_column', None)
    sheet_name = config.get('sheet_name', 0)
    na_vals = config.get('na_values', ['', 'NULL', '缺考', '缺席', 'nan', 'NaN', 'None'])

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    try:
        df = _read_excel_file(file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
    except Exception as e:
        raise ValueError(f"无法读取文件 {file_path}: {str(e)}")

    if df.empty:
        raise ValueError(f"文件 {file_path} 为空")

    df.columns = df.columns.str.strip()
    df.columns = [re.sub(r'[（）]', '()', col) for col in df.columns]

    class_col = class_col_given
    name_col = name_col_given
    zuohao_col = zuohao_col_given

    if class_col is None:
        class_col = _find_column_flexible(df, _get_class_name_variants())
    
    if class_col is None:
        cols_list = ', '.join(df.columns.tolist())
        raise ValueError(f"总表中无法识别班级列。找到的列有：{cols_list}\n"
                         f"请确保表头包含'班级'、'班'等字样，或在配置中指定class_column。")

    if name_col is None:
        name_col = _find_column_flexible(df, _get_name_variants())
    
    if name_col is None:
        cols_list = ', '.join(df.columns.tolist())
        raise ValueError(f"总表中无法识别姓名列。找到的列有：{cols_list}\n"
                         f"请确保表头包含'姓名'、'名字'等字样，或在配置中指定name_column。")

    if zuohao_col is None:
        zuohao_col = _find_column_flexible(df, _get_zuohao_variants())

    exclude_cols = [class_col, name_col]
    if zuohao_col:
        exclude_cols.append(zuohao_col)
    
    potential_subjects = [col for col in df.columns if col not in exclude_cols]
    subject_cols = []
    for col in potential_subjects:
        series = df[col].replace(na_vals, pd.NA)
        numeric_series = pd.to_numeric(series, errors='coerce')
        if numeric_series.notna().any():
            subject_cols.append(col)

    if not subject_cols:
        print(f"警告：文件 {file_path} 中未识别出任何成绩列，将只包含班级、姓名和座号（如果有）。")

    needed_cols = [class_col, name_col]
    if zuohao_col:
        needed_cols.append(zuohao_col)
    needed_cols += subject_cols

    df_sub = df[needed_cols].copy()

    rename_dict = {class_col: '班级', name_col: '姓名'}
    if zuohao_col:
        rename_dict[zuohao_col] = '座号'
    df_sub.rename(columns=rename_dict, inplace=True)

    for col in subject_cols:
        df_sub[col] = df_sub[col].replace(na_vals, pd.NA)
        df_sub[col] = pd.to_numeric(df_sub[col], errors='coerce')

    df_sub.dropna(how='all', subset=['班级', '姓名'], inplace=True)

    return df_sub

def _get_class_name_variants() -> List[str]:
    return [
        '班级', '班', 'class', 'Class', '班级名称', '班别', '教学班',
        '行政班', '班级名称', '所属班级', '班号', 'classname', 'className'
    ]

def _get_name_variants() -> List[str]:
    return [
        '姓名', '名字', 'name', 'Name', '学生姓名', '学生名字', '考生姓名',
        '姓名全称', 'fullname', 'FullName', 'student name', 'StudentName'
    ]

def _get_zuohao_variants() -> List[str]:
    return [
        '座号', '座位号', '座次', '座位', '序号', '编号', '学号',
        'zuohao', 'seat', 'Seat', 'seat number', 'SeatNumber'
    ]

def _find_column_flexible(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    df_cols_lower = {col: col.strip().lower() for col in df.columns}
    candidates_lower = [c.strip().lower() for c in candidates]

    for col, col_lower in df_cols_lower.items():
        if col_lower in candidates_lower:
            return col

    candidates_sorted = sorted(candidates_lower, key=len, reverse=True)
    for cand in candidates_sorted:
        if len(cand) < 2:
            continue
        for col, col_lower in df_cols_lower.items():
            if cand in col_lower:
                return col

    return None