"""教师帮助模块

包含保存教师信息、填充教师信息等教师相关的帮助函数
"""

import pandas as pd
import re
from PyQt5.QtWidgets import QMessageBox


def extract_class_number(class_name):
    """
    从班级名称中提取数字
    
    Args:
        class_name: 班级名称字符串
    
    Returns:
        int: 提取的班级数字，如果没有找到则返回0
    """
    if not class_name:
        return 0
    
    # 尝试从各种格式中提取数字
    class_str = str(class_name).strip()
    
    # 匹配数字，包括可能的括号、尖括号等
    match = re.search(r'(\d+)', class_str)
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return 0
    
    return 0


def get_class_name_mapping(known_classes):
    """
    创建班级名称映射，将不同格式的班级名称映射到标准格式
    
    Args:
        known_classes: 已知的班级名称列表
    
    Returns:
        dict: 映射字典，键为原始班级名称，值为标准班级名称
    """
    mapping = {}
    
    # 为每个已知班级创建映射
    for cls in known_classes:
        mapping[cls] = cls
    
    return mapping


def normalize_class_name(class_name, known_classes):
    """
    标准化班级名称，将不同格式的班级名称转换为已有的标准格式
    
    Args:
        class_name: 原始班级名称
        known_classes: 已知的班级名称列表
    
    Returns:
        str: 标准化后的班级名称，如果无法匹配则返回原始名称
    """
    if not class_name:
        return ''
    
    # 提取原始班级的数字
    class_num = extract_class_number(class_name)
    if class_num == 0:
        return str(class_name).strip()
    
    # 尝试匹配已知班级
    for cls in known_classes:
        if extract_class_number(cls) == class_num:
            return cls
    
    # 如果没有匹配到，返回原始名称
    return str(class_name).strip()


def save_teachers(app):
    """将当前表格数据保存到 self.teachers_df，统一班级格式
    
    Args:
        app: GradeAnalyzerApp实例
    """
    data = []
    columns = [app.teacher_table.horizontalHeaderItem(col).text() for col in range(app.teacher_table.columnCount())]
    
    for row in range(app.teacher_table.rowCount()):
        row_data = []
        for col in range(app.teacher_table.columnCount()):
            item = app.teacher_table.item(row, col)
            row_data.append(item.text() if item else '')
        data.append(row_data)
    
    if not columns or not data:
        return
    
    app.teachers_df = pd.DataFrame(data, columns=columns)
    if '班级' in app.teachers_df.columns:
        # 获取已知的班级名称（从原始数据或班级汇总中）
        known_classes = []
        if hasattr(app, 'raw_data') and app.raw_data is not None:
            known_classes = app.raw_data['班级'].astype(str).unique().tolist()
        elif hasattr(app, 'class_summary') and app.class_summary is not None:
            known_classes = app.class_summary['班级'].astype(str).unique().tolist()
        
        # 标准化班级名称
        if known_classes:
            app.teachers_df['班级'] = app.teachers_df['班级'].apply(
                lambda x: normalize_class_name(x, known_classes)
            )
        else:
            app.teachers_df['班级'] = app.teachers_df['班级'].astype(str).str.strip()
    
    QMessageBox.information(app, "保存成功", "教师配置已保存到内存，重新计算后生效。")


def fill_teacher_info(app):
    """将教师信息填充到 subject_details 的'任课教师'列，并返回班主任映射
    
    Args:
        app: GradeAnalyzerApp实例
    
    Returns:
        班主任映射字典 {班级: 班主任}
    """
    if app.teachers_df is None or app.teachers_df.empty:
        return {}
    
    teacher_map = {}
    head_map = {}
    
    # 获取已知的班级名称（从原始数据或班级汇总中）
    known_classes = []
    if hasattr(app, 'raw_data') and app.raw_data is not None:
        known_classes = app.raw_data['班级'].astype(str).unique().tolist()
    elif hasattr(app, 'class_summary') and app.class_summary is not None:
        known_classes = app.class_summary['班级'].astype(str).unique().tolist()
    
    # 处理教师数据，标准化班级名称
    for _, row in app.teachers_df.iterrows():
        cls = str(row.get('班级', '')).strip()
        if not cls:
            continue
        
        # 标准化班级名称
        normalized_cls = normalize_class_name(cls, known_classes)
        
        head = str(row.get('班主任', '')).strip()
        if head:
            head_map[normalized_cls] = head
        
        for subj in app.params['subjects']:
            teacher = str(row.get(subj, '')).strip()
            if teacher:
                teacher_map[(normalized_cls, subj)] = teacher
    
    for subj, df in app.subject_details.items():
        if '任课教师' not in df.columns:
            df['任课教师'] = ''
        
        df['班级_clean'] = df['班级'].astype(str).str.strip()
        for idx, row in df.iterrows():
            cls_clean = row['班级_clean']
            key = (cls_clean, subj)
            if key in teacher_map:
                df.at[idx, '任课教师'] = teacher_map[key]
        df.drop(columns=['班级_clean'], inplace=True)
    
    # 为总分的任课教师列填充班主任信息
    if '总分' in app.subject_details:
        total_df = app.subject_details['总分']
        if '任课教师' not in total_df.columns:
            total_df['任课教师'] = ''
        
        total_df['班级_clean'] = total_df['班级'].astype(str).str.strip()
        for idx, row in total_df.iterrows():
            cls_clean = row['班级_clean']
            if cls_clean in head_map:
                total_df.at[idx, '任课教师'] = head_map[cls_clean]
        total_df.drop(columns=['班级_clean'], inplace=True)
    
    return head_map
