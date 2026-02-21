"""教师帮助模块

包含保存教师信息、填充教师信息等教师相关的帮助函数
"""

import pandas as pd
from PyQt5.QtWidgets import QMessageBox


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
    
    for _, row in app.teachers_df.iterrows():
        cls = str(row.get('班级', '')).strip()
        if not cls:
            continue
        
        head = str(row.get('班主任', '')).strip()
        if head:
            head_map[cls] = head
        
        for subj in app.params['subjects']:
            teacher = str(row.get(subj, '')).strip()
            if teacher:
                teacher_map[(cls, subj)] = teacher
    
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
    
    return head_map
