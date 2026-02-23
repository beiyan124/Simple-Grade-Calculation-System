"""表格处理模块

包含更新各种表格的函数
"""

import pandas as pd
from PyQt5.QtWidgets import QTableWidgetItem
from PyQt5.QtCore import Qt


def update_student_table(app):
    """更新学生排名表格
    
    Args:
        app: GradeAnalyzerApp实例
    """
    if app.student_rank is None:
        return
    
    df = app.student_rank
    table = app.student_table
    
    # 优化：预先获取数据，避免在循环中重复访问
    rows, cols = df.shape
    columns = list(df.columns)
    data = df.values.tolist()
    
    # 设置表格行列
    table.setRowCount(rows)
    table.setColumnCount(cols)
    
    # 设置表头
    table.setHorizontalHeaderLabels(columns)
    
    # 优化：批量填充数据，减少表格操作次数
    for idx, row_data in enumerate(data):
        for col_idx, value in enumerate(row_data):
            col_name = columns[col_idx]
            if pd.notna(value):
                # 检测比差列并显示为百分数形式
                if '比差' in col_name:
                    item = QTableWidgetItem(f"{value:.2f}%")
                else:
                    item = QTableWidgetItem(str(value))
            else:
                item = QTableWidgetItem("")
            item.setTextAlignment(Qt.AlignCenter)
            table.setItem(idx, col_idx, item)


def update_class_table(app):
    """更新班级汇总表格
    
    Args:
        app: GradeAnalyzerApp实例
    """
    if app.class_summary is None:
        return
    
    df = app.class_summary
    table = app.class_table
    
    # 设置表格行列
    table.setRowCount(len(df))
    table.setColumnCount(len(df.columns))
    
    # 设置表头
    table.setHorizontalHeaderLabels(list(df.columns))
    
    # 填充数据
    for idx, (_, row) in enumerate(df.iterrows()):
        for col_idx, col in enumerate(df.columns):
            value = row[col]
            item = QTableWidgetItem(str(value) if pd.notna(value) else "")
            item.setTextAlignment(Qt.AlignCenter)
            table.setItem(idx, col_idx, item)


def update_subject_table(app, subject):
    """更新单科班级分析表格
    
    Args:
        app: GradeAnalyzerApp实例
        subject: 要显示的科目
    """
    if subject not in app.subject_details:
        return
    
    df = app.subject_details[subject]
    table = app.subject_table
    
    # 设置表格行列
    table.setRowCount(len(df))
    table.setColumnCount(len(df.columns))
    
    # 设置表头
    table.setHorizontalHeaderLabels(list(df.columns))
    
    # 填充数据
    for idx, (_, row) in enumerate(df.iterrows()):
        for col_idx, col in enumerate(df.columns):
            value = row[col]
            item = QTableWidgetItem(str(value) if pd.notna(value) else "")
            item.setTextAlignment(Qt.AlignCenter)
            table.setItem(idx, col_idx, item)


def update_subject_ranking_table(app, subject):
    """更新单科班级排名表格
    
    Args:
        app: GradeAnalyzerApp实例
        subject: 要显示的科目
    """
    if subject not in app.subject_rankings:
        return
    
    df = app.subject_rankings[subject]
    table = app.ranking_table
    
    # 设置表格行列
    table.setRowCount(len(df))
    table.setColumnCount(len(df.columns))
    
    # 设置表头
    table.setHorizontalHeaderLabels(list(df.columns))
    
    # 填充数据
    for idx, (_, row) in enumerate(df.iterrows()):
        for col_idx, col in enumerate(df.columns):
            value = row[col]
            item = QTableWidgetItem(str(value) if pd.notna(value) else "")
            item.setTextAlignment(Qt.AlignCenter)
            table.setItem(idx, col_idx, item)
