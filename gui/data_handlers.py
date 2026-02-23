"""数据处理模块

包含更新参数、构建科目输入、刷新教师表格等数据处理相关的函数
"""

import pandas as pd
import unicodedata
import config
from PyQt5.QtWidgets import QTableWidgetItem


def update_params_from_inputs(app):
    """根据用户输入更新参数
    
    Args:
        app: GradeAnalyzerApp实例
    """
    # 从阈值表格中获取科目阈值设置
    full_marks = {}
    passing = {}
    excellent = {}
    
    for row in range(app.threshold_table.rowCount()):
        subject_item = app.threshold_table.item(row, 0)
        if not subject_item:
            continue
        
        subject = subject_item.text()
        
        # 获取满分
        full_mark_item = app.threshold_table.item(row, 1)
        if full_mark_item:
            try:
                full_marks[subject] = float(full_mark_item.text())
            except ValueError:
                full_marks[subject] = config.DEFAULT_THRESHOLDS['满分']
        else:
            full_marks[subject] = config.DEFAULT_THRESHOLDS['满分']
        
        # 获取及格线
        pass_item = app.threshold_table.item(row, 2)
        if pass_item:
            try:
                passing[subject] = float(pass_item.text())
            except ValueError:
                passing[subject] = config.DEFAULT_THRESHOLDS['及格线']
        else:
            passing[subject] = config.DEFAULT_THRESHOLDS['及格线']
        
        # 获取优秀线
        excellent_item = app.threshold_table.item(row, 3)
        if excellent_item:
            try:
                excellent[subject] = float(excellent_item.text())
            except ValueError:
                excellent[subject] = config.DEFAULT_THRESHOLDS['优秀线']
        else:
            excellent[subject] = config.DEFAULT_THRESHOLDS['优秀线']
    
    app.params['full_marks'] = full_marks
    app.params['passing_score'] = passing
    app.params['excellent_score'] = excellent
    app.params['rank_method'] = app.rank_combo.currentText()
    app.params['exclude_missing'] = app.exclude_missing_checkbox.isChecked()
    
    # 总分计算默认开启，权重使用等权
    subjects = app.params['subjects']
    app.params['calc_total'] = True
    app.params['weights'] = {s: 1.0 for s in subjects}


def build_subject_inputs(app):
    """根据导入的数据构建科目输入控件
    
    Args:
        app: GradeAnalyzerApp实例
    """
    if app.raw_data is None:
        return
    
    exclude_cols = [
        config.DEFAULT_COLUMN_MAPPING.get('班级', '班级'),
        config.DEFAULT_COLUMN_MAPPING.get('姓名', '姓名'),
        config.DEFAULT_COLUMN_MAPPING.get('座号', '座号')
    ]
    subjects = [col for col in app.raw_data.columns if col not in exclude_cols]
    
    subjects_clean = []
    for s in subjects:
        s = s.strip()
        s = s.replace('\xa0', '').replace('\u3000', '')
        s = unicodedata.normalize('NFKC', s)
        subjects_clean.append(s)
    
    app.params['subjects'] = subjects
    app.subject_to_clean = dict(zip(subjects, subjects_clean))
    
    # 清空subject_vars
    app.subject_vars.clear()
    
    # 直接更新阈值表格
    app.threshold_table.setRowCount(0)
    app.threshold_table.setColumnCount(4)
    app.threshold_table.setHorizontalHeaderLabels(['科目', '满分', '及格线', '优秀线'])
    
    # 添加科目行到阈值表格
    for row, subject in enumerate(subjects):
        # 计算清洁后的科目名称
        clean_subj = app.subject_to_clean[subject]
        # 获取默认阈值
        if clean_subj in config.SUBJECT_SPECIFIC_THRESHOLDS:
            subj_default = config.SUBJECT_SPECIFIC_THRESHOLDS[clean_subj]
        else:
            subj_default = config.DEFAULT_THRESHOLDS
        
        # 插入行
        app.threshold_table.insertRow(row)
        
        # 设置科目名称
        app.threshold_table.setItem(row, 0, QTableWidgetItem(subject))
        
        # 设置满分、及格线、优秀线
        # 获取满分值
        full_mark = subj_default.get('满分', config.DEFAULT_THRESHOLDS.get('满分', 100))
        app.threshold_table.setItem(row, 1, QTableWidgetItem(str(full_mark)))
        
        # 计算及格线值（满分*0.6）
        pass_mark = full_mark * 0.6
        app.threshold_table.setItem(row, 2, QTableWidgetItem(str(pass_mark)))
        
        # 计算优秀线值（满分*0.85）
        excellent_mark = full_mark * 0.85
        app.threshold_table.setItem(row, 3, QTableWidgetItem(str(excellent_mark)))
    
    # 更新参数
    update_params_from_inputs(app)


def refresh_teacher_table(app):
    """根据当前科目列表刷新教师表格的列结构，并加载已有数据
    
    Args:
        app: GradeAnalyzerApp实例
    """
    if not app.params['subjects']:
        # 没有科目时清空表格
        if app.teacher_table:
            app.teacher_table.setRowCount(0)
            app.teacher_table.setColumnCount(0)
        return
    
    subjects = app.params['subjects']
    classes = sorted(app.raw_data['班级'].unique()) if app.raw_data is not None else []
    columns = ['班级', '班主任'] + subjects
    
    # 设置表格列
    app.teacher_table.setColumnCount(len(columns))
    app.teacher_table.setHorizontalHeaderLabels(columns)
    
    # 清空现有行
    app.teacher_table.setRowCount(0)
    
    if app.teachers_df is not None and not app.teachers_df.empty:
        for _, row in app.teachers_df.iterrows():
            table_row = app.teacher_table.rowCount()
            app.teacher_table.insertRow(table_row)
            
            for col_idx, col in enumerate(columns):
                val = row.get(col, '')
                if col == '班级':
                    val = str(val)
                item = QTableWidgetItem(val)
                app.teacher_table.setItem(table_row, col_idx, item)
    else:
        for cls in classes:
            table_row = app.teacher_table.rowCount()
            app.teacher_table.insertRow(table_row)
            
            app.teacher_table.setItem(table_row, 0, QTableWidgetItem(cls))
            for col_idx in range(1, len(columns)):
                app.teacher_table.setItem(table_row, col_idx, QTableWidgetItem(''))
