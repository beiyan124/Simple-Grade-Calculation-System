"""
主窗口模块

包含GradeAnalyzerApp主应用类，负责整个应用的初始化和管理
"""

import sys
import os
import pandas as pd
from typing import Dict, List, Optional, Any

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QCheckBox, QLineEdit, QComboBox, QTextEdit,
    QTableWidget, QTableWidgetItem, QTabWidget, QFrame, QSplitter,
    QFileDialog, QMessageBox, QScrollArea, QHeaderView
)
from PyQt5.QtCore import Qt, QSize, QEvent
from PyQt5.QtGui import QFont, QColor, QPalette, QBrush, QIcon
from PyQt5.QtWidgets import QStyleFactory

import data_loader
import calculator
import exporter
import config
from .ui_utils import create_table, load_help_text, apply_frosted_glass_effect
from .table_handlers import update_student_table, update_class_table, update_subject_table, update_subject_ranking_table
from .data_handlers import update_params_from_inputs, build_subject_inputs, refresh_teacher_table
from .teacher_helpers import save_teachers, fill_teacher_info


class GradeAnalyzerApp(QMainWindow):
    """年段一分三率计算系统主应用类（PyQt5版本）"""
    
    def __init__(self):
        """初始化应用"""
        super().__init__()
        self.setWindowTitle("年段一分三率计算系统")
        
        # 设置窗口大小并支持分辨率适配
        self.resize(1280, 720)
        self.setMinimumSize(1000, 600)
        
        # 内部数据存储
        self.raw_data: Optional[pd.DataFrame] = None
        self.history_df: Optional[pd.DataFrame] = None      # 历史排名数据
        self.student_rank: Optional[pd.DataFrame] = None
        self.class_summary: Optional[pd.DataFrame] = None
        self.subject_details: Dict[str, pd.DataFrame] = {}
        self.subject_rankings: Dict[str, pd.DataFrame] = {}
        self.teachers_df: Optional[pd.DataFrame] = None
        self.params: Dict[str, Any] = self._get_default_params()
        
        # 界面控件变量
        self.subject_vars: Dict[str, Dict[str, str]] = {}
        self.total_var = True
        self.rank_method_var = config.DEFAULT_RANK_METHOD
        self.exclude_missing_var = config.DEFAULT_EXCLUDE_MISSING
        # 进退步相关变量
        self.progress_var = False
        self.history_file_label = "未导入历史文件"
        
        # 单科班级分析相关变量
        self.selected_subject_var = ""
        self.subject_combobox: Optional[QComboBox] = None
        
        # 单科班级排名相关变量
        self.ranking_subject_var = ""
        self.ranking_combobox: Optional[QComboBox] = None
        self.ranking_table: Optional[QTableWidget] = None
        
        # 教师配置页面控件
        self.threshold_table: Optional[QTableWidget] = None
        self.teacher_table: Optional[QTableWidget] = None
        
        # 构建界面
        self._setup_ui()
        apply_frosted_glass_effect(self)
        
    def _get_default_params(self) -> Dict[str, Any]:
        """获取默认参数设置
        
        Returns:
            默认参数字典
        """
        return {
            'subjects': [],
            'full_marks': {},
            'passing_score': {},
            'excellent_score': {},
            'calc_total': True,
            'weights': {},
            'rank_method': config.DEFAULT_RANK_METHOD,
            'exclude_missing': config.DEFAULT_EXCLUDE_MISSING
        }
    
    def _setup_ui(self):
        """设置用户界面"""
        # 创建中央窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        
        # 创建菜单栏
        self._create_menu()
        
        # 创建工具栏
        self._create_toolbar()
        
        # 创建主面板
        self._create_main_panel(main_layout)
        
        # 创建状态栏
        self.statusBar().showMessage("就绪")
    
    def _create_menu(self):
        """创建菜单栏"""
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu("文件")
        import_action = file_menu.addAction("导入文件")
        import_action.triggered.connect(self.load_files)
        import_action.setShortcut("Ctrl+O")
        
        file_menu.addSeparator()
        exit_action = file_menu.addAction("退出")
        exit_action.triggered.connect(self.close)
        
        # 帮助菜单
        help_menu = menubar.addMenu("帮助")
        about_action = help_menu.addAction("关于")
        about_action.triggered.connect(self._show_about)
    
    def _create_toolbar(self):
        """创建工具栏"""
        toolbar = self.addToolBar("工具栏")
        toolbar.setMovable(False)
        
        import_btn = QPushButton("导入文件")
        import_btn.clicked.connect(self.load_files)
        toolbar.addWidget(import_btn)
        
        calc_btn = QPushButton("计算")
        calc_btn.clicked.connect(self.apply_params_and_calculate)
        toolbar.addWidget(calc_btn)
        
        export_btn = QPushButton("导出结果")
        export_btn.clicked.connect(self.export_results)
        toolbar.addWidget(export_btn)
    
    def _create_main_panel(self, parent_layout):
        """创建主面板"""
        # 创建分割器
        splitter = QSplitter(Qt.Horizontal)
        parent_layout.addWidget(splitter)
        
        # 左侧参数设置面板
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_panel.setMinimumWidth(350)
        left_panel.setMaximumWidth(500)
        
        # 右侧结果显示面板
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        # 创建左侧面板内容
        self._create_left_panel(left_layout)
        
        # 创建右侧面板内容
        self._create_right_panel(right_layout)
        
        # 将面板添加到分割器
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(1, 1)
    
    def _create_left_panel(self, parent_layout):
        """构建左侧参数设置面板
        
        Args:
            parent_layout: 父布局
        """
        # 标题
        title_label = QLabel("参数设置")
        title_label.setFont(QFont('Arial', 12, QFont.Bold))
        parent_layout.addWidget(title_label)
        
        # 1. 排名规则
        rank_group = QFrame()
        rank_group.setFrameShape(QFrame.StyledPanel)
        rank_group.setFrameShadow(QFrame.Raised)
        rank_layout = QVBoxLayout(rank_group)
        
        rank_title = QLabel("排名规则")
        rank_title.setFont(QFont('Arial', 10, QFont.Bold))
        rank_layout.addWidget(rank_title)
        
        rank_label = QLabel("同分处理：")
        rank_layout.addWidget(rank_label)
        
        self.rank_combo = QComboBox()
        self.rank_combo.addItems(['min', 'dense', 'average'])
        self.rank_combo.setCurrentText(config.DEFAULT_RANK_METHOD)
        rank_layout.addWidget(self.rank_combo)
        
        rank_note = QLabel("min=占用名次, dense=不占用, average=平均名次")
        rank_note.setFont(QFont('Arial', 8))
        rank_layout.addWidget(rank_note)
        
        parent_layout.addWidget(rank_group)
        
        # 2. 缺考处理
        missing_group = QFrame()
        missing_group.setFrameShape(QFrame.StyledPanel)
        missing_group.setFrameShadow(QFrame.Raised)
        missing_layout = QVBoxLayout(missing_group)
        
        missing_title = QLabel("缺考处理")
        missing_title.setFont(QFont('Arial', 10, QFont.Bold))
        missing_layout.addWidget(missing_title)
        
        self.exclude_missing_checkbox = QCheckBox("排除缺考（空值/非数值）计算平均分和率")
        self.exclude_missing_checkbox.setChecked(config.DEFAULT_EXCLUDE_MISSING)
        missing_layout.addWidget(self.exclude_missing_checkbox)
        
        parent_layout.addWidget(missing_group)
        
        # 3. 进退步计算区域
        progress_group = QFrame()
        progress_group.setFrameShape(QFrame.StyledPanel)
        progress_group.setFrameShadow(QFrame.Raised)
        progress_layout = QVBoxLayout(progress_group)
        
        progress_title = QLabel("进退步计算")
        progress_title.setFont(QFont('Arial', 10, QFont.Bold))
        progress_layout.addWidget(progress_title)
        
        self.progress_checkbox = QCheckBox("启用进退步计算（需导入历史总表）")
        progress_layout.addWidget(self.progress_checkbox)
        
        btn_frame = QWidget()
        btn_layout = QHBoxLayout(btn_frame)
        
        load_history_btn = QPushButton("导入历史总表")
        load_history_btn.clicked.connect(self.load_history_file)
        btn_layout.addWidget(load_history_btn)
        
        self.history_label = QLabel("未导入历史文件")
        self.history_label.setFont(QFont('Arial', 8))
        btn_layout.addWidget(self.history_label)
        
        progress_layout.addWidget(btn_frame)
        parent_layout.addWidget(progress_group)
        
        # 4. 考试设置
        exam_group = QFrame()
        exam_group.setFrameShape(QFrame.StyledPanel)
        exam_group.setFrameShadow(QFrame.Raised)
        exam_layout = QVBoxLayout(exam_group)
        
        exam_title = QLabel("考试设置")
        exam_title.setFont(QFont('Arial', 10, QFont.Bold))
        exam_layout.addWidget(exam_title)
        
        # 考试名称
        exam_name_layout = QHBoxLayout()
        exam_name_label = QLabel("考试名称：")
        exam_name_label.setFixedWidth(80)
        self.exam_name_entry = QLineEdit()
        exam_name_layout.addWidget(exam_name_label)
        exam_name_layout.addWidget(self.exam_name_entry)
        exam_layout.addLayout(exam_name_layout)
        
        # 考试时间日期
        exam_date_layout = QHBoxLayout()
        exam_date_label = QLabel("考试时间：")
        exam_date_label.setFixedWidth(80)
        self.exam_date_entry = QLineEdit()
        self.exam_date_entry.setPlaceholderText("YYYY-MM-DD")
        exam_date_layout.addWidget(exam_date_label)
        exam_date_layout.addWidget(self.exam_date_entry)
        exam_layout.addLayout(exam_date_layout)
        
        # 出卷人
        exam_author_layout = QHBoxLayout()
        exam_author_label = QLabel("出卷人：")
        exam_author_label.setFixedWidth(80)
        self.exam_author_entry = QLineEdit()
        exam_author_layout.addWidget(exam_author_label)
        exam_author_layout.addWidget(self.exam_author_entry)
        exam_layout.addLayout(exam_author_layout)
        
        parent_layout.addWidget(exam_group)
        
        # 5. 操作按钮
        btn_frame2 = QWidget()
        btn_layout2 = QHBoxLayout(btn_frame2)
        
        reset_btn = QPushButton("重置为默认")
        reset_btn.clicked.connect(self._reset_params)
        btn_layout2.addWidget(reset_btn)
        
        parent_layout.addWidget(btn_frame2)
        
        # 添加伸缩空间
        parent_layout.addStretch()
    
    def _create_right_panel(self, parent_layout):
        """构建右侧结果显示面板
        
        Args:
            parent_layout: 父布局
        """
        # 创建标签页控件
        self.notebook = QTabWidget()
        
        # 学生排名表页
        self.student_frame = QWidget()
        student_layout = QVBoxLayout(self.student_frame)
        self.student_table = create_table()
        student_layout.addWidget(self.student_table)
        self.notebook.addTab(self.student_frame, "学生排名表")
        
        # 班级汇总表页
        self.class_frame = QWidget()
        class_layout = QVBoxLayout(self.class_frame)
        self.class_table = create_table()
        class_layout.addWidget(self.class_table)
        self.notebook.addTab(self.class_frame, "班级汇总表")
        
        # 单科班级分析页
        self.subject_frame = QWidget()
        subject_layout = QVBoxLayout(self.subject_frame)
        
        top_frame = QWidget()
        top_layout = QHBoxLayout(top_frame)
        
        subject_label = QLabel("选择科目：")
        top_layout.addWidget(subject_label)
        
        self.subject_combobox = QComboBox()
        self.subject_combobox.currentTextChanged.connect(self._on_subject_selected)
        top_layout.addWidget(self.subject_combobox)
        
        subject_layout.addWidget(top_frame)
        
        self.subject_table = create_table()
        subject_layout.addWidget(self.subject_table)
        
        self.notebook.addTab(self.subject_frame, "单科班级分析")
        
        # 单科班级排名页
        self.ranking_frame = QWidget()
        ranking_layout = QVBoxLayout(self.ranking_frame)
        
        top_frame2 = QWidget()
        top_layout2 = QHBoxLayout(top_frame2)
        
        ranking_label = QLabel("选择科目：")
        top_layout2.addWidget(ranking_label)
        
        self.ranking_combobox = QComboBox()
        self.ranking_combobox.currentTextChanged.connect(self._on_ranking_subject_selected)
        top_layout2.addWidget(self.ranking_combobox)
        
        ranking_layout.addWidget(top_frame2)
        
        self.ranking_table = create_table()
        ranking_layout.addWidget(self.ranking_table)
        
        self.notebook.addTab(self.ranking_frame, "单科班级排名")
        
        # 使用说明页
        self.help_frame = QWidget()
        help_layout = QVBoxLayout(self.help_frame)
        
        self.help_text = QTextEdit()
        self.help_text.setReadOnly(True)
        self.help_text.setFont(QFont('微软雅黑', 10))
        load_help_text(self.help_text)
        help_layout.addWidget(self.help_text)
        
        self.notebook.addTab(self.help_frame, "使用说明")
        
        # 教师配置页
        self.teacher_config_frame = QWidget()
        teacher_layout = QVBoxLayout(self.teacher_config_frame)
        
        # 上方：科目阈值预览
        threshold_frame = QFrame()
        threshold_frame.setFrameShape(QFrame.StyledPanel)
        threshold_frame.setFrameShadow(QFrame.Raised)
        threshold_layout = QVBoxLayout(threshold_frame)
        
        threshold_title = QLabel("科目阈值设置（可编辑后点击应用）")
        threshold_title.setFont(QFont('Arial', 10, QFont.Bold))
        threshold_layout.addWidget(threshold_title)
        
        self.threshold_table = create_table()
        threshold_layout.addWidget(self.threshold_table)
        
        btn_threshold_frame = QWidget()
        btn_threshold_layout = QHBoxLayout(btn_threshold_frame)
        
        sync_left_btn = QPushButton("从左侧同步阈值")
        sync_left_btn.clicked.connect(self._sync_thresholds_from_left)
        btn_threshold_layout.addWidget(sync_left_btn)
        
        sync_right_btn = QPushButton("应用阈值到左侧")
        sync_right_btn.clicked.connect(self._sync_thresholds_to_left)
        btn_threshold_layout.addWidget(sync_right_btn)
        
        threshold_layout.addWidget(btn_threshold_frame)
        teacher_layout.addWidget(threshold_frame)
        
        # 下方：班级教师配置表格
        teacher_frame = QFrame()
        teacher_frame.setFrameShape(QFrame.StyledPanel)
        teacher_frame.setFrameShadow(QFrame.Raised)
        teacher_layout2 = QVBoxLayout(teacher_frame)
        
        teacher_title = QLabel("班级教师配置（双击单元格可编辑）")
        teacher_title.setFont(QFont('Arial', 10, QFont.Bold))
        teacher_layout2.addWidget(teacher_title)
        
        self.teacher_table = create_table()
        teacher_layout2.addWidget(self.teacher_table)
        
        btn_teacher_frame = QWidget()
        btn_teacher_layout = QHBoxLayout(btn_teacher_frame)
        
        import_teacher_btn = QPushButton("导入教师配置")
        import_teacher_btn.clicked.connect(self._import_teachers)
        btn_teacher_layout.addWidget(import_teacher_btn)
        
        save_teacher_btn = QPushButton("保存教师配置")
        save_teacher_btn.clicked.connect(self._save_teachers)
        btn_teacher_layout.addWidget(save_teacher_btn)
        
        refresh_teacher_btn = QPushButton("刷新表格")
        refresh_teacher_btn.clicked.connect(self._refresh_teacher_table)
        btn_teacher_layout.addWidget(refresh_teacher_btn)
        
        teacher_layout2.addWidget(btn_teacher_frame)
        teacher_layout.addWidget(teacher_frame)
        
        self.notebook.addTab(self.teacher_config_frame, "科目与教师设置")
        
        parent_layout.addWidget(self.notebook)
    

    
    def _reset_params(self):
        """重置参数为默认值"""
        self.rank_combo.setCurrentText(config.DEFAULT_RANK_METHOD)
        self.exclude_missing_checkbox.setChecked(config.DEFAULT_EXCLUDE_MISSING)
        self.progress_checkbox.setChecked(False)
        self.history_df = None
        self.history_label.setText("未导入历史文件")
        
        # 重置考试设置
        if hasattr(self, 'exam_name_entry'):
            self.exam_name_entry.clear()
        if hasattr(self, 'exam_date_entry'):
            self.exam_date_entry.clear()
        if hasattr(self, 'exam_author_entry'):
            self.exam_author_entry.clear()
        
        self.params = self._get_default_params()
        self.subject_details.clear()
        self.subject_rankings.clear()
        self.teachers_df = None
        
        # 清空下拉框
        if self.subject_combobox:
            self.subject_combobox.clear()
            self.selected_subject_var = ""
        if self.ranking_combobox:
            self.ranking_combobox.clear()
            self.ranking_subject_var = ""
        
        # 清空表格
        for table in [self.student_table, self.class_table, self.subject_table, 
                     self.ranking_table, self.threshold_table, self.teacher_table]:
            if table:
                table.setRowCount(0)
                table.setColumnCount(0)
        
        # 重新构建科目输入
        if self.raw_data is not None:
            self._refresh_teacher_table()
        
        self.statusBar().showMessage("参数已重置")
    
    def _build_subject_inputs(self):
        """根据导入的数据构建科目输入控件"""
        build_subject_inputs(self)
    
    def _update_params_from_inputs(self):
        """根据用户输入更新参数"""
        update_params_from_inputs(self)
    
    def load_files(self):
        """导入Excel文件"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        
        if not file_paths:
            return
        
        self.statusBar().showMessage("正在加载文件...")
        
        try:
            self.raw_data = data_loader.load_excel_files(file_paths, {})
            self._build_subject_inputs()
            self._refresh_teacher_table()
            self.statusBar().showMessage(f"已加载 {len(self.raw_data)} 条学生记录")
            QMessageBox.information(self, "导入成功", f"成功加载 {len(file_paths)} 个文件，共 {len(self.raw_data)} 条记录。")
        except Exception as e:
            self.statusBar().showMessage("加载失败")
            QMessageBox.critical(self, "导入错误", f"加载文件时出错：\n{str(e)}")
    
    def load_history_file(self):
        """导入历史总表文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择历史总表Excel文件", "", "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        
        if not file_path:
            return
        
        self.statusBar().showMessage("正在加载历史文件...")
        
        try:
            self.history_df = data_loader.load_history_file(file_path, {})
            if self.history_df.empty:
                QMessageBox.warning(self, "警告", "历史文件无有效数据。")
                self.history_label.setText("导入失败或无数据")
            else:
                self.history_label.setText(f"已导入 {len(self.history_df)} 条记录")
                self.statusBar().showMessage("历史文件加载成功")
        except Exception as e:
            self.history_df = None
            self.history_label.setText("导入失败")
            self.statusBar().showMessage("历史文件加载失败")
            QMessageBox.critical(self, "导入错误", f"加载历史文件时出错：\n{str(e)}")
    
    def apply_params_and_calculate(self):
        """应用参数并执行计算"""
        if self.raw_data is None:
            QMessageBox.warning(self, "无数据", "请先导入成绩文件。")
            return
        
        if self.teacher_table and self.teacher_table.rowCount() > 0:
            self._save_teachers()
        
        self._update_params_from_inputs()
        self.statusBar().showMessage("正在计算...")
        
        try:
            result = calculator.calculate(
                self.raw_data,
                self.params,
                history_df=self.history_df if self.progress_checkbox.isChecked() else None,
                calc_progress=self.progress_checkbox.isChecked()
            )
            
            if len(result) == 3:
                self.student_rank, self.class_summary, self.subject_details = result
            else:
                self.student_rank, self.class_summary = result
                self.subject_details = {}
            
            head_map = self._fill_teacher_info()
            
            if head_map:
                self.class_summary['班级_clean'] = self.class_summary['班级'].astype(str).str.strip()
                cols = list(self.class_summary.columns)
                class_pos = cols.index('班级_clean') if '班级_clean' in cols else 0
                head_col = self.class_summary['班级_clean'].map(head_map).fillna('')
                self.class_summary.insert(class_pos + 1, '班主任', head_col)
                self.class_summary.drop(columns=['班级_clean'], inplace=True)
            else:
                self.class_summary['班主任'] = ''
            
            self._build_subject_rankings()
            self._update_student_table()
            self._update_class_table()
            self._update_subject_combobox()
            self._update_ranking_combobox()
            
            self.statusBar().showMessage("计算完成")
        except Exception as e:
            self.statusBar().showMessage("计算失败")
            QMessageBox.critical(self, "计算错误", f"计算过程中出错：\n{str(e)}")
    
    def export_results(self):
        """导出计算结果到Excel文件"""
        if self.student_rank is None or self.class_summary is None:
            QMessageBox.warning(self, "无结果", "请先计算后再导出。")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存Excel文件", "", "Excel文件 (*.xlsx);;所有文件 (*.*)"
        )
        
        if not file_path:
            return
        
        self.statusBar().showMessage("正在导出...")
        
        try:
            # 收集考试信息
            exam_info = {
                '考试名称': self.exam_name_entry.text() if hasattr(self, 'exam_name_entry') else '',
                '考试时间日期': self.exam_date_entry.text() if hasattr(self, 'exam_date_entry') else '',
                '出卷人': self.exam_author_entry.text() if hasattr(self, 'exam_author_entry') else ''
            }
            
            # 直接使用用户选择的文件路径，允许覆盖已存在的文件
            
            exporter.export_to_excel(self.student_rank, self.class_summary,
                                      self.subject_details, self.subject_rankings, file_path, exam_info=exam_info)
            self.statusBar().showMessage(f"导出成功：{os.path.basename(file_path)}")
            QMessageBox.information(self, "导出完成", f"结果已保存到：\n{file_path}")
        except PermissionError as e:
            self.statusBar().showMessage("导出失败")
            QMessageBox.critical(self, "权限错误", f"无法写入文件，可能是因为：\n1. 没有写入目标目录的权限\n2. 文件已被其他程序打开\n3. 文件是只读的\n\n请尝试保存到其他目录。\n\n错误信息：\n{str(e)}")
        except Exception as e:
            self.statusBar().showMessage("导出失败")
            QMessageBox.critical(self, "导出错误", f"导出文件时出错：\n{str(e)}")
    
    def _update_student_table(self):
        """更新学生排名表格"""
        update_student_table(self)
    
    def _update_class_table(self):
        """更新班级汇总表格"""
        update_class_table(self)
    
    def _update_subject_combobox(self):
        """更新科目选择下拉框"""
        subjects = list(self.subject_details.keys())
        if subjects:
            self.subject_combobox.clear()
            self.subject_combobox.addItems(subjects)
            self.selected_subject_var = subjects[0]
            self._update_subject_table(subjects[0])
        else:
            self.subject_combobox.clear()
            self.selected_subject_var = ""
            self.subject_table.setRowCount(0)
            self.subject_table.setColumnCount(0)
    
    def _update_subject_table(self, subject):
        """更新单科班级分析表格
        
        Args:
            subject: 要显示的科目
        """
        update_subject_table(self, subject)
    
    def _on_subject_selected(self, subject):
        """科目选择事件处理
        
        Args:
            subject: 选择的科目
        """
        if subject and subject in self.subject_details:
            self._update_subject_table(subject)
        else:
            self.subject_table.setRowCount(0)
            self.subject_table.setColumnCount(0)
    
    def _build_subject_rankings(self):
        """构建单科班级排名数据"""
        self.subject_rankings.clear()
        for subject, df in self.subject_details.items():
            grade_row = df[df['班级'] == '年段'].iloc[0] if '年段' in df['班级'].values else None
            if grade_row is None:
                continue
            class_rows = df[df['班级'] != '年段'].copy()
            if class_rows.empty:
                continue
            data = class_rows[['班级', '任课教师', '平均分', '及格率', '优生率']].copy()
            grade_avg = grade_row['平均分']
            grade_pass = grade_row['及格率']
            grade_excel = grade_row['优生率']
            data['平均分比差'] = (data['平均分'] - grade_avg) / grade_avg * 10 if grade_avg > 0 else 0
            data['及格率比差'] = (data['及格率'] - grade_pass) * 10
            data['优生率比差'] = (data['优生率'] - grade_excel) * 10
            data['平均分名次'] = data['平均分'].rank(method='min', ascending=False).astype('Int64')
            data['及格率名次'] = data['及格率'].rank(method='min', ascending=False).astype('Int64')
            data['优生率名次'] = data['优生率'].rank(method='min', ascending=False).astype('Int64')
            data.sort_values('班级', inplace=True)
            cols_order = [
                '班级', '任课教师',
                '平均分', '平均分比差', '平均分名次',
                '及格率', '及格率比差', '及格率名次',
                '优生率', '优生率比差', '优生率名次'
            ]
            ranking_df = data[cols_order].copy()
            grade_row_data = {
                '班级': '年段',
                '任课教师': '',
                '平均分': grade_avg,
                '平均分比差': 0.0,
                '平均分名次': pd.NA,
                '及格率': grade_pass,
                '及格率比差': 0.0,
                '及格率名次': pd.NA,
                '优生率': grade_excel,
                '优生率比差': 0.0,
                '优生率名次': pd.NA
            }
            ranking_df = pd.concat([ranking_df, pd.DataFrame([grade_row_data])], ignore_index=True)
            self.subject_rankings[subject] = ranking_df
    
    def _update_ranking_combobox(self):
        """更新排名科目选择下拉框"""
        subjects = list(self.subject_rankings.keys())
        if subjects:
            self.ranking_combobox.clear()
            self.ranking_combobox.addItems(subjects)
            self.ranking_subject_var = subjects[0]
            self._update_subject_ranking_table(subjects[0])
        else:
            self.ranking_combobox.clear()
            self.ranking_subject_var = ""
            if self.ranking_table:
                self.ranking_table.setRowCount(0)
                self.ranking_table.setColumnCount(0)
    
    def _on_ranking_subject_selected(self, subject):
        """排名科目选择事件处理
        
        Args:
            subject: 选择的科目
        """
        if subject and subject in self.subject_rankings:
            self._update_subject_ranking_table(subject)
        else:
            if self.ranking_table:
                self.ranking_table.setRowCount(0)
                self.ranking_table.setColumnCount(0)
    
    def _update_subject_ranking_table(self, subject):
        """更新单科班级排名表格
        
        Args:
            subject: 要显示的科目
        """
        update_subject_ranking_table(self, subject)
    
    def _sync_thresholds_from_left(self):
        """将左侧面板的阈值同步到阈值表格"""
        # 左侧面板不再有科目阈值设置，此方法不再需要
        pass

    def _sync_thresholds_to_left(self):
        """将阈值表格的值更新到左侧面板"""
        # 左侧面板不再有科目阈值设置，此方法不再需要
        # 直接更新参数即可
        self._update_params_from_inputs()
        QMessageBox.information(self, "成功", "科目阈值已更新到左侧面板，请重新计算生效。")
    
    def _refresh_teacher_table(self):
        """根据当前科目列表刷新教师表格的列结构，并加载已有数据"""
        refresh_teacher_table(self)
    
    def _import_teachers(self):
        """从Excel文件导入教师配置"""
        if not self.params['subjects']:
            QMessageBox.warning(self, "无科目", "请先导入成绩文件并生成科目列表，再导入教师配置。")
            return
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择教师配置Excel文件", "", "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        
        if not file_path:
            return
        
        try:
            df = pd.read_excel(file_path, dtype=str)  # pandas自动选择引擎
            if df.shape[1] < 2:
                QMessageBox.critical(self, "导入错误", "文件至少需要两列（班级和班主任）")
                return
            
            subjects = self.params['subjects']
            expected_cols = ['班级', '班主任'] + subjects
            if df.shape[1] > len(expected_cols):
                df = df.iloc[:, :len(expected_cols)]
            df.columns = expected_cols[:df.shape[1]]
            df['班级'] = df['班级'].astype(str).str.strip()
            
            self.teachers_df = df
            self._refresh_teacher_table()
            QMessageBox.information(self, "导入成功", f"已导入 {len(df)} 个班级的教师配置。")
        except Exception as e:
            QMessageBox.critical(self, "导入错误", f"导入失败：{str(e)}")
    
    def _save_teachers(self):
        """将当前表格数据保存到 self.teachers_df，统一班级格式"""
        save_teachers(self)
    
    def _fill_teacher_info(self):
        """将教师信息填充到 subject_details 的'任课教师'列，并返回班主任映射
        
        Returns:
            班主任映射字典 {班级: 班主任}
        """
        return fill_teacher_info(self)
    
    def _show_about(self):
        """显示关于对话框"""
        about_text = """年段一分三率计算系统

版本 4.8

功能：导入班级成绩Excel，计算平均分、及格率、优秀率，
并进行班级评比、学生总排位以及单科班级详细排名。

新增特性：
- 增加座号识别与显示
- 班级汇总表增加总分分数段动态显示
- 单科班级排名表（含比差和名次）
- 语数英总分及排名
- 进退步计算（需导入历史总表）
- 教师配置功能完整实现（导入、保存、刷新、双击编辑）

使用步骤：
1. 导入Excel文件（可多选）
2. （可选）导入历史总表，勾选"启用进退步计算"
3. 设置各科分数线及参数
4. 配置教师信息
5. 点击计算查看结果
6. 导出汇总Excel

技术支持：Python + PyQt5 + Pandas
"""
        QMessageBox.information(self, "关于", about_text)
    
    def resizeEvent(self, event):
        """窗口大小改变事件处理，实现分辨率适配"""
        super().resizeEvent(event)
        
        # 可以在这里添加自定义的分辨率适配逻辑
        # 例如根据窗口大小调整字体大小、控件位置等
        
        # 示例：根据窗口宽度调整左侧面板宽度
        if hasattr(self, 'centralWidget'):
            pass
