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
    QFileDialog, QMessageBox, QScrollArea, QHeaderView, QProgressBar
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
        self.class_order: List[str] = []
        
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
        
        # 进度条
        self.progress_bar: Optional[QProgressBar] = None
        
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
        
        # 创建进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        main_layout.addWidget(self.progress_bar)
        
        # 创建状态栏
        self.statusBar().showMessage("就绪")
    
    def _create_menu(self):
        """创建菜单栏"""
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu("文件")
        import_class_action = file_menu.addAction("从班级导入")
        import_class_action.triggered.connect(lambda: self._import_grade_file(use_class_file=True))
        import_total_action = file_menu.addAction("从总表导入")
        import_total_action.triggered.connect(lambda: self._import_grade_file(use_class_file=False))
        file_menu.addSeparator()
        exit_action = file_menu.addAction("退出")
        exit_action.triggered.connect(self.close)
        
        # 帮助菜单
        help_menu = menubar.addMenu("帮助")
        about_action = help_menu.addAction("关于")
        about_action.triggered.connect(self._show_about)
        
        # 添加项目地址菜单项
        project_action = help_menu.addAction("项目地址")
        project_action.triggered.connect(self._open_project_url)
    
    def _create_toolbar(self):
        """创建工具栏"""
        toolbar = self.addToolBar("工具栏")
        toolbar.setMovable(False)
        
        import_btn = QPushButton("从班级导入")
        import_btn.clicked.connect(self.load_files)
        toolbar.addWidget(import_btn)
        
        import_total_btn = QPushButton("从总表导入")
        import_total_btn.clicked.connect(self.load_total_file)
        toolbar.addWidget(import_total_btn)
        
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

        
        # 5. 操作按钮
        btn_frame2 = QWidget()
        btn_layout2 = QHBoxLayout(btn_frame2)
        
        clear_btn = QPushButton("清除内存数据")
        clear_btn.clicked.connect(self._clear_memory_data)
        btn_layout2.addWidget(clear_btn)
        
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
        
        # 班级详情页
        self.class_detail_frame = QWidget()
        class_detail_layout = QVBoxLayout(self.class_detail_frame)
        
        top_frame3 = QWidget()
        top_layout3 = QHBoxLayout(top_frame3)
        
        class_detail_label = QLabel("选择班级：")
        top_layout3.addWidget(class_detail_label)
        
        self.class_detail_combobox = QComboBox()
        self.class_detail_combobox.currentTextChanged.connect(self._on_class_detail_selected)
        top_layout3.addWidget(self.class_detail_combobox)
        
        class_detail_layout.addWidget(top_frame3)
        
        self.class_detail_table = create_table()
        class_detail_layout.addWidget(self.class_detail_table)
        
        self.notebook.addTab(self.class_detail_frame, "班级详情")
        
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
    

    
    def _clear_memory_data(self):
        """清除内存数据"""
        # 重置参数
        self.rank_combo.setCurrentText(config.DEFAULT_RANK_METHOD)
        self.exclude_missing_checkbox.setChecked(config.DEFAULT_EXCLUDE_MISSING)
        self.progress_checkbox.setChecked(False)
        
        # 清除历史数据
        self.history_df = None
        self.history_label.setText("未导入历史文件")
        
        # 重置核心数据
        self.params = self._get_default_params()
        self.raw_data = None  # 清除原始数据
        self.student_rank = None  # 清除学生排名
        self.class_summary = None  # 清除班级汇总
        self.subject_details.clear()  # 清除科目详情
        self.subject_rankings.clear()  # 清除科目排名
        self.teachers_df = None  # 清除教师数据
        self.class_order = []  # 清除班级顺序
        
        # 清空下拉框
        if self.subject_combobox:
            self.subject_combobox.clear()
            self.selected_subject_var = ""
        if self.ranking_combobox:
            self.ranking_combobox.clear()
            self.ranking_subject_var = ""
        if hasattr(self, 'class_detail_combobox'):
            self.class_detail_combobox.clear()
        
        # 清空表格
        for table in [self.student_table, self.class_table, self.subject_table, 
                     self.ranking_table, self.threshold_table, self.teacher_table]:
            if table:
                table.setRowCount(0)
                table.setColumnCount(0)
        
        # 清空班级详细表格
        if hasattr(self, 'class_detail_table') and self.class_detail_table:
            self.class_detail_table.setRowCount(0)
            self.class_detail_table.setColumnCount(0)
        
        # 清空教师表格
        self._refresh_teacher_table()
        
        # 清空表头
        self.statusBar().showMessage("内存数据已清除")
    
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
        
        self._update_progress(0, "正在加载文件...")
        
        try:
            total_files = len(file_paths)
            for i, file_path in enumerate(file_paths):
                progress = int((i + 1) / total_files * 70)  # 70% 用于加载文件
                self._update_progress(progress, f"正在加载文件 {i+1}/{total_files}...")
                
            self.raw_data = data_loader.load_excel_files(file_paths, {})
            
            self._update_progress(80, "正在构建科目输入...")
            self._build_subject_inputs()
            
            self._update_progress(90, "正在刷新教师表格...")
            self._refresh_teacher_table()
            
            self._update_progress(100, f"已加载 {len(self.raw_data)} 条学生记录")
            QMessageBox.information(self, "导入成功", f"成功加载 {len(file_paths)} 个文件，共 {len(self.raw_data)} 条记录。")
        except Exception as e:
            self._update_progress(0, "加载失败")
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
    
    def load_total_file(self):
        """从年段总表导入数据，自动根据班级列分好班级"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择年段总表Excel文件", "", "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        
        if not file_path:
            return
        
        self._update_progress(0, "正在加载总表文件...")
        
        try:
            self._update_progress(30, "正在读取文件...")
            self.raw_data = data_loader.load_total_score_file(file_path, {})
            
            self._update_progress(70, "正在构建科目输入...")
            self._build_subject_inputs()
            
            self._update_progress(90, "正在刷新教师表格...")
            self._refresh_teacher_table()
            
            self._update_progress(100, f"已加载 {len(self.raw_data)} 条学生记录")
            QMessageBox.information(self, "导入成功", f"成功从总表导入 {len(self.raw_data)} 条记录。\n班级数量：{self.raw_data['班级'].nunique()}")
        except Exception as e:
            self._update_progress(0, "加载失败")
            QMessageBox.critical(self, "导入错误", f"加载总表时出错：\n{str(e)}")
    
    def apply_params_and_calculate(self):
        """应用参数并执行计算"""
        if self.raw_data is None:
            QMessageBox.warning(self, "无数据", "请先导入成绩文件。")
            return
        
        self._update_progress(0, "正在准备计算...")
        
        # 清理之前的计算结果，释放内存
        self.student_rank = None
        self.class_summary = None
        self.subject_details.clear()
        self.subject_rankings.clear()
        
        # 清理班级详情表格
        if hasattr(self, 'class_detail_table') and self.class_detail_table:
            self.class_detail_table.setRowCount(0)
            self.class_detail_table.setColumnCount(0)
        
        # 强制垃圾回收
        import gc
        gc.collect()
        
        if self.teacher_table and self.teacher_table.rowCount() > 0:
            self._update_progress(10, "正在保存教师信息...")
            self._save_teachers()
        
        self._update_progress(20, "正在更新参数...")
        self._update_params_from_inputs()
        
        self._update_progress(30, "正在执行计算...")
        try:
            result = calculator.calculate(
                self.raw_data,
                self.params,
                history_df=self.history_df if self.progress_checkbox.isChecked() else None,
                calc_progress=self.progress_checkbox.isChecked()
            )
            
            self._update_progress(50, "正在处理计算结果...")
            if len(result) == 4:
                self.student_rank, self.class_summary, self.subject_details, self.class_order = result
            else:
                self.student_rank, self.class_summary, self.subject_details = result
                self.class_order = []
            
            self._update_progress(60, "正在填充教师信息...")
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
            
            if self.class_order:
                self._update_progress(70, "正在排序表格...")
                self._sort_tables_by_class_order()
            
            self._update_progress(75, "正在构建科目排名...")
            self._build_subject_rankings()
            
            self._update_progress(80, "正在更新学生表格...")
            self._update_student_table()
            
            self._update_progress(85, "正在更新班级表格...")
            self._update_class_table()
            
            self._update_progress(90, "正在更新下拉框...")
            self._update_subject_combobox()
            self._update_ranking_combobox()
            # 只更新班级详情下拉框，不自动更新表格
            if self.student_rank is not None:
                classes = self._sort_classes_numerically(self.student_rank['班级'].unique())
                if classes:
                    self.class_detail_combobox.clear()
                    self.class_detail_combobox.addItems(classes)
                else:
                    self.class_detail_combobox.clear()
            
            self._update_progress(100, "计算完成")
            
            # 再次强制垃圾回收，释放内存
            gc.collect()
        except Exception as e:
            self._update_progress(0, "计算失败")
            QMessageBox.critical(self, "计算错误", f"计算过程中出错：\n{str(e)}")
            
            # 发生错误时清理数据，释放内存
            self.student_rank = None
            self.class_summary = None
            self.subject_details.clear()
            self.subject_rankings.clear()
            gc.collect()
    
    def export_results(self):
        """导出计算结果到Excel文件"""
        if self.student_rank is None or self.class_summary is None:
            QMessageBox.warning(self, "无数据", "请先计算成绩")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存Excel文件", "年段成绩分析.xlsx", "Excel文件 (*.xlsx);;所有文件 (*.*)"
        )
        
        if not file_path:
            return
        
        if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'
        
        self._update_progress(0, "正在导出...")
        
        try:
            self._update_progress(20, "正在准备数据...")
            # 不再收集考试信息，已移除考试设置UI
            
            self._update_progress(40, "正在写入Excel文件...")
            # 直接使用用户选择的文件路径，允许覆盖已存在的文件
            exporter.export_to_excel(self.student_rank, self.class_summary,
                                      self.subject_details, self.subject_rankings, file_path,
                                      passing_score=self.params.get('passing_score', {}),
                                      excellent_score=self.params.get('excellent_score', {}))
            
            self._update_progress(70, "正在为分数段上色...")
            # 此处不需要额外代码，因为染色逻辑已经集成在export_to_excel函数中
            
            self._update_progress(90, "正在调整格式...")
            # 此处不需要额外代码，因为格式调整逻辑已经集成在export_to_excel函数中
            
            self._update_progress(100, f"导出成功：{os.path.basename(file_path)}")
            QMessageBox.information(self, "导出完成", f"结果已保存到：\n{file_path}")
        except PermissionError as e:
            self._update_progress(0, "导出失败")
            QMessageBox.critical(self, "权限错误", f"无法写入文件，可能是因为：\n1. 没有写入目标目录的权限\n2. 文件已被其他程序打开\n3. 文件是只读的\n\n请尝试保存到其他目录。\n\n错误信息：\n{str(e)}")
        except Exception as e:
            self._update_progress(0, "导出失败")
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
    
    def _sort_tables_by_class_order(self):
        """按照班级顺序对班级汇总表、单科班级分析和单科班级排名进行排序"""
        if not self.class_order:
            return
        
        class_order_map = {cls: idx for idx, cls in enumerate(self.class_order)}
        
        def sort_key(cls):
            cls_str = str(cls)
            if cls_str == '年段':
                return (1, 0)
            return (0, class_order_map.get(cls_str, 999))
        
        if self.class_summary is not None and '班级' in self.class_summary.columns:
            self.class_summary['排序键'] = self.class_summary['班级'].astype(str).apply(sort_key)
            self.class_summary = self.class_summary.sort_values('排序键').drop(columns=['排序键']).reset_index(drop=True)
        
        if self.subject_details:
            for subj in self.subject_details:
                df = self.subject_details[subj]
                if '班级' in df.columns:
                    df['排序键'] = df['班级'].astype(str).apply(sort_key)
                    self.subject_details[subj] = df.sort_values('排序键').drop(columns=['排序键']).reset_index(drop=True)
        
        if self.subject_rankings:
            for subj in self.subject_rankings:
                df = self.subject_rankings[subj]
                if '班级' in df.columns:
                    df['排序键'] = df['班级'].astype(str).apply(sort_key)
                    self.subject_rankings[subj] = df.sort_values('排序键').drop(columns=['排序键']).reset_index(drop=True)
        
        if self.teachers_df is not None and '班级' in self.teachers_df.columns:
            self.teachers_df['排序键'] = self.teachers_df['班级'].astype(str).apply(sort_key)
            self.teachers_df = self.teachers_df.sort_values('排序键').drop(columns=['排序键']).reset_index(drop=True)
            self._refresh_teacher_table()
    
    def _build_subject_rankings(self):
        """构建单科班级排名数据"""
        self.subject_rankings.clear()
        
        # 获取班主任映射
        head_map = {}
        if hasattr(self, 'teachers_df') and self.teachers_df is not None:
            for _, row in self.teachers_df.iterrows():
                cls = str(row.get('班级', '')).strip()
                head = str(row.get('班主任', '')).strip()
                if cls and head:
                    head_map[cls] = head
        
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
        
        # 添加总分栏，任课教师为班主任
        if '总分' in self.subject_details:
            total_df = self.subject_details['总分']
            grade_row = total_df[total_df['班级'] == '年段'].iloc[0] if '年段' in total_df['班级'].values else None
            if grade_row is not None:
                class_rows = total_df[total_df['班级'] != '年段'].copy()
                if not class_rows.empty:
                    data = class_rows[['班级', '平均分', '及格率', '优生率']].copy()
                    # 添加班主任作为任课教师
                    data['任课教师'] = data['班级'].map(head_map).fillna('')
                    grade_avg = grade_row['平均分']
                    grade_pass = grade_row['及格率']
                    grade_excel = grade_row['优生率']
                    data['平均分比差'] = (data['平均分'] - grade_avg) / grade_avg * 10 if grade_avg > 0 else 0
                    data['及格率比差'] = (data['及格率'] - grade_pass) * 10
                    data['优生率比差'] = (data['优生率'] - grade_excel) * 10
                    data['平均分名次'] = data['平均分'].rank(method='min', ascending=False).astype('Int64')
                    data['及格率名次'] = data['及格率'].rank(method='min', ascending=False).astype('Int64')
                    data['优生率名次'] = data['优生率'].rank(method='min', ascending=False).astype('Int64')
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
                    self.subject_rankings['总分'] = ranking_df
    
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
    
    def _sort_classes_numerically(self, classes):
        """按照班级名称中的数字顺序排序班级"""
        import re
        
        def get_class_number(cls):
            match = re.search(r'(\d+)', str(cls))
            if match:
                return int(match.group(1))
            return 999999
        
        return sorted(classes, key=get_class_number)
    
    def _update_class_detail_combobox(self):
        """更新班级详情下拉框"""
        if self.student_rank is not None:
            classes = self._sort_classes_numerically(self.student_rank['班级'].unique())
            if classes:
                self.class_detail_combobox.clear()
                self.class_detail_combobox.addItems(classes)
                # 始终更新班级详情表格，确保数据的一致性
                self._on_class_detail_selected(classes[0])
            else:
                self.class_detail_combobox.clear()
                self.class_detail_table.setRowCount(0)
                self.class_detail_table.setColumnCount(0)
        else:
            self.class_detail_combobox.clear()
            self.class_detail_table.setRowCount(0)
            self.class_detail_table.setColumnCount(0)
    
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
    
    def _on_class_detail_selected(self, class_name):
        """班级详情选择事件处理
        
        Args:
            class_name: 选择的班级
        """
        if class_name:
            self._update_class_detail_table(class_name)
        else:
            self.class_detail_table.setRowCount(0)
            self.class_detail_table.setColumnCount(0)
    
    def _update_class_detail_table(self, class_name):
        """更新班级详情表格
        
        Args:
            class_name: 要显示的班级
        """
        if self.student_rank is None:
            return
        
        # 优化：使用布尔索引过滤数据，避免copy操作
        class_df = self.student_rank[self.student_rank['班级'] == class_name]
        
        if class_df.empty:
            self.class_detail_table.setRowCount(0)
            self.class_detail_table.setColumnCount(0)
            return
        
        # 优化：预先获取数据，避免在循环中重复访问
        rows, cols = class_df.shape
        columns = list(class_df.columns)
        data = class_df.values.tolist()
        
        # 设置表格行列
        self.class_detail_table.setRowCount(rows)
        self.class_detail_table.setColumnCount(cols)
        
        # 设置表头
        self.class_detail_table.setHorizontalHeaderLabels(columns)
        
        # 优化：批量填充数据，减少表格操作次数
        for idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data):
                if pd.notna(value):
                    item = QTableWidgetItem(str(value))
                else:
                    item = QTableWidgetItem("")
                item.setTextAlignment(Qt.AlignCenter)
                self.class_detail_table.setItem(idx, col_idx, item)
    
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
            # 清理之前的教师数据，释放内存
            self.teachers_df = None
            
            # 强制垃圾回收
            import gc
            gc.collect()
            
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
            
            # 再次强制垃圾回收，释放内存
            gc.collect()
        except Exception as e:
            QMessageBox.critical(self, "导入错误", f"导入失败：{str(e)}")
            
            # 发生错误时清理数据，释放内存
            self.teachers_df = None
            gc.collect()
    
    def _save_teachers(self):
        """将当前表格数据保存到 self.teachers_df，统一班级格式"""
        save_teachers(self)
    
    def _fill_teacher_info(self):
        """将教师信息填充到 subject_details 的'任课教师'列，并返回班主任映射
        
        Returns:
            班主任映射字典 {班级: 班主任}
        """
        return fill_teacher_info(self)
    
    def _open_project_url(self):
        """打开项目GitHub地址"""
        import webbrowser
        webbrowser.open("https://github.com/beiyan124/Simple-Grade-Calculation-System")

    def _show_about(self):
        """显示关于对话框"""
        about_text = """年段一分三率计算系统

版本 0.6

功能：
1. 导入Excel成绩文件
2. 计算学生总分、排名和进退步
3. 统计班级一分三率（平均分、及格率、优秀率）
4. 生成单科班级分析和排名
5. 导入教师信息，支持自动识别班级格式
6. 导出格式化的Excel结果报表，包含染色功能

新增特性：
- 更新配色方案，优化输出表格的染色机制
- 为班级汇总表分数段添加绿色系染色方案（拿破仑绿、中绿、中浅绿、浅绿、浅浅绿）
- 优化学生排名表的染色颜色，使层次更加分明（中蓝、浅天蓝、浅蓝、中橙）
- 为年级排名、班级排名、语数英总分、语数英排名、上次排名、进退步添加不同颜色标识
- 为输出表格添加冻结窗格功能，提高数据可读性
- 增加导入教师过程中的兼容性，自动识别班级中的数字并替换为已分班级的文字形式
- 在单科班级排名的总分一栏添加班主任信息
- 调整学生排名表的列顺序，将座号放在姓名之前
- 实现学生排名的染色功能，基于科目阈值的及格线和优秀线
- 实现用中橙色标记科目最高分的功能

技术支持：Python + PyQt5 + Pandas

作者：beiyan124与Trae,deepseek
"""
        QMessageBox.information(self, "关于", about_text)
    
    def _update_progress(self, value: int, message: str = ""):
        """更新进度条
        
        Args:
            value: 进度值（0-100）
            message: 进度消息
        """
        if self.progress_bar:
            self.progress_bar.setValue(value)
        if message:
            self.statusBar().showMessage(message)
        # 强制刷新界面
        QApplication.processEvents()
    
    def resizeEvent(self, event):
        """窗口大小改变事件处理，实现分辨率适配"""
        super().resizeEvent(event)
        
        # 可以在这里添加自定义的分辨率适配逻辑
        # 例如根据窗口大小调整字体大小、控件位置等
        
        # 示例：根据窗口宽度调整左侧面板宽度
        if hasattr(self, 'centralWidget'):
            pass
