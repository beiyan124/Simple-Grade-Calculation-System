"""
配置模块

存储系统默认参数和常量，包括科目特定默认阈值、配色方案、列名变体、分数段设置等
新增：历史文件列名变体
"""

# ==================== 基础配置 ====================
DEFAULT_COLUMN_MAPPING = {
    '班级': '班级',
    '姓名': '姓名',
    '座号': '座号'
}
DEFAULT_WEIGHTS = {}
DEFAULT_RANK_METHOD = 'min'
DEFAULT_NA_VALUES = ['', 'NULL', '缺考', '缺席', 'nan', 'NaN', 'NA']
DEFAULT_EXCLUDE_MISSING = True
DEFAULT_OUTPUT_PREFIX = '年级成绩汇总'
DEFAULT_WINDOW_SIZE = '1300x800'
DEFAULT_SHEET_NAME = 0
DEFAULT_INFER_CLASS_FROM_FILENAME = True

# ==================== 列名变体 ====================
CLASS_COLUMN_VARIANTS = [
    '班级', '班', 'class', 'Class', '班级名称', '班别', '教学班',
    '行政班', '班级名称', '所属班级', '班号', 'classname', 'className'
]
NAME_COLUMN_VARIANTS = [
    '姓名', '名字', 'name', 'Name', '学生姓名', '学生名字', '考生姓名',
    '姓名全称', 'fullname', 'FullName', 'student name', 'StudentName'
]
ZUOHAO_COLUMN_VARIANTS = [
    '座号', '座位号', '座次', '座位', '序号', '编号', '学号',
    'zuohao', 'seat', 'Seat', 'seat number', 'SeatNumber'
]

# ==================== 历史文件列名变体（新增）====================
HISTORY_NAME_VARIANTS = NAME_COLUMN_VARIANTS + ['学生', '考生']
HISTORY_RANK_VARIANTS = [
    '排名', '名次', '段名', '年级排名', '上次排名', 'rank', 'Rank',
    '年级名次', '年段排名', '段次'
]

# ==================== 阈值设置 ====================
DEFAULT_THRESHOLDS = {
    '满分': 100,
    '及格线': 60,
    '优秀线': 80
}
SUBJECT_SPECIFIC_THRESHOLDS = {
    '语文': {'满分': 150, '及格线': 90, '优秀线': 120},
    '数学': {'满分': 150, '及格线': 90, '优秀线': 120},
    '英语': {'满分': 150, '及格线': 90, '优秀线': 120},
    '物理': {'满分': 100, '及格线': 60, '优秀线': 80},
    '化学': {'满分': 100, '及格线': 60, '优秀线': 80},
    '历史': {'满分': 100, '及格线': 60, '优秀线': 80},
    '政治': {'满分': 100, '及格线': 60, '优秀线': 80},
    '地理': {'满分': 100, '及格线': 60, '优秀线': 80},
    '生物': {'满分': 100, '及格线': 60, '优秀线': 80},
    '体育': {'满分': 40, '及格线': 0, '优秀线': 0}
}

# ==================== 分数段相关常量 ====================
NUM_SCORE_SEGMENTS = 10
DEFAULT_SCORE_BINS = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
DEFAULT_SCORE_LABELS = [
    '0-10', '10-20', '20-30', '30-40', '40-50',
    '50-60', '60-70', '70-80', '80-90', '90-100'
]
SUBJECT_DETAIL_COLUMNS = ['班级', '参考人数'] + DEFAULT_SCORE_LABELS + \
                         ['及格率', '优生率', '平均分', '最高分', '最低分', '任课教师']

# ==================== 配色方案 ====================
COLOR_STUDENT = {
    'header_bg': '#B3E5FC',
    'header_fg': '#01579B',
    'odd_row_bg': '#FFFFFF',
    'even_row_bg': '#E1F5FE',
    'border': '#81D4FA',
    'text': '#212121',
    'highlight_bg': '#FFECB3'
}
COLOR_CLASS = {
    'header_bg': '#C8E6C9',
    'header_fg': '#1B5E20',
    'odd_row_bg': '#FFFFFF',
    'even_row_bg': '#F1F8E9',
    'border': '#A5D6A7',
    'text': '#212121',
    'highlight_bg': '#FFF9C4'
}
COLOR_SUBJECT = {
    'header_bg': '#FFE0B2',
    'header_fg': '#E65100',
    'odd_row_bg': '#FFFFFF',
    'even_row_bg': '#FFF3E0',
    'border': '#FFCC80',
    'text': '#212121',
    'highlight_bg': '#F8BBD0'
}