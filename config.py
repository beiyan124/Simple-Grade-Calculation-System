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

# ==================== 新增色表（从图片识别）====================
# 左列：蓝色系（从上到下）
COLOR_BLUE_SCALE = [
    '#E3F2FD',  # 浅蓝
    '#BBDEFB',  # 稍浅蓝
    '#64B5F6',  # 中蓝
    '#1E88E5',  # 深蓝
    '#0D47A1'   # 最深蓝
]

# 右列：橙色/棕色系（从上到下）
COLOR_ORANGE_SCALE = [
    '#FFF3E0',  # 浅橙
    '#FFE0B2',  # 稍浅橙
    '#FFB74D',  # 中橙
    '#FF8F00',  # 深橙
    '#E65100'   # 最深棕
]

# 综合色表（可用于扩展配色方案）
COLOR_CUSTOM = {
    'header_bg': '#64B5F6',      # 中蓝
    'header_fg': '#0D47A1',      # 最深蓝
    'odd_row_bg': '#FFFFFF',     # 白色
    'even_row_bg': '#E3F2FD',    # 浅蓝
    'border': '#BBDEFB',         # 稍浅蓝
    'text': '#212121',           # 黑色
    'highlight_bg': '#FFB74D'    # 中橙
}

# ==================== 识别色（新增）====================
# 基于用户提供的图片识别的配色方案
COLOR_IDENTIFICATION = {
    'header_bg': '#1976D2',      # 深蓝色（右侧）
    'header_fg': '#FFFFFF',      # 白色文字
    'odd_row_bg': '#FFFFFF',     # 白色
    'even_row_bg': '#E3F2FD',    # 浅蓝色（左侧）
    'border': '#81D4FA',         # 浅蓝色边框
    'text': '#212121',           # 黑色文字
    'highlight_bg': '#CE93D8'    # 浅紫色/粉色（中间）
}

# 识别色表（从图片提取的三种颜色）
IDENTIFICATION_COLORS = [
    '#81D4FA',  # 浅蓝色（左侧）
    '#CE93D8',  # 浅紫色/粉色（中间）
    '#1976D2'   # 深蓝色（右侧）
]

# ==================== 学生排名染色方案（新增）====================# 学生排名染色方案（新增）
# 学生排名表中根据分数范围的染色方案
COLOR_SCORE_RANGE = {
    'excellent': '#64B5F6',  # 优秀线往上 - 中蓝
    'passing': '#BBDEFB',     # 优秀线到及格线 - 浅天蓝
    'failing': '#E1F5FE',     # 及格线往下 - 浅蓝
    'highest': '#FFB74D'      # 科目最高分 - 中橙
}

# ==================== 班级汇总分数段染色方案（新增）====================
# 班级汇总表中分数段的染色方案
COLOR_CLASS_SCORE_SEGMENT = [
    '#EBF1DE',  # 浅浅绿
    '#D8E4BC',  # 浅绿
    '#C4D79B',  # 中浅绿
    '#A3BF5F',  # 中绿
    '#76933C'   # 拿破仑绿（最深）
]