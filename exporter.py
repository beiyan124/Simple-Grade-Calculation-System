"""
输出模块

负责将计算结果导出为格式化的Excel文件，包含学生排名、班级汇总、单科班级分析和单科班级排名
分析和排名表分别合并为单个工作表，各科目纵向排列，每个科目块前添加科目名称行，块间空一行
使用配色方案设置表头、奇偶行、边框和高亮行

版本：4.6.2
新增：每个科目块前添加科目名称行
修复：处理pandas的<NA>值
"""

import pandas as pd
import numpy as np
import os
import warnings
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from typing import Dict, List, Tuple

# 忽略libpng警告（可能由openpyxl或Pillow内部触发）
warnings.filterwarnings("ignore", message="libpng warning: iCCP: known incorrect sRGB profile")

# 常见科目顺序（可根据需要调整）
COMMON_SUBJECT_ORDER = ['语文', '数学', '英语', '物理', '化学', '生物', '历史', '政治', '地理', '体育']

def export_to_excel(student_df: pd.DataFrame, class_df: pd.DataFrame,
                    subject_details: Dict[str, pd.DataFrame],
                    subject_rankings: Dict[str, pd.DataFrame],
                    output_path: str, exam_info: Dict[str, str] = None) -> None:
    """
    将学生排名表、班级汇总表、单科班级分析（纵向堆叠）和单科班级排名（纵向堆叠）导出为Excel文件

    参数:
        student_df: 学生排名表DataFrame
        class_df: 班级汇总表DataFrame
        subject_details: 字典，键为科目名，值为该科目的班级分析表DataFrame
        subject_rankings: 字典，键为科目名，值为该科目的班级排名表DataFrame
        output_path: 输出文件路径
        exam_info: 字典，包含考试相关信息，如考试名称、考试时间、出卷人等
    """
    # 确保输出目录存在
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    # 处理DataFrame中的<NA>值，替换为np.nan，以便openpyxl能处理
    student_df = student_df.replace(pd.NA, np.nan)
    class_df = class_df.replace(pd.NA, np.nan)

    # 使用openpyxl引擎写入Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 先创建所有工作表
        student_sheet_name = '学生排名'
        class_sheet_name = '班级汇总'
        analysis_sheet_name = '单科班级分析'
        ranking_sheet_name = '单科班级排名'
        
        # 写入基本表数据
        student_df.to_excel(writer, sheet_name=student_sheet_name, index=False)
        class_df.to_excel(writer, sheet_name=class_sheet_name, index=False)

        # 获取工作簿对象
        workbook = writer.book
        
        # 1. 创建考试信息首页（如果有考试信息）
        if exam_info:
            info_sheet = workbook.create_sheet('考试信息', 0)  # 插入到最前面
            _create_info_sheet(info_sheet, exam_info)
        
        # 2. 为学生排名表添加标题并格式化
        student_sheet = workbook[student_sheet_name]
        _add_sheet_title(student_sheet, exam_info)
        _format_student_sheet(student_sheet)
        
        # 3. 为班级汇总表添加标题并格式化
        class_sheet = workbook[class_sheet_name]
        _add_sheet_title(class_sheet, exam_info)
        _format_class_sheet(class_sheet, class_df)
        
        # 4. 创建并处理单科班级分析工作表
        if subject_details:
            analysis_sheet = workbook.create_sheet(analysis_sheet_name, len(workbook.sheetnames))
            _add_sheet_title(analysis_sheet, exam_info)
            analysis_blocks = _write_blocks(analysis_sheet, subject_details, is_analysis=True)
            _format_block_sheet(analysis_sheet, analysis_blocks, is_analysis=True)
        
        # 5. 创建并处理单科班级排名工作表
        if subject_rankings:
            ranking_sheet = workbook.create_sheet(ranking_sheet_name, len(workbook.sheetnames))
            _add_sheet_title(ranking_sheet, exam_info)
            ranking_blocks = _write_blocks(ranking_sheet, subject_rankings, is_analysis=False)
            _format_block_sheet(ranking_sheet, ranking_blocks, is_analysis=False)
        
        # 6. 调整所有工作表的列宽
        for sheet_name in workbook.sheetnames:
            _adjust_column_width(workbook[sheet_name])


def _create_info_sheet(sheet, exam_info):
    """
    创建考试信息首页
    
    参数:
        sheet: openpyxl工作表对象
        exam_info: 字典，包含考试相关信息
    """
    # 设置页面背景色为浅黄色
    sheet.sheet_properties.tabColor = "FFFFCC"
    
    # 设置整个页面的背景色为浅黄色
    for row in range(1, 21):
        for col in range(1, 11):
            cell = sheet.cell(row=row, column=col)
            cell.fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
    
    # 设置学年信息
    sheet.merge_cells(start_row=5, start_column=1, end_row=5, end_column=10)
    year_cell = sheet.cell(row=5, column=1)
    year_cell.value = exam_info.get('考试名称', '').split(' ')[0] if exam_info.get('考试名称', '') else ''
    year_cell.font = Font(bold=True, size=16)
    year_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 设置学校和年段信息
    sheet.merge_cells(start_row=7, start_column=1, end_row=7, end_column=10)
    school_cell = sheet.cell(row=7, column=1)
    school_cell.value = exam_info.get('考试名称', '').split(' ')[1] if len(exam_info.get('考试名称', '').split(' ')) > 1 else ''
    school_cell.font = Font(bold=True, size=16)
    school_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 设置表格标题
    sheet.merge_cells(start_row=9, start_column=1, end_row=9, end_column=10)
    table_title_cell = sheet.cell(row=9, column=1)
    table_title_cell.value = exam_info.get('考试名称', '').split(' ')[2] if len(exam_info.get('考试名称', '').split(' ')) > 2 else '质量分析表'
    table_title_cell.font = Font(bold=True, size=20)
    table_title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 设置内部资料提示
    sheet.merge_cells(start_row=11, start_column=1, end_row=11, end_column=10)
    note_cell = sheet.cell(row=11, column=1)
    note_cell.value = "(内部资料注意保存)"
    note_cell.font = Font(size=12)
    note_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 设置落款
    sheet.merge_cells(start_row=17, start_column=1, end_row=17, end_column=10)
    signature_cell = sheet.cell(row=17, column=1)
    signature_cell.value = exam_info.get('出卷人', 'xx中学教务处')
    signature_cell.font = Font(size=14)
    signature_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 设置日期
    sheet.merge_cells(start_row=19, start_column=1, end_row=19, end_column=10)
    date_cell = sheet.cell(row=19, column=1)
    date_cell.value = exam_info.get('考试时间日期', '')
    date_cell.font = Font(size=14)
    date_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 调整列宽
    for col in range(1, 11):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 12
    
    # 调整行高
    for row in range(1, 21):
        sheet.row_dimensions[row].height = 25


def _add_sheet_title(sheet, exam_info):
    """
    为工作表添加标题
    
    参数:
        sheet: openpyxl工作表对象
        exam_info: 字典，包含考试相关信息
    """
    if not exam_info or '考试名称' not in exam_info:
        return
    
    # 插入标题行
    sheet.insert_rows(1, 3)
    
    # 合并标题区域
    sheet.merge_cells(start_row=1, start_column=1, end_row=3, end_column=10)
    
    # 设置标题背景色
    for row in range(1, 4):
        for col in range(1, 11):
            cell = sheet.cell(row=row, column=col)
            cell.fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
    
    # 设置标题文本 - 只对合并单元格的左上角单元格（第1行第1列）进行操作
    title_cell = sheet.cell(row=1, column=1)
    title_cell.value = exam_info['考试名称']
    title_cell.font = Font(bold=True, size=16)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 调整标题行高
    for row in range(1, 4):
        sheet.row_dimensions[row].height = 20

def _sort_subjects(subjects: List[str]) -> List[str]:
    """按常见科目顺序排序，未知科目放在后面"""
    def key_func(s):
        try:
            return COMMON_SUBJECT_ORDER.index(s)
        except ValueError:
            return len(COMMON_SUBJECT_ORDER)  # 未知科目排在最后
    return sorted(subjects, key=key_func)

def _write_blocks(worksheet, subject_dict: Dict[str, pd.DataFrame], is_analysis: bool) -> List[Tuple[str, int, int]]:
    """
    将多个科目的DataFrame纵向写入工作表，每个块前添加科目名称行，块间空一行

    参数:
        worksheet: openpyxl工作表对象
        subject_dict: 科目名到DataFrame的映射
        is_analysis: 是否分析表（仅用于标识，此处未使用）

    返回:
        blocks: 列表，每个元素为 (科目名, 表头行号, 数据最后一行行号)
    """
    subjects_sorted = _sort_subjects(list(subject_dict.keys()))
    current_row = 4  # 从标题下方开始（标题占了3行）
    blocks = []

    for subject in subjects_sorted:
        df = subject_dict[subject]

        # 如果不是第一个块，先插入一个空行
        if current_row > 4:
            current_row += 1

        # 写入科目名称行（第一列写科目名，其余单元格留空）
        worksheet.cell(row=current_row, column=1, value=subject)
        title_row = current_row
        current_row += 1

        # 写入表头（列名）
        header_row = current_row
        for col_idx, col_name in enumerate(df.columns, start=1):
            try:
                cell = worksheet.cell(row=header_row, column=col_idx)
                cell.value = col_name
            except AttributeError:
                # 忽略合并单元格的只读属性错误
                pass

        # 写入数据行
        data_start_row = header_row + 1
        for row_idx, (_, row_data) in enumerate(df.iterrows(), start=data_start_row):
            for col_idx, value in enumerate(row_data, start=1):
                try:
                    if pd.isna(value):
                        cell_value = None
                    else:
                        cell_value = value
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.value = cell_value
                except AttributeError:
                    # 忽略合并单元格的只读属性错误
                    pass

        data_end_row = data_start_row + len(df) - 1
        # 记录此块信息
        blocks.append((subject, header_row, data_end_row))

        # 更新当前行为下一个块的起始行
        current_row = data_end_row + 1

    return blocks

def _format_block_sheet(worksheet, blocks: List[Tuple[str, int, int]], is_analysis: bool) -> None:
    """
    对纵向堆叠的块工作表应用样式：每个块内表头特殊，数据行奇偶交替，边框，科目名称行特殊样式

    参数:
        worksheet: 已写入数据的工作表
        blocks: 块信息列表，每个元素为 (科目名, 表头行号, 数据最后一行行号)
        is_analysis: 是否分析表（用于确定百分比列）
    """
    from config import COLOR_SUBJECT as color

    # 边框
    border = _get_border(color['border'])
    # 表头样式
    header_fill = _get_fill(color['header_bg'])
    header_font = Font(bold=True, color=color['header_fg'][1:])
    header_alignment = Alignment(horizontal='center', vertical='center')
    # 数据行样式
    odd_fill = _get_fill(color['odd_row_bg'])
    even_fill = _get_fill(color['even_row_bg'])
    data_alignment = Alignment(horizontal='center', vertical='center')
    # 科目名称行样式（加粗，不同背景色，居中）
    title_fill = _get_fill(color['header_bg'])  # 与表头相同背景，或可自定义
    title_font = Font(bold=True, color=color['header_fg'][1:])
    title_alignment = Alignment(horizontal='left', vertical='center')

    # 百分比列名称（用于分析和排名表）
    pct_columns = ['及格率', '优生率', '平均分比差', '及格率比差', '优生率比差']

    # 不需要额外的行号偏移，因为_write_blocks函数返回的行号已经考虑了标题行的位置

    for subject, header_row, data_end_row in blocks:
        title_row = header_row - 1
        # 直接使用blocks中返回的行号，不需要再偏移
        adjusted_title_row = title_row
        adjusted_header_row = header_row
        adjusted_data_end_row = data_end_row
        
        # 格式化科目名称行
        for col_idx in range(1, worksheet.max_column + 1):
            try:
                cell = worksheet.cell(row=adjusted_title_row, column=col_idx)
                cell.fill = title_fill
                cell.font = title_font
                cell.alignment = title_alignment
                cell.border = border
            except AttributeError:
                # 忽略合并单元格的只读属性错误
                pass

        # 格式化表头行
        for col_idx in range(1, worksheet.max_column + 1):
            try:
                cell = worksheet.cell(row=adjusted_header_row, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment
                cell.border = border
            except AttributeError:
                # 忽略合并单元格的只读属性错误
                pass

        # 格式化数据行（奇偶背景）
        data_start_row = adjusted_header_row + 1
        for row_idx in range(data_start_row, adjusted_data_end_row + 1):
            row_fill = even_fill if (row_idx - data_start_row) % 2 == 0 else odd_fill
            for col_idx in range(1, worksheet.max_column + 1):
                try:
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.fill = row_fill
                    cell.alignment = data_alignment
                    cell.border = border

                    # 对百分比列应用数字格式
                    try:
                        col_name = worksheet.cell(row=adjusted_header_row, column=col_idx).value
                        if col_name in pct_columns and isinstance(cell.value, (int, float)):
                            cell.number_format = '0.00%'
                    except AttributeError:
                        # 忽略合并单元格的只读属性错误
                        pass
                except AttributeError:
                    # 忽略合并单元格的只读属性错误
                    pass

def _get_fill(color_hex: str) -> PatternFill:
    """根据六位十六进制颜色代码创建填充对象（自动去掉#，并确保为6位）"""
    if color_hex.startswith('#'):
        color_hex = color_hex[1:]
    if len(color_hex) == 6:
        return PatternFill(start_color=color_hex, end_color=color_hex, fill_type='solid')
    else:
        return PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

def _get_border(color_hex: str) -> Border:
    """根据边框颜色创建边框对象"""
    if color_hex.startswith('#'):
        color_hex = color_hex[1:]
    side = Side(style='thin', color=color_hex)
    return Border(left=side, right=side, top=side, bottom=side)

def _format_student_sheet(worksheet):
    """设置学生排名表的格式：表头颜色、奇偶行颜色、边框、第一名高亮"""
    from config import COLOR_STUDENT as color

    header_fill = _get_fill(color['header_bg'])
    header_font = Font(bold=True, color=color['header_fg'][1:])
    header_alignment = Alignment(horizontal='center', vertical='center')
    border = _get_border(color['border'])

    # 表头现在在第4行（因为插入了3行标题）
    header_row = 4
    
    # 格式化表头行
    for col_idx in range(1, worksheet.max_column + 1):
        try:
            cell = worksheet.cell(row=header_row, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = border
        except AttributeError:
            # 忽略合并单元格的只读属性错误
            pass

    odd_fill = _get_fill(color['odd_row_bg'])
    even_fill = _get_fill(color['even_row_bg'])
    highlight_fill = _get_fill(color['highlight_bg'])
    data_alignment = Alignment(horizontal='center', vertical='center')

    # 获取表头，处理可能的合并单元格
    header = []
    for col_idx in range(1, worksheet.max_column + 1):
        try:
            cell = worksheet.cell(row=header_row, column=col_idx)
            header.append(cell.value)
        except AttributeError:
            # 忽略合并单元格的只读属性错误
            header.append(None)
    rank_col_idx = None
    if '年级排名' in header:
        rank_col_idx = header.index('年级排名') + 1

    # 格式化数据行
    for row_idx in range(header_row + 1, worksheet.max_row + 1):
        row_fill = even_fill if (row_idx - header_row - 1) % 2 == 0 else odd_fill
        is_highlight = False
        if rank_col_idx:
            try:
                cell = worksheet.cell(row=row_idx, column=rank_col_idx)
                if isinstance(cell.value, (int, float)) and cell.value == 1:
                    is_highlight = True
            except AttributeError:
                # 忽略合并单元格的只读属性错误
                pass

        for col_idx in range(1, worksheet.max_column + 1):
            try:
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.fill = highlight_fill if is_highlight else row_fill
                cell.alignment = data_alignment
                cell.border = border
            except AttributeError:
                # 忽略合并单元格的只读属性错误
                pass

def _format_class_sheet(worksheet, df):
    """设置班级汇总表的格式：表头颜色、奇偶行颜色、边框"""
    from config import COLOR_CLASS as color

    header_fill = _get_fill(color['header_bg'])
    header_font = Font(bold=True, color=color['header_fg'][1:])
    header_alignment = Alignment(horizontal='center', vertical='center')
    border = _get_border(color['border'])

    # 表头现在在第4行（因为插入了3行标题）
    header_row = 4
    
    # 格式化表头行
    for col_idx in range(1, worksheet.max_column + 1):
        try:
            cell = worksheet.cell(row=header_row, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = border
        except AttributeError:
            # 忽略合并单元格的只读属性错误
            pass

    odd_fill = _get_fill(color['odd_row_bg'])
    even_fill = _get_fill(color['even_row_bg'])
    data_alignment = Alignment(horizontal='center', vertical='center')

    # 格式化数据行
    for row_idx in range(header_row + 1, worksheet.max_row + 1):
        row_fill = even_fill if (row_idx - header_row - 1) % 2 == 0 else odd_fill
        for col_idx in range(1, worksheet.max_column + 1):
            try:
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.fill = row_fill
                cell.alignment = data_alignment
                cell.border = border
            except AttributeError:
                # 忽略合并单元格的只读属性错误
                pass

def _adjust_column_width(worksheet):
    """根据内容自动调整列宽"""
    for col in worksheet.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                # 检查单元格是否是合并单元格的一部分
                # 只有左上角的单元格可以访问value属性
                if cell.value:
                    length = len(str(cell.value))
                    if length > max_length:
                        max_length = length
            except AttributeError:
                # 忽略合并单元格的只读属性错误
                pass
        adjusted_width = min(max(max_length + 2, 8), 40)
        worksheet.column_dimensions[col_letter].width = adjusted_width