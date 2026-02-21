"""
计算模块

负责学生总分计算、排名以及班级一分三率统计，并生成单科班级详细统计表（含分数段分布）
新增：语数英总分及排名，进退步计算（基于历史排名）
"""

import pandas as pd
import numpy as np
from typing import Tuple, Dict, Any, List, Optional

def calculate(df: pd.DataFrame, params: Dict[str, Any],
              history_df: Optional[pd.DataFrame] = None,
              calc_progress: bool = False) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, pd.DataFrame]]:
    """
    计算学生排名、班级汇总表（含总分分数段统计），以及各科目班级详细统计表
    新增语数英总分及排名，进退步计算

    参数:
        df: 原始成绩DataFrame，必须包含'班级'、'姓名'及各科成绩列
        params: 参数字典，包含：
            - subjects: 科目列表
            - full_marks: 各科满分字典
            - passing_score: 及格线字典
            - excellent_score: 优秀线字典
            - calc_total: 是否计算总分
            - weights: 各科权重字典（若calc_total为True）
            - rank_method: 排名方法（'min', 'dense', 'average'等）
            - exclude_missing: 是否排除缺考（NaN）计算平均分和率
        history_df: 历史排名DataFrame，包含'姓名'和'上次排名'列
        calc_progress: 是否计算进退步

    返回:
        (student_rank_df, class_summary_df, subject_details_dict)
        student_rank_df: 学生排名表，包含总分、年级排名、班级排名、
                         语数英总分、语数英排名、上次排名、进退步（如果启用）
        class_summary_df: 班级汇总表（含总分分数段）
        subject_details_dict: 各科目班级详细统计表
    """
    # 复制数据，避免修改原始
    df = df.copy()

    # 确保班级列为字符串类型，方便分组
    df['班级'] = df['班级'].astype(str)

    # 确保各科成绩为数值类型（若无法转换则设为NaN）
    for subj in params['subjects']:
        df[subj] = pd.to_numeric(df[subj], errors='coerce')

    # 参数解包
    subjects = params['subjects']
    full_marks = params['full_marks']
    passing = params['passing_score']
    excellent = params['excellent_score']
    calc_total = params['calc_total']
    weights = params.get('weights', {})
    rank_method = params.get('rank_method', 'min')
    exclude_missing = params.get('exclude_missing', True)
    
    # 添加总分科目
    if calc_total:
        # 计算总分满分线（所有科目满分之和）
        total_full_mark = sum(full_marks.get(subj, 100) for subj in subjects)
        # 将总分添加到科目列表
        subjects_with_total = subjects + ['总分']
        # 更新满分线字典，添加总分满分线
        full_marks['总分'] = total_full_mark
        # 总分没有优生线和及格线
        passing['总分'] = 0
        excellent['总分'] = 0
    else:
        subjects_with_total = subjects

    # ------------------------------------------------------------------
    # 1. 计算总分（如果需要）
    # ------------------------------------------------------------------
    if calc_total:
        # 确保所有科目都在weights中，若缺失则默认为1
        for subj in subjects:
            if subj not in weights:
                weights[subj] = 1.0

        # 计算加权总分：将缺考(NaN)视为0分，乘以权重后求和
        total_score = pd.Series(0.0, index=df.index)
        for subj in subjects:
            subject_score = df[subj].fillna(0) * weights[subj]
            total_score += subject_score

        df['总分'] = total_score
        # 确保总分是数值类型
        df['总分'] = pd.to_numeric(df['总分'], errors='coerce')
    else:
        # 若不计算总分，后续排名跳过
        pass

    # ------------------------------------------------------------------
    # 2. 计算语数英总分及排名
    # ------------------------------------------------------------------
    # 识别语数英科目（基于常见名称，可扩展）
    yuwen_subj = None
    shuxue_subj = None
    yingyu_subj = None
    for subj in subjects:
        cleaned = subj.strip().lower()
        if cleaned in ['语文', 'chinese', 'yuwen']:
            yuwen_subj = subj
        elif cleaned in ['数学', 'math', 'mathematics', 'shuxue']:
            shuxue_subj = subj
        elif cleaned in ['英语', 'english', 'yingyu']:
            yingyu_subj = subj

    # 计算语数英总分（缺考计0分）
    yuwen_score = df[yuwen_subj].fillna(0) if yuwen_subj else 0
    shuxue_score = df[shuxue_subj].fillna(0) if shuxue_subj else 0
    yingyu_score = df[yingyu_subj].fillna(0) if yingyu_subj else 0
    df['语数英总分'] = yuwen_score + shuxue_score + yingyu_score

    # 语数英年级排名
    if df['语数英总分'].notna().any():
        df['语数英排名'] = df['语数英总分'].rank(method=rank_method, ascending=False, na_option='bottom').astype('Int64')
    else:
        df['语数英排名'] = pd.NA

    # ------------------------------------------------------------------
    # 3. 学生排名（原有总分排名）
    # ------------------------------------------------------------------
    if calc_total:
        # 年级排名（仅对有效总分排名，若总分全为NaN则排名为NaN）
        if df['总分'].notna().any():
            df['年级排名'] = df['总分'].rank(method=rank_method, ascending=False, na_option='bottom').astype('Int64')
        else:
            df['年级排名'] = pd.NA

        # 班级排名（按班级分组）
        df['班级排名'] = df.groupby('班级')['总分'].rank(method=rank_method, ascending=False, na_option='bottom').astype('Int64')
    else:
        df['年级排名'] = pd.NA
        df['班级排名'] = pd.NA

    student_rank_df = df

    # ------------------------------------------------------------------
    # 4. 进退步计算（如果启用）
    # ------------------------------------------------------------------
    if calc_progress and history_df is not None and not history_df.empty:
        # 合并历史排名
        history_df_clean = history_df[['姓名', '上次排名']].copy()
        history_df_clean['姓名'] = history_df_clean['姓名'].astype(str).str.strip()
        df['姓名_clean'] = df['姓名'].astype(str).str.strip()
        df = df.merge(history_df_clean, left_on='姓名_clean', right_on='姓名', how='left', suffixes=('', '_hist'))
        # 删除多余列
        df.drop(columns=['姓名_clean', '姓名_hist'], inplace=True, errors='ignore')
        # 计算进退步：本次排名 - 上次排名（负数表示进步）
        df['进退步'] = df['年级排名'] - df['上次排名']
    else:
        df['上次排名'] = pd.NA
        df['进退步'] = pd.NA

    # 更新 student_rank_df
    student_rank_df = df

    # ------------------------------------------------------------------
    # 5. 班级汇总表（基础：班级、总分平均分、班级排名）
    # ------------------------------------------------------------------
    classes = sorted(df['班级'].unique())
    class_stats = []

    for cls in classes:
        class_data = df[df['班级'] == cls]
        class_stat = {'班级': cls}

        if calc_total and '总分' in df.columns:
            total_scores = pd.to_numeric(class_data['总分'], errors='coerce')
            if exclude_missing:
                valid_total = total_scores.dropna()
            else:
                valid_total = total_scores.fillna(0)
            if len(valid_total) > 0:
                class_stat['总分平均分'] = round(valid_total.mean(), 2)
            else:
                class_stat['总分平均分'] = np.nan
        else:
            class_stat['总分平均分'] = np.nan

        class_stats.append(class_stat)

    class_summary_df = pd.DataFrame(class_stats)

    if '总分平均分' in class_summary_df.columns:
        class_summary_df['班级排名'] = class_summary_df['总分平均分'].rank(method='min', ascending=False, na_option='bottom').astype('Int64')
    else:
        class_summary_df['班级排名'] = pd.NA

    # ------------------------------------------------------------------
    # 6. 总分分数段统计（仅当计算总分时进行）
    # ------------------------------------------------------------------
    if calc_total and '总分' in df.columns:
        total_scores = df['总分']
        total_min, total_max = total_scores.min(), total_scores.max()

        if total_min == total_max:
            bins = [total_min, total_max]
            if total_min.is_integer():
                labels = [f"{int(total_min)}"]
            else:
                labels = [f"{total_min:.1f}"]
        else:
            # 使用半百区间方法计算分数段
            # 找到最高分所在的半百区间
            max_half_hundred = ((int(total_max) + 49) // 50) * 50
            # 找到最低分所在的半百区间
            min_half_hundred = (int(total_min) // 50) * 50
            # 取最大与最小值
            range_start = min_half_hundred
            range_end = max_half_hundred
            # 计算分数段间隔
            step = (range_end - range_start) / 10
            # 确保step至少为1
            step = max(step, 1)
            # 生成bins
            bins = np.arange(range_start, range_end + step, step)
            # 确保bins包含最大值
            if bins[-1] < total_max:
                bins = np.append(bins, range_end + step)
            # 生成标签
            if step.is_integer():
                labels = [f"{int(bins[i])}-{int(bins[i+1])}" for i in range(len(bins)-1)]
            else:
                labels = [f"{bins[i]:.1f}-{bins[i+1]:.1f}" for i in range(len(bins)-1)]

        seg_data = []
        for cls in classes:
            class_scores = df[df['班级'] == cls]['总分']
            cuts = pd.cut(class_scores, bins=bins, labels=labels, right=True, include_lowest=True)
            seg_counts = cuts.value_counts().reindex(labels, fill_value=0)
            row = {'班级': cls}
            for label in labels:
                row[label] = seg_counts[label]
            seg_data.append(row)

        class_seg_df = pd.DataFrame(seg_data)
        class_summary_df = pd.merge(class_summary_df, class_seg_df, on='班级', how='left')

        grade_cuts = pd.cut(total_scores, bins=bins, labels=labels, right=True, include_lowest=True)
        grade_seg_counts = grade_cuts.value_counts().reindex(labels, fill_value=0)
        grade_seg_dict = {label: grade_seg_counts[label] for label in labels}
    else:
        labels = []
        grade_seg_dict = {}

    # ------------------------------------------------------------------
    # 7. 单科班级详细统计表生成
    # ------------------------------------------------------------------
    subject_details_dict = {}
    # 批量计算所有科目
    for subj in subjects:
        subject_df = _calculate_subject_class_details(df, subj, params)
        subject_details_dict[subj] = subject_df
    
    # 添加总分分析（如果计算总分）
    if calc_total and '总分' in df.columns:
        # 创建总分参数
        total_params = params.copy()
        total_params['subjects'] = ['总分']
        # 计算总分班级分析
        total_df = _calculate_subject_class_details(df, '总分', total_params)
        subject_details_dict['总分'] = total_df

    # ------------------------------------------------------------------
    # 8. 学生排名表按年级排名升序排序
    # ------------------------------------------------------------------
    if calc_total and '年级排名' in student_rank_df.columns:
        student_rank_df = student_rank_df.sort_values('年级排名', ascending=True, na_position='last')
    elif '总分' in student_rank_df.columns:
        student_rank_df = student_rank_df.sort_values('总分', ascending=False, na_position='last')

    # ------------------------------------------------------------------
    # 9. 添加年段行到班级汇总表和单科排名表
    # ------------------------------------------------------------------
    # 班级汇总表年段行
    grade_row = {'班级': '年段'}
    if calc_total:
        all_scores = pd.to_numeric(df['总分'], errors='coerce')
        if exclude_missing:
            valid_all = all_scores.dropna()
        else:
            valid_all = all_scores.fillna(0)
        if len(valid_all) > 0:
            grade_row['总分平均分'] = round(valid_all.mean(), 2)
        else:
            grade_row['总分平均分'] = np.nan
        for label in labels:
            grade_row[label] = grade_seg_dict.get(label, 0)
    else:
        grade_row['总分平均分'] = np.nan
    grade_row['班级排名'] = pd.NA

    class_summary_df = pd.concat([class_summary_df, pd.DataFrame([grade_row])], ignore_index=True)

    # 单科排名表年段行（每个科目都添加）
    for subj, df_subj in subject_details_dict.items():
        grade_row_subj = _calculate_grade_subject_details(df, subj, params)
        subject_details_dict[subj] = pd.concat([df_subj, pd.DataFrame([grade_row_subj])], ignore_index=True)

    return student_rank_df, class_summary_df, subject_details_dict

def _calculate_subject_class_details(df: pd.DataFrame, subject: str, params: Dict[str, Any]) -> pd.DataFrame:
    """计算单个科目的班级详细统计表（含分数段分布）"""


    full = params['full_marks'].get(subject, 100)
    num_segments = 10
    bins = np.linspace(0, full, num_segments + 1)
    step = bins[1] - bins[0]
    if step.is_integer():
        labels = [f"{int(bins[i])}-{int(bins[i+1])}" for i in range(num_segments)]
    else:
        labels = [f"{bins[i]:.1f}-{bins[i+1]:.1f}" for i in range(num_segments)]

    exclude_missing = params.get('exclude_missing', True)
    passing_score = params['passing_score'].get(subject, 60)
    excellent_score = params['excellent_score'].get(subject, 80)

    # 确保成绩列为数值类型
    scores = pd.to_numeric(df[subject], errors='coerce')
    df_copy = df.copy()
    df_copy['score'] = scores

    # 按班级分组计算
    classes = sorted(df['班级'].unique())
    rows = []

    # 向量化计算每个班级
    for cls in classes:
        class_data = df_copy[df_copy['班级'] == cls]
        class_scores = class_data['score']

        if exclude_missing:
            valid_scores = class_scores.dropna()
            ref_count = len(valid_scores)
        else:
            valid_scores = class_scores.fillna(0)
            ref_count = len(class_data)

        if ref_count == 0:
            row = {
                '班级': cls,
                '参考人数': 0,
                '及格率': np.nan,
                '优生率': np.nan,
                '平均分': np.nan,
                '最高分': np.nan,
                '最低分': np.nan,
                '任课教师': ''
            }
            for label in labels:
                row[label] = 0
        else:
            # 向量化计算分数段
            cuts = pd.cut(valid_scores, bins=bins, labels=labels, right=True, include_lowest=True)
            seg_counts = cuts.value_counts().reindex(labels, fill_value=0)

            # 向量化计算统计指标
            pass_count = (valid_scores >= passing_score).sum()
            pass_rate = pass_count / ref_count if ref_count > 0 else np.nan
            excellent_count = (valid_scores >= excellent_score).sum()
            excellent_rate = excellent_count / ref_count if ref_count > 0 else np.nan

            avg_score = valid_scores.mean()
            max_score = valid_scores.max()
            min_score = valid_scores.min()

            row = {
                '班级': cls,
                '参考人数': ref_count,
                '及格率': round(pass_rate, 4),
                '优生率': round(excellent_rate, 4),
                '平均分': round(avg_score, 2),
                '最高分': max_score,
                '最低分': min_score,
                '任课教师': ''
            }
            for label in labels:
                row[label] = seg_counts[label]

        rows.append(row)

    columns = ['班级', '参考人数'] + labels + ['及格率', '优生率', '平均分', '最高分', '最低分', '任课教师']
    result_df = pd.DataFrame(rows, columns=columns)

    # 批量转换分数段列为整数类型
    for label in labels:
        result_df[label] = result_df[label].astype('Int64')

    return result_df

def _calculate_grade_subject_details(df: pd.DataFrame, subject: str, params: Dict[str, Any]) -> dict:
    """计算年级（年段）的单科统计，返回一行字典"""


    full = params['full_marks'].get(subject, 100)
    num_segments = 10
    bins = np.linspace(0, full, num_segments + 1)
    step = bins[1] - bins[0]
    if step.is_integer():
        labels = [f"{int(bins[i])}-{int(bins[i+1])}" for i in range(num_segments)]
    else:
        labels = [f"{bins[i]:.1f}-{bins[i+1]:.1f}" for i in range(num_segments)]

    exclude_missing = params.get('exclude_missing', True)
    passing_score = params['passing_score'].get(subject, 60)
    excellent_score = params['excellent_score'].get(subject, 80)

    scores = pd.to_numeric(df[subject], errors='coerce')

    if exclude_missing:
        valid_scores = scores.dropna()
        ref_count = len(valid_scores)
    else:
        valid_scores = scores.fillna(0)
        ref_count = len(df)

    if ref_count == 0:
        row = {
            '班级': '年段',
            '参考人数': 0,
            '及格率': np.nan,
            '优生率': np.nan,
            '平均分': np.nan,
            '最高分': np.nan,
            '最低分': np.nan,
            '任课教师': ''
        }
        for label in labels:
            row[label] = 0
    else:
        cuts = pd.cut(valid_scores, bins=bins, labels=labels, right=True, include_lowest=True)
        seg_counts = cuts.value_counts().reindex(labels, fill_value=0)

        pass_count = (valid_scores >= passing_score).sum()
        pass_rate = pass_count / ref_count if ref_count > 0 else np.nan
        excellent_count = (valid_scores >= excellent_score).sum()
        excellent_rate = excellent_count / ref_count if ref_count > 0 else np.nan

        avg_score = valid_scores.mean()
        max_score = valid_scores.max()
        min_score = valid_scores.min()

        row = {
            '班级': '年段',
            '参考人数': ref_count,
            '及格率': round(pass_rate, 4),
            '优生率': round(excellent_rate, 4),
            '平均分': round(avg_score, 2),
            '最高分': max_score,
            '最低分': min_score,
            '任课教师': ''
        }
        for label in labels:
            val = seg_counts[label]
            row[label] = int(val) if not pd.isna(val) else 0

    return row