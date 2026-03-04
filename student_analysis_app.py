#!/usr/bin/env python3
"""
学生成绩分析系统 - Streamlit自动化仪表板
上传Excel文件，自动生成分析仪表板
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import os
from datetime import datetime

# ==================== 配置 ====================
st.set_page_config(
    page_title="学生成绩分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 样式定制
st.markdown("""
<style>
    .main-header {
        font-size: 2rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 1rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .stAlert {
        padding: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 常量配置 ====================

# 科目满分配置
MAX_SCORES = {
    '语文': 150,
    '数学': 150,
    '英语': 150,
    '首选科目': 100,
    '四选二科目': 100,
    '总分': 750
}

# 默认分数线配置
DEFAULT_THRESHOLDS = {
    '总分': {'985': 600, '211': 550, '一本': 500},
    'default': {'985': 90, '211': 80, '一本': 60}  # 百分比
}

# 进步/退步阈值（分数变化小于此值视为持平）
PROGRESS_THRESHOLD = 3

# 允许访问的根目录（安全限制）
ALLOWED_ROOT_PATHS = [
    "E:\\op\\op4\\data",
    "E:\\op\\op4\\data_raw",
    "E:\\op\\op4\\data_raw_progress",
    ".",
    ""
]

# ==================== 工具函数 ====================

def validate_folder_path(folder_path: str) -> tuple[bool, str]:
    """验证文件夹路径的安全性
    
    Args:
        folder_path: 用户输入的文件夹路径
        
    Returns:
        (is_valid, error_message): 是否有效及错误信息
    """
    if not folder_path:
        return False, "路径不能为空"
    
    # 转换为绝对路径
    try:
        abs_path = os.path.abspath(folder_path)
    except Exception:
        return False, "路径格式无效"
    
    # 检查路径是否包含危险字符
    dangerous_chars = ['..', '~', '$', '`', ';', '|', '&']
    for char in dangerous_chars:
        if char in folder_path:
            return False, f"路径包含不安全字符: {char}"
    
    # 检查是否为绝对路径，在允许列表中
    is_allowed = False
    for allowed in ALLOWED_ROOT_PATHS:
        if allowed and abs_path.startswith(os.path.abspath(allowed)):
            is_allowed = True
            break
    
    # 如果不在允许列表中，但路径存在且可访问，也允许
    if not is_allowed and os.path.isdir(folder_path):
        # 检查路径深度，防止越界访问
        path_parts = abs_path.split(os.sep)
        if len(path_parts) > 4:  # 限制路径深度
            return False, "路径深度超出允许范围"
        is_allowed = True
    
    if not is_allowed:
        return False, f"路径不在允许范围内。建议使用: {', '.join([p for p in ALLOWED_ROOT_PATHS if p])}"
    
    return True, ""

# ==================== 数据处理模块 ====================

def safe_col(columns, names):
    """智能列匹配 - 返回第一个匹配的列名"""
    cols = [str(c) for c in columns]
    for nm in names:
        for c in cols:
            if nm.lower() == c.lower():
                return c
    # 模糊匹配
    for nm in names:
        for c in cols:
            if nm.lower() in c.lower():
                return c
    return None

def locate(cols, targets):
    """定位包含目标关键字的列"""
    for c in cols:
        cl = str(c)
        for t in targets:
            if t and t in cl:
                return c
    return None

def parse_excel_file(file, exam_name: str = None) -> tuple:
    """解析单个Excel文件
    
    Returns:
        tuple: (DataFrame, error_message)
    """
    try:
        df = pd.read_excel(file)
        if exam_name is None:
            exam_name = os.path.splitext(file.name)[0]
        
        cols = df.columns.tolist()
        
        # 智能识别列
        id_col = safe_col(cols, ['学号', 'id', 'student_id', 'stu_id'])
        name_col = safe_col(cols, ['姓名', 'name'])
        class_col = safe_col(cols, ['班级', 'class'])
        exam_subj_col = safe_col(cols, ['选考科目', '科目'])
        foreign_col = safe_col(cols, ['外语类型', 'language'])
        total_col = safe_col(cols, ['总分', 'score', 'total'])
        class_rank_col = safe_col(cols, ['班级排名', 'class_rank'])
        grade_rank_col = safe_col(cols, ['年级排名', 'grade_rank', 'grade'])
        
        # 科目列
        you = locate(cols, ['语文'])
        math = locate(cols, ['数学'])
        eng = locate(cols, ['英语'])
        ph_hist = locate(cols, ['物理(历史)', '物理', '历史'])
        chem_poli = locate(cols, ['化学(政治)', '化学', '政治'])
        bio_geo = locate(cols, ['生物(地理)', '生物', '地理'])
        chem_poli_f = locate(cols, ['化学(政治)赋分', '赋分'])
        bio_geo_f = locate(cols, ['生物(地理)赋分', '赋分'])
        
        if id_col is None:
            return None, f"无法识别学号列: {file.name}"
        
        records = []
        for idx, row in df.iterrows():
            rec = {
                '考试名称': exam_name,
            }
            
            # 保留所有原始列
            for col in df.columns:
                rec[col] = row.get(col)
            
            # 标准化列名
            if id_col and id_col in rec:
                rec['学号'] = rec.pop(id_col)
            if name_col and name_col in rec:
                rec['姓名'] = rec.pop(name_col)
            if class_col and class_col in rec:
                rec['班级'] = rec.pop(class_col)
            if exam_subj_col and exam_subj_col in rec:
                rec['选考科目'] = rec.pop(exam_subj_col)
            if foreign_col and foreign_col in rec:
                rec['外语类型'] = rec.pop(foreign_col)
            if total_col and total_col in rec:
                rec['总分'] = rec.pop(total_col)
            if class_rank_col and class_rank_col in rec:
                rec['班级排名'] = rec.pop(class_rank_col)
            if grade_rank_col and grade_rank_col in rec:
                rec['年级排名'] = rec.pop(grade_rank_col)
            if you and you in rec:
                rec['语文'] = rec.pop(you)
            if math and math in rec:
                rec['数学'] = rec.pop(math)
            if eng and eng in rec:
                rec['英语'] = rec.pop(eng)
            if ph_hist and ph_hist in rec:
                rec['物理(历史)'] = rec.pop(ph_hist)
            if chem_poli and chem_poli in rec:
                rec['化学(政治)'] = rec.pop(chem_poli)
            if bio_geo and bio_geo in rec:
                rec['生物(地理)'] = rec.pop(bio_geo)
            if chem_poli_f and chem_poli_f in rec:
                rec['化学(政治)赋分'] = rec.pop(chem_poli_f)
            if bio_geo_f and bio_geo_f in rec:
                rec['生物(地理)赋分'] = rec.pop(bio_geo_f)
            
            records.append(rec)
        
        return pd.DataFrame(records), None
    except Exception as e:
        return None, str(e)

def merge_excel_files(files) -> tuple:
    """合并多个Excel文件
    
    Returns:
        tuple: (merged_DataFrame, list_of_errors)
    """
    all_records = []
    errors = []
    
    for file in files:
        df, error = parse_excel_file(file)
        if error:
            errors.append(error)
        elif df is not None:
            all_records.append(df)
    
    if not all_records:
        return None, errors
    
    merged = pd.concat(all_records, ignore_index=True)
    
    # 确保所有必要列存在
    expected_cols = [
        '考试名称', '学号', '姓名', '班级', '选考科目', '外语类型', '总分',
        '班级排名', '年级排名', '语文', '数学', '英语',
        '物理(历史)', '化学(政治)', '生物(地理)',
        '化学(政治)赋分', '生物(地理)赋分'
    ]
    for col in expected_cols:
        if col not in merged.columns:
            merged[col] = np.nan
    
    return merged, errors

def calculate_subject_stats(df):
    """计算各科目统计数据"""
    df = df.copy()
    
    # 检查是否有选考科目列
    has_exam_subject = '选考科目' in df.columns
    
    # 确保选考科目列为字符串类型
    if has_exam_subject:
        df['选考科目'] = df['选考科目'].astype(str)
    
    stats = []
    
    # 语文数学英语
    basic_subjects = ['语文', '数学', '英语']
    for subj in basic_subjects:
        if subj in df.columns:
            data = pd.to_numeric(df[subj], errors='coerce').dropna()
            if len(data) > 0:
                stats.append({
                    '科目': subj,
                    '平均分': round(data.mean(), 2),
                    '最高分': data.max(),
                    '最低分': data.min(),
                    '标准差': round(data.std(), 2),
                    '人数': len(data)
                })
    
    # 首选科目（物理/历史）
    if '物理(历史)' in df.columns:
        if has_exam_subject:
            # 首选物理的学生（包含"物"）
            physics_students = df[df['选考科目'].str.contains('物', na=False)]
            # 首选历史的学生（包含"史"）
            history_students = df[df['选考科目'].str.contains('史', na=False)]
        else:
            physics_students = df[df['物理(历史)'].notna()]
            history_students = pd.DataFrame()
        
        # 物理统计
        if not physics_students.empty:
            physics_data = pd.to_numeric(physics_students['物理(历史)'], errors='coerce').dropna()
            if len(physics_data) > 0:
                stats.append({
                    '科目': '物理(首选)',
                    '平均分': round(physics_data.mean(), 2),
                    '最高分': physics_data.max(),
                    '最低分': physics_data.min(),
                    '标准差': round(physics_data.std(), 2),
                    '人数': len(physics_data)
                })
        
        # 历史统计
        if not history_students.empty:
            history_data = pd.to_numeric(history_students['物理(历史)'], errors='coerce').dropna()
            if len(history_data) > 0:
                stats.append({
                    '科目': '历史(首选)',
                    '平均分': round(history_data.mean(), 2),
                    '最高分': history_data.max(),
                    '最低分': history_data.min(),
                    '标准差': round(history_data.std(), 2),
                    '人数': len(history_data)
                })
    
    # 四选二科目：根据选考科目列动态识别
    if has_exam_subject:
        # 检查各科目是否有学生选择
        has_chem = df['选考科目'].str.contains('化', na=False).any()
        has_bio = df['选考科目'].str.contains('生', na=False).any()
        has_poli = df['选考科目'].str.contains('政', na=False).any()
        has_geo = df['选考科目'].str.contains('地', na=False).any()
        
        # 化学
        if has_chem and '化学(政治)' in df.columns:
            chem_students = df[df['选考科目'].str.contains('化', na=False)]
            chem_data = pd.to_numeric(chem_students['化学(政治)'], errors='coerce').dropna()
            if len(chem_data) > 0:
                stats.append({
                    '科目': '化学',
                    '平均分': round(chem_data.mean(), 2),
                    '最高分': chem_data.max(),
                    '最低分': chem_data.min(),
                    '标准差': round(chem_data.std(), 2),
                    '人数': len(chem_data)
                })
        
        # 生物
        if has_bio and '生物(地理)' in df.columns:
            bio_students = df[df['选考科目'].str.contains('生', na=False)]
            bio_data = pd.to_numeric(bio_students['生物(地理)'], errors='coerce').dropna()
            if len(bio_data) > 0:
                stats.append({
                    '科目': '生物',
                    '平均分': round(bio_data.mean(), 2),
                    '最高分': bio_data.max(),
                    '最低分': bio_data.min(),
                    '标准差': round(bio_data.std(), 2),
                    '人数': len(bio_data)
                })
        
        # 政治
        if has_poli and '化学(政治)' in df.columns:
            poli_students = df[df['选考科目'].str.contains('政', na=False)]
            poli_data = pd.to_numeric(poli_students['化学(政治)'], errors='coerce').dropna()
            if len(poli_data) > 0:
                stats.append({
                    '科目': '政治',
                    '平均分': round(poli_data.mean(), 2),
                    '最高分': poli_data.max(),
                    '最低分': poli_data.min(),
                    '标准差': round(poli_data.std(), 2),
                    '人数': len(poli_data)
                })
        
        # 地理
        if has_geo and '生物(地理)' in df.columns:
            geo_students = df[df['选考科目'].str.contains('地', na=False)]
            geo_data = pd.to_numeric(geo_students['生物(地理)'], errors='coerce').dropna()
            if len(geo_data) > 0:
                stats.append({
                    '科目': '地理',
                    '平均分': round(geo_data.mean(), 2),
                    '最高分': geo_data.max(),
                    '最低分': geo_data.min(),
                    '标准差': round(geo_data.std(), 2),
                    '人数': len(geo_data)
                })
    else:
        # 没有选考科目列，使用原来的方式
        four_choice = {
            '化学(政治)': '化学/四选二',
            '生物(地理)': '生物/四选二',
        }
        for col, subj in four_choice.items():
            if col in df.columns:
                data = pd.to_numeric(df[col], errors='coerce').dropna()
                if len(data) > 0:
                    stats.append({
                        '科目': subj,
                        '平均分': round(data.mean(), 2),
                        '最高分': data.max(),
                        '最低分': data.min(),
                        '标准差': round(data.std(), 2),
                        '人数': len(data)
                    })
    
    return pd.DataFrame(stats)

def calculate_class_stats(df):
    """计算各班级统计数据"""
    if '班级' not in df.columns or '总分' not in df.columns:
        return None
    
    df = df.copy()
    df['总分'] = pd.to_numeric(df['总分'], errors='coerce')
    df = df.dropna(subset=['班级', '总分'])
    
    if df.empty:
        return None
    
    stats = df.groupby('班级').agg({
        '总分': ['mean', 'max', 'min', 'count'],
        '学号': 'nunique'
    }).round(2)
    stats.columns = ['平均分', '最高分', '最低分', '参考人数', '实际人数']
    stats = stats.reset_index()
    return stats

def calculate_class_subject_avg(df):
    """计算各班级各科目的平均分
    
    Returns:
        DataFrame with columns: 班级, 语文, 数学, 英语, 物理(首选), 历史(首选), 化学, 生物, 政治, 地理
    """
    if '班级' not in df.columns:
        return None
    
    df = df.copy()
    
    # 确保选考科目列为字符串类型
    if '选考科目' in df.columns:
        df['选考科目'] = df['选考科目'].astype(str)
    
    # 确定存在的科目列
    subject_cols = {}
    
    # 基础科目
    for subj in ['语文', '数学', '英语']:
        if subj in df.columns:
            df[subj] = pd.to_numeric(df[subj], errors='coerce')
            subject_cols[subj] = subj
    
    # 首选科目（物理/历史）
    if '物理(历史)' in df.columns:
        df['物理(历史)'] = pd.to_numeric(df['物理(历史)'], errors='coerce')
        
        if '选考科目' in df.columns:
            # 按选考科目分组计算
            physics_mask = df['选考科目'].str.contains('物', na=False)
            history_mask = df['选考科目'].str.contains('史', na=False)
            
            # 创建物理/历史平均分列
            df['物理(首选)'] = df['物理(历史)'].where(physics_mask)
            df['历史(首选)'] = df['物理(历史)'].where(history_mask)
            
            subject_cols['物理(首选)'] = '物理(首选)'
            subject_cols['历史(首选)'] = '历史(首选)'
        else:
            # 没有选考科目列，全部当作物理(首选)
            df['物理(首选)'] = df['物理(历史)']
            subject_cols['物理(首选)'] = '物理(首选)'
    
    # 四选二科目
    four_two_subjects = [
        ('化学(政治)', '化学'),
        ('生物(地理)', '生物'),
    ]
    
    has_exam_subject = '选考科目' in df.columns
    
    for col, name in four_two_subjects:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
            if has_exam_subject:
                # 根据选考科目确定是哪个科目
                char_map = {'化': '化学', '生': '生物', '政': '政治', '地': '地理'}
                
                for char, subj_name in char_map.items():
                    mask = df['选考科目'].str.contains(char, na=False)
                    if subj_name not in subject_cols:
                        df[subj_name] = df[col].where(mask)
                        subject_cols[subj_name] = subj_name
                    else:
                        # 多个学生选同一科目
                        df[subj_name] = df[subj_name].fillna(df[col].where(mask))
            else:
                # 没有选考科目列，根据列名判断
                if '化学' in col:
                    df['化学'] = df[col]
                    subject_cols['化学'] = '化学'
                elif '生物' in col:
                    df['生物'] = df[col]
                    subject_cols['生物'] = '生物'
    
    # 直接处理剩余的四选二科目列
    if '化学(政治)' in df.columns and '化学' not in subject_cols:
        df['化学'] = pd.to_numeric(df['化学(政治)'], errors='coerce')
        subject_cols['化学'] = '化学'
    
    if '生物(地理)' in df.columns and '生物' not in subject_cols:
        df['生物'] = pd.to_numeric(df['生物(地理)'], errors='coerce')
        subject_cols['生物'] = '生物'
    
    # 按班级分组计算各科平均分
    if not subject_cols:
        return None
    
    # 选择需要统计的列
    cols_to_agg = ['班级'] + list(subject_cols.keys())
    available_cols = [c for c in cols_to_agg if c in df.columns]
    
    if len(available_cols) <= 1:  # 只有班级列
        return None
    
    # 计算各科平均分
    agg_dict = {col: 'mean' for col in available_cols if col != '班级'}
    
    try:
        class_subj_avg = df.groupby('班级')[list(agg_dict.keys())].mean().round(2)
        class_subj_avg = class_subj_avg.reset_index()
        
        # 重命名列
        rename_dict = {}
        for col in class_subj_avg.columns:
            if col in subject_cols.values():
                rename_dict[col] = col
            elif col in subject_cols:
                rename_dict[col] = subject_cols[col]
        
        class_subj_avg = class_subj_avg.rename(columns=rename_dict)
        
        return class_subj_avg
    except Exception as e:
        return None

def calculate_class_thresholds_stats(df, thresholds):
    """计算各班级总分上三线的人数
    
    Args:
        df: 数据框
        thresholds: 字典，包含总分分数线设置
    """
    if '班级' not in df.columns or '总分' not in df.columns:
        return None
    
    df = df.copy()
    df['总分'] = pd.to_numeric(df['总分'], errors='coerce')
    df = df.dropna(subset=['班级', '总分'])
    
    if df.empty:
        return None
    
    # 获取总分分数线
    if '总分' in thresholds:
        t = thresholds['总分']
    else:
        t = {'985': 600, '211': 550, '一本': 500}
    
    # 计算各班上线人数
    stats = df.groupby('班级').apply(
        lambda x: pd.Series({
            '班级总人数': len(x),
            '985人数': (x['总分'] >= t['985']).sum(),
            '985率': round(100 * (x['总分'] >= t['985']).sum() / len(x), 2) if len(x) > 0 else 0,
            '211人数': (x['总分'] >= t['211']).sum(),
            '211率': round(100 * (x['总分'] >= t['211']).sum() / len(x), 2) if len(x) > 0 else 0,
            '一本人数': (x['总分'] >= t['一本']).sum(),
            '一本率': round(100 * (x['总分'] >= t['一本']).sum() / len(x), 2) if len(x) > 0 else 0,
        })
    ).reset_index()
    
    return stats

def calculate_progress_between_exams(df, prev_exam, curr_exam):
    """计算指定两次考试之间的进步/退步分析
    
    Args:
        df: 数据框
        prev_exam: 上次考试名称
        curr_exam: 本次考试名称
    
    Returns:
        DataFrame with progress data between two specific exams
    """
    if '学号' not in df.columns or '考试名称' not in df.columns or '总分' not in df.columns:
        return None
    
    if prev_exam == curr_exam:
        return None
    
    df = df.copy()
    df['总分'] = pd.to_numeric(df['总分'], errors='coerce')
    
    progress_data = []
    
    for student_id in df['学号'].unique():
        student_df = df[df['学号'] == student_id]
        
        prev_data = student_df[student_df['考试名称'] == prev_exam]
        curr_data = student_df[student_df['考试名称'] == curr_exam]
        
        if prev_data.empty or curr_data.empty:
            continue
        
        prev_total = prev_data['总分'].values[0]
        curr_total = curr_data['总分'].values[0]
        
        # 计算排名变化
        prev_rank = prev_data['年级排名'].values[0] if '年级排名' in prev_data.columns and not pd.isna(prev_data['年级排名'].values[0]) else None
        curr_rank = curr_data['年级排名'].values[0] if '年级排名' in curr_data.columns and not pd.isna(curr_data['年级排名'].values[0]) else None
        
        score_diff = curr_total - prev_total
        
        if prev_rank and curr_rank:
            rank_change = prev_rank - curr_rank  # 正值表示进步
        else:
            rank_change = None
        
        name = student_df['姓名'].values[0] if '姓名' in student_df.columns and not student_df['姓名'].isna().all() else None
        class_name = student_df['班级'].values[0] if '班级' in student_df.columns and not student_df['班级'].isna().all() else None
        
        progress_data.append({
            '学号': student_id,
            '姓名': name,
            '班级': class_name,
            '上次考试': prev_exam,
            '本次考试': curr_exam,
            '上次总分': prev_total,
            '本次总分': curr_total,
            '分数变化': score_diff,
            '上次排名': prev_rank,
            '本次排名': curr_rank,
            '排名变化': rank_change,
            '状态': '进步' if score_diff > st.session_state.progress_threshold else ('退步' if score_diff < -st.session_state.progress_threshold else '持平')
        })
    
    return pd.DataFrame(progress_data)

def calculate_grade_rates(df, thresholds):
    """计算各科985率、211率、一本率
    
    Args:
        df: 数据框
        thresholds: 字典，包含总分和单科的分数线设置
    """
    df = df.copy()
    
    # 确保选考科目列为字符串类型
    if '选考科目' in df.columns:
        df['选考科目'] = df['选考科目'].astype(str)
    
    # 检查是否有选考科目列，用于区分四选二科目
    has_exam_subject = '选考科目' in df.columns
    
    # 科目列对应关系：原始列名 -> 显示名称
    subject_mapping = {
        '语文': ('语文', MAX_SCORES['语文']),
        '数学': ('数学', MAX_SCORES['语文']),
        '英语': ('英语', MAX_SCORES['语文']),
    }
    
    results = []
    
    # 先计算总分的三率
    if '总分' in df.columns and '总分' in thresholds:
        data = pd.to_numeric(df['总分'], errors='coerce').dropna()
        if len(data) > 0:
            t = thresholds['总分']
            level_985 = (data >= t['985']).sum()
            level_211 = ((data >= t['211']) & (data < t['985'])).sum()
            level_yiben = ((data >= t['一本']) & (data < t['211'])).sum()
            below_yiben = (data < t['一本']).sum()
            results.append({
                '科目': '总分',
                '满分': MAX_SCORES['总分'],
                '总人数': len(data),
                '985人数': level_985,
                '985率': round(100 * level_985 / len(data), 2),
                '211人数': level_211,
                '211率': round(100 * level_211 / len(data), 2),
                '一本人数': level_yiben,
                '一本率': round(100 * level_yiben / len(data), 2),
                '一本以下人数': below_yiben,
                '一本以下率': round(100 * below_yiben / len(data), 2)
            })
    
    # 计算语文数学英语的三率
    for col, (subj, max_score) in subject_mapping.items():
        if col in df.columns:
            data = pd.to_numeric(df[col], errors='coerce').dropna()
            if len(data) > 0:
                # 获取对应科目的分数线
                if col in thresholds:
                    t = thresholds[col]
                elif '单科' in thresholds:
                    t = thresholds['单科']
                else:
                    t = thresholds.get('default', {'985': 120, '211': 100, '一本': 90})
                
                level_985 = (data >= t['985']).sum()
                level_211 = ((data >= t['211']) & (data < t['985'])).sum()
                level_yiben = ((data >= t['一本']) & (data < t['211'])).sum()
                below_yiben = (data < t['一本']).sum()
                
                results.append({
                    '科目': subj,
                    '满分': max_score,
                    '总人数': len(data),
                    '985人数': level_985,
                    '985率': round(100 * level_985 / len(data), 2),
                    '211人数': level_211,
                    '211率': round(100 * level_211 / len(data), 2),
                    '一本人数': level_yiben,
                    '一本率': round(100 * level_yiben / len(data), 2),
                    '一本以下人数': below_yiben,
                    '一本以下率': round(100 * below_yiben / len(data), 2)
                })
    
    # 处理首选科目（物理/历史）
    if '物理(历史)' in df.columns:
        # 根据选考科目判断是物理还是历史
        if has_exam_subject:
            # 首选物理的学生（包含"物"）
            physics_students = df[df['选考科目'].str.contains('物', na=False)]
            # 首选历史的学生（包含"史"）
            history_students = df[df['选考科目'].str.contains('史', na=False)]
        else:
            physics_students = df[df['物理(历史)'].notna()]
            history_students = pd.DataFrame()
        
        # 物理三率
        if not physics_students.empty:
            physics_data = pd.to_numeric(physics_students['物理(历史)'], errors='coerce').dropna()
            if len(physics_data) > 0:
                t = thresholds.get('物理', thresholds.get('物理(历史)', thresholds.get('单科', thresholds.get('default'))))
                level_985 = (physics_data >= t['985']).sum()
                level_211 = ((physics_data >= t['211']) & (physics_data < t['985'])).sum()
                level_yiben = ((physics_data >= t['一本']) & (physics_data < t['211'])).sum()
                below_yiben = (physics_data < t['一本']).sum()
                results.append({
                    '科目': '物理(首选)',
                    '满分': 100,
                    '总人数': len(physics_data),
                    '985人数': level_985,
                    '985率': round(100 * level_985 / len(physics_data), 2),
                    '211人数': level_211,
                    '211率': round(100 * level_211 / len(physics_data), 2),
                    '一本人数': level_yiben,
                    '一本率': round(100 * level_yiben / len(physics_data), 2),
                    '一本以下人数': below_yiben,
                    '一本以下率': round(100 * below_yiben / len(physics_data), 2)
                })
        
        # 历史三率
        if not history_students.empty:
            history_data = pd.to_numeric(history_students['物理(历史)'], errors='coerce').dropna()
            if len(history_data) > 0:
                t = thresholds.get('历史', thresholds.get('物理(历史)', thresholds.get('单科', thresholds.get('default'))))
                level_985 = (history_data >= t['985']).sum()
                level_211 = ((history_data >= t['211']) & (history_data < t['985'])).sum()
                level_yiben = ((history_data >= t['一本']) & (history_data < t['211'])).sum()
                below_yiben = (history_data < t['一本']).sum()
                results.append({
                    '科目': '历史(首选)',
                    '满分': 100,
                    '总人数': len(history_data),
                    '985人数': level_985,
                    '985率': round(100 * level_985 / len(history_data), 2),
                    '211人数': level_211,
                    '211率': round(100 * level_211 / len(history_data), 2),
                    '一本人数': level_yiben,
                    '一本率': round(100 * level_yiben / len(history_data), 2),
                    '一本以下人数': below_yiben,
                    '一本以下率': round(100 * below_yiben / len(history_data), 2)
                })
    
    # 处理四选二科目（化学/生物/政治/地理）
    # 根据选考科目动态识别：不再按固定组合，而是按实际选择科目
    if has_exam_subject:
        # 检查各科目是否有学生选择（动态识别）
        has_chem = df['选考科目'].str.contains('化', na=False).any()
        has_bio = df['选考科目'].str.contains('生', na=False).any()
        has_poli = df['选考科目'].str.contains('政', na=False).any()
        has_geo = df['选考科目'].str.contains('地', na=False).any()
        
        # 化学三率
        if has_chem and '化学(政治)' in df.columns:
            chem_students = df[df['选考科目'].str.contains('化', na=False)]
            chem_data = pd.to_numeric(chem_students['化学(政治)'], errors='coerce').dropna()
            if len(chem_data) > 0:
                t = thresholds.get('化学', thresholds.get('单科', thresholds.get('default')))
                level_985 = (chem_data >= t['985']).sum()
                level_211 = ((chem_data >= t['211']) & (chem_data < t['985'])).sum()
                level_yiben = ((chem_data >= t['一本']) & (chem_data < t['211'])).sum()
                below_yiben = (chem_data < t['一本']).sum()
                results.append({
                    '科目': '化学',
                    '满分': 100,
                    '总人数': len(chem_data),
                    '985人数': level_985,
                    '985率': round(100 * level_985 / len(chem_data), 2),
                    '211人数': level_211,
                    '211率': round(100 * level_211 / len(chem_data), 2),
                    '一本人数': level_yiben,
                    '一本率': round(100 * level_yiben / len(chem_data), 2),
                    '一本以下人数': below_yiben,
                    '一本以下率': round(100 * below_yiben / len(chem_data), 2)
                })
        
        # 生物三率
        if has_bio and '生物(地理)' in df.columns:
            bio_students = df[df['选考科目'].str.contains('生', na=False)]
            bio_data = pd.to_numeric(bio_students['生物(地理)'], errors='coerce').dropna()
            if len(bio_data) > 0:
                t = thresholds.get('生物', thresholds.get('单科', thresholds.get('default')))
                level_985 = (bio_data >= t['985']).sum()
                level_211 = ((bio_data >= t['211']) & (bio_data < t['985'])).sum()
                level_yiben = ((bio_data >= t['一本']) & (bio_data < t['211'])).sum()
                below_yiben = (bio_data < t['一本']).sum()
                results.append({
                    '科目': '生物',
                    '满分': 100,
                    '总人数': len(bio_data),
                    '985人数': level_985,
                    '985率': round(100 * level_985 / len(bio_data), 2),
                    '211人数': level_211,
                    '211率': round(100 * level_211 / len(bio_data), 2),
                    '一本人数': level_yiben,
                    '一本率': round(100 * level_yiben / len(bio_data), 2),
                    '一本以下人数': below_yiben,
                    '一本以下率': round(100 * below_yiben / len(bio_data), 2)
                })
        
        # 政治三率
        if has_poli and '化学(政治)' in df.columns:
            poli_students = df[df['选考科目'].str.contains('政', na=False)]
            poli_data = pd.to_numeric(poli_students['化学(政治)'], errors='coerce').dropna()
            if len(poli_data) > 0:
                t = thresholds.get('政治', thresholds.get('单科', thresholds.get('default')))
                level_985 = (poli_data >= t['985']).sum()
                level_211 = ((poli_data >= t['211']) & (poli_data < t['985'])).sum()
                level_yiben = ((poli_data >= t['一本']) & (poli_data < t['211'])).sum()
                below_yiben = (poli_data < t['一本']).sum()
                results.append({
                    '科目': '政治',
                    '满分': 100,
                    '总人数': len(poli_data),
                    '985人数': level_985,
                    '985率': round(100 * level_985 / len(poli_data), 2),
                    '211人数': level_211,
                    '211率': round(100 * level_211 / len(poli_data), 2),
                    '一本人数': level_yiben,
                    '一本率': round(100 * level_yiben / len(poli_data), 2),
                    '一本以下人数': below_yiben,
                    '一本以下率': round(100 * below_yiben / len(poli_data), 2)
                })
        
        # 地理三率
        if has_geo and '生物(地理)' in df.columns:
            geo_students = df[df['选考科目'].str.contains('地', na=False)]
            geo_data = pd.to_numeric(geo_students['生物(地理)'], errors='coerce').dropna()
            if len(geo_data) > 0:
                t = thresholds.get('地理', thresholds.get('单科', thresholds.get('default')))
                level_985 = (geo_data >= t['985']).sum()
                level_211 = ((geo_data >= t['211']) & (geo_data < t['985'])).sum()
                level_yiben = ((geo_data >= t['一本']) & (geo_data < t['211'])).sum()
                below_yiben = (geo_data < t['一本']).sum()
                results.append({
                    '科目': '地理',
                    '满分': 100,
                    '总人数': len(geo_data),
                    '985人数': level_985,
                    '985率': round(100 * level_985 / len(geo_data), 2),
                    '211人数': level_211,
                    '211率': round(100 * level_211 / len(geo_data), 2),
                    '一本人数': level_yiben,
                    '一本率': round(100 * level_yiben / len(geo_data), 2),
                    '一本以下人数': below_yiben,
                    '一本以下率': round(100 * below_yiben / len(geo_data), 2)
                })
    else:
        # 没有选考科目列，使用原来的方式
        subject_mapping_2 = {
            '化学(政治)': ('化学/四选二', 100),
            '生物(地理)': ('生物/四选二', 100),
        }
        for col, (subj, max_score) in subject_mapping_2.items():
            if col in df.columns:
                data = pd.to_numeric(df[col], errors='coerce').dropna()
                if len(data) > 0:
                    if col in thresholds:
                        t = thresholds[col]
                    elif '单科' in thresholds:
                        t = thresholds['单科']
                    else:
                        t = thresholds.get('default', {'985': 90, '211': 80, '一本': 60})
                    
                    level_985 = (data >= t['985']).sum()
                    level_211 = ((data >= t['211']) & (data < t['985'])).sum()
                    level_yiben = ((data >= t['一本']) & (data < t['211'])).sum()
                    below_yiben = (data < t['一本']).sum()
                    
                    results.append({
                        '科目': subj,
                        '满分': max_score,
                        '总人数': len(data),
                        '985人数': level_985,
                        '985率': round(100 * level_985 / len(data), 2),
                        '211人数': level_211,
                        '211率': round(100 * level_211 / len(data), 2),
                        '一本人数': level_yiben,
                        '一本率': round(100 * level_yiben / len(data), 2),
                        '一本以下人数': below_yiben,
                        '一本以下率': round(100 * below_yiben / len(data), 2)
                    })
    
    return pd.DataFrame(results)

# ==================== 可视化模块 ====================

def plot_subject_comparison(stats_df):
    """科目对比柱状图"""
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=stats_df['科目'],
        y=stats_df['平均分'],
        name='平均分',
        marker_color='#1f77b4',
        text=stats_df['平均分'],
        textposition='outside'
    ))
    fig.add_trace(go.Bar(
        x=stats_df['科目'],
        y=stats_df['最高分'],
        name='最高分',
        marker_color='#2ca02c',
        text=stats_df['最高分'],
        textposition='outside'
    ))
    fig.add_trace(go.Bar(
        x=stats_df['科目'],
        y=stats_df['最低分'],
        name='最低分',
        marker_color='#d62728',
        text=stats_df['最低分'],
        textposition='outside'
    ))
    fig.update_layout(
        title='各科目成绩对比',
        xaxis_title='科目',
        yaxis_title='分数',
        barmode='group',
        height=400
    )
    return fig

def plot_class_comparison(class_stats):
    """班级对比图"""
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=class_stats['班级'],
        y=class_stats['平均分'],
        name='班级平均分',
        marker_color='#1f77b4',
        text=class_stats['平均分'],
        textposition='outside'
    ))
    fig.update_layout(
        title='各班级平均分对比',
        xaxis_title='班级',
        yaxis_title='平均分',
        height=400
    )
    return fig

def plot_class_thresholds_comparison(class_thresholds_stats):
    """班级三线人数对比图"""
    if class_thresholds_stats is None or class_thresholds_stats.empty:
        return None
    
    # 计算动态高度：每个班级至少50像素
    num_classes = len(class_thresholds_stats)
    dynamic_height = max(400, num_classes * 50 + 100)
    
    fig = go.Figure()
    
    # 985人数
    fig.add_trace(go.Bar(
        x=class_thresholds_stats['班级'],
        y=class_thresholds_stats['985人数'],
        name='985人数',
        marker_color='#e74c3c',
        text=class_thresholds_stats['985人数'],
        textposition='outside'
    ))
    
    # 211人数
    fig.add_trace(go.Bar(
        x=class_thresholds_stats['班级'],
        y=class_thresholds_stats['211人数'],
        name='211人数',
        marker_color='#f39c12',
        text=class_thresholds_stats['211人数'],
        textposition='outside'
    ))
    
    # 一本人数
    fig.add_trace(go.Bar(
        x=class_thresholds_stats['班级'],
        y=class_thresholds_stats['一本人数'],
        name='一本人数',
        marker_color='#27ae60',
        text=class_thresholds_stats['一本人数'],
        textposition='outside'
    ))
    
    fig.update_layout(
        title='各班级总分上线人数对比',
        xaxis_title='班级',
        yaxis_title='人数',
        barmode='group',
        height=dynamic_height
    )
    return fig

def plot_class_subject_heatmap(class_subject_avg):
    """班级各科平均分热力图"""
    if class_subject_avg is None or class_subject_avg.empty:
        return None
    
    # 准备热力图数据
    df = class_subject_avg.copy()
    
    # 确保有班级列
    if '班级' not in df.columns:
        return None
    
    # 设置班级为索引
    df = df.set_index('班级')
    
    # 转换数值类型
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # 计算动态高度：每个班级至少30像素，加上标题和边距
    num_classes = len(df)
    dynamic_height = max(400, num_classes * 40 + 100)
    
    # 字体大小也根据班级数量调整
    font_size = max(10, min(12, 200 // num_classes)) if num_classes > 0 else 12
    
    # 创建热力图
    fig = go.Figure(data=go.Heatmap(
        z=df.values,
        x=df.columns,
        y=df.index,
        colorscale='RdYlGn',
        text=df.values,
        texttemplate='%{text:.1f}',
        textfont={"size": font_size},
        colorbar=dict(title='平均分')
    ))
    
    fig.update_layout(
        title='各班级各科目平均分热力图',
        xaxis_title='科目',
        yaxis_title='班级',
        height=dynamic_height,
        font=dict(size=font_size)
    )
    
    return fig

def plot_progress_distribution(progress_df):
    """进步/退步分布饼图"""
    if progress_df is None or progress_df.empty:
        return None
    
    status_counts = progress_df['状态'].value_counts()
    
    colors = {'进步': '#2ca02c', '退步': '#d62728', '持平': '#ff7f0e'}
    fig = go.Figure(data=[go.Pie(
        labels=status_counts.index,
        values=status_counts.values,
        marker=dict(colors=[colors.get(s, '#1f77b4') for s in status_counts.index]),
        textinfo='label+percent',
        hole=0.4
    )])
    fig.update_layout(title='进步/退步分布', height=350)
    return fig

def plot_student_trend(df, student_id):
    """学生成绩趋势图"""
    student_df = df[df['学号'] == student_id].copy()
    student_df = student_df.sort_values('考试名称')
    
    if student_df.empty or '总分' not in student_df.columns:
        return None
    
    subjects = ['语文', '数学', '英语', '物理(历史)', '化学(政治)', '生物(地理)']
    
    fig = go.Figure()
    
    # 添加总分趋势线
    fig.add_trace(go.Scatter(
        x=student_df['考试名称'],
        y=pd.to_numeric(student_df['总分'], errors='coerce'),
        mode='lines+markers',
        name='总分',
        line=dict(width=3, color='#1f77b4')
    ))
    
    # 添加各科目趋势线
    for subj in subjects:
        if subj in student_df.columns:
            fig.add_trace(go.Scatter(
                x=student_df['考试名称'],
                y=pd.to_numeric(student_df[subj], errors='coerce'),
                mode='lines+markers',
                name=subj.replace('(历史)', '/历史').replace('(政治)', '/政治').replace('(地理)', '/地理'),
                line=dict(width=1)
            ))
    
    student_name = student_df['姓名'].iloc[0] if '姓名' in student_df.columns else student_id
    fig.update_layout(
        title=f'{student_name} 成绩趋势',
        xaxis_title='考试',
        yaxis_title='分数',
        height=450,
        hovermode='x unified'
    )
    return fig

def plot_score_distribution(df, subject='总分'):
    """分数分布直方图"""
    if subject not in df.columns:
        return None
    
    data = pd.to_numeric(df[subject], errors='coerce').dropna()
    if data.empty:
        return None
    
    fig = go.Figure(data=[go.Histogram(
        x=data,
        nbinsx=20,
        marker_color='#1f77b4',
        opacity=0.75
    )])
    fig.update_layout(
        title=f'{subject}分数分布',
        xaxis_title='分数',
        yaxis_title='人数',
        height=350
    )
    return fig

def plot_grade_rates_heatmap(grade_df):
    """三率热力图 - 985率、211率、一本率"""
    if grade_df is None or grade_df.empty:
        return None
    
    # 准备数据
    subjects = grade_df['科目'].tolist()
    rates = ['985率', '211率', '一本率']
    z_data = []
    
    for rate in rates:
        z_data.append(grade_df[rate].tolist())
    
    # 创建热力图
    fig = go.Figure(data=go.Heatmap(
        z=z_data,
        x=subjects,
        y=rates,
        colorscale=[
            [0, '#ffebee'],      # 浅红 (低)
            [0.3, '#fff3e0'],   # 浅橙
            [0.6, '#e8f5e9'],    # 浅绿
            [1, '#c8e6c9']       # 深绿 (高)
        ],
        text=[[f'{v:.1f}%' for v in row] for row in z_data],
        texttemplate='%{text}',
        showscale=True,
        colorbar=dict(title='百分比')
    ))
    
    fig.update_layout(
        title='各科目985/211/一本比率分布',
        xaxis_title='科目',
        yaxis_title='比率',
        height=300
    )
    return fig

# ==================== 主应用 ====================

def main():
    # 初始化session_state
    if 'progress_threshold' not in st.session_state:
        st.session_state.progress_threshold = PROGRESS_THRESHOLD
    
    # 标题
    st.markdown('<div class="main-header">📊 学生成绩分析系统</div>', unsafe_allow_html=True)
    
    # 侧边栏 - 文件上传
    with st.sidebar:
        st.header("📁 数据上传")
        
        # 上传方式选择 - 默认为文件夹路径（更稳定）
        upload_method = st.radio(
            "选择上传方式",
            ["📂 文件夹路径", "🔼 点击上传"],
            horizontal=True,
            help="如点击上传不可用，请使用文件夹路径方式"
        )
        
        uploaded_files = None
        
        if upload_method == "📂 文件夹路径":
            # 预设常用路径
            default_paths = [
                "E:\\op\\op4\\data_raw_progress",
                "E:\\op\\op4\\data_raw",
                ".",
                ""
            ]
            
            # 文件夹路径方式
            folder_path = st.text_input(
                "📂 输入Excel文件夹路径",
                value="E:\\op\\op4\\data_raw_progress",
                help="输入包含Excel文件的文件夹路径"
            )
            
            # 安全验证
            is_valid, error_msg = validate_folder_path(folder_path)
            
            if folder_path and not is_valid:
                st.error(f"❌ {error_msg}")
            elif folder_path and os.path.isdir(folder_path):
                # 获取文件夹中的Excel文件
                import glob as glob_module
                excel_files = glob_module.glob(os.path.join(folder_path, "*.xlsx")) + \
                              glob_module.glob(os.path.join(folder_path, "*.xls"))
                
                if excel_files:
                    st.success(f"✅ 找到 {len(excel_files)} 个Excel文件")
                    # 显示文件列表
                    with st.expander(f"📋 查看文件列表 ({len(excel_files)}个)"):
                        for f in excel_files:
                            st.text(os.path.basename(f))
                    
                    # 读取文件
                    uploaded_files = []
                    for f in excel_files:
                        try:
                            with open(f, 'rb') as file:
                                from io import BytesIO
                                uploaded_files.append(BytesIO(file.read()))
                                # 设置文件名
                                uploaded_files[-1].name = os.path.basename(f)
                        except Exception as e:
                            st.warning(f"⚠️ 读取失败: {os.path.basename(f)}")
                else:
                    st.error("❌ 文件夹中没有找到Excel文件")
            elif folder_path:
                st.error("❌ 路径无效或不存在，请检查路径是否正确")
        else:
            # 点击上传方式
            st.warning("⚠️ 如果点击上传按钮无响应，请切换到「文件夹路径」模式")
            uploaded_files = st.file_uploader(
                "上传Excel文件",
                type=['xlsx', 'xls'],
                accept_multiple_files=True,
                help="支持xlsx和xls格式，可一次性选择多个文件"
            )
        
        st.divider()
        
        # 分析选项
        st.header("⚙️ 分析选项")
        
        # 分学科分数线设置
        st.subheader("📏 分数线设置")
        
        use_global_thresholds = st.toggle("统一分数线", value=True, help="关闭后可分科目设置")
        
        # 科目满分设置（使用常量）
        
        if use_global_thresholds:
            # 总分分数线设置
            st.markdown("**总分分数线**")
            col1, col2, col3 = st.columns(3)
            with col1:
                score_985 = st.slider("985线", 300, MAX_SCORES['总分'], 600, 10, key="score_985")
            with col2:
                score_211 = st.slider("211线", 300, MAX_SCORES['总分'], 550, 10, key="score_211")
            with col3:
                score_yiben = st.slider("一本线", 300, MAX_SCORES['总分'], 500, 10, key="score_yiben")
            
            thresholds = {
                '总分': {'985': score_985, '211': score_211, '一本': score_yiben},
                'default': {'985': 90, '211': 80, '一本': 60}
            }
            
            st.markdown("---")
            st.markdown("**单科分数线**")
            col1, col2, col3 = st.columns(3)
            with col1:
                excellent_threshold = st.slider("985线", 60, MAX_SCORES['语文'], 120, 5, key="excellent_global")
            with col2:
                good_threshold = st.slider("211线", 60, MAX_SCORES['语文'], 100, 5, key="good_global")
            with col3:
                pass_threshold = st.slider("一本线", 30, MAX_SCORES['语文'], 90, 5, key="pass_global")
            
            thresholds['单科'] = {'985': excellent_threshold, '211': good_threshold, '一本': pass_threshold}
        else:
            st.markdown("**总分分数线**")
            col1, col2, col3 = st.columns(3)
            with col1:
                score_985 = st.slider("985线", 300, MAX_SCORES['总分'], 600, 10, key="total_985")
            with col2:
                score_211 = st.slider("211线", 300, MAX_SCORES['总分'], 550, 10, key="total_211")
            with col3:
                score_yiben = st.slider("一本线", 300, MAX_SCORES['总分'], 500, 10, key="total_yiben")
            
            thresholds = {
                '总分': {'985': score_985, '211': score_211, '一本': score_yiben},
                'default': {'985': 120, '211': 100, '一本': 90}
            }
            
            st.markdown("---")
            st.markdown("**分学科设置：**")
            
            # 语数外满分MAX_SCORES['语文']
            st.markdown("📚 语数外 (满分MAX_SCORES['语文'])")
            col1, col2, col3 = st.columns(3)
            with col1:
                thresholds['语文'] = {'985': st.slider("语文985线", 60, MAX_SCORES['语文'], 120, 5, key="yuwen_985")}
            with col2:
                thresholds['语文']['211'] = st.slider("语文211线", 60, MAX_SCORES['语文'], 100, 5, key="yuwen_211")
            with col3:
                thresholds['语文']['一本'] = st.slider("语文一本线", 30, MAX_SCORES['语文'], 90, 5, key="yuwen_yiben")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                thresholds['数学'] = {'985': st.slider("数学985线", 60, MAX_SCORES['语文'], 120, 5, key="shuxue_985")}
            with col2:
                thresholds['数学']['211'] = st.slider("数学211线", 60, MAX_SCORES['语文'], 100, 5, key="shuxue_211")
            with col3:
                thresholds['数学']['一本'] = st.slider("数学一本线", 30, MAX_SCORES['语文'], 90, 5, key="shuxue_yiben")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                thresholds['英语'] = {'985': st.slider("英语985线", 60, MAX_SCORES['语文'], 120, 5, key="yingyu_985")}
            with col2:
                thresholds['英语']['211'] = st.slider("英语211线", 60, MAX_SCORES['语文'], 100, 5, key="yingyu_211")
            with col3:
                thresholds['英语']['一本'] = st.slider("英语一本线", 30, MAX_SCORES['语文'], 90, 5, key="yingyu_yiben")
            
            # 首选科目满分100 (物理/历史二选一)
            st.markdown("📋 首选科目 (满分100, 物理/历史二选一)")
            
            # 首选科目选择按钮
            first_choice = st.radio(
                "选择首选科目",
                ["物理", "历史"],
                horizontal=True,
                help="选择本次考试学生的首选科目类型"
            )
            
            if first_choice == "物理":
                first_col = "物理"
                st.success("📌 当前首选：物理")
            else:
                first_col = "历史"
                st.success("📌 当前首选：历史")
            
            # 首选科目分数线设置
            col1, col2, col3 = st.columns(3)
            with col1:
                thresholds[first_col] = {'985': st.slider(f"{first_choice} 985线", 60, 100, 90, 5, key=f"{first_col}_985")}
            with col2:
                thresholds[first_col]['211'] = st.slider(f"{first_choice} 211线", 60, 100, 80, 5, key=f"{first_col}_211")
            with col3:
                thresholds[first_col]['一本'] = st.slider(f"{first_choice} 一本线", 30, 100, 60, 5, key=f"{first_col}_yiben")
            
            # 同步到物理(历史)列名（用于数据匹配）
            thresholds['物理(历史)'] = thresholds[first_col].copy()
            
            # 四选二科目满分100 (化学/生物/政治/地理)
            st.markdown("📋 四选二科目 (满分100, 分别设置)")
            
            # 分别设置化学、生物、政治、地理的分数线
            four_choice_subjects = ['化学', '生物', '政治', '地理']
            
            for subj in four_choice_subjects:
                with st.expander(f"📌 {subj} 分数线设置"):
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        thresholds[subj] = {'985': st.slider(f"{subj} 985线", 60, 100, 90, 5, key=f"{subj}_985")}
                    with col2:
                        thresholds[subj]['211'] = st.slider(f"{subj} 211线", 60, 100, 80, 5, key=f"{subj}_211")
                    with col3:
                        thresholds[subj]['一本'] = st.slider(f"{subj} 一本线", 30, 100, 60, 5, key=f"{subj}_yiben")
        
        st.divider()
        
        # 进步/退步阈值设置
        st.subheader("📈 进步分析设置")
        progress_threshold = st.slider(
            "进步/退步阈值",
            min_value=1,
            max_value=20,
            value=st.session_state.progress_threshold,
            help="分数变化超过此值才被标记为进步或退步，否则视为持平"
        )
        st.session_state.progress_threshold = progress_threshold
        
        st.divider()
        
        top_n = st.slider(
            "Top学生数量",
            min_value=5,
            max_value=50,
            value=10
        )
        
        show_raw_data = st.checkbox("显示原始数据", value=False)
    
    # 主内容区
    if not uploaded_files:
        st.info("👈 请在侧边栏上传Excel文件开始分析")
        
        # 显示示例数据
        st.subheader("📋 示例数据结构")
        sample_data = {
            '学号': ['001', '002', '003'],
            '姓名': ['张三', '李四', '王五'],
            '班级': ['高一(1)班', '高一(1)班', '高一(2)班'],
            '总分': [650, 620, 580],
            '语文': [110, 105, 100],
            '数学': [120, 115, 95],
            '英语': [130, 125, 110],
            '物理(历史)': [85, 90, 80],
            '化学(政治)': [88, 82, 78],
            '生物(地理)': [92, 88, 85]
        }
        st.dataframe(pd.DataFrame(sample_data), use_container_width=True)
        return
    
    # 数据处理
    with st.spinner('正在处理数据...'):
        df, errors = merge_excel_files(uploaded_files)
    
    if errors:
        st.warning(f"处理过程中有 {len(errors)} 个问题:")
        for err in errors[:5]:
            st.text(f"  - {err}")
    
    if df is None or df.empty:
        st.error("❌ 无法解析上传的文件，请检查格式")
        return
    
    # 数据概览
    st.success(f"✅ 成功加载 {len(uploaded_files)} 个文件，共 {len(df)} 条记录")
    
    exams = df['考试名称'].unique()
    classes = df['班级'].dropna().unique() if '班级' in df.columns else []
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("考试次数", len(exams))
    with col2:
        st.metric("学生人数", df['学号'].nunique())
    with col3:
        st.metric("班级数量", len(classes))
    with col4:
        selected_exam = df['考试名称'].max()
        avg_score = pd.to_numeric(df[df['考试名称'] == selected_exam]['总分'], errors='coerce').mean()
        st.metric("最新平均分", f"{avg_score:.1f}" if pd.notna(avg_score) else "N/A")
    
    # 考试选择器
    st.subheader("🎯 选择分析考试")
    exam_options = sorted(df['考试名称'].unique(), key=lambda x: str(x))
    selected_exam = st.selectbox(
        "选择要分析的考试",
        options=exam_options,
        index=len(exam_options) - 1,
        help="选择本次分析针对的考试场次"
    )
    
    # 根据选择的考试筛选数据
    exam_df = df[df['考试名称'] == selected_exam].copy()
    
    # 显示选中考试的统计信息
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("参考人数", len(exam_df))
    with col2:
        exam_avg = pd.to_numeric(exam_df['总分'], errors='coerce').mean()
        st.metric("平均分", f"{exam_avg:.1f}" if pd.notna(exam_avg) else "N/A")
    with col3:
        exam_max = pd.to_numeric(exam_df['总分'], errors='coerce').max()
        st.metric("最高分", f"{exam_max:.0f}" if pd.notna(exam_max) else "N/A")
    
    # 原始数据展示
    if show_raw_data:
        with st.expander("📋 查看原始数据"):
            st.dataframe(df, use_container_width=True)
            # 导出功能
            csv = df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                "📥 下载合并后的CSV",
                data=csv,
                file_name=f"学生成绩_{datetime.now().strftime('%Y%m%d')}.csv",
                mime='text/csv'
            )
    
    st.divider()
    
    # ==================== 分析仪表板 ====================
    
    # Tab布局
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📈 总体概览", 
        "🏫 班级分析",
        "👨‍🎓 学生详情",
        "📊 进步分析",
        "📋 数据报表"
    ])
    
    # 预先计算进步分析数据（供数据报表使用）
    all_exams_for_progress = sorted(df['考试名称'].unique(), key=lambda x: str(x))
    progress_df = None
    if len(all_exams_for_progress) >= 2:
        progress_df = calculate_progress_between_exams(
            df, 
            all_exams_for_progress[0], 
            all_exams_for_progress[-1]
        )
    
    with tab1:
        st.subheader("📈 总体成绩概览")
        
        # 科目统计 - 使用选中的考试数据
        subject_stats = calculate_subject_stats(exam_df)
        
        col1, col2 = st.columns([2, 1])
        with col1:
            if subject_stats is not None and not subject_stats.empty:
                st.plotly_chart(plot_subject_comparison(subject_stats), use_container_width=True)
        with col2:
            st.dataframe(subject_stats, use_container_width=True, hide_index=True)
        
        # 分数分布
        st.subheader("📊 分数分布")
        col1, col2 = st.columns(2)
        
        with col1:
            # 总分分布
            fig_dist = plot_score_distribution(exam_df, '总分')
            if fig_dist:
                st.plotly_chart(fig_dist, use_container_width=True)
        
        with col2:
            # 优秀率/良好率/及格率
            grade_df = calculate_grade_rates(exam_df, thresholds)
            if grade_df is not None and not grade_df.empty:
                # 显示三率热力图
                fig_rates = plot_grade_rates_heatmap(grade_df)
                if fig_rates:
                    st.plotly_chart(fig_rates, use_container_width=True)
                
                # 显示详细数据表格
                st.dataframe(grade_df, use_container_width=True, hide_index=True)
    
    with tab2:
        st.subheader("🏫 班级成绩对比")
        
        class_stats = calculate_class_stats(df)
        class_subject_avg = calculate_class_subject_avg(exam_df)
        
        if class_stats is not None and not class_stats.empty:
            # 班级平均分对比
            st.subheader("📊 班级平均分对比")
            col1, col2 = st.columns([2, 1])
            with col1:
                st.plotly_chart(plot_class_comparison(class_stats), use_container_width=True)
            with col2:
                st.dataframe(class_stats, use_container_width=True, hide_index=True)
            
            st.divider()
            
            # 班级各科平均分对比 - 可开关
            show_subject_avg = st.checkbox("📚 显示班级各科平均分对比", value=True, key="show_subject_avg")
            
            if show_subject_avg and class_subject_avg is not None and not class_subject_avg.empty:
                st.subheader("📚 班级各科平均分对比")
                
                # 调整列顺序，把班级放在第一列
                cols = class_subject_avg.columns.tolist()
                if '班级' in cols:
                    cols.remove('班级')
                    cols = ['班级'] + cols
                    class_subject_avg = class_subject_avg[cols]
                
                st.dataframe(class_subject_avg, use_container_width=True, hide_index=True)
                
                # 可视化：各班各科平均分热力图
                fig_subject_heatmap = plot_class_subject_heatmap(class_subject_avg)
                if fig_subject_heatmap:
                    st.plotly_chart(fig_subject_heatmap, use_container_width=True)
                
                st.divider()
            
            # 班级三线人数对比 - 可开关
            show_thresholds = st.checkbox("🎯 显示班级总分上线人数对比", value=True, key="show_thresholds")
            
            if show_thresholds:
                st.subheader("🎯 班级总分上线人数对比")
                
                class_thresholds_stats = calculate_class_thresholds_stats(df, thresholds)
            
                if class_thresholds_stats is not None and not class_thresholds_stats.empty:
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        fig_thresholds = plot_class_thresholds_comparison(class_thresholds_stats)
                        if fig_thresholds:
                            st.plotly_chart(fig_thresholds, use_container_width=True)
                    with col2:
                        # 显示详细数据
                        st.dataframe(class_thresholds_stats, use_container_width=True, hide_index=True)
                        
                        # 班级详细统计
                        st.subheader("📋 班级详细数据")
                        selected_class = st.selectbox("选择班级", class_thresholds_stats['班级'].unique() if len(class_thresholds_stats) > 0 else [])
                        
                        if selected_class:
                            class_df = df[(df['班级'] == selected_class) & (df['考试名称'] == selected_exam)]
                            class_df = class_df.copy()
                            class_df['总分'] = pd.to_numeric(class_df['总分'], errors='coerce')
                            class_df = class_df.sort_values('总分', ascending=False)
                            st.dataframe(
                                class_df[['学号', '姓名', '总分', '班级排名', '年级排名']].head(20),
                                use_container_width=True,
                                hide_index=True
                            )
            elif show_thresholds:
                st.info("无法计算班级分数线数据")
        else:
            st.warning("没有足够的班级数据进行对比")
    
    with tab3:
        st.subheader("👨‍🎓 学生成绩详情")
        
        # 选择学生
        student_list = df[['学号', '姓名']].drop_duplicates()
        student_list = student_list.dropna(subset=['学号'])
        
        if not student_list.empty:
            # 学号显示优化 - 只显示有选中考试数据的学生
            exam_students = exam_df['学号'].unique()
            student_list = student_list[student_list['学号'].isin(exam_students)]
            
            if student_list.empty:
                st.warning("当前考试没有学生数据")
            else:
                student_options = [
                    f"{row['学号']} - {row['姓名']}" if pd.notna(row['姓名']) else str(row['学号'])
                    for _, row in student_list.iterrows()
                ]
                selected = st.selectbox("选择学生", student_options)
                
                if selected:
                    student_id = selected.split(' - ')[0] if ' - ' in selected else selected
                    
                    # 学生基本信息 - 使用全部数据展示历史趋势
                    student_df = df[df['学号'] == student_id]
                    
                    if student_df.empty:
                        st.warning("该学生没有成绩数据")
                    else:
                        # 获取学生在选中考试中的成绩
                        exam_student_df = exam_df[exam_df['学号'] == student_id]
                        
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            name = student_df['姓名'].iloc[0] if not student_df.empty else 'N/A'
                            st.metric("姓名", name if pd.notna(name) else 'N/A')
                        with col2:
                            class_name = student_df['班级'].iloc[0] if not student_df.empty else 'N/A'
                            st.metric("班级", class_name if pd.notna(class_name) else 'N/A')
                        with col3:
                            if not exam_student_df.empty:
                                total = exam_student_df['总分'].iloc[0]
                                st.metric("总分", f"{total:.0f}" if pd.notna(total) else 'N/A')
                            else:
                                st.metric("总分", 'N/A')
                        with col4:
                            if not exam_student_df.empty:
                                rank = exam_student_df['年级排名'].iloc[0]
                                st.metric("排名", f"{rank:.0f}" if pd.notna(rank) else 'N/A')
                            else:
                                st.metric("排名", 'N/A')
                        
                        # 成绩趋势图 - 使用全部数据展示趋势
                        fig_trend = plot_student_trend(df, student_id)
                        if fig_trend:
                            st.plotly_chart(fig_trend, use_container_width=True)
                        
                        # 学生历史成绩表
                        history = student_df[['考试名称', '总分', '年级排名', '班级排名']].sort_values('考试名称')
                        st.dataframe(history, use_container_width=True, hide_index=True)
        else:
            st.warning("没有找到学生数据")
    
    with tab4:
        st.subheader("📊 进步/退步分析")
        
        # 获取所有考试列表
        all_exams = sorted(df['考试名称'].unique(), key=lambda x: str(x))
        
        if len(all_exams) < 2:
            st.info("需要至少2次考试数据才能分析进步/退步趋势")
        else:
            # 考试选择器
            col1, col2 = st.columns(2)
            with col1:
                prev_exam = st.selectbox(
                    "选择上次考试",
                    options=all_exams,
                    index=0,
                    help="选择作为基准的考试"
                )
            with col2:
                curr_exam = st.selectbox(
                    "选择本次考试",
                    options=all_exams,
                    index=len(all_exams) - 1,
                    help="选择要比较的考试"
                )
            
            # 计算指定两次考试的进步/退步
            progress_df = calculate_progress_between_exams(df, prev_exam, curr_exam)
            
            if progress_df is not None and not progress_df.empty:
                st.markdown(f"**比较**: {prev_exam} → {curr_exam}")
                
                col1, col2 = st.columns([1, 2])
                with col1:
                    fig_pie = plot_progress_distribution(progress_df)
                    if fig_pie:
                        st.plotly_chart(fig_pie, use_container_width=True)
                
                with col2:
                    # 统计数据
                    status_counts = progress_df['状态'].value_counts()
                    col_a, col_b, col_c = st.columns(3)
                    with col_a:
                        st.metric("进步人数", status_counts.get('进步', 0))
                    with col_b:
                        st.metric("退步人数", status_counts.get('退步', 0))
                    with col_c:
                        st.metric("持平人数", status_counts.get('持平', 0))
                
                # 进步排行榜
                st.subheader("🏆 进步排行榜")
                
                progress_top = progress_df[progress_df['状态'] == '进步'].copy()
                if not progress_top.empty:
                    # 按班级排序，再按排名变化排序（进步大的在前）
                    progress_top = progress_top.sort_values(
                        by=['班级', '排名变化'], 
                        ascending=[True, False]
                    ).head(10)
                    st.dataframe(
                        progress_top[['姓名', '班级', '分数变化', '排名变化']].rename(columns={
                            '姓名': '学生', '班级': '班级',
                            '分数变化': '分数+', '排名变化': '排名↑'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )
                
                # 退步提醒
                st.subheader("⚠️ 需要关注的学生")
                regress_top = progress_df[progress_df['状态'] == '退步'].copy()
                if not regress_top.empty:
                    # 按班级排序，再按排名变化排序（退步多的在前，即排名变化为负且绝对值大的）
                    regress_top = regress_top.sort_values(
                        by=['班级', '排名变化'], 
                        ascending=[True, True]
                    ).head(10)
                    st.dataframe(
                        regress_top[['姓名', '班级', '分数变化', '排名变化']].rename(columns={
                            '姓名': '学生', '班级': '班级',
                            '分数变化': '分数-', '排名变化': '排名↓'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )
            else:
                st.warning(f"在 {prev_exam} 和 {curr_exam} 中没有找到相同学生的数据")
    
    with tab5:
        st.subheader("📋 数据报表")
        
        # 定义居中显示的CSS样式
        st.markdown("""
        <style>
        .dataframe td {
            text-align: center !important;
        }
        .dataframe th {
            text-align: center !important;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # 报表选项
        report_type = st.selectbox(
            "选择报表类型",
            ["班级成绩报表", "学生成绩报表", "科目成绩报表", "进步分析报表"]
        )
        
        if report_type == "班级成绩报表":
            if class_stats is not None:
                st.dataframe(class_stats.style.set_properties(**{'text-align': 'center'}), use_container_width=True, hide_index=True)
                csv = class_stats.to_csv(index=False, encoding='utf-8-sig')
                st.download_button("📥 下载报表", csv, "班级成绩报表.csv", "text/csv")
        
        elif report_type == "学生成绩报表":
            exam_df = df[df['考试名称'] == selected_exam].copy()
            exam_df['总分'] = pd.to_numeric(exam_df['总分'], errors='coerce')
            exam_df = exam_df.sort_values('总分', ascending=False)
            display_cols = ['学号', '姓名', '班级', '总分', '班级排名', '年级排名']
            available_cols = [c for c in display_cols if c in exam_df.columns]
            st.dataframe(exam_df[available_cols].style.set_properties(**{'text-align': 'center'}), use_container_width=True, hide_index=True)
            csv = exam_df[available_cols].to_csv(index=False, encoding='utf-8-sig')
            st.download_button("📥 下载报表", csv, "学生成绩报表.csv", "text/csv")
        
        elif report_type == "科目成绩报表":
            if subject_stats is not None:
                st.dataframe(subject_stats.style.set_properties(**{'text-align': 'center'}), use_container_width=True, hide_index=True)
                csv = subject_stats.to_csv(index=False, encoding='utf-8-sig')
                st.download_button("📥 下载报表", csv, "科目成绩报表.csv", "text/csv")
        
        elif report_type == "进步分析报表":
            if progress_df is not None:
                st.dataframe(progress_df.style.set_properties(**{'text-align': 'center'}), use_container_width=True)
                csv = progress_df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button("📥 下载报表", csv, "进步分析报表.csv", "text/csv")

if __name__ == '__main__':
    main()
