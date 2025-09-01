import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from typing import List, Dict, Tuple, Set, Optional
from collections import defaultdict

# 导入自定义模块
from utils import (
    load_names_from_excel,
    load_preferences_from_excel,
    parse_custom_weights,
    compute_pair_weights,
    generate_seats,
    generate_adjacent_edges,
    visualize_layout,
    validate_layout,
    solve_top_n_assignments,
    compute_satisfaction_metrics,
    export_assignment_to_excel,
    export_assignment_to_image,
    get_seat_info
)

# 设置matplotlib中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 页面配置
st.set_page_config(page_title="智能座位分配系统", layout="wide", 
                   page_icon="🪑", initial_sidebar_state="expanded")

# 自定义CSS样式
st.markdown("""
<style>
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .stExpander {
        border: 1px solid #f0f2f6;
        border-radius: 0.5rem;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }
    .result-card {
        padding: 1rem;
        border-radius: 0.5rem;
        background: #f8f9fa;
        margin-bottom: 1rem;
    }
    .highlight-text {
        font-weight: bold;
        color: #4361ee;
    }
    .title-container {
        display: flex;
        align-items: center;
        margin-bottom: 1rem;
    }
    .title-icon {
        font-size: 2rem;
        margin-right: 1rem;
    }
    .metric-container {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# 页面标题和介绍
st.markdown("""
<div class="title-container">
    <div class="title-icon">🪑</div>
    <div>
        <h1 style="margin:0;">智能座位分配系统</h1>
        <p style="margin:0;color:#666;">基于喜好关系的座位优化分配</p>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div style="background-color:#f0f7ff;padding:10px;border-radius:5px;margin-bottom:20px;">
    <p style="margin-bottom:5px;"><b>系统功能</b>：根据喜好关系优化座位分配，让喜欢坐在一起的人相邻，不喜欢坐在一起的人分开。</p>
    <p style="margin:0;"><b>使用方法</b>：上传名单和喜好关系Excel，设置权重，配置教室布局，然后生成最优座位方案。</p>
</div>
""", unsafe_allow_html=True)

# 初始化session state变量
if 'willing_pairs_by_rank' not in st.session_state:
    st.session_state.willing_pairs_by_rank = defaultdict(set)
if 'unwilling_pairs_by_rank' not in st.session_state:
    st.session_state.unwilling_pairs_by_rank = defaultdict(set)
if 'willing_headers' not in st.session_state:
    st.session_state.willing_headers = []
if 'unwilling_headers' not in st.session_state:
    st.session_state.unwilling_headers = []

# 侧边栏配置
with st.sidebar:
    st.header("📋 配置参数")
    
    # 数据输入
    st.subheader("1. 数据输入")
    
    # 上传名单文件
    names_file = st.file_uploader(
        "上传名单Excel文件", 
        type=['xlsx', 'xls'],
        help="Excel文件第一列为姓名，可包含表头"
    )
    
    # 上传喜好关系文件
    preferences_file = st.file_uploader(
        "上传喜好关系Excel文件（可选）", 
        type=['xlsx', 'xls'],
        help="支持自定义数据范围的Excel文件"
    )
    
    # 数据范围配置
    if preferences_file is not None:
        st.subheader("📊 数据范围配置")
        
        # 自动识别选项
        auto_detect_range = st.checkbox(
            "🔍 自动识别单元格范围",
            value=True,
            help="自动识别喜好名单的单元格范围，中间的空列用来隔开愿意和不愿意，系统会查找最长的那列以确定行数"
        )
        
        # 初始化session_state中的检测结果
        if 'detected_willing_range' not in st.session_state:
            st.session_state.detected_willing_range = "A1:B10"
        if 'detected_unwilling_range' not in st.session_state:
            st.session_state.detected_unwilling_range = "D1:E10"
            
        if not auto_detect_range:
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**喜欢关系数据范围**")
                willing_range = st.text_input(
                    "喜欢关系单元格范围",
                    value=st.session_state.detected_willing_range,
                    help="例如：A1:B10 表示从A1到B10的区域，每两列为一个等级"
                )
                
            with col2:
                st.write("**不喜欢关系数据范围**")
                unwilling_range = st.text_input(
                    "不喜欢关系单元格范围",
                    value=st.session_state.detected_unwilling_range,
                    help="例如：D1:E10 表示从D1到E10的区域，每两列为一个等级"
                )
        else:
            willing_range = st.session_state.detected_willing_range
            unwilling_range = st.session_state.detected_unwilling_range
        
        # 工作表选择
        sheet_name_input = st.text_input(
            "工作表名称（可选）",
            value="",
            help="留空则使用第一个工作表"
        )
        
        st.info("💡 **数据格式说明**：\n- 每一列为一个等级的喜好关系\n- 每个单元格包含用逗号分隔的人名对（如：张三,李四）\n- 第一行可以是表头（会自动跳过）\n- 从第二行开始是实际的人名对数据")
    
    # 手动输入喜好关系
    st.subheader("手动输入喜好关系")
    manual_preferences = st.text_area(
        "输入格式：姓名1,喜欢/不喜欢,姓名2（每行一个关系）",
        height=100,
        help="例如：张三,喜欢,李四\n王五,不喜欢,赵六"
    )
    
    # 权重设置
    st.subheader("权重设置")
    
    # 动态生成权重设置控件
    like_weights = []
    dislike_weights = []
    
    # 如果有Excel数据，根据等级数量生成权重控件
    if st.session_state.willing_headers:
        st.write("**喜欢关系权重设置**")
        for i, (header1, header2) in enumerate(st.session_state.willing_headers):
            label = f"等级{i+1}权重" if not header1 or not header2 else f"{header1}-{header2}权重"
            weight = st.slider(
                label, 
                1.0, 10.0, 5.0 - i * 0.5, 0.5, 
                key=f"like_weight_{i}",
                help=f"等级{i+1}的喜欢关系权重，数值越大影响越大"
            )
            like_weights.append(weight)
    else:
        # 默认单一权重
        like_weight = st.slider("喜欢权重", 1.0, 10.0, 5.0, 0.5, help="数值越大，喜欢的人越倾向于坐在一起")
        like_weights = [like_weight]
    
    if st.session_state.unwilling_headers:
        st.write("**不喜欢关系权重设置**")
        for i, (header1, header2) in enumerate(st.session_state.unwilling_headers):
            label = f"等级{i+1}权重" if not header1 or not header2 else f"{header1}-{header2}权重"
            weight = st.slider(
                label, 
                -10.0, -1.0, -5.0 - i * 0.5, 0.5, 
                key=f"dislike_weight_{i}",
                help=f"等级{i+1}的不喜欢关系权重，绝对值越大影响越大"
            )
            dislike_weights.append(weight)
    else:
        # 默认单一权重
        dislike_weight = st.slider("不喜欢权重", -10.0, -1.0, -5.0, 0.5, help="数值越小（绝对值越大），不喜欢的人越倾向于分开")
        dislike_weights = [dislike_weight]
    
    # 显示权重配置信息
    if st.session_state.willing_headers or st.session_state.unwilling_headers:
        st.info(f"📊 权重配置：{len(like_weights)}个喜欢等级权重，{len(dislike_weights)}个不喜欢等级权重")
    
    # 座位布局配置
    st.header("2. 座位布局")
    
    # 基本布局
    st.subheader("基本布局")
    n_cols = st.number_input("列数", min_value=1, max_value=20, value=4, step=1)
    
    # 每列的行数配置
    st.subheader("每列行数配置")
    layout_mode = st.radio(
        "布局模式",
        ["统一行数", "自定义每列行数"],
        help="统一行数：所有列都有相同的行数；自定义：可以为每列设置不同的行数"
    )
    
    if layout_mode == "统一行数":
        # 根据学生数量动态计算默认行数
        default_rows = 15
        if names_file is not None:
            try:
                temp_names = load_names_from_excel(names_file)
                if temp_names:
                    # 计算每列平均行数，向上取整
                    default_rows = max(1, (len(temp_names) + n_cols - 1) // n_cols)
            except:
                pass
        uniform_rows = st.number_input("每列行数", min_value=1, max_value=30, value=default_rows, step=1)
        col_rows = [uniform_rows] * n_cols
    else:
        col_rows = []
        # 根据学生数量动态计算默认行数
        default_rows = 15
        if names_file is not None:
            try:
                temp_names = load_names_from_excel(names_file)
                if temp_names:
                    # 计算每列平均行数，向上取整
                    default_rows = max(1, (len(temp_names) + n_cols - 1) // n_cols)
            except:
                pass
        for i in range(n_cols):
            rows = st.number_input(f"第{i+1}列行数", min_value=1, max_value=30, value=default_rows, step=1, key=f"col_{i}")
            col_rows.append(rows)
    
    # 过道配置
    st.subheader("过道配置")
    enable_aisles = st.checkbox("启用过道", value=False, help="过道会阻断座位间的邻接关系")
    
    aisles = []
    if enable_aisles:
        n_aisles = st.slider("过道数量", min_value=1, max_value=min(5, n_cols-1), value=1, step=1)
        
        for i in range(n_aisles):
            st.write(f"**过道 {i+1}**")
            col1, col2 = st.columns(2)
            with col1:
                left_col = st.selectbox(
                    f"左侧列", 
                    options=list(range(1, n_cols+1)), 
                    index=min(i*2, n_cols-2),
                    key=f"aisle_{i}_left"
                )
            with col2:
                right_col = st.selectbox(
                    f"右侧列", 
                    options=list(range(left_col+1, n_cols+1)), 
                    index=0,
                    key=f"aisle_{i}_right"
                )
            aisles.append((left_col-1, right_col-1))  # 转换为0索引
    
    # 显示座位布局预览
    total_seats = sum(col_rows)
    aisle_info = f"，{len(aisles)}条过道" if enable_aisles and aisles else ""
    st.info(f"座位布局：{n_cols}列，共{total_seats}个座位{aisle_info}")
    
    include_diag = st.checkbox("包含对角线邻座关系", True, help="选中则认为对角线方向也是邻座")
    
    st.header("3. 计算设置")
    # 优化参数
    st.subheader("优化参数")
    top_n = st.slider("生成方案数", 1, 10, 3, 1)
    time_limit = st.slider("每个方案计算时间限制(秒)", 1.0, 60.0, 10.0, 0.5)
    
    # 可视化选项
    st.subheader("可视化选项")
    show_visualization = st.checkbox("显示座位分配图", value=True)
    split_visualization = st.checkbox("拆分正负关系为两张图", value=False, help="将愿意在一起和不愿意在一起的关系分别显示在两张图中")
    show_all_lines = st.checkbox("显示所有关系线", value=False, help="开启后显示所有喜好关系线，关闭则仅显示实际满足的关系线")
    if show_visualization:
        viz_dpi = st.slider("图像分辨率 (DPI)", min_value=50, max_value=300, value=100, step=10)
        viz_figsize_w = st.slider("图像宽度", min_value=8, max_value=20, value=12, step=1)
        viz_figsize_h = st.slider("图像高度", min_value=6, max_value=16, value=8, step=1)
    
    # 显示选项
    st.subheader("显示选项")
    debug_mode = st.checkbox("启用调试模式", value=False, help="显示详细的生成过程信息")
    more_info_mode = st.checkbox("更多信息", value=False, help="显示详细的满足率统计和分析信息")

# 主区域
tab1, tab2, tab3 = st.tabs(["数据预览", "结果查看", "使用帮助"])

with tab1:
    st.header("📊 数据预览")
    
    # 处理名单数据
    names = []
    if names_file is not None:
        try:
            names = load_names_from_excel(names_file)
            st.success(f"✅ 成功加载 {len(names)} 个姓名")
            
            # 显示名单预览
            with st.expander("👥 查看名单详情", expanded=False):
                cols = st.columns(4)
                for i, name in enumerate(names):
                    with cols[i % 4]:
                        st.write(f"{i+1}. {name}")
        except Exception as e:
            st.error(f"❌ 名单文件读取失败：{str(e)}")
    
    # 处理喜好关系数据
    preferences = []
    if preferences_file is not None:
        try:
            # 获取用户配置的参数
            sheet_name = sheet_name_input.strip() if sheet_name_input.strip() else None
            
            # load_preferences_from_excel返回元组，需要正确解析
            st.session_state.willing_pairs_by_rank, st.session_state.unwilling_pairs_by_rank, st.session_state.willing_headers, st.session_state.unwilling_headers, detection_results = load_preferences_from_excel(
                preferences_file,
                sheet_name=sheet_name,
                willing_range_spec=willing_range,
                unwilling_range_spec=unwilling_range,
                auto_detect=auto_detect_range
            )
            
            # 显示自动识别结果并更新session_state
            if auto_detect_range and detection_results:
                # 更新session_state中的检测结果
                if detection_results[0]:
                    st.session_state.detected_willing_range = detection_results[0]
                if detection_results[1]:
                    st.session_state.detected_unwilling_range = detection_results[1]
                    
                st.success("✅ 自动识别成功！")
                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"🟢 喜欢关系范围: {detection_results[0]}")
                with col2:
                    st.info(f"🔴 不喜欢关系范围: {detection_results[1]}")
            elif auto_detect_range:
                st.warning("⚠️ 未能自动识别到有效的数据范围，请检查Excel文件格式。")
            
            # 将解析的数据转换为preferences格式
            for rank, pairs in st.session_state.willing_pairs_by_rank.items():
                for pair in pairs:
                    names_list = list(pair)
                    if len(names_list) >= 2:
                        preferences.append([names_list[0], "喜欢", names_list[1]])
            
            for rank, pairs in st.session_state.unwilling_pairs_by_rank.items():
                for pair in pairs:
                    names_list = list(pair)
                    if len(names_list) >= 2:
                        preferences.append([names_list[0], "不喜欢", names_list[1]])
            
            st.success(f"✅ 成功加载 {len(preferences)} 条喜好关系")
            
            # 显示数据范围信息
            if st.session_state.willing_pairs_by_rank or st.session_state.unwilling_pairs_by_rank:
                # 使用实际使用的范围（自动检测时使用检测结果，否则使用用户输入）
                actual_willing_range = st.session_state.detected_willing_range if auto_detect_range and hasattr(st.session_state, 'detected_willing_range') else willing_range
                actual_unwilling_range = st.session_state.detected_unwilling_range if auto_detect_range and hasattr(st.session_state, 'detected_unwilling_range') else unwilling_range
                st.info(f"📊 数据解析详情：\n- 喜欢关系范围：{actual_willing_range}\n- 不喜欢关系范围：{actual_unwilling_range}\n- 工作表：{sheet_name or '第一个工作表'}")
                
                # 显示解析的等级信息
                level_info = []
                for i, (h1, h2) in enumerate(st.session_state.willing_headers):
                    level_info.append(h1 or f"喜欢等级{i+1}")
                for i, (h1, h2) in enumerate(st.session_state.unwilling_headers):
                    level_info.append(h1 or f"不喜欢等级{i+1}")
                
                if level_info:
                    st.success("✅ 识别的等级列：\n" + "\n".join(level_info))
        except Exception as e:
            st.error(f"❌ 喜好关系文件读取失败：{str(e)}")
            st.error("请检查数据范围设置是否正确，确保Excel文件格式符合要求。")
    
    # 处理手动输入的喜好关系
    if manual_preferences.strip():
        try:
            manual_prefs = parse_custom_weights(manual_preferences)
            # 将手动输入的权重对转换为preferences格式
            for name1, name2, weight in manual_prefs:
                pref_type = "喜欢" if weight > 0 else "不喜欢"
                preferences.append([name1, pref_type, name2])
            st.success(f"✅ 成功解析 {len(manual_prefs)} 条手动输入的喜好关系")
        except Exception as e:
            st.error(f"❌ 手动输入解析失败：{str(e)}")
    
    # 显示喜好关系预览
    if preferences:
        with st.expander("💕 查看喜好关系详情", expanded=False):
            # 统计信息
            like_count = sum(1 for p in preferences if len(p) > 1 and p[1] == "喜欢")
            dislike_count = sum(1 for p in preferences if len(p) > 1 and p[1] == "不喜欢")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("总关系数", len(preferences))
            with col2:
                st.metric("喜欢关系", like_count)
            with col3:
                st.metric("不喜欢关系", dislike_count)
            
            # 按等级分组显示喜好关系
            if hasattr(st.session_state, 'willing_pairs_by_rank') and hasattr(st.session_state, 'unwilling_pairs_by_rank'):
                # 使用Excel数据的等级分组
                willing_pairs = st.session_state.willing_pairs_by_rank
                unwilling_pairs = st.session_state.unwilling_pairs_by_rank
                willing_headers = getattr(st.session_state, 'willing_headers', [])
                unwilling_headers = getattr(st.session_state, 'unwilling_headers', [])
                
                # 创建标签页
                tab_names = []
                tab_contents = []
                
                # 添加喜欢等级标签页
                for rank in sorted(willing_pairs.keys()):
                    if willing_pairs[rank]:
                        header_name = willing_headers[rank-1][0] if rank-1 < len(willing_headers) else f"喜欢等级{rank}"
                        tab_names.append(f"😊 {header_name}")
                        pairs_list = [[list(pair)[0], "喜欢", list(pair)[1]] for pair in willing_pairs[rank]]
                        tab_contents.append(("like", pairs_list, rank))
                
                # 添加不喜欢等级标签页
                for rank in sorted(unwilling_pairs.keys()):
                    if unwilling_pairs[rank]:
                        header_name = unwilling_headers[rank-1][0] if rank-1 < len(unwilling_headers) else f"不喜欢等级{rank}"
                        tab_names.append(f"😤 {header_name}")
                        pairs_list = [[list(pair)[0], "不喜欢", list(pair)[1]] for pair in unwilling_pairs[rank]]
                        tab_contents.append(("dislike", pairs_list, rank))
                
                if tab_names:
                    tabs = st.tabs(tab_names)
                    for i, (tab_type, pairs_list, rank) in enumerate(tab_contents):
                        with tabs[i]:
                            if pairs_list:
                                df_level = pd.DataFrame(pairs_list, columns=["姓名1", "关系类型", "姓名2"])
                                st.dataframe(df_level, use_container_width=True)
                                st.info(f"本等级共有 {len(pairs_list)} 对关系")
                            else:
                                st.info("本等级暂无关系数据")
                else:
                    st.info("暂无等级分组数据")
            else:
                # 手动输入数据的简单显示
                df_prefs = pd.DataFrame(preferences, columns=["姓名1", "关系类型", "姓名2"])
                
                # 按关系类型分组显示
                like_prefs = [p for p in preferences if len(p) > 1 and p[1] == "喜欢"]
                dislike_prefs = [p for p in preferences if len(p) > 1 and p[1] == "不喜欢"]
                
                tab_names = []
                if like_prefs:
                    tab_names.append("😊 喜欢关系")
                if dislike_prefs:
                    tab_names.append("😤 不喜欢关系")
                
                if tab_names:
                    tabs = st.tabs(tab_names)
                    tab_idx = 0
                    
                    if like_prefs:
                        with tabs[tab_idx]:
                            df_like = pd.DataFrame(like_prefs, columns=["姓名1", "关系类型", "姓名2"])
                            st.dataframe(df_like, use_container_width=True)
                            st.info(f"共有 {len(like_prefs)} 对喜欢关系")
                        tab_idx += 1
                    
                    if dislike_prefs:
                        with tabs[tab_idx]:
                            df_dislike = pd.DataFrame(dislike_prefs, columns=["姓名1", "关系类型", "姓名2"])
                            st.dataframe(df_dislike, use_container_width=True)
                            st.info(f"共有 {len(dislike_prefs)} 对不喜欢关系")
                else:
                    st.dataframe(df_prefs, use_container_width=True)
    
    # 座位布局预览
    if names:
        st.subheader("🪑 座位布局预览")
        
        # 生成座位和邻接关系
        seats = generate_seats(n_cols, col_rows)
        edges = generate_adjacent_edges(n_cols, col_rows, include_diag, aisles if enable_aisles else [])
        
        # 显示布局信息
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"👥 人数：{len(names)}")
            st.info(f"🪑 座位数：{len(seats)}")
            if len(seats) >= len(names):
                st.success(f"✅ 空余座位：{len(seats) - len(names)}")
            else:
                st.error(f"❌ 座位不足：缺少 {len(names) - len(seats)} 个座位")
        
        with col2:
            st.info(f"🔗 邻接边数：{len(edges)}")
            if preferences:
                # 如果有Excel数据，使用已解析的等级数据
                if st.session_state.willing_pairs_by_rank or st.session_state.unwilling_pairs_by_rank:
                    # 使用Excel中的多等级数据
                    current_willing = st.session_state.willing_pairs_by_rank
                    current_unwilling = st.session_state.unwilling_pairs_by_rank
                else:
                    # 解析手动输入的偏好数据为按等级分组的格式
                    current_willing = defaultdict(set)
                    current_unwilling = defaultdict(set)
                    
                    for p in preferences:
                        if len(p) > 2:
                            name1, pref_type, name2 = p[0], p[1], p[2]
                            pair = frozenset([name1, name2])
                            if pref_type == "喜欢":
                                # 如果有多个等级，需要根据实际等级分配
                                level = 0 if len(like_weights) == 1 else 1
                                current_willing[level].add(pair)
                            elif pref_type == "不喜欢":
                                # 如果有多个等级，需要根据实际等级分配
                                level = 0 if len(dislike_weights) == 1 else 1
                                current_unwilling[level].add(pair)
                
                pair_weights = compute_pair_weights(
                    current_willing, 
                    current_unwilling,
                    like_weights, 
                    dislike_weights
                )
                st.info(f"⚖️ 权重对数：{len(pair_weights)}")
                
                # 显示等级统计信息
                if len(like_weights) > 1 or len(dislike_weights) > 1:
                    willing_count = sum(len(pairs) for pairs in current_willing.values())
                    unwilling_count = sum(len(pairs) for pairs in current_unwilling.values())
                    st.info(f"📊 等级统计：{len(current_willing)}个喜欢等级({willing_count}对)，{len(current_unwilling)}个不喜欢等级({unwilling_count}对)")
                    
                    # 显示详细等级统计
                    level_stats = []
                    for rank, pairs in current_willing.items():
                        level_stats.append(f"喜欢等级{rank}: {len(pairs)}对")
                    for rank, pairs in current_unwilling.items():
                        level_stats.append(f"不喜欢等级{rank}: {len(pairs)}对")
                    
                    if level_stats:
                        st.info("📊 详细等级统计：\n" + "\n".join(level_stats))
        
        # 可视化布局
        try:
            fig = visualize_layout(n_cols, col_rows, edges, aisles if enable_aisles else None)
            
            st.pyplot(fig, use_container_width=True)
        except Exception as e:
            st.error(f"❌ 布局可视化失败：{str(e)}")

with tab2:
    st.header("🎯 座位分配结果")
    
    if not names:
        st.warning("⚠️ 请先在左侧上传名单文件")
    elif len(names) > sum(col_rows):
        st.error(f"❌ 座位不足：需要 {len(names)} 个座位，但只有 {sum(col_rows)} 个")
    else:
        # 计算按钮
        if st.button("🚀 计算最优座位分配", type="primary", use_container_width=True):
            with st.spinner("🔄 正在计算最优座位分配..."):
                try:
                    # 生成座位和邻接关系
                    seats = generate_seats(n_cols, col_rows)
                    edges = generate_adjacent_edges(n_cols, col_rows, include_diag, aisles if enable_aisles else [])
                    
                    # 计算权重
                    pair_weights = {}
                    if preferences:
                        # 如果有Excel数据，使用已解析的等级数据
                        if st.session_state.willing_pairs_by_rank or st.session_state.unwilling_pairs_by_rank:
                            # 使用Excel中的多等级数据
                            current_willing = st.session_state.willing_pairs_by_rank
                            current_unwilling = st.session_state.unwilling_pairs_by_rank
                            
                            # 确保数据结构正确
                            if not isinstance(current_willing, dict):
                                current_willing = defaultdict(set)
                            if not isinstance(current_unwilling, dict):
                                current_unwilling = defaultdict(set)
                        else:
                            # 解析手动输入的偏好数据为按等级分组的格式
                            current_willing = defaultdict(set)
                            current_unwilling = defaultdict(set)
                            
                            for p in preferences:
                                if len(p) > 2:
                                    name1, pref_type, name2 = p[0], p[1], p[2]
                                    pair = frozenset([name1, name2])
                                    if pref_type == "喜欢":
                                        # 如果有多个等级，需要根据实际等级分配
                                        level = 0 if len(like_weights) == 1 else 1
                                        current_willing[level].add(pair)
                                    elif pref_type == "不喜欢":
                                        # 如果有多个等级，需要根据实际等级分配
                                        level = 0 if len(dislike_weights) == 1 else 1
                                        current_unwilling[level].add(pair)
                        
                        pair_weights = compute_pair_weights(
                            current_willing, 
                            current_unwilling,
                            like_weights, 
                            dislike_weights
                        )
                    
                    # 存储到session state
                    st.session_state.seats = seats
                    st.session_state.edges = edges
                    st.session_state.pair_weights = pair_weights
                    st.session_state.names = names
                    
                    # 求解
                    # 创建进度条和状态显示
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    def progress_callback(progress: float, message: str):
                        progress_bar.progress(progress)
                        status_text.text(f"🔄 {message}")
                    
                    if debug_mode:
                        debug_container = st.empty()
                        debug_info = []
                        
                        # 重定向调试输出
                        import io
                        import sys
                        from contextlib import redirect_stdout
                        
                        debug_output = io.StringIO()
                        with redirect_stdout(debug_output):
                            results = solve_top_n_assignments(
                                names, seats, pair_weights, edges,
                                top_n=top_n, time_limit_s=time_limit,
                                progress_callback=progress_callback,
                                debug_mode=debug_mode
                            )
                        
                        # 显示调试信息
                        debug_text = debug_output.getvalue()
                        if debug_text:
                            with st.expander("🔍 调试信息", expanded=True):
                                st.text(debug_text)
                    else:
                        results = solve_top_n_assignments(
                            names, seats, pair_weights, edges,
                            top_n=top_n, time_limit_s=time_limit,
                            progress_callback=progress_callback,
                            debug_mode=debug_mode
                        )
                    
                    # 清除进度显示
                    progress_bar.empty()
                    status_text.empty()
                    
                    if results:
                        st.session_state.results = results
                        st.success(f"✅ 成功生成 {len(results)} 个座位方案！")
                    else:
                        st.error("❌ 未能找到可行的座位分配方案")
                        
                except Exception as e:
                    st.error(f"❌ 计算过程中出现错误：{str(e)}")
                    if debug_mode:
                        st.exception(e)
        
        # 显示结果
        if hasattr(st.session_state, 'results') and st.session_state.results:
            st.subheader("📋 分配方案")
            
            # 计算各等级满足率统计（只计算一次，使用第一个方案）
            first_assignment = st.session_state.results[0].assignment
            willing_pairs_data = st.session_state.get('willing_pairs_by_rank', {})
            unwilling_pairs_data = st.session_state.get('unwilling_pairs_by_rank', {})
            
            # 计算愿意关系各等级满足率
            willing_satisfaction = []
            willing_names = []
            
            if willing_pairs_data and len(willing_pairs_data) > 0:
                for rank in sorted(willing_pairs_data.keys()):
                    pairs = willing_pairs_data[rank]
                    if pairs:
                        satisfied_in_level = 0
                        total_pairs = len(pairs)
                        
                        for pair in pairs:
                             pair_list = list(pair)
                             if len(pair_list) >= 2:
                                 name1, name2 = pair_list[0], pair_list[1]
                                 if name1 in first_assignment and name2 in first_assignment:
                                     seat1_idx = first_assignment[name1]
                                     seat2_idx = first_assignment[name2]
                                     # 检查这两个座位是否相邻
                                     seat1_coord = st.session_state.seats[seat1_idx]
                                     seat2_coord = st.session_state.seats[seat2_idx]
                                     is_adjacent = False
                                     for (c1, r1), (c2, r2) in st.session_state.edges:
                                         if ((c1, r1) == seat1_coord and (c2, r2) == seat2_coord) or \
                                            ((c1, r1) == seat2_coord and (c2, r2) == seat1_coord):
                                             is_adjacent = True
                                             break
                                     if is_adjacent:
                                         satisfied_in_level += 1
                        
                        level_rate = (satisfied_in_level / total_pairs) * 100 if total_pairs > 0 else 0
                        willing_satisfaction.append(level_rate)
                        willing_names.append(f'第{rank}顺位愿意')
            
            # 计算不愿意关系各等级满足率（满足率指成功分开的比例）
            unwilling_satisfaction = []
            unwilling_names = []
            
            if unwilling_pairs_data and len(unwilling_pairs_data) > 0:
                for rank in sorted(unwilling_pairs_data.keys()):
                    pairs = unwilling_pairs_data[rank]
                    if pairs:
                        separated_in_level = 0
                        total_pairs = len(pairs)
                        
                        for pair in pairs:
                            pair_list = list(pair)
                            if len(pair_list) >= 2:
                                name1, name2 = pair_list[0], pair_list[1]
                                if name1 in first_assignment and name2 in first_assignment:
                                    seat1_idx = first_assignment[name1]
                                    seat2_idx = first_assignment[name2]
                                    # 检查这两个座位是否相邻
                                    seat1_coord = st.session_state.seats[seat1_idx]
                                    seat2_coord = st.session_state.seats[seat2_idx]
                                    is_adjacent = False
                                    for (c1, r1), (c2, r2) in st.session_state.edges:
                                        if ((c1, r1) == seat1_coord and (c2, r2) == seat2_coord) or \
                                           ((c1, r1) == seat2_coord and (c2, r2) == seat1_coord):
                                            is_adjacent = True
                                            break
                                    # 不愿意关系的满足是指没有相邻
                                    if not is_adjacent:
                                        separated_in_level += 1
                        
                        level_rate = (separated_in_level / total_pairs) * 100 if total_pairs > 0 else 0
                        unwilling_satisfaction.append(level_rate)
                        unwilling_names.append(f'第{rank}顺位不愿意')
            
            # 显示各等级满足率统计（只显示一次）
            if willing_satisfaction or unwilling_satisfaction:
                st.markdown("### 📊 各等级满足率统计")
                col1, col2 = st.columns(2)
                
                with col1:
                    # 显示愿意关系满足率
                    if willing_satisfaction:
                        for idx, (name, rate) in enumerate(zip(willing_names, willing_satisfaction)):
                            color = '#28a745' if rate >= 70 else '#ffc107' if rate >= 40 else '#dc3545'
                            st.markdown(f"""
                            <div class="metric-container">
                                <h4>💚 {name}</h4>
                                <h2 style="color: {color};">{rate:.1f}%</h2>
                                <small>愿意关系满足率</small>
                            </div>
                            """, unsafe_allow_html=True)
                
                with col2:
                    # 显示不愿意关系满足率
                    if unwilling_satisfaction:
                        for idx, (name, rate) in enumerate(zip(unwilling_names, unwilling_satisfaction)):
                            color = '#28a745' if rate >= 70 else '#ffc107' if rate >= 40 else '#dc3545'
                            st.markdown(f"""
                            <div class="metric-container">
                                <h4>🚫 {name}</h4>
                                <h2 style="color: {color};">{rate:.1f}%</h2>
                                <small>不愿意关系满足率（成功分开）</small>
                            </div>
                            """, unsafe_allow_html=True)
            else:
                st.info("💡 当前数据为单一等级，如需查看多等级满足率统计，请在Excel中按列分别填写不同等级的喜好关系")
            
            for i, result in enumerate(st.session_state.results):
                assignment = result.assignment
                objective = result.objective
                
                # 计算满足度指标
                # 提取正向关系对
                positive_pairs = set()
                for pair, weight in st.session_state.pair_weights.items():
                    if weight > 0:
                        positive_pairs.add(pair)
                
                metrics = compute_satisfaction_metrics(
                    assignment, 
                    st.session_state.names, 
                    st.session_state.seats, 
                    positive_pairs, 
                    st.session_state.edges,
                    st.session_state.get('willing_pairs_by_rank', {})
                )
                
                # metrics返回元组: (满足的人数, 有喜好关系的总人数, 满足的对数, 第一意愿满足率)
                n_satisfied = metrics[0]
                n_total_people_with_pref = metrics[1]
                n_satisfied_pairs = metrics[2]
                first_preference_rate = metrics[3]
                
                # 计算满足率
                satisfaction_rate = round((n_satisfied / n_total_people_with_pref * 100) if n_total_people_with_pref > 0 else 0)
                
                with st.expander(f"🎯 方案 {i+1} - 满足率 {satisfaction_rate}%", expanded=(i==0)):
                    
                    # 显示可视化结果
                    if show_visualization:
                        # 提取过道列索引（使用左侧列索引，因为过道绘制在左侧列的右边）
                        aisle_cols = [left for left, right in aisles] if enable_aisles and aisles else None
                        img_bytes = export_assignment_to_image(
                            assignment, st.session_state.seats, n_cols, col_rows, st.session_state.pair_weights,
                            figsize=(viz_figsize_w, viz_figsize_h), dpi=viz_dpi, split_visualization=split_visualization,
                            aisles=aisle_cols, show_all_lines=show_all_lines
                        )
                        if split_visualization:
                            # 拆分可视化返回ZIP文件，需要特殊处理
                            import zipfile
                            import io
                            
                            # 从ZIP中提取图片并显示
                            try:
                                with zipfile.ZipFile(io.BytesIO(img_bytes), 'r') as zip_file:
                                    # 显示正向关系图
                                    if 'positive_relationships.png' in zip_file.namelist():
                                        pos_img = zip_file.read('positive_relationships.png')
                                        st.image(pos_img, caption=f"方案 {i+1} - 正向关系图", use_container_width=True)
                                    
                                    # 显示负向关系图
                                    if 'negative_relationships.png' in zip_file.namelist():
                                        neg_img = zip_file.read('negative_relationships.png')
                                        st.image(neg_img, caption=f"方案 {i+1} - 负向关系图", use_container_width=True)
                            except Exception as e:
                                st.error(f"图片预览失败: {str(e)}")
                                st.info("请使用下载按钮获取分离的图片文件。")
                        else:
                            st.image(img_bytes, caption=f"方案 {i+1} 座位分配图", use_container_width=True)
                    else:
                        st.info("可视化已关闭，如需查看座位分配图请在左侧开启可视化选项。")
                    
                    # 基本统计信息（始终显示）
                    st.subheader("📊 基本统计")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.markdown(f"""
                        <div class="metric-container">
                            <h4>😊 满足率</h4>
                            <h2 style="color: {'#28a745' if satisfaction_rate >= 70 else '#ffc107' if satisfaction_rate >= 40 else '#dc3545'};">{satisfaction_rate}%</h2>
                            <small>有喜好关系且满足的人数占比</small>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-container">
                            <h4>🎯 目标函数值</h4>
                            <h2 style="color: #4361ee;">{objective:.2f}</h2>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-container">
                            <h4>💑 满足对数</h4>
                            <h2 style="color: #17a2b8;">{n_satisfied_pairs} 对</h2>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # 下载按钮（始终显示）
                    st.subheader("📥 下载选项")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        excel_bytes = export_assignment_to_excel(
                            assignment, st.session_state.seats, n_cols, col_rows
                        )
                        
                        st.download_button(
                            label="📊 下载Excel座位表",
                            data=excel_bytes,
                            file_name=f"座位方案_{i+1}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key=f"download_excel_summary_{i}"
                        )
                    
                    with col2:
                        if show_visualization:
                            # 提取过道列索引（使用左侧列索引，因为过道绘制在左侧列的右边）
                            aisle_cols = [left for left, right in aisles] if enable_aisles and aisles else None
                            download_img_bytes = export_assignment_to_image(
                                assignment, st.session_state.seats, n_cols, col_rows, st.session_state.pair_weights,
                                figsize=(viz_figsize_w, viz_figsize_h), dpi=viz_dpi, split_visualization=split_visualization,
                                aisles=aisle_cols, show_all_lines=show_all_lines
                            )
                            
                            if split_visualization:
                                st.download_button(
                                    label="📦 下载分离图片包 (ZIP)",
                                    data=download_img_bytes,
                                    file_name=f"座位方案_{i+1}_分离图片.zip",
                                    mime="application/zip",
                                    use_container_width=True,
                                    key=f"download_zip_summary_{i}"
                                )
                            else:
                                st.download_button(
                                    label="🖼️ 下载座位分配图",
                                    data=download_img_bytes,
                                    file_name=f"座位方案_{i+1}.png",
                                    mime="image/png",
                                    use_container_width=True,
                                    key=f"download_img_summary_{i}"
                                )
                    
                    # 更多信息（仅在启用时显示）
                    if more_info_mode:
                        # 显示详细统计信息
                        st.subheader("📊 详细统计")
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.markdown(f"""
                            <div class="metric-container">
                                <h4>👥 有喜好关系的人数</h4>
                                <h2>{n_total_people_with_pref}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col2:
                            st.markdown(f"""
                            <div class="metric-container">
                                <h4>✅ 喜好被满足的人数</h4>
                                <h2 style="color: #28a745;">{n_satisfied}</h2>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col3:
                            st.markdown(f"""
                            <div class="metric-container">
                                <h4>🥇 第一意愿满足率</h4>
                                <h2 style="color: {'#28a745' if first_preference_rate >= 70 else '#ffc107' if first_preference_rate >= 40 else '#dc3545'};">{first_preference_rate:.1f}%</h2>
                                <small>最高等级喜好关系的满足率</small>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        # 添加统计图表
                        st.subheader("📈 统计图表")
                        
                        # 创建满足率饼图
                        col1, col2 = st.columns(2)
                        with col1:
                            fig_pie, ax_pie = plt.subplots(figsize=(6, 4))
                            satisfied_count = n_satisfied
                            unsatisfied_count = n_total_people_with_pref - n_satisfied
                            
                            if n_total_people_with_pref > 0:
                                labels = ['满足', '未满足']
                                sizes = [satisfied_count, unsatisfied_count]
                                colors = ['#28a745', '#dc3545']
                                
                                ax_pie.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
                                ax_pie.set_title('喜好关系满足情况', fontsize=12, fontweight='bold')
                                
                                st.pyplot(fig_pie, use_container_width=True)
                        

                        

                        
                        # 下载按钮
                        excel_bytes = export_assignment_to_excel(
                            assignment, st.session_state.seats, n_cols, col_rows
                        )
                        
                        st.download_button(
                            label="📊 下载Excel座位表",
                            data=excel_bytes,
                            file_name=f"座位方案_{i+1}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key=f"download_excel_detail_{i}"
                        )
                        
                        if show_visualization:
                            # 提取过道列索引（使用左侧列索引，因为过道绘制在左侧列的右边）
                            aisle_cols = [left for left, right in aisles] if enable_aisles and aisles else None
                            download_img_bytes = export_assignment_to_image(
                                assignment, st.session_state.seats, n_cols, col_rows, st.session_state.pair_weights,
                                figsize=(viz_figsize_w, viz_figsize_h), dpi=viz_dpi, split_visualization=split_visualization,
                                aisles=aisle_cols, show_all_lines=show_all_lines
                            )
                            
                            if split_visualization:
                                st.download_button(
                                    label="📦 下载分离图片包 (ZIP)",
                                    data=download_img_bytes,
                                    file_name=f"座位方案_{i+1}_分离图片.zip",
                                    mime="application/zip",
                                    use_container_width=True
                                )
                            else:
                                st.download_button(
                                    label="🖼️ 下载座位分配图",
                                    data=download_img_bytes,
                                    file_name=f"座位方案_{i+1}.png",
                                    mime="image/png",
                                    use_container_width=True
                                )
                    
                    # 显示详细分配信息
                    st.subheader("📝 详细座位分配")
                    
                    # 创建座位分配表格
                    assignment_data = []
                    for name, seat_idx in assignment.items():
                        seat_info = get_seat_info(seat_idx, st.session_state.seats)
                        if 0 <= seat_idx < len(st.session_state.seats):
                            col, row = st.session_state.seats[seat_idx]
                            assignment_data.append({
                                "姓名": name,
                                "座位编号": seat_info,
                                "列": col + 1,
                                "行": row + 1
                            })
                        else:
                            assignment_data.append({
                                "姓名": name,
                                "座位编号": "未知座位",
                                "列": 0,
                                "行": 0
                            })
                    
                    df_assignment = pd.DataFrame(assignment_data)
                    st.dataframe(df_assignment, use_container_width=True, hide_index=True)

with tab3:
    st.header("📖 使用帮助")
    
    st.markdown("""
    ### 🎯 系统简介
    
    该系统根据人员间的喜好关系自动生成最优座位方案，尽量让喜欢坐在一起的人相邻，不喜欢坐在一起的人分开。
    
    **核心特性：**
    - 🧠 智能优化算法，基于OR-Tools约束求解器
    - 📊 支持多种数据输入格式
    - 🎨 可视化座位分配结果
    - 📈 详细的满足度统计分析
    - 📥 多格式结果导出
    
    ### 📋 使用步骤
    
    #### 1. 准备数据
    - 准备包含人员名单的Excel文件（每行一个姓名）
    - 准备包含喜好关系的Excel文件（可选，也可以直接在系统中输入）
    
    #### 2. 数据输入
    - 上传人员名单Excel
    - 上传喜好关系Excel（可选）
    - 设置关系权重（数值越大，影响越大）
    
    #### 3. 配置布局
    - 设置教室的列数和每列的排数
    - 选择是否包含对角线邻座关系
    
    #### 4. 优化设置
    - 设置需要生成的方案数量
    - 设置计算时间限制
    
    #### 5. 生成结果
    - 点击"计算最优座位分配"按钮
    - 查看生成的座位方案
    - 下载Excel座位表或座位分配图
    
    ### 📄 文件格式说明
    
    #### 名单文件格式
    - 默认读取第一列的所有非空值作为姓名
    - 可以在第一行添加表头（会自动跳过）
    - 支持.xlsx和.xls格式
    
    #### 喜好关系文件格式
    - 第一列：学生姓名
    - 第二列：偏好类型（喜欢/不喜欢）
    - 第三列：目标学生姓名
    
    ### ⚙️ 参数说明
    
    - **喜欢权重**：正数，数值越大，喜欢的人越倾向于坐在一起
    - **不喜欢权重**：负数，绝对值越大，不喜欢的人越倾向于分开
    - **对角线邻座**：是否将对角线方向的座位也视为相邻
    - **生成方案数**：系统会生成多个方案供选择
    - **时间限制**：每个方案的最大计算时间
    
    ### 🔍 结果解读
    
    - **目标函数值**：数值越大表示整体满足度越高
    - **满足率**：有喜好关系且得到满足的人数占比
    - **满足对数**：成功满足的喜好关系对数
    
    ### ❓ 常见问题
    
    **Q: 为什么有些人的喜好没有被满足？**
    A: 系统会尽力满足所有喜好，但在座位有限的情况下，可能无法满足所有要求。系统会优先满足权重更高的关系。
    
    **Q: 如何提高满足率？**
    A: 可以尝试调整权重设置、增加座位数量、或减少冲突的喜好关系。
    
    **Q: 支持哪些文件格式？**
    A: 目前支持Excel格式（.xlsx和.xls），未来会支持更多格式。
    """)