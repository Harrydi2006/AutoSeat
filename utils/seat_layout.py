"""座位布局模块

包含座位生成、邻座关系计算和布局可视化功能。
"""

import matplotlib.pyplot as plt
import matplotlib.patches as patches
from typing import List, Tuple, Dict, Set
import streamlit as st


def generate_seats(n_cols: int, rows_per_col: List[int]) -> List[Tuple[int, int]]:
    """生成座位列表
    
    Args:
        n_cols: 列数
        rows_per_col: 每列的排数列表
        
    Returns:
        List[Tuple[int, int]]: 座位坐标列表 [(列, 排), ...]
    """
    seats = []
    for c in range(n_cols):
        for r in range(rows_per_col[c]):
            seats.append((c, r))
    return seats


def generate_adjacent_edges(n_cols: int, rows_per_col: List[int], include_diag: bool = True, aisles: List[Tuple[int, int]] = None) -> List[Tuple[Tuple[int, int], Tuple[int, int]]]:
    """生成邻座位的有向边
    
    Args:
        n_cols: 列数
        rows_per_col: 每列的排数列表
        include_diag: 是否包含对角线邻座关系
        aisles: 过道列表，每个元素为(左列索引, 右列索引)，表示这两列之间有过道
        
    Returns:
        List[Tuple[Tuple[int, int], Tuple[int, int]]]: 邻座边列表
    """
    edges = []
    if aisles is None:
        aisles = []
    
    def has_aisle_between(col1: int, col2: int) -> bool:
        """检查两列之间是否有过道"""
        for left_col, right_col in aisles:
            if (col1 == left_col and col2 == right_col) or (col1 == right_col and col2 == left_col):
                return True
        return False
    
    # 检查每个座位，添加与相邻座位的边
    for c in range(n_cols):
        for r in range(rows_per_col[c]):
            # 这个座位的坐标
            curr = (c, r)
            
            # 右边的座位 (如果存在且没有过道隔开)
            if c + 1 < n_cols and r < rows_per_col[c+1] and not has_aisle_between(c, c+1):
                edges.append((curr, (c+1, r)))
                edges.append(((c+1, r), curr))  # 双向
            
            # 上面的座位 (如果存在)
            if r + 1 < rows_per_col[c]:
                edges.append((curr, (c, r+1)))
                edges.append(((c, r+1), curr))  # 双向
            
            # 如果包含对角线边
            if include_diag:
                # 右上 (如果存在且没有过道隔开)
                if c + 1 < n_cols and r + 1 < rows_per_col[c+1] and not has_aisle_between(c, c+1):
                    edges.append((curr, (c+1, r+1)))
                    edges.append(((c+1, r+1), curr))  # 双向
                
                # 右下 (如果存在且没有过道隔开)
                if c + 1 < n_cols and r - 1 >= 0 and r - 1 < rows_per_col[c+1] and not has_aisle_between(c, c+1):
                    edges.append((curr, (c+1, r-1)))
                    edges.append(((c+1, r-1), curr))  # 双向
    
    return edges


def visualize_layout(n_cols: int, rows_per_col: List[int], edges: List[Tuple[Tuple[int, int], Tuple[int, int]]], aisles: List[Tuple[int, int]] = None) -> plt.Figure:
    """可视化教室布局
    
    Args:
        n_cols: 列数
        rows_per_col: 每列的排数列表
        edges: 邻座边列表
        aisles: 过道列表，每个元素为(左列索引, 右列索引)
        
    Returns:
        plt.Figure: matplotlib图形对象
    """
    max_rows = max(rows_per_col) if rows_per_col else 0
    
    # 计算每列的实际x位置，考虑过道间距
    col_x_positions = []
    current_x = 0
    aisle_spacing = 0.5  # 过道间距
    
    if aisles is None:
        aisles = []
    
    # 创建过道集合，便于快速查找
    aisle_set = set()
    for left_col, right_col in aisles:
        aisle_set.add((left_col, right_col))
    
    for c in range(n_cols):
        col_x_positions.append(current_x)
        current_x += 1  # 基本列宽
        
        # 检查当前列后面是否有过道
        for left_col, right_col in aisles:
            if c == left_col:
                current_x += aisle_spacing
                break
    
    # 计算图形尺寸
    total_width = current_x if col_x_positions else n_cols
    fig, ax = plt.subplots(figsize=(max(8, total_width * 1.2), max(6, max_rows * 1.2)))
    
    # 绘制座位
    for c in range(n_cols):
        x_pos = col_x_positions[c]
        for r in range(rows_per_col[c]):
            rect = patches.Rectangle((x_pos, r), 0.9, 0.9, fill=True, 
                                   edgecolor='black', facecolor='lightgray', alpha=0.5)
            ax.add_patch(rect)
            ax.text(x_pos + 0.45, r + 0.45, f"C{c+1}-R{r+1}", ha='center', va='center', fontsize=8)
    
    # 绘制过道线
    if aisles:
        for left_col, right_col in aisles:
            # 在两列之间画过道线
            x_pos = col_x_positions[left_col] + 1
            ax.axvline(x=x_pos, color='orange', linestyle='--', linewidth=2, alpha=0.7, label='过道')
    
    # 绘制邻座关系
    for (c1, r1), (c2, r2) in edges:
        x1_pos = col_x_positions[c1] if c1 < len(col_x_positions) else c1
        x2_pos = col_x_positions[c2] if c2 < len(col_x_positions) else c2
        ax.plot([x1_pos + 0.45, x2_pos + 0.45], [r1 + 0.45, r2 + 0.45], 'b-', alpha=0.3)
    
    ax.set_xlim(-0.5, total_width + 0.5)
    ax.set_ylim(-0.5, max_rows + 0.5)
    ax.set_xticks([col_x_positions[i] + 0.45 for i in range(n_cols)])
    ax.set_yticks(range(max_rows))
    ax.set_xticklabels([f'C{i+1}' for i in range(n_cols)])
    ax.set_yticklabels([f'R{i+1}' for i in range(max_rows)])
    ax.set_title('教室座位布局')
    plt.tight_layout()
    
    return fig


def get_adjacent_seat_pairs(seats: List[Tuple[int, int]], edges: List[Tuple[Tuple[int, int], Tuple[int, int]]]) -> Set[frozenset]:
    """获取所有相邻座位对
    
    Args:
        seats: 座位列表
        edges: 邻座边列表
        
    Returns:
        Set[frozenset]: 相邻座位对集合
    """
    adjacent_pairs = set()
    
    for (c1, r1), (c2, r2) in edges:
        try:
            idx1 = seats.index((c1, r1))
            idx2 = seats.index((c2, r2))
            adjacent_pairs.add(frozenset([idx1, idx2]))
        except ValueError:
            continue
    
    return adjacent_pairs


def validate_layout(n_cols: int, rows_per_col: List[int], num_people: int) -> Tuple[bool, str]:
    """验证布局配置是否合理
    
    Args:
        n_cols: 列数
        rows_per_col: 每列的排数列表
        num_people: 人数
        
    Returns:
        Tuple[bool, str]: (是否有效, 错误信息)
    """
    if n_cols <= 0:
        return False, "列数必须大于0"
    
    if len(rows_per_col) != n_cols:
        return False, f"每列排数配置数量({len(rows_per_col)})与列数({n_cols})不匹配"
    
    if any(r <= 0 for r in rows_per_col):
        return False, "每列排数必须大于0"
    
    total_seats = sum(rows_per_col)
    if total_seats < num_people:
        return False, f"座位总数({total_seats})少于人数({num_people})"
    
    return True, ""


def get_seat_info(seat_idx: int, seats: List[Tuple[int, int]]) -> str:
    """获取座位信息字符串
    
    Args:
        seat_idx: 座位索引
        seats: 座位列表
        
    Returns:
        str: 座位信息，如"C1-R2"
    """
    if 0 <= seat_idx < len(seats):
        c, r = seats[seat_idx]
        return f"C{c+1}-R{r+1}"
    return "未知座位"