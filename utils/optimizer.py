"""优化求解模块

包含座位分配优化算法和约束求解功能。
"""

from ortools.sat.python import cp_model
from typing import List, Dict, Tuple, Set, Optional, Callable
from collections import defaultdict
import time
import random


class SeatAssignmentResult:
    """座位分配结果类"""
    
    def __init__(self, assignment: Dict[str, int], objective: float, status: int):
        self.assignment = assignment
        self.objective = objective
        self.status = status
        self.satisfaction_rate = 0.0
        self.n_satisfied = 0
        self.n_total_people_with_pref = 0
        self.n_satisfied_pairs = 0


def solve_top_n_assignments(
    people: List[str],
    seats: List[Tuple[int, int]],
    pair_weights: Dict[frozenset, float],
    oriented_edges: List[Tuple[Tuple[int, int], Tuple[int, int]]],
    top_n: int,
    time_limit_s: float = 10.0,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    debug_mode: bool = False
) -> List[SeatAssignmentResult]:
    """求解并返回Top-N个座位方案
    
    Args:
        people: 人员名单
        seats: 座位列表
        pair_weights: 人员对权重字典
        oriented_edges: 邻座边列表
        top_n: 需要生成的方案数量
        time_limit_s: 每个方案的时间限制（秒）
        progress_callback: 进度回调函数
        
    Returns:
        List[SeatAssignmentResult]: 座位分配结果列表
    """
    # 仅对非零权重的pair建变量，降低规模
    weighted_pairs = [(list(p)[0], list(p)[1], w) for p, w in pair_weights.items() if abs(w) > 1e-9]
    name_to_idx = {p: i for i, p in enumerate(people)}
    P = len(people)
    S = len(seats)
    
    if progress_callback:
        progress_callback(0.0, "初始化求解器...")
    
    if debug_mode:
        print(f"🔍 [调试] 开始求解座位分配问题")
        print(f"🔍 [调试] 人员数量: {P}, 座位数量: {S}")
        print(f"🔍 [调试] 权重对数量: {len(weighted_pairs)}")
        print(f"🔍 [调试] 邻座边数量: {len(oriented_edges)}")
        print(f"🔍 [调试] 需要生成方案数: {top_n}")
        print(f"🔍 [调试] 时间限制: {time_limit_s}秒")
    
    if P > S:
        raise ValueError(f"座位数({S})不足以容纳全部人员({P})")

    # 预转索引
    weighted_idx_pairs = []
    for a, b, w in weighted_pairs:
        if a in name_to_idx and b in name_to_idx:
            ia, ib = name_to_idx[a], name_to_idx[b]
            if ia == ib:
                continue
            if ia > ib:
                ia, ib = ib, ia
            weighted_idx_pairs.append((ia, ib, w))
    
    if debug_mode:
        print(f"🔍 [调试] 有效权重对数量: {len(weighted_idx_pairs)}")
        if weighted_idx_pairs:
            pos_weights = [w for _, _, w in weighted_idx_pairs if w > 0]
            neg_weights = [w for _, _, w in weighted_idx_pairs if w < 0]
            print(f"🔍 [调试] 正权重对数: {len(pos_weights)}, 负权重对数: {len(neg_weights)}")
            if pos_weights:
                print(f"🔍 [调试] 正权重范围: {min(pos_weights):.2f} ~ {max(pos_weights):.2f}")
            if neg_weights:
                print(f"🔍 [调试] 负权重范围: {min(neg_weights):.2f} ~ {max(neg_weights):.2f}")

    results = []

    # 迭代求K解：每次加一个no-good cut
    if progress_callback:
        progress_callback(0.1, "创建约束模型...")
    model = cp_model.CpModel()
    if progress_callback:
        progress_callback(0.2, "创建决策变量...")
    x = {}
    for i in range(P):
        for s in range(S):
            x[i, s] = model.NewBoolVar(f"x_{i}_{s}")

    # 每人一个座位
    if progress_callback:
        progress_callback(0.3, "添加座位分配约束...")
    for i in range(P):
        model.Add(sum(x[i, s] for s in range(S)) == 1)
    # 每座位至多一人
    for s in range(S):
        model.Add(sum(x[i, s] for i in range(P)) <= 1)

    # 目标：邻座对的权重和。用乘积变量 m = x[i,s] * x[j,t]，只在相邻座位上考虑
    if progress_callback:
        progress_callback(0.4, "构建目标函数...")
    m_vars = []
    m_coeffs = []
    
    # 构建座位索引映射
    seat_to_idx = {seat: idx for idx, seat in enumerate(seats)}
    
    for (i, j, w) in weighted_idx_pairs:
        for (seat1, seat2) in oriented_edges:
            if seat1 in seat_to_idx and seat2 in seat_to_idx:
                s = seat_to_idx[seat1]
                t = seat_to_idx[seat2]
                m = model.NewBoolVar(f"m_{i}_{j}_{s}_{t}")
                model.AddMultiplicationEquality(m, [x[i, s], x[j, t]])
                m_vars.append(m)
                m_coeffs.append(w)

    if m_vars:  # 只有在有权重对时才设置目标函数
        model.Maximize(sum(c * v for c, v in zip(m_coeffs, m_vars)))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_s
    solver.parameters.num_search_workers = min(8, 4)  # 限制线程数
    
    used_solutions = []

    for k in range(top_n):
        if progress_callback:
            progress_callback(0.5 + 0.4 * k / top_n, f"求解第 {k+1}/{top_n} 个方案...")
        
        if debug_mode:
            print(f"🔍 [调试] 开始求解第 {k+1} 个方案...")
            
        status = solver.Solve(model)
        
        if debug_mode:
            print(f"🔍 [调试] 第 {k+1} 个方案求解状态: {status}")
            if status == cp_model.OPTIMAL:
                print(f"🔍 [调试] 找到最优解")
            elif status == cp_model.FEASIBLE:
                print(f"🔍 [调试] 找到可行解")
            else:
                print(f"🔍 [调试] 未找到解，停止搜索")
        
        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            break

        # 提取解
        assign_idx = {}
        for i in range(P):
            for s in range(S):
                if solver.Value(x[i, s]) == 1:
                    assign_idx[i] = s
                    break
        obj_value = solver.ObjectiveValue() if m_vars else 0
        
        if debug_mode:
            print(f"🔍 [调试] 第 {k+1} 个方案目标函数值: {obj_value:.2f}")
            print(f"🔍 [调试] 座位分配: {[(people[i], seats[assign_idx[i]]) for i in range(P)]}")

        # 转人名
        assignment = {people[i]: assign_idx[i] for i in range(P)}

        # 保存并加no-good cut
        used_solutions.append(assign_idx.copy())
        # sum over i x[i, assigned[i]] <= P-1
        model.Add(sum(x[i, assign_idx[i]] for i in range(P)) <= P - 1)
        
        if debug_mode:
            print(f"🔍 [调试] 已添加no-good约束，排除当前解")


        
        results.append(SeatAssignmentResult(
            assignment=assignment,
            objective=obj_value,
            status=status
        ))
        
    if progress_callback:
        progress_callback(1.0, "计算完成!")

    return results


def compute_satisfaction_metrics(
    assignment: Dict[str, int], 
    people: List[str], 
    seats: List[Tuple[int, int]], 
    positive_pairs: Set[frozenset], 
    oriented_edges: List[Tuple[Tuple[int, int], Tuple[int, int]]],
    willing_pairs_by_rank: Optional[Dict[int, Set]] = None
) -> Tuple[int, int, int, float]:
    """计算满足喜好的人数和对数
    
    Args:
        assignment: 座位分配方案
        people: 人员列表
        seats: 座位列表
        positive_pairs: 正向关系对集合
        oriented_edges: 邻座边列表
        willing_pairs_by_rank: 按等级分组的喜好关系（可选）
        
    Returns:
        Tuple[int, int, int, float]: (满足的人数, 有喜好关系的总人数, 满足的对数, 第一意愿满足率)
    """
    # 座位索引到坐标的映射
    idx_to_seat = {i: seat for i, seat in enumerate(seats)}
    # 人名到座位索引的映射
    person_to_sidx = {p: sidx for p, sidx in assignment.items()}
    
    # 邻座关系
    adjacent_seats = set()
    for (c1, r1), (c2, r2) in oriented_edges:
        try:
            idx1 = seats.index((c1, r1))
            idx2 = seats.index((c2, r2))
            adjacent_seats.add(frozenset([idx1, idx2]))
        except ValueError:
            continue
    
    # 统计满足的人和对数
    satisfied_people = set()
    satisfied_pairs = 0
    
    for pair in positive_pairs:
        person1, person2 = list(pair)
        if person1 not in person_to_sidx or person2 not in person_to_sidx:
            continue
            
        seat_idx1 = person_to_sidx[person1]
        seat_idx2 = person_to_sidx[person2]
        
        # 检查这两个座位是否相邻
        if frozenset([seat_idx1, seat_idx2]) in adjacent_seats:
            satisfied_people.add(person1)
            satisfied_people.add(person2)
            satisfied_pairs += 1
    
    total_people_with_pref = len(set().union(*positive_pairs)) if positive_pairs else 0
    
    # 计算第一意愿满足率
    first_preference_rate = 0.0
    if willing_pairs_by_rank and 1 in willing_pairs_by_rank:
        first_rank_pairs = willing_pairs_by_rank[1]
        satisfied_first_pairs = 0
        
        for pair in first_rank_pairs:
            person1, person2 = list(pair)
            if person1 not in person_to_sidx or person2 not in person_to_sidx:
                continue
                
            seat_idx1 = person_to_sidx[person1]
            seat_idx2 = person_to_sidx[person2]
            
            # 检查这两个座位是否相邻
            if frozenset([seat_idx1, seat_idx2]) in adjacent_seats:
                satisfied_first_pairs += 1
        
        if len(first_rank_pairs) > 0:
            first_preference_rate = (satisfied_first_pairs / len(first_rank_pairs)) * 100
    
    return len(satisfied_people), total_people_with_pref, satisfied_pairs, first_preference_rate


def validate_assignment(
    assignment: Dict[str, int], 
    people: List[str], 
    seats: List[Tuple[int, int]]
) -> Tuple[bool, str]:
    """验证座位分配方案的有效性
    
    Args:
        assignment: 座位分配方案
        people: 人员列表
        seats: 座位列表
        
    Returns:
        Tuple[bool, str]: (是否有效, 错误信息)
    """
    # 检查是否所有人都有座位
    if set(assignment.keys()) != set(people):
        missing = set(people) - set(assignment.keys())
        extra = set(assignment.keys()) - set(people)
        msg = ""
        if missing:
            msg += f"缺少人员: {missing}. "
        if extra:
            msg += f"多余人员: {extra}. "
        return False, msg
    
    # 检查座位索引是否有效
    invalid_seats = []
    for person, seat_idx in assignment.items():
        if not (0 <= seat_idx < len(seats)):
            invalid_seats.append((person, seat_idx))
    
    if invalid_seats:
        return False, f"无效座位索引: {invalid_seats}"
    
    # 检查是否有重复座位分配
    seat_counts = {}
    for person, seat_idx in assignment.items():
        if seat_idx in seat_counts:
            return False, f"座位{seat_idx}被重复分配给{seat_counts[seat_idx]}和{person}"
        seat_counts[seat_idx] = person
    
    return True, ""


def get_assignment_summary(result: SeatAssignmentResult) -> Dict[str, any]:
    """获取分配方案摘要信息
    
    Args:
        result: 座位分配结果
        
    Returns:
        Dict: 摘要信息字典
    """
    return {
        "总人数": len(result.assignment),
        "目标函数值": round(result.objective, 2),
        "满足率": f"{result.satisfaction_rate}%",
        "满足的人数": result.n_satisfied,
        "有喜好关系的人数": result.n_total_people_with_pref,
        "满足的关系对数": result.n_satisfied_pairs,
        "求解状态": "最优解" if result.status == cp_model.OPTIMAL else "可行解"
    }