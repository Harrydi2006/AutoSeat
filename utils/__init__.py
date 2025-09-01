"""智能座位分配系统工具模块

包含数据处理、座位布局、优化求解和导出功能的工具函数。
"""

from .data_processor import (
    load_names_from_excel,
    load_preferences_from_excel,
    parse_custom_weights,
    compute_pair_weights,
    parse_cell_range
)

from .seat_layout import (
    generate_seats,
    generate_adjacent_edges,
    visualize_layout,
    get_adjacent_seat_pairs,
    validate_layout,
    get_seat_info
)

from .optimizer import (
    SeatAssignmentResult,
    solve_top_n_assignments,
    compute_satisfaction_metrics,
    validate_assignment,
    get_assignment_summary
)

from .exporter import (
    export_assignment_to_excel,
    export_assignment_to_image,
    create_assignment_summary_excel,
    export_layout_preview
)

__all__ = [
    # 数据处理
    'load_names_from_excel',
    'load_preferences_from_excel', 
    'parse_custom_weights',
    'compute_pair_weights',
    'parse_cell_range',
    
    # 座位布局
    'generate_seats',
    'generate_adjacent_edges',
    'visualize_layout',
    'get_adjacent_seat_pairs',
    'validate_layout',
    'get_seat_info',
    
    # 优化求解
    'SeatAssignmentResult',
    'solve_top_n_assignments',
    'compute_satisfaction_metrics',
    'validate_assignment',
    'get_assignment_summary',
    
    # 导出功能
    'export_assignment_to_excel',
    'export_assignment_to_image',
    'create_assignment_summary_excel',
    'export_layout_preview'
]

__version__ = '1.0.0'
__author__ = 'Seat Assignment System'
__description__ = '智能座位分配系统工具模块'