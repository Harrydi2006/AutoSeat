#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试导入模块
"""

try:
    import streamlit as st
    print("✓ Streamlit 导入成功")
except ImportError as e:
    print(f"✗ Streamlit 导入失败: {e}")

try:
    import pandas as pd
    print("✓ Pandas 导入成功")
except ImportError as e:
    print(f"✗ Pandas 导入失败: {e}")

try:
    import numpy as np
    print("✓ NumPy 导入成功")
except ImportError as e:
    print(f"✗ NumPy 导入失败: {e}")

try:
    import matplotlib.pyplot as plt
    print("✓ Matplotlib 导入成功")
except ImportError as e:
    print(f"✗ Matplotlib 导入失败: {e}")

try:
    from ortools.sat.python import cp_model
    print("✓ OR-Tools 导入成功")
except ImportError as e:
    print(f"✗ OR-Tools 导入失败: {e}")

try:
    import openpyxl
    print("✓ OpenPyXL 导入成功")
except ImportError as e:
    print(f"✗ OpenPyXL 导入失败: {e}")

try:
    from PIL import Image
    print("✓ Pillow 导入成功")
except ImportError as e:
    print(f"✗ Pillow 导入失败: {e}")

try:
    from utils import (
        load_names_from_excel,
        load_preferences_from_excel,
        compute_pair_weights,
        generate_seats,
        generate_adjacent_edges,
        solve_top_n_assignments,
        export_assignment_to_excel,
        export_assignment_to_image
    )
    print("✓ Utils 模块导入成功")
except ImportError as e:
    print(f"✗ Utils 模块导入失败: {e}")

print("\n导入测试完成!")